#!/usr/bin/env python3
"""Evaluator-side PDF evidence collector for Agent PromptBench.

The Agent never receives this file. It intentionally uses parsers and a renderer
that are independent from the PyMuPDF mutation provider under evaluation.
"""

from __future__ import annotations

import hashlib
import json
import math
import mimetypes
import pathlib
import re
import shutil
import subprocess
import sys
from typing import Any

import pdfplumber
import pypdf
import reportlab
from PIL import Image, ImageChops, ImageDraw
from pypdf.generic import IndirectObject, NullObject
from reportlab.pdfbase import pdfmetrics


def term_count(haystack: str | bytes, needle: str) -> int:
    encoded = needle.encode("utf-8")
    return haystack.count(encoded if isinstance(haystack, bytes) else needle)


def decoded_stream_evidence(reader: pypdf.PdfReader) -> tuple[bytes, list[dict[str, Any]]]:
    chunks: list[bytes] = []
    errors: list[dict[str, Any]] = []
    for generation, object_numbers in getattr(reader, "xref", {}).items():
        for object_number in object_numbers:
            try:
                value = reader.get_object(IndirectObject(object_number, generation, reader))
                getter = getattr(value, "get_data", None)
                if callable(getter):
                    data = getter()
                    if isinstance(data, bytes):
                        chunks.append(data)
            except Exception as error:
                errors.append(
                    {
                        "object": object_number,
                        "generation": generation,
                        "error": f"{type(error).__name__}: {error}"[:500],
                    }
                )
    return b"\n".join(chunks), errors


def inspect_pdf(file_path: pathlib.Path, terms: list[str]) -> dict[str, Any]:
    raw = file_path.read_bytes()
    reader = pypdf.PdfReader(str(file_path), strict=True)
    if reader.is_encrypted:
        raise ValueError(f"oracle cannot inspect encrypted fixture without credentials: {file_path}")
    pages: list[dict[str, Any]] = []
    full_text_parts: list[str] = []
    for index, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        full_text_parts.append(text)
        media_box = page.mediabox
        pages.append(
            {
                "page": index + 1,
                "width": float(media_box.width),
                "height": float(media_box.height),
                "rotation": int(page.get("/Rotate", 0) or 0),
                "termCounts": {term: term_count(text, term) for term in terms},
            }
        )
    full_text = "\n".join(full_text_parts)
    decoded, decoded_errors = decoded_stream_evidence(reader)
    metadata_text = "\n".join(f"{key}={value}" for key, value in (reader.metadata or {}).items())
    return {
        "path": str(file_path),
        "bytes": len(raw),
        "sha256": hashlib.sha256(raw).hexdigest(),
        "pageCount": len(pages),
        "pages": pages,
        "termCounts": {term: term_count(full_text, term) for term in terms},
        "rawTermCounts": {term: term_count(raw, term) for term in terms},
        "decodedStreamTermCounts": {term: term_count(decoded, term) for term in terms},
        "decodedStreamErrors": decoded_errors,
        "metadataTermCounts": {term: term_count(metadata_text, term) for term in terms},
        "startxrefCount": len(re.findall(rb"(?m)^startxref\s*$", raw)),
        "eofCount": len(re.findall(rb"%%EOF", raw)),
    }


def resolve_pdf_value(value: Any) -> Any:
    while isinstance(value, IndirectObject):
        value = value.get_object()
    return value


def active_structure_evidence(file_path: pathlib.Path, terms: list[str]) -> dict[str, Any]:
    reader = pypdf.PdfReader(str(file_path), strict=True)
    forbidden_names = {
        "/AA",
        "/EmbeddedFiles",
        "/JS",
        "/JavaScript",
        "/Launch",
        "/OpenAction",
    }
    action_types = {"/ImportData", "/JavaScript", "/Launch", "/SubmitForm"}
    name_counts = {name: 0 for name in sorted(forbidden_names)}
    action_type_counts = {name: 0 for name in sorted(action_types)}
    visited: set[tuple[int, int]] = set()
    scalar_values: list[str] = []

    def visit(value: Any) -> None:
        if isinstance(value, IndirectObject):
            identity = (value.idnum, value.generation)
            if identity in visited:
                return
            visited.add(identity)
            value = value.get_object()
        if isinstance(value, dict):
            for key, child in value.items():
                token = str(key)
                resolved = resolve_pdf_value(child)
                if token in name_counts and not isinstance(resolved, NullObject):
                    name_counts[token] += 1
                if token == "/S" and str(resolved) in action_type_counts:
                    action_type_counts[str(resolved)] += 1
                visit(child)
        elif isinstance(value, (list, tuple)):
            for child in value:
                visit(child)
        elif not isinstance(value, NullObject):
            scalar_values.append(str(value))

    visit(reader.trailer)
    attachments = []
    attachment_term_counts = {term: 0 for term in terms}
    try:
        attachment_list = list(reader.attachment_list)
    except TypeError:
        if name_counts.get("/EmbeddedFiles", 0):
            raise
        attachment_list = []
    for attachment in attachment_list:
        payload = attachment.content
        attachments.append({
            "name": attachment.name,
            "bytes": len(payload),
            "sha256": hashlib.sha256(payload).hexdigest(),
        })
        for term in terms:
            attachment_term_counts[term] += term_count(payload, term)

    annotations = []
    widgets = []
    for page_index, page in enumerate(reader.pages, 1):
        for reference in page.get("/Annots", []) or []:
            annotation = resolve_pdf_value(reference)
            subtype = str(annotation.get("/Subtype", ""))
            record = {
                "page": page_index,
                "subtype": subtype,
                "contents": str(resolve_pdf_value(annotation.get("/Contents", "")) or ""),
            }
            if subtype == "/Widget":
                values = {}
                for key in ("/V", "/DV"):
                    resolved = resolve_pdf_value(annotation.get(key, ""))
                    values[key] = "" if isinstance(resolved, NullObject) else str(resolved or "")
                widgets.append({**record, "name": str(resolve_pdf_value(annotation.get("/T", "")) or ""), "values": values})
            else:
                annotations.append(record)

    metadata = {str(key): str(value or "") for key, value in (reader.metadata or {}).items()}
    personal_metadata = {
        key: value
        for key, value in metadata.items()
        if key in {"/Author", "/Creator", "/Keywords", "/Subject", "/Title"}
        and value.strip()
        and not (key == "/Title" and value.strip().casefold() == "untitled")
    }
    populated_widgets = [
        widget
        for widget in widgets
        if any(value.strip().casefold() not in {"", "off", "false", "none"} for value in widget["values"].values())
    ]
    scalar_text = "\n".join(scalar_values)
    return {
        "structuralNameCounts": name_counts,
        "actionTypeCounts": action_type_counts,
        "attachments": attachments,
        "attachmentTermCounts": attachment_term_counts,
        "structureTermCounts": {term: term_count(scalar_text, term) for term in terms},
        "commentAnnotations": annotations,
        "widgets": widgets,
        "populatedWidgets": populated_widgets,
        "metadata": metadata,
        "personalMetadata": personal_metadata,
    }


def pdf_name_text(value: Any) -> str | None:
    if value is None:
        return None
    raw = str(resolve_pdf_value(value))
    if raw.startswith("/"):
        raw = raw[1:]
    return re.sub(r"#([0-9A-Fa-f]{2})", lambda match: chr(int(match.group(1), 16)), raw) or None


def expected_attachment_evidence(file_path: pathlib.Path) -> list[dict[str, Any]]:
    reader = pypdf.PdfReader(str(file_path), strict=True)
    records: list[dict[str, Any]] = []
    for ordinal, attachment in enumerate(reader.attachment_list, 1):
        payload = bytes(attachment.content)
        display_name = str(attachment.alternative_name or attachment.name or "attachment")
        declared_mime = pdf_name_text(attachment.subtype)
        records.append({
            "scope": "document",
            "page": None,
            "annotationIndex": None,
            "internalKey": str(attachment.name),
            "displayName": display_name,
            "mime": declared_mime or mimetypes.guess_type(display_name, strict=False)[0] or "application/octet-stream",
            "bytes": len(payload),
            "sha256": hashlib.sha256(payload).hexdigest(),
            "ordinal": ordinal,
        })
    for page_number, page in enumerate(reader.pages, 1):
        for annotation_index, reference in enumerate(page.get("/Annots", []) or []):
            annotation = resolve_pdf_value(reference)
            if not isinstance(annotation, dict) or str(annotation.get("/Subtype", "")) != "/FileAttachment":
                continue
            filespec = resolve_pdf_value(annotation.get("/FS"))
            if not isinstance(filespec, dict):
                raise ValueError("page FileAttachment has no FileSpec dictionary")
            display_name = str(resolve_pdf_value(filespec.get("/UF") or filespec.get("/F") or "attachment"))
            internal_key = str(resolve_pdf_value(filespec.get("/F") or filespec.get("/UF") or display_name))
            embedded = resolve_pdf_value(filespec.get("/EF"))
            if not isinstance(embedded, dict):
                raise ValueError(f"page attachment {display_name!r} has no /EF dictionary")
            stream = resolve_pdf_value(embedded.get("/UF") or embedded.get("/F"))
            if stream is None or not callable(getattr(stream, "get_data", None)):
                raise ValueError(f"page attachment {display_name!r} has no readable stream")
            payload = bytes(stream.get_data())
            declared_mime = pdf_name_text(stream.get("/Subtype"))
            records.append({
                "scope": "page",
                "page": page_number,
                "annotationIndex": annotation_index,
                "internalKey": internal_key,
                "displayName": display_name,
                "mime": declared_mime or mimetypes.guess_type(display_name, strict=False)[0] or "application/octet-stream",
                "bytes": len(payload),
                "sha256": hashlib.sha256(payload).hexdigest(),
                "ordinal": len(records) + 1,
            })
    return records


def quarantine_directory_evidence(directory: pathlib.Path) -> dict[str, Any]:
    if not directory.is_dir():
        raise ValueError(f"quarantine directory is missing: {directory}")
    files: list[dict[str, Any]] = []
    invalid: list[dict[str, str]] = []
    for candidate in sorted(directory.rglob("*")):
        relative = candidate.relative_to(directory).as_posix()
        if candidate.is_symlink():
            invalid.append({"path": relative, "type": "symlink"})
        elif candidate.is_file():
            payload = candidate.read_bytes()
            files.append({
                "path": relative,
                "bytes": len(payload),
                "sha256": hashlib.sha256(payload).hexdigest(),
                "flat": candidate.parent == directory,
            })
        elif not candidate.is_dir():
            invalid.append({"path": relative, "type": "non-regular"})
    return {"path": str(directory), "files": files, "invalid": invalid}


def attachment_quarantine(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    manifest_path = pathlib.Path(payload["manifest"])
    quarantine = pathlib.Path(payload["quarantine"])
    manifest_bytes = manifest_path.read_bytes()
    manifest = json.loads(manifest_bytes)
    expected = expected_attachment_evidence(source)
    quarantine_resolved = quarantine.resolve()
    unsafe_raw_paths = []
    for record in expected:
        raw_target = (quarantine / record["displayName"]).resolve()
        if quarantine_resolved not in raw_target.parents:
            unsafe_raw_paths.append({"displayName": record["displayName"], "resolved": str(raw_target)})
    return {
        "kind": "attachment-quarantine",
        "source": inspect_pdf(source, []),
        "expectedAttachments": expected,
        "unsafeRawPaths": unsafe_raw_paths,
        "manifest": manifest,
        "manifestFile": {
            "path": str(manifest_path),
            "bytes": len(manifest_bytes),
            "sha256": hashlib.sha256(manifest_bytes).hexdigest(),
        },
        "quarantine": quarantine_directory_evidence(quarantine),
    }


def pdf_boolean(value: Any) -> bool:
    value = resolve_pdf_value(value)
    return bool(getattr(value, "value", value))


def qualified_widget_name(widget: Any) -> str:
    parts: list[str] = []
    current = resolve_pdf_value(widget)
    visited: set[tuple[int, int] | int] = set()
    while isinstance(current, dict):
        reference = getattr(current, "indirect_reference", None)
        identity: tuple[int, int] | int = (
            (int(reference.idnum), int(reference.generation)) if reference is not None else id(current)
        )
        if identity in visited:
            raise ValueError("cyclic AcroForm parent chain")
        visited.add(identity)
        partial_name = current.get("/T")
        if partial_name is not None:
            parts.append(str(resolve_pdf_value(partial_name)))
        parent = current.get("/Parent")
        if parent is None:
            break
        current = resolve_pdf_value(parent)
    return ".".join(reversed(parts))


def form_structure_evidence(file_path: pathlib.Path) -> dict[str, Any]:
    reader = pypdf.PdfReader(str(file_path), strict=True)
    root = resolve_pdf_value(reader.trailer["/Root"])
    acro_form = resolve_pdf_value(root.get("/AcroForm")) if isinstance(root, dict) else None
    fields = reader.get_fields() or {}
    field_records: dict[str, Any] = {}
    for name, field in fields.items():
        value = resolve_pdf_value(field.get("/V", ""))
        default_value = resolve_pdf_value(field.get("/DV", ""))
        states = [str(state) for state in list(field.get("/_States_", []) or [])]
        flags = int(field.get("/Ff", 0) or 0)
        field_records[str(name)] = {
            "fieldType": str(field.get("/FT", "")),
            "value": "" if isinstance(value, NullObject) else str(value or ""),
            "defaultValue": "" if isinstance(default_value, NullObject) else str(default_value or ""),
            "flags": flags,
            "readOnly": bool(flags & 1),
            "states": states,
            "kids": len(list(field.get("/Kids", []) or [])),
        }

    widgets: list[dict[str, Any]] = []
    for page_number, page in enumerate(reader.pages, 1):
        page_height = float(page.mediabox.height)
        for reference in page.get("/Annots", []) or []:
            widget = resolve_pdf_value(reference)
            if not isinstance(widget, dict) or str(widget.get("/Subtype", "")) != "/Widget":
                continue
            name = qualified_widget_name(widget)
            parent = resolve_pdf_value(widget.get("/Parent")) if widget.get("/Parent") is not None else widget
            field_type = str(widget.get("/FT") or parent.get("/FT") or "")
            flags = int(widget.get("/Ff", parent.get("/Ff", 0)) or 0)
            rectangle = [float(value) for value in widget.get("/Rect", [])]
            if len(rectangle) != 4:
                raise ValueError(f"widget {name!r} has invalid /Rect")
            appearance = resolve_pdf_value(widget.get("/AP")) if widget.get("/AP") is not None else None
            normal = resolve_pdf_value(appearance.get("/N")) if isinstance(appearance, dict) and appearance.get("/N") is not None else None
            normal_is_stream = callable(getattr(normal, "get_data", None))
            appearance_states = sorted(str(state) for state in normal.keys()) if isinstance(normal, dict) and not normal_is_stream else []
            widgets.append({
                "page": page_number,
                "name": name,
                "fieldType": field_type,
                "rect": [rectangle[0], page_height - rectangle[3], rectangle[2], page_height - rectangle[1]],
                "appearancePresent": normal is not None,
                "appearanceStreamBytes": len(normal.get_data()) if normal_is_stream else None,
                "appearanceStates": appearance_states,
                "selectedState": str(resolve_pdf_value(widget.get("/AS", "")) or ""),
                "readOnly": bool(flags & 1),
            })
    return {
        "acroFormPresent": isinstance(acro_form, dict),
        "needAppearances": pdf_boolean(acro_form.get("/NeedAppearances", False)) if isinstance(acro_form, dict) else None,
        "fieldTreeRoots": len(list(acro_form.get("/Fields", []) or [])) if isinstance(acro_form, dict) else 0,
        "fields": field_records,
        "widgets": widgets,
    }


def literal_style(file_path: pathlib.Path, page_number: int, term: str) -> dict[str, Any]:
    with pdfplumber.open(str(file_path)) as document:
        page = document.pages[page_number - 1]
        characters = page.chars
        character_text = "".join(str(character.get("text", "")) for character in characters)
        start = character_text.find(term)
        if start < 0:
            return {"found": False, "term": term, "page": page_number}
        selected = characters[start : start + len(term)]
        if "".join(str(character.get("text", "")) for character in selected) != term:
            return {"found": False, "term": term, "page": page_number}
        sizes = sorted({round(float(character.get("size", 0)), 6) for character in selected})
        fonts = sorted({str(character.get("fontname", "")) for character in selected})
        return {
            "found": True,
            "term": term,
            "page": page_number,
            "bbox": [
                min(float(character["x0"]) for character in selected),
                min(float(character["top"]) for character in selected),
                max(float(character["x1"]) for character in selected),
                max(float(character["bottom"]) for character in selected),
            ],
            "fonts": fonts,
            "sizes": sizes,
        }


def run_poppler(executable: str, source: pathlib.Path, prefix: pathlib.Path, dpi: int) -> list[pathlib.Path]:
    prefix.parent.mkdir(parents=True, exist_ok=True)
    result = subprocess.run(
        [executable, "-png", "-r", str(dpi), str(source), str(prefix)],
        text=True,
        capture_output=True,
        timeout=90,
        check=False,
    )
    if result.returncode != 0:
        raise RuntimeError(f"pdftoppm failed for {source}: {result.stderr.strip()}")
    candidates = list(prefix.parent.glob(f"{prefix.name}-*.png"))
    return sorted(candidates, key=lambda value: int(value.stem.rsplit("-", 1)[-1]))


def bbox_within(inner: tuple[int, int, int, int] | None, outer: list[int]) -> bool:
    return bool(
        inner
        and inner[0] >= outer[0]
        and inner[1] >= outer[1]
        and inner[2] <= outer[2]
        and inner[3] <= outer[3]
    )


def compare_rendered_pages(
    source: pathlib.Path,
    output: pathlib.Path,
    render_root: pathlib.Path,
    poppler: str,
    target_style: dict[str, Any],
    target_page: int,
    dpi: int = 144,
) -> dict[str, Any]:
    shutil.rmtree(render_root, ignore_errors=True)
    source_pages = run_poppler(poppler, source, render_root / "source" / "page", dpi)
    output_pages = run_poppler(poppler, output, render_root / "output" / "page", dpi)
    scale = dpi / 72.0
    padding = 8
    target_bbox = target_style.get("bbox") or [0, 0, 0, 0]
    allowed_mask = [
        math.floor(target_bbox[0] * scale) - padding,
        math.floor(target_bbox[1] * scale) - padding,
        math.ceil(target_bbox[2] * scale) + padding,
        math.ceil(target_bbox[3] * scale) + padding,
    ]
    pages: list[dict[str, Any]] = []
    for index in range(max(len(source_pages), len(output_pages))):
        if index >= len(source_pages) or index >= len(output_pages):
            pages.append({"page": index + 1, "missing": True})
            continue
        with Image.open(source_pages[index]) as source_image_raw, Image.open(output_pages[index]) as output_image_raw:
            source_image = source_image_raw.convert("RGB")
            output_image = output_image_raw.convert("RGB")
            same_dimensions = source_image.size == output_image.size
            difference = ImageChops.difference(source_image, output_image) if same_dimensions else None
            difference_bbox = difference.getbbox() if difference else None
            non_white_bbox = ImageChops.difference(output_image, Image.new("RGB", output_image.size, "white")).getbbox()
            pages.append(
                {
                    "page": index + 1,
                    "sameDimensions": same_dimensions,
                    "nonBlank": non_white_bbox is not None,
                    "changedPixelsBBox": list(difference_bbox) if difference_bbox else None,
                    "changedWithinAllowedMask": index + 1 == target_page and bbox_within(difference_bbox, allowed_mask),
                }
            )
    return {
        "renderer": "poppler-pdftoppm",
        "dpi": dpi,
        "sourcePageCount": len(source_pages),
        "outputPageCount": len(output_pages),
        "allowedMask": {"page": target_page, "bboxPx": allowed_mask},
        "pages": pages,
    }


def compare_rendered_pages_with_masks(
    source: pathlib.Path,
    output: pathlib.Path,
    render_root: pathlib.Path,
    poppler: str,
    allowed_masks: list[dict[str, Any]],
    dpi: int = 144,
) -> dict[str, Any]:
    shutil.rmtree(render_root, ignore_errors=True)
    source_pages = run_poppler(poppler, source, render_root / "source" / "page", dpi)
    output_pages = run_poppler(poppler, output, render_root / "output" / "page", dpi)
    scale = dpi / 72.0
    padding = 8
    masks_by_page: dict[int, list[list[int]]] = {}
    for mask in allowed_masks:
        x0, top, x1, bottom = mask["bbox"]
        masks_by_page.setdefault(int(mask["page"]), []).append([
            math.floor(x0 * scale) - padding,
            math.floor(top * scale) - padding,
            math.ceil(x1 * scale) + padding,
            math.ceil(bottom * scale) + padding,
        ])
    pages = []
    for index in range(max(len(source_pages), len(output_pages))):
        if index >= len(source_pages) or index >= len(output_pages):
            pages.append({"page": index + 1, "missing": True})
            continue
        with Image.open(source_pages[index]) as source_image_raw, Image.open(output_pages[index]) as output_image_raw:
            source_image = source_image_raw.convert("RGB")
            output_image = output_image_raw.convert("RGB")
            same_dimensions = source_image.size == output_image.size
            difference = ImageChops.difference(source_image, output_image) if same_dimensions else None
            difference_bbox = difference.getbbox() if difference else None
            outside = difference.copy() if difference else None
            if outside:
                painter = ImageDraw.Draw(outside)
                for mask in masks_by_page.get(index + 1, []):
                    painter.rectangle(mask, fill=(0, 0, 0))
            outside_bbox = outside.getbbox() if outside else None
            non_white_bbox = ImageChops.difference(output_image, Image.new("RGB", output_image.size, "white")).getbbox()
            pages.append({
                "page": index + 1,
                "sameDimensions": same_dimensions,
                "nonBlank": non_white_bbox is not None,
                "changedPixelsBBox": list(difference_bbox) if difference_bbox else None,
                "changedOutsideAllowedMasksBBox": list(outside_bbox) if outside_bbox else None,
                "changedOnlyWithinAllowedMasks": outside_bbox is None,
            })
    return {
        "renderer": "poppler-pdftoppm",
        "dpi": dpi,
        "sourcePageCount": len(source_pages),
        "outputPageCount": len(output_pages),
        "allowedMasks": [{"page": page, "bboxPx": bbox} for page, boxes in masks_by_page.items() for bbox in boxes],
        "pages": pages,
    }


def compare_form_render(
    source: pathlib.Path,
    output: pathlib.Path,
    render_root: pathlib.Path,
    poppler: str,
    source_structure: dict[str, Any],
    expected_fields: dict[str, str],
    dpi: int = 144,
) -> dict[str, Any]:
    expected_widgets: list[dict[str, Any]] = []
    for widget in source_structure["widgets"]:
        expected_value = str(expected_fields.get(widget["name"], ""))
        if not expected_value:
            expected_change = False
        elif widget["fieldType"] == "/Btn":
            target_state = f"/{expected_value.lstrip('/')}"
            expected_change = target_state in widget["appearanceStates"]
        else:
            expected_change = True
        expected_widgets.append({**widget, "expectedChange": expected_change})

    allowed_masks = [
        {"page": widget["page"], "bbox": widget["rect"]}
        for widget in expected_widgets
        if widget["expectedChange"]
    ]
    visual = compare_rendered_pages_with_masks(source, output, render_root, poppler, allowed_masks, dpi=dpi)
    scale = dpi / 72.0
    source_pages = run_poppler(poppler, source, render_root / "detail-source" / "page", dpi)
    output_pages = run_poppler(poppler, output, render_root / "detail-output" / "page", dpi)
    widget_changes = []
    for widget in expected_widgets:
        page_index = int(widget["page"]) - 1
        if page_index >= len(source_pages) or page_index >= len(output_pages):
            widget_changes.append({"name": widget["name"], "page": widget["page"], "missing": True})
            continue
        with Image.open(source_pages[page_index]) as source_raw, Image.open(output_pages[page_index]) as output_raw:
            source_image = source_raw.convert("RGB")
            output_image = output_raw.convert("RGB")
            x0, top, x1, bottom = widget["rect"]
            pixel_box = (
                max(0, math.floor(x0 * scale)),
                max(0, math.floor(top * scale)),
                min(output_image.width, math.ceil(x1 * scale)),
                min(output_image.height, math.ceil(bottom * scale)),
            )
            source_crop = source_image.crop(pixel_box)
            output_crop = output_image.crop(pixel_box)
            difference = ImageChops.difference(source_crop, output_crop)
            difference_bbox = difference.getbbox()
            inset = min(6, max(1, min(output_crop.size) // 4))
            interior_box = (inset, inset, max(inset, output_crop.width - inset), max(inset, output_crop.height - inset))
            interior_difference = difference.crop(interior_box)
            interior_bbox = interior_difference.getbbox()
            widget_changes.append({
                "name": widget["name"],
                "page": widget["page"],
                "fieldType": widget["fieldType"],
                "appearanceStates": widget["appearanceStates"],
                "expectedChange": widget["expectedChange"],
                "pixelBox": list(pixel_box),
                "changedPixelsBBox": list(difference_bbox) if difference_bbox else None,
                "changedInteriorPixelsBBox": list(interior_bbox) if interior_bbox else None,
            })
    return {**visual, "widgetChanges": widget_changes}


def bounded_replace(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    output = pathlib.Path(payload["output"])
    old = payload["old"]
    new = payload["new"]
    target_page = int(payload["targetPage"])
    source_style = literal_style(source, target_page, old)
    output_style = literal_style(output, target_page, new)
    return {
        "kind": "bounded-replace",
        "source": inspect_pdf(source, [old, new]),
        "output": inspect_pdf(output, [old, new]),
        "sourceStyle": source_style,
        "outputStyle": output_style,
        "visual": compare_rendered_pages(
            source,
            output,
            pathlib.Path(payload["renderRoot"]),
            payload["poppler"],
            source_style,
            target_page,
        ),
    }


def overflow_refusal(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    old = payload["old"]
    replacement = payload["new"]
    source_style = literal_style(source, 1, old)
    with pdfplumber.open(str(source)) as document:
        page = document.pages[0]
        containing_rectangles = [
            rectangle
            for rectangle in page.rects
            if source_style.get("found")
            and float(rectangle["x0"]) <= source_style["bbox"][0]
            and float(rectangle["x1"]) >= source_style["bbox"][2]
            and float(rectangle["top"]) <= source_style["bbox"][1]
            and float(rectangle["bottom"]) >= source_style["bbox"][3]
        ]
    target_rectangle = min(containing_rectangles, key=lambda value: float(value["width"])) if containing_rectangles else None
    font = source_style.get("fonts", [""])[0] if source_style.get("fonts") else ""
    size = source_style.get("sizes", [0])[0] if source_style.get("sizes") else 0
    normalized_font = "Helvetica" if "helvetica" in font.lower() else font
    replacement_width = pdfmetrics.stringWidth(replacement, normalized_font, size) if normalized_font and size else None
    available_width = (
        float(target_rectangle["x1"]) - float(source_style["bbox"][0])
        if target_rectangle and source_style.get("found")
        else None
    )
    return {
        "kind": "overflow-refusal",
        "source": inspect_pdf(source, [old, replacement]),
        "sourceStyle": source_style,
        "geometry": {
            "targetRectangle": {
                "x0": float(target_rectangle["x0"]),
                "x1": float(target_rectangle["x1"]),
                "top": float(target_rectangle["top"]),
                "bottom": float(target_rectangle["bottom"]),
                "width": float(target_rectangle["width"]),
                "height": float(target_rectangle["height"]),
            }
            if target_rectangle
            else None,
            "font": normalized_font,
            "fontSize": size,
            "replacementWidth": replacement_width,
            "availableWidth": available_width,
            "fits": replacement_width <= available_width if replacement_width is not None and available_width is not None else None,
        },
    }


def active_content_sanitize(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    output = pathlib.Path(payload["output"])
    terms = list(payload.get("terms") or [])
    source_bytes = source.read_bytes()
    output_bytes = output.read_bytes()
    return {
        "kind": "active-content-sanitize",
        "source": inspect_pdf(source, terms),
        "output": inspect_pdf(output, terms),
        "sourceStructure": active_structure_evidence(source, terms),
        "outputStructure": active_structure_evidence(output, terms),
        "originalPrefixPreserved": output_bytes.startswith(source_bytes),
        "visual": compare_rendered_pages_with_masks(
            source,
            output,
            pathlib.Path(payload["renderRoot"]),
            payload["poppler"],
            [
                {"page": 1, "bbox": [72, 148, 252, 172]},
                {"page": 1, "bbox": [500, 92, 520, 112]},
            ],
        ),
    }


def acroform_visible(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    output = pathlib.Path(payload["output"])
    expected_fields = {str(name): str(value) for name, value in dict(payload.get("fields") or {}).items()}
    source_structure = form_structure_evidence(source)
    output_structure = form_structure_evidence(output)
    source_bytes = source.read_bytes()
    output_bytes = output.read_bytes()
    return {
        "kind": "acroform-visible",
        "source": inspect_pdf(source, [value for value in expected_fields.values() if value]),
        "output": inspect_pdf(output, [value for value in expected_fields.values() if value]),
        "sourceForm": source_structure,
        "outputForm": output_structure,
        "originalPrefixPreserved": output_bytes.startswith(source_bytes),
        "visual": compare_form_render(
            source,
            output,
            pathlib.Path(payload["renderRoot"]),
            payload["poppler"],
            source_structure,
            expected_fields,
        ),
    }


def tagged_structure_evidence(file_path: pathlib.Path) -> dict[str, Any]:
    reader = pypdf.PdfReader(str(file_path), strict=True)
    root = resolve_pdf_value(reader.trailer["/Root"])
    structure_root = resolve_pdf_value(root.get("/StructTreeRoot")) if isinstance(root, dict) else None
    mark_info = resolve_pdf_value(root.get("/MarkInfo")) if isinstance(root, dict) else None
    page_numbers: dict[tuple[int, int], int] = {}
    for page_number, page in enumerate(reader.pages, 1):
        reference = getattr(page, "indirect_reference", None)
        if reference is not None:
            page_numbers[(int(reference.idnum), int(reference.generation))] = page_number

    def page_number_for(value: Any, inherited: int | None = None) -> int | None:
        if isinstance(value, IndirectObject):
            return page_numbers.get((int(value.idnum), int(value.generation)), inherited)
        return inherited

    roles: dict[str, int] = {}
    records: list[dict[str, Any]] = []
    tables: list[dict[str, Any]] = []
    root_ids: list[str] = []
    objr_count = 0
    visited: set[tuple[int, int] | int] = set()

    def visit(value: Any, inherited_page: int | None = None, root_child: bool = False) -> set[int]:
        nonlocal objr_count
        identity: tuple[int, int] | int
        if isinstance(value, IndirectObject):
            identity = (int(value.idnum), int(value.generation))
        else:
            identity = id(value)
        if identity in visited:
            return set()
        resolved = resolve_pdf_value(value)
        if not isinstance(resolved, dict):
            return set()
        visited.add(identity)
        if str(resolved.get("/Type", "")) == "/OBJR":
            objr_count += 1
            page_number = page_number_for(resolved.get("/Pg"), inherited_page)
            return {page_number} if page_number else set()
        if str(resolved.get("/Type", "")) != "/StructElem":
            return set()
        role = str(resolved.get("/S", "")).lstrip("/")
        roles[role] = roles.get(role, 0) + 1
        page_number = page_number_for(resolved.get("/Pg"), inherited_page)
        structure_id = str(resolve_pdf_value(resolved.get("/ID", "")) or "")
        if root_child and structure_id:
            root_ids.append(structure_id)
        record = {
            "role": role,
            "id": structure_id,
            "page": page_number,
            "alt": str(resolve_pdf_value(resolved.get("/Alt", "")) or ""),
        }
        records.append(record)
        kid_value = resolve_pdf_value(resolved.get("/K"))
        kids = list(kid_value) if isinstance(kid_value, (list, tuple)) else [kid_value]
        descendant_pages = {page_number} if page_number else set()
        roles_before = dict(roles)
        for kid in kids:
            descendant_pages.update(visit(kid, page_number))
        record["descendantPages"] = sorted(descendant_pages)
        if role == "Table":
            tables.append({
                "id": structure_id,
                "pages": sorted(descendant_pages),
                "rows": roles.get("TR", 0) - roles_before.get("TR", 0),
                "headers": roles.get("TH", 0) - roles_before.get("TH", 0),
                "dataCells": roles.get("TD", 0) - roles_before.get("TD", 0),
            })
        return descendant_pages

    root_kids = resolve_pdf_value(structure_root.get("/K")) if isinstance(structure_root, dict) else []
    for kid in list(root_kids) if isinstance(root_kids, (list, tuple)) else [root_kids]:
        visit(kid, root_child=True)

    links: list[dict[str, Any]] = []
    artifact_markers = 0
    page_text: list[str] = []
    for page_number, page in enumerate(reader.pages, 1):
        page_text.append(page.extract_text() or "")
        contents = page.get_contents()
        if contents is not None:
            artifact_markers += len(re.findall(rb"/Artifact\s+BMC\b", contents.get_data()))
        for reference in page.get("/Annots", []) or []:
            annotation = resolve_pdf_value(reference)
            if not isinstance(annotation, dict) or str(annotation.get("/Subtype", "")) != "/Link":
                continue
            action = resolve_pdf_value(annotation.get("/A")) if annotation.get("/A") is not None else None
            links.append({
                "page": page_number,
                "uri": str(resolve_pdf_value(action.get("/URI", "")) or "") if isinstance(action, dict) else "",
                "structParent": int(annotation.get("/StructParent")) if annotation.get("/StructParent") is not None else None,
                "rect": [float(value) for value in annotation.get("/Rect", [])],
            })

    title = str(getattr(reader.metadata, "title", "") or "")
    return {
        "tagged": isinstance(structure_root, dict) and bool(resolve_pdf_value(mark_info.get("/Marked"))) if isinstance(mark_info, dict) else False,
        "language": str(resolve_pdf_value(root.get("/Lang", "")) or "") if isinstance(root, dict) else "",
        "title": title,
        "roles": roles,
        "records": records,
        "rootIds": root_ids,
        "tables": tables,
        "figuresWithAlt": sum(1 for record in records if record["role"] == "Figure" and record["alt"].strip()),
        "links": links,
        "linkObjrAssociations": objr_count,
        "artifactMarkers": artifact_markers,
        "pageStructParents": [int(page.get("/StructParents")) if page.get("/StructParents") is not None else None for page in reader.pages],
        "parentTreePresent": isinstance(structure_root, dict) and structure_root.get("/ParentTree") is not None,
        "pageText": page_text,
    }


def render_created_pdf(file_path: pathlib.Path, render_root: pathlib.Path, poppler: str, dpi: int = 144) -> dict[str, Any]:
    shutil.rmtree(render_root, ignore_errors=True)
    rendered = run_poppler(poppler, file_path, render_root / "output" / "page", dpi)
    pages: list[dict[str, Any]] = []
    for page_number, page_path in enumerate(rendered, 1):
        with Image.open(page_path) as raw:
            image = raw.convert("RGB")
            difference = ImageChops.difference(image, Image.new("RGB", image.size, "white"))
            ink_bbox = difference.getbbox()
            touches_edge = bool(ink_bbox and (ink_bbox[0] <= 2 or ink_bbox[1] <= 2 or ink_bbox[2] >= image.width - 2 or ink_bbox[3] >= image.height - 2))
            pages.append({
                "page": page_number,
                "width": image.width,
                "height": image.height,
                "nonBlank": ink_bbox is not None,
                "inkBBox": list(ink_bbox) if ink_bbox else None,
                "touchesEdge": touches_edge,
                "bytes": page_path.stat().st_size,
            })
    return {"renderer": "poppler-pdftoppm", "dpi": dpi, "pageCount": len(rendered), "pages": pages}


def accessible_report(payload: dict[str, Any]) -> dict[str, Any]:
    source = pathlib.Path(payload["source"])
    output = pathlib.Path(payload["output"])
    source_bytes = source.read_bytes()
    structure = tagged_structure_evidence(output)
    return {
        "kind": "accessible-report",
        "source": {"path": str(source), "bytes": len(source_bytes), "sha256": hashlib.sha256(source_bytes).hexdigest()},
        "output": inspect_pdf(output, []),
        "structure": structure,
        "visual": render_created_pdf(output, pathlib.Path(payload["renderRoot"]), payload["poppler"]),
    }


def main() -> None:
    payload = json.load(sys.stdin)
    kind = payload.get("kind")
    if kind == "bounded-replace":
        evidence = bounded_replace(payload)
    elif kind == "overflow-refusal":
        evidence = overflow_refusal(payload)
    elif kind == "active-content-sanitize":
        evidence = active_content_sanitize(payload)
    elif kind == "acroform-visible":
        evidence = acroform_visible(payload)
    elif kind == "attachment-quarantine":
        evidence = attachment_quarantine(payload)
    elif kind == "accessible-report":
        evidence = accessible_report(payload)
    else:
        raise ValueError(f"unsupported PDF oracle kind: {kind}")
    evidence["toolchain"] = {
        "pypdf": pypdf.__version__,
        "pdfplumber": pdfplumber.__version__,
        "reportlab": reportlab.Version,
        "pillow": Image.__version__,
    }
    json.dump(evidence, sys.stdout, ensure_ascii=False, separators=(",", ":"))
    sys.stdout.write("\n")


if __name__ == "__main__":
    main()
