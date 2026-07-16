#!/usr/bin/env python3
"""Evaluator-side PDF evidence collector for Agent PromptBench.

The Agent never receives this file. It intentionally uses parsers and a renderer
that are independent from the PyMuPDF mutation provider under evaluation.
"""

from __future__ import annotations

import hashlib
import json
import math
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


def main() -> None:
    payload = json.load(sys.stdin)
    kind = payload.get("kind")
    if kind == "bounded-replace":
        evidence = bounded_replace(payload)
    elif kind == "overflow-refusal":
        evidence = overflow_refusal(payload)
    elif kind == "active-content-sanitize":
        evidence = active_content_sanitize(payload)
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
