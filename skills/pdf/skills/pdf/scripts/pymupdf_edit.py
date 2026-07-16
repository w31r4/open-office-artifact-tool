#!/usr/bin/env python3
"""Inspect or directly edit an existing PDF through the explicit PyMuPDF provider."""

from __future__ import annotations

import argparse
import hashlib
import json
import math
import os
from pathlib import Path
import shutil
import sys
import tempfile
from typing import Any


class ProviderError(RuntimeError):
    pass


TEXT_FIT_ABSOLUTE_TOLERANCE_PT = 0.0001
TEXT_FIT_RELATIVE_TOLERANCE = 0.000001
TEXT_FIT_MAX_TOLERANCE_PT = 0.0005
ACTIVE_CONTENT_NULL_KEYS = frozenset({
    "AA",
    "EmbeddedFiles",
    "ImportData",
    "JavaScript",
    "JS",
    "Launch",
    "OpenAction",
    "RichMedia",
    "SubmitForm",
    "XFA",
})


def text_width_fit(text_width: float, available_width: float) -> dict[str, float | bool]:
    """Accept float-quantization noise without creating visible layout slack."""
    if not math.isfinite(text_width) or not math.isfinite(available_width) or text_width < 0 or available_width < 0:
        raise ProviderError("replacement text and box widths must be finite non-negative numbers")
    overflow = max(0.0, text_width - available_width)
    tolerance = min(
        TEXT_FIT_MAX_TOLERANCE_PT,
        max(
            TEXT_FIT_ABSOLUTE_TOLERANCE_PT,
            max(abs(text_width), abs(available_width)) * TEXT_FIT_RELATIVE_TOLERANCE,
        ),
    )
    return {
        "fits": overflow <= tolerance,
        "textWidth": text_width,
        "availableWidth": available_width,
        "overflow": overflow,
        "tolerance": tolerance,
    }


def replacement_source_span(pymupdf, page, term: str, bounds) -> dict[str, Any]:
    candidates = []
    for block in page.get_text("dict").get("blocks", []):
        if block.get("type") != 0:
            continue
        for line in block.get("lines", []):
            for span in line.get("spans", []):
                if term not in str(span.get("text") or ""):
                    continue
                span_bounds = pymupdf.Rect(span.get("bbox"))
                if not span_bounds.intersects(bounds):
                    continue
                candidates.append({"span": span, "direction": tuple(line.get("dir") or ())})
    if len(candidates) != 1:
        raise ProviderError(
            "replace_text requires each match to resolve to exactly one source text span; "
            f"found {len(candidates)} candidates"
        )
    candidate = candidates[0]
    direction = candidate["direction"]
    if len(direction) != 2 or not math.isclose(direction[0], 1.0, abs_tol=0.000001) or not math.isclose(direction[1], 0.0, abs_tol=0.000001):
        raise ProviderError("replace_text supports only horizontal source text; rotated or skewed text requires an explicit specialist workflow")
    span = candidate["span"]
    origin = span.get("origin")
    if not isinstance(origin, (list, tuple)) or len(origin) != 2:
        raise ProviderError("replace_text source span has no usable baseline origin")
    return span


def builtin_font_alias(source_font: str) -> str | None:
    base_name = source_font.rsplit("+", 1)[-1]
    aliases = {
        "Courier": "cour",
        "Courier-Bold": "cobo",
        "Courier-BoldOblique": "cobi",
        "Courier-Oblique": "coit",
        "Helvetica": "helv",
        "Helvetica-Bold": "hebo",
        "Helvetica-BoldOblique": "hebi",
        "Helvetica-Oblique": "heit",
        "Times-Bold": "tibo",
        "Times-BoldItalic": "tibi",
        "Times-Italic": "tiit",
        "Times-Roman": "tiro",
    }
    return aliases.get(base_name)


def accepted_license(value: str | None) -> str:
    selected = str(value or os.environ.get("OPEN_OFFICE_PDF_PYMUPDF_LICENSE", "")).strip().lower()
    if selected not in {"agpl", "commercial"}:
        raise ProviderError("PyMuPDF use requires --accept-license agpl|commercial or OPEN_OFFICE_PDF_PYMUPDF_LICENSE")
    return selected


def require_pymupdf():
    try:
        import pymupdf
    except ImportError as exc:
        raise ProviderError("PyMuPDF is required; install the documented tested version in the selected Python environment") from exc
    return pymupdf


def file_record(path: Path) -> dict[str, Any]:
    payload = path.read_bytes()
    return {"path": str(path), "bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}


def open_document(pymupdf, source: Path, password_env: str | None):
    doc = pymupdf.open(str(source))
    if doc.needs_pass:
        if not password_env:
            doc.close()
            raise ProviderError("encrypted input requires --password-env naming an authorized environment variable")
        password = os.environ.get(password_env)
        if password is None:
            doc.close()
            raise ProviderError(f"password environment variable {password_env!r} is not set")
        if doc.authenticate(password) <= 0:
            doc.close()
            raise ProviderError("authorized password did not decrypt the PDF")
    return doc


def validate_limits(source: Path, doc, args: argparse.Namespace) -> None:
    if args.max_bytes < 1 or args.max_pages < 1 or args.max_xrefs < 1:
        raise ProviderError("max-bytes, max-pages, and max-xrefs must be positive")
    if source.stat().st_size > args.max_bytes:
        raise ProviderError(f"PDF is {source.stat().st_size} bytes; max-bytes is {args.max_bytes}")
    if doc.page_count > args.max_pages:
        raise ProviderError(f"PDF has {doc.page_count} pages; max-pages is {args.max_pages}")
    if doc.xref_length() > args.max_xrefs:
        raise ProviderError(f"PDF has {doc.xref_length()} xrefs; max-xrefs is {args.max_xrefs}")


def signature_policy(doc) -> dict[str, Any]:
    signature_xrefs = []
    docmdp_xrefs = []
    fieldmdp_xrefs = []
    perms_xrefs = []
    byte_range_xrefs = []
    for xref in range(1, doc.xref_length()):
        try:
            obj = doc.xref_object(xref, compressed=False, ascii=True) or ""
        except Exception:
            continue
        if "/FT/Sig" in obj.replace(" ", "") or "/Type/Sig" in obj.replace(" ", ""):
            signature_xrefs.append(xref)
        if "/DocMDP" in obj:
            docmdp_xrefs.append(xref)
        if "/FieldMDP" in obj:
            fieldmdp_xrefs.append(xref)
        if "/Perms" in obj:
            perms_xrefs.append(xref)
        if "/ByteRange" in obj:
            byte_range_xrefs.append(xref)
    sigflags = int(doc.get_sigflags() or 0)
    return {
        "sigflags": sigflags,
        "signatureXrefs": signature_xrefs,
        "byteRangeXrefs": byte_range_xrefs,
        "docMDPXrefs": docmdp_xrefs,
        "fieldMDPXrefs": fieldmdp_xrefs,
        "permsXrefs": perms_xrefs,
        "hasSignatureEvidence": bool(sigflags > 0 or signature_xrefs or byte_range_xrefs),
        "hasDocMDP": bool(docmdp_xrefs or perms_xrefs),
        "hasFieldMDP": bool(fieldmdp_xrefs),
    }


def document_inventory(doc, source: Path, version: str) -> dict[str, Any]:
    pages = []
    for index in range(doc.page_count):
        page = doc.load_page(index)
        annotations = []
        for annotation in page.annots() or []:
            annotations.append({"xref": annotation.xref, "type": annotation.type, "rect": list(annotation.rect), "info": dict(annotation.info or {})})
        widgets = []
        for widget in page.widgets() or []:
            widgets.append({
                "xref": widget.xref,
                "name": widget.field_name,
                "value": widget.field_value,
                "type": widget.field_type_string,
                "rect": list(widget.rect),
            })
        fonts = page.get_fonts(full=True) or []
        images = page.get_images(full=True) or []
        pages.append({
            "page": index + 1,
            "width": float(page.rect.width),
            "height": float(page.rect.height),
            "rotation": int(page.rotation),
            "textChars": len(page.get_text("text") or ""),
            "fonts": len(fonts),
            "images": len(images),
            "annotations": annotations,
            "widgets": widgets,
        })
    try:
        attachments = list(doc.embfile_names() or [])
    except Exception:
        attachments = []
    return {
        "provider": "pymupdf",
        "providerVersion": version,
        "strategy": "read-only",
        "silentFallback": False,
        "source": file_record(source),
        "summary": {
            "pages": doc.page_count,
            "xrefs": doc.xref_length(),
            "encrypted": bool(doc.needs_pass),
            "attachments": len(attachments),
            "annotations": sum(len(page["annotations"]) for page in pages),
            "widgets": sum(len(page["widgets"]) for page in pages),
            "images": sum(page["images"] for page in pages),
        },
        "metadata": dict(doc.metadata or {}),
        "xmlMetadataChars": len(doc.get_xml_metadata() or ""),
        "attachments": attachments,
        "signaturePolicy": signature_policy(doc),
        "pages": pages,
    }


def validate_signed_mutation(policy: dict[str, Any], strategy: str, args: argparse.Namespace) -> None:
    constrained = policy["hasSignatureEvidence"] or policy["hasDocMDP"] or policy["hasFieldMDP"]
    if not constrained:
        return
    if strategy == "incremental" and not args.allow_signed:
        raise ProviderError("signed/signature-constrained input requires --allow-signed after pyHanko and DocMDP review")
    if strategy in {"rewrite", "sanitize"} and not args.invalidate_signatures:
        raise ProviderError(f"{strategy} on signed/signature-constrained input requires --invalidate-signatures")


def sanitize_active_content(pymupdf, doc) -> dict[str, Any]:
    """Remove the bounded inert-publication surface before PyMuPDF's full scrub.

    PyMuPDF's Document.scrub() is still the package cleanup authority. These
    explicit steps cover document actions, comments, populated widgets, and
    isolated invisible text that scrub() does not remove in every supported
    provider build. Invisible text overlapping visible text fails closed
    instead of redacting legitimate page content.
    """
    hidden_spans = []
    annotations_removed = 0
    widgets_cleared = 0
    action_keys_removed = []
    name_tree_keys_removed = []
    for page_index in range(doc.page_count):
        page = doc.load_page(page_index)
        try:
            traces = list(page.get_texttrace() or [])
        except Exception as exc:
            raise ProviderError(f"cannot inspect hidden text on page {page_index + 1}: {exc}") from exc
        visible_rects = [
            pymupdf.Rect(trace["bbox"])
            for trace in traces
            if int(trace.get("type", 0)) != 3 and float(trace.get("opacity", 1)) != 0
        ]
        page_hidden = [
            pymupdf.Rect(trace["bbox"])
            for trace in traces
            if int(trace.get("type", 0)) == 3 or float(trace.get("opacity", 1)) == 0
        ]
        for bounds in page_hidden:
            if any((bounds & visible).get_area() > 0.01 for visible in visible_rects):
                raise ProviderError(
                    f"hidden text on page {page_index + 1} overlaps visible text; "
                    "safe removal cannot preserve the visible page"
                )
            page.add_redact_annot(bounds, fill=False, cross_out=False)
            hidden_spans.append({"page": page_index + 1, "rect": list(bounds)})
        if page_hidden:
            page.apply_redactions(images=0, graphics=0, text=0)

        for annotation in list(page.annots() or []):
            page.delete_annot(annotation)
            annotations_removed += 1

        for widget in list(page.widgets() or []):
            value = str(widget.field_value or "").strip()
            if value.casefold() not in {"", "off", "false", "none"}:
                widget.field_value = ""
                widget.update()
                widgets_cleared += 1
            for key in ("V", "DV"):
                if key in (doc.xref_get_keys(widget.xref) or []):
                    doc.xref_set_key(widget.xref, key, "null")

    for xref in range(1, doc.xref_length()):
        try:
            keys = doc.xref_get_keys(xref) or []
        except Exception:
            continue
        for key in ("AA", "OpenAction"):
            if key in keys:
                value_type, _ = doc.xref_get_key(xref, key)
                if value_type != "null":
                    doc.xref_set_key(xref, key, "null")
                    action_keys_removed.append({"xref": xref, "key": key})
    catalog = doc.pdf_catalog()
    if catalog > 0:
        value_type, value = doc.xref_get_key(catalog, "Names")
        if value_type == "xref":
            try:
                names_xref = int(str(value).split()[0])
                names_keys = doc.xref_get_keys(names_xref) or []
                for key in ("EmbeddedFiles", "JavaScript"):
                    if key in names_keys:
                        nested_type, _ = doc.xref_get_key(names_xref, key)
                        if nested_type != "null":
                            doc.xref_set_key(names_xref, key, "null")
                            name_tree_keys_removed.append({"xref": names_xref, "key": key})
            except (TypeError, ValueError):
                raise ProviderError("catalog /Names reference could not be resolved for active-content cleanup")
    return {
        "type": "active_content_cleanup",
        "hiddenTextSpansRemoved": hidden_spans,
        "annotationsRemoved": annotations_removed,
        "widgetsCleared": widgets_cleared,
        "actionKeysRemoved": action_keys_removed,
        "nameTreeKeysRemoved": name_tree_keys_removed,
    }


def remove_null_active_content_keys(doc) -> list[dict[str, Any]]:
    """Physically remove active-content dictionary keys left as null by scrub.

    PyMuPDF's public key setter represents deletion as a null value. A strict
    public-copy residue scan must not leave those active-content names in the
    rewritten object graph, even when their values are inert. PyMuPDF renders
    indirect dictionaries one key/value pair per line; refuse unfamiliar
    serialization instead of applying a broad textual rewrite.
    """
    removed_records = []
    for xref in range(1, doc.xref_length()):
        try:
            existing_keys = set(doc.xref_get_keys(xref) or [])
        except Exception:
            continue
        candidates = {
            key
            for key in ACTIVE_CONTENT_NULL_KEYS & existing_keys
            if doc.xref_get_key(xref, key)[0] == "null"
        }
        if not candidates:
            continue

        original = doc.xref_object(xref, compressed=False) or ""
        retained_lines = []
        removed_keys = []
        for line in original.splitlines():
            tokens = line.strip().split()
            if len(tokens) == 2 and tokens[0].startswith("/") and tokens[1] == "null":
                key = tokens[0][1:]
                if key in candidates:
                    removed_keys.append(key)
                    continue
            retained_lines.append(line)
        if set(removed_keys) != candidates:
            missing = sorted(candidates - set(removed_keys))
            raise ProviderError(
                f"cannot safely remove null active-content keys {missing} from xref {xref}; "
                "provider object serialization was not the expected one-entry-per-line form"
            )

        doc.update_object(xref, "\n".join(retained_lines))
        remaining = ACTIVE_CONTENT_NULL_KEYS & set(doc.xref_get_keys(xref) or [])
        if remaining:
            raise ProviderError(f"active-content keys remain in xref {xref}: {sorted(remaining)}")
        removed_records.append({"xref": xref, "keys": sorted(removed_keys)})
    return removed_records


def load_operations(path: Path) -> list[dict[str, Any]]:
    source = path.expanduser().resolve()
    if not source.is_file():
        raise ProviderError("--operations must be an existing JSON file")
    if source.stat().st_size > 10 * 1024 * 1024:
        raise ProviderError("operations JSON exceeds the 10 MiB limit")
    payload = json.loads(source.read_text("utf-8"))
    operations = payload.get("operations") if isinstance(payload, dict) else payload
    if not isinstance(operations, list) or not operations:
        raise ProviderError("operations JSON must be a non-empty list or an object with an operations list")
    if len(operations) > 10_000:
        raise ProviderError("operations list exceeds the 10,000 operation limit")
    if not all(isinstance(operation, dict) for operation in operations):
        raise ProviderError("every operation must be a JSON object")
    return operations


def number(value: Any, label: str) -> float:
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        raise ProviderError(f"{label} must be a number")
    return float(value)


def rect(value: Any, label: str = "rect") -> tuple[float, float, float, float]:
    if not isinstance(value, list) or len(value) != 4:
        raise ProviderError(f"{label} must be [x0, y0, x1, y1]")
    values = tuple(number(item, label) for item in value)
    if values[2] <= values[0] or values[3] <= values[1]:
        raise ProviderError(f"{label} must have positive width and height")
    return values


def point(value: Any, label: str = "point") -> tuple[float, float]:
    if not isinstance(value, list) or len(value) != 2:
        raise ProviderError(f"{label} must be [x, y]")
    return number(value[0], label), number(value[1], label)


def color(value: Any, label: str = "color") -> tuple[float, ...] | None:
    if value is None:
        return None
    if not isinstance(value, list) or len(value) not in {1, 3, 4}:
        raise ProviderError(f"{label} must contain 1, 3, or 4 values in the range 0..1")
    values = tuple(number(item, label) for item in value)
    if any(item < 0 or item > 1 for item in values):
        raise ProviderError(f"{label} values must be in the range 0..1")
    return values


def page_for(doc, operation: dict[str, Any]):
    page_number = operation.get("page")
    if not isinstance(page_number, int) or isinstance(page_number, bool) or page_number < 1 or page_number > doc.page_count:
        raise ProviderError(f"operation page must be between 1 and {doc.page_count}")
    return page_number, doc.load_page(page_number - 1)


def ensure_inside(page, bounds: tuple[float, float, float, float], label: str) -> None:
    page_rect = page.rect
    if bounds[0] < page_rect.x0 or bounds[1] < page_rect.y0 or bounds[2] > page_rect.x1 or bounds[3] > page_rect.y1:
        raise ProviderError(f"{label} lies outside page bounds {list(page_rect)}")


def apply_operations(pymupdf, doc, operations: list[dict[str, Any]], strategy: str) -> list[dict[str, Any]]:
    allowed = {
        "insert_textbox",
        "insert_image",
        "replace_image",
        "add_text_annotation",
        "fill_form",
        "rotate_page",
        "delete_page",
        "redact_text",
        "redact_rect",
        "replace_text",
        "scrub",
    }
    if strategy == "sanitize":
        forbidden = [operation.get("type") for operation in operations if operation.get("type") not in {"redact_text", "redact_rect", "replace_text", "scrub"}]
        if forbidden:
            raise ProviderError(f"sanitize accepts only redact_text, redact_rect, replace_text, and scrub operations; found {forbidden}")
    if strategy != "sanitize" and any(operation.get("type") in {"redact_text", "redact_rect", "replace_text", "scrub"} for operation in operations):
        raise ProviderError("redaction and scrub operations require --strategy sanitize")

    applied: list[dict[str, Any]] = []
    redaction_pages: set[int] = set()
    replacement_overlays: list[dict[str, Any]] = []
    for index, operation in enumerate(operations, 1):
        operation_type = operation.get("type")
        if operation_type not in allowed:
            raise ProviderError(f"operation {index} has unsupported type {operation_type!r}")

        if operation_type == "insert_textbox":
            page_number, page = page_for(doc, operation)
            bounds = rect(operation.get("rect"))
            ensure_inside(page, bounds, "textbox")
            text = operation.get("text")
            if not isinstance(text, str) or not text or len(text) > 2_000_000:
                raise ProviderError("insert_textbox text must be a non-empty string under 2,000,000 characters")
            fontfile = operation.get("font_file")
            if fontfile:
                fontfile = str(Path(fontfile).expanduser().resolve())
                if not Path(fontfile).is_file():
                    raise ProviderError(f"font_file does not exist: {fontfile}")
            remaining = page.insert_textbox(
                bounds,
                text,
                fontname=str(operation.get("font_name") or "helv"),
                fontfile=fontfile,
                fontsize=number(operation.get("font_size", 11), "font_size"),
                color=color(operation.get("color")),
                fill=color(operation.get("fill"), "fill"),
                align=int(operation.get("align", 0)),
                rotate=int(operation.get("rotation", 0)),
                overlay=bool(operation.get("overlay", True)),
            )
            if remaining < 0:
                raise ProviderError(f"insert_textbox on page {page_number} does not fit; provider returned {remaining}")
            applied.append({"type": operation_type, "page": page_number, "remainingHeight": remaining})

        elif operation_type == "insert_image":
            page_number, page = page_for(doc, operation)
            bounds = rect(operation.get("rect"))
            ensure_inside(page, bounds, "image rectangle")
            image_path = Path(str(operation.get("path") or "")).expanduser().resolve()
            if not image_path.is_file():
                raise ProviderError(f"insert_image path does not exist: {image_path}")
            if image_path.stat().st_size > 100 * 1024 * 1024:
                raise ProviderError("insert_image source exceeds the 100 MiB limit")
            xref = page.insert_image(
                bounds,
                filename=str(image_path),
                keep_proportion=bool(operation.get("keep_proportion", True)),
                overlay=bool(operation.get("overlay", True)),
                rotate=int(operation.get("rotation", 0)),
            )
            applied.append({"type": operation_type, "page": page_number, "xref": xref, "imageSha256": file_record(image_path)["sha256"]})

        elif operation_type == "replace_image":
            page_number, page = page_for(doc, operation)
            xref = operation.get("xref")
            if not isinstance(xref, int) or isinstance(xref, bool) or xref <= 0:
                raise ProviderError("replace_image requires a positive integer xref")
            page_xrefs = {int(image[0]) for image in page.get_images(full=True) or []}
            if xref not in page_xrefs:
                raise ProviderError(f"image xref {xref} is not used on page {page_number}")
            image_path = Path(str(operation.get("path") or "")).expanduser().resolve()
            if not image_path.is_file():
                raise ProviderError(f"replace_image path does not exist: {image_path}")
            if image_path.stat().st_size > 100 * 1024 * 1024:
                raise ProviderError("replace_image source exceeds the 100 MiB limit")
            page.replace_image(xref, filename=str(image_path))
            applied.append({"type": operation_type, "page": page_number, "xref": xref, "imageSha256": file_record(image_path)["sha256"]})

        elif operation_type == "add_text_annotation":
            page_number, page = page_for(doc, operation)
            location = point(operation.get("point"))
            text = operation.get("text")
            if not isinstance(text, str) or not text:
                raise ProviderError("add_text_annotation requires non-empty text")
            annotation = page.add_text_annot(location, text, icon=str(operation.get("icon") or "Note"))
            info = {}
            if operation.get("title") is not None:
                info["title"] = str(operation["title"])
            if operation.get("subject") is not None:
                info["subject"] = str(operation["subject"])
            if info:
                annotation.set_info(info)
            annotation.update()
            applied.append({"type": operation_type, "page": page_number, "xref": annotation.xref})

        elif operation_type == "fill_form":
            name = operation.get("field")
            if not isinstance(name, str) or not name:
                raise ProviderError("fill_form requires a non-empty field name")
            requested_page = operation.get("page")
            matched = []
            page_indexes = range(doc.page_count) if requested_page is None else [page_for(doc, operation)[0] - 1]
            for page_index in page_indexes:
                page = doc.load_page(page_index)
                for widget in page.widgets() or []:
                    if widget.field_name == name:
                        widget.field_value = str(operation.get("value", ""))
                        widget.update()
                        matched.append({"page": page_index + 1, "xref": widget.xref})
            if not matched:
                raise ProviderError(f"form field not found: {name}")
            applied.append({"type": operation_type, "field": name, "widgets": matched})

        elif operation_type == "rotate_page":
            page_number, page = page_for(doc, operation)
            rotation = operation.get("rotation")
            if rotation not in {0, 90, 180, 270}:
                raise ProviderError("rotate_page rotation must be 0, 90, 180, or 270")
            page.set_rotation(rotation)
            applied.append({"type": operation_type, "page": page_number, "rotation": rotation})

        elif operation_type == "delete_page":
            page_number, _ = page_for(doc, operation)
            if doc.page_count <= 1:
                raise ProviderError("delete_page cannot remove the final page")
            doc.delete_page(page_number - 1)
            applied.append({"type": operation_type, "page": page_number})

        elif operation_type == "redact_text":
            term = operation.get("term")
            if not isinstance(term, str) or not term:
                raise ProviderError("redact_text requires a non-empty term")
            page_indexes = range(doc.page_count)
            if operation.get("page") is not None:
                page_indexes = [page_for(doc, operation)[0] - 1]
            matches = 0
            page_counts = []
            for page_index in page_indexes:
                page = doc.load_page(page_index)
                quads = page.search_for(term, quads=True)
                for quad in quads:
                    page.add_redact_annot(quad, fill=color(operation.get("fill", [0, 0, 0]), "fill"), cross_out=False)
                if quads:
                    redaction_pages.add(page_index)
                    matches += len(quads)
                    page_counts.append({"page": page_index + 1, "matches": len(quads)})
            if not matches:
                raise ProviderError(f"redact_text term was not found in selectable page text: {term!r}")
            applied.append({"type": operation_type, "termSha256": hashlib.sha256(term.encode("utf-8")).hexdigest(), "matches": matches, "pages": page_counts})

        elif operation_type == "redact_rect":
            page_number, page = page_for(doc, operation)
            bounds = rect(operation.get("rect"))
            ensure_inside(page, bounds, "redaction rectangle")
            page.add_redact_annot(bounds, fill=color(operation.get("fill", [0, 0, 0]), "fill"), cross_out=False)
            redaction_pages.add(page_number - 1)
            applied.append({"type": operation_type, "page": page_number, "rect": list(bounds)})

        elif operation_type == "replace_text":
            term = operation.get("term")
            replacement = operation.get("replacement")
            if not isinstance(term, str) or not term:
                raise ProviderError("replace_text requires a non-empty term")
            if not isinstance(replacement, str) or not replacement:
                raise ProviderError("replace_text requires a non-empty replacement")
            page_indexes = range(doc.page_count)
            if operation.get("page") is not None:
                page_indexes = [page_for(doc, operation)[0] - 1]
            matches = 0
            page_counts = []
            for page_index in page_indexes:
                page = doc.load_page(page_index)
                quads = page.search_for(term, quads=True)
                for quad in quads:
                    bounds = pymupdf.Rect(quad.rect)
                    source_span = replacement_source_span(pymupdf, page, term, bounds)
                    source_font = str(source_span.get("font") or "")
                    font_name = operation.get("font_name") or builtin_font_alias(source_font)
                    if not font_name:
                        raise ProviderError(
                            f"replace_text cannot preserve source font {source_font!r}; provide a supported font_name in an explicitly reviewed operation"
                        )
                    page.add_redact_annot(bounds, fill=color(operation.get("fill", [1, 1, 1]), "fill"), cross_out=False)
                    replacement_overlays.append({
                        "pageIndex": page_index,
                        "rect": bounds,
                        "replacement": replacement,
                        "fontName": str(font_name),
                        "fontSize": number(operation.get("font_size", source_span.get("size")), "font_size"),
                        "color": color(operation["color"]) if operation.get("color") is not None else pymupdf.sRGB_to_pdf(int(source_span.get("color") or 0)),
                        "baseline": float(source_span["origin"][1]),
                        "sourceFont": source_font,
                        "sourceFontSize": float(source_span.get("size") or 0),
                    })
                if quads:
                    redaction_pages.add(page_index)
                    matches += len(quads)
                    page_counts.append({"page": page_index + 1, "matches": len(quads)})
            if not matches:
                raise ProviderError(f"replace_text term was not found in selectable page text: {term!r}")
            applied.append({"type": operation_type, "termSha256": hashlib.sha256(term.encode("utf-8")).hexdigest(), "matches": matches, "pages": page_counts})

        elif operation_type == "scrub":
            applied.append({"type": "scrub", "strictDefaults": True})

    for page_index in sorted(redaction_pages):
        page = doc.load_page(page_index)
        page.apply_redactions(images=2, graphics=1, text=0)
    if redaction_pages:
        applied.append({"type": "apply_redactions", "pages": [page + 1 for page in sorted(redaction_pages)]})
    fit_checks = []
    for overlay in replacement_overlays:
        page = doc.load_page(overlay["pageIndex"])
        available_width = overlay["rect"].width
        text_width = pymupdf.get_text_length(overlay["replacement"], fontname=overlay["fontName"], fontsize=overlay["fontSize"])
        fit = text_width_fit(text_width, available_width)
        fit_checks.append({
            "page": overlay["pageIndex"] + 1,
            "sourceFont": overlay["sourceFont"],
            "sourceFontSize": overlay["sourceFontSize"],
            "outputFont": overlay["fontName"],
            "outputFontSize": overlay["fontSize"],
            "baseline": overlay["baseline"],
            **fit,
        })
        if not fit["fits"]:
            raise ProviderError(
                f"replacement text width {text_width:.6f} exceeds the original box width {available_width:.6f} "
                f"by {fit['overflow']:.6f}pt (numeric tolerance {fit['tolerance']:.6f}pt); "
                "general PDF reflow is intentionally unsupported"
            )
        page.insert_text(
            (overlay["rect"].x0, overlay["baseline"]),
            overlay["replacement"],
            fontname=overlay["fontName"],
            fontsize=overlay["fontSize"],
            color=overlay["color"],
            overlay=True,
        )
    if replacement_overlays:
        applied.append({"type": "replacement_overlays", "count": len(replacement_overlays), "reflow": False, "fitChecks": fit_checks})
    return applied


def save_document(pymupdf, doc, strategy: str, temporary: Path) -> dict[str, Any]:
    if strategy == "incremental":
        if not doc.can_save_incrementally():
            raise ProviderError("Document.can_save_incrementally() is false; refusing a rewrite fallback")
        doc.saveIncr()
        return {"nullActiveContentKeysRemoved": []}
    temporary.unlink(missing_ok=True)
    null_active_content_keys_removed = []
    if strategy == "sanitize":
        doc.scrub(
            attached_files=True,
            clean_pages=True,
            embedded_files=True,
            hidden_text=True,
            javascript=True,
            metadata=True,
            redactions=True,
            redact_images=2,
            remove_links=True,
            reset_fields=True,
            reset_responses=True,
            thumbnails=True,
            xml_metadata=True,
        )
        catalog = doc.pdf_catalog()
        if catalog > 0:
            doc.xref_set_key(catalog, "PageMode", "/UseNone")
        null_active_content_keys_removed = remove_null_active_content_keys(doc)
        doc.save(
            str(temporary),
            garbage=4,
            clean=True,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            preserve_metadata=False,
            use_objstms=1,
            encryption=pymupdf.PDF_ENCRYPT_KEEP,
        )
    else:
        doc.save(
            str(temporary),
            garbage=3,
            clean=False,
            deflate=True,
            deflate_images=True,
            deflate_fonts=True,
            preserve_metadata=True,
            use_objstms=1,
            encryption=pymupdf.PDF_ENCRYPT_KEEP,
        )
    return {"nullActiveContentKeysRemoved": null_active_content_keys_removed}


def edit(args: argparse.Namespace, pymupdf, license_choice: str) -> dict[str, Any]:
    source = args.input.expanduser().resolve()
    output = args.output.expanduser().resolve()
    operations_path = args.operations.expanduser().resolve()
    if not source.is_file():
        raise ProviderError("input must be an existing PDF")
    if source == output:
        raise ProviderError("input and output must differ; in-place source overwrite is forbidden")
    if operations_path in {source, output}:
        raise ProviderError("operations JSON must differ from input and output paths")
    operations = load_operations(operations_path)
    if args.strategy == "sanitize" and not args.invalidate_signatures:
        raise ProviderError("sanitize requires explicit --invalidate-signatures acknowledgement")
    residue_sensitive_operations = {
        "redact_text",
        "redact_rect",
        "replace_text",
    }
    if args.strategy == "sanitize" and any(operation.get("type") in residue_sensitive_operations for operation in operations) and not args.sensitive_term:
        raise ProviderError("sanitize redaction/replacement requires at least one --sensitive-term for the strict residue gate")
    if args.strategy == "sanitize":
        operation_terms = {
            str(operation["term"])
            for operation in operations
            if operation.get("type") in {"redact_text", "replace_text"} and operation.get("term")
        }
        missing_terms = sorted(operation_terms - set(args.sensitive_term))
        if missing_terms:
            raise ProviderError("every redact_text/replace_text term must also be supplied as --sensitive-term")
    output.parent.mkdir(parents=True, exist_ok=True)
    source_bytes = source.read_bytes()

    temporary: Path | None = None
    doc = None
    residue_report = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{output.name}.", suffix=".pdf", dir=output.parent, delete=False) as stream:
            temporary = Path(stream.name)
        if args.strategy == "incremental":
            shutil.copyfile(source, temporary)
            doc = open_document(pymupdf, temporary, args.password_env)
        else:
            doc = open_document(pymupdf, source, args.password_env)
        validate_limits(source, doc, args)
        before_policy = signature_policy(doc)
        validate_signed_mutation(before_policy, args.strategy, args)
        applied = apply_operations(pymupdf, doc, operations, args.strategy)
        active_cleanup = None
        if args.strategy == "sanitize":
            active_cleanup = sanitize_active_content(pymupdf, doc)
            applied.append(active_cleanup)
        save_details = save_document(pymupdf, doc, args.strategy, temporary)
        if active_cleanup is not None:
            active_cleanup.update(save_details)
        doc.close()
        doc = None

        result_bytes = temporary.read_bytes()
        if not result_bytes.startswith(b"%PDF-"):
            raise ProviderError("provider output is not a PDF")
        prefix_preserved = result_bytes.startswith(source_bytes)
        if args.strategy == "incremental" and not prefix_preserved:
            raise ProviderError("incremental output does not preserve the exact original byte prefix")
        if args.strategy == "sanitize" and prefix_preserved:
            raise ProviderError("sanitized output retains the complete original byte prefix")

        if args.strategy == "sanitize":
            from residue_scan import scan_pdf

            residue_report = scan_pdf(
                temporary,
                args.sensitive_term,
                require_ocr=True,
                require_single_revision=True,
                require_inert=True,
                ocr_language=args.ocr_language,
                ocr_dpi=args.ocr_dpi,
            )
            if not residue_report["ok"]:
                raise ProviderError(
                    f"strict residue scan failed: matches={residue_report['summary']['matches']}, "
                    f"incomplete={residue_report['summary']['incompleteChecks']}, "
                    f"inert={residue_report.get('activeContent', {}).get('violations', [])}"
                )

        result_doc = open_document(pymupdf, temporary, args.password_env)
        try:
            after_policy = signature_policy(result_doc)
            page_count = result_doc.page_count
        finally:
            result_doc.close()
        temporary.replace(output)
        temporary = None
        return {
            "provider": "pymupdf",
            "providerVersion": pymupdf.__version__,
            "licenseChoice": license_choice,
            "strategy": args.strategy,
            "silentFallback": False,
            "source": file_record(source),
            "output": file_record(output),
            "originalPrefixPreserved": prefix_preserved,
            "signaturePolicyBefore": before_policy,
            "signaturePolicyAfter": after_policy,
            "operations": applied,
            "pages": page_count,
            "residueScan": residue_report,
            "requiredNextGates": ["qpdf --check when installed", "pdfinfo", "Poppler render every page", "manual visual review"],
        }
    finally:
        if doc is not None:
            doc.close()
        if temporary is not None:
            temporary.unlink(missing_ok=True)


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    subparsers = root.add_subparsers(dest="command", required=True)

    probe = subparsers.add_parser("probe", help="report the explicit provider capability surface")
    probe.add_argument("--accept-license", choices=["agpl", "commercial"])

    inspect = subparsers.add_parser("inspect", help="inspect native pages, objects, fonts, images, forms, and signature constraints")
    inspect.add_argument("input", type=Path)
    inspect.add_argument("--output", type=Path)
    inspect.add_argument("--accept-license", choices=["agpl", "commercial"])
    inspect.add_argument("--password-env")
    inspect.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    inspect.add_argument("--max-pages", type=int, default=2_000)
    inspect.add_argument("--max-xrefs", type=int, default=1_000_000)

    mutation = subparsers.add_parser("edit", help="apply a bounded operation list directly to original PDF bytes")
    mutation.add_argument("input", type=Path)
    mutation.add_argument("output", type=Path)
    mutation.add_argument("--strategy", choices=["rewrite", "incremental", "sanitize"], required=True)
    mutation.add_argument("--operations", type=Path, required=True)
    mutation.add_argument("--accept-license", choices=["agpl", "commercial"])
    mutation.add_argument("--password-env")
    mutation.add_argument("--allow-signed", action="store_true")
    mutation.add_argument("--invalidate-signatures", action="store_true")
    mutation.add_argument("--sensitive-term", action="append", default=[])
    mutation.add_argument("--ocr-language", default="eng")
    mutation.add_argument("--ocr-dpi", type=int, default=150)
    mutation.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    mutation.add_argument("--max-pages", type=int, default=2_000)
    mutation.add_argument("--max-xrefs", type=int, default=1_000_000)
    return root


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parser().parse_args()
    output_on_failure = getattr(args, "output", None) if args.command == "edit" else None
    try:
        license_choice = accepted_license(args.accept_license)
        pymupdf = require_pymupdf()
        if args.command == "probe":
            print(json.dumps({
                "provider": "pymupdf",
                "providerVersion": pymupdf.__version__,
                "licenseChoice": license_choice,
                "available": True,
                "silentFallback": False,
                "strategies": ["read-only", "rewrite", "incremental", "sanitize"],
                "operations": ["insert_textbox", "insert_image", "replace_image", "add_text_annotation", "fill_form", "rotate_page", "delete_page", "redact_text", "redact_rect", "replace_text", "scrub"],
            }, indent=2, sort_keys=True))
            return 0
        if args.command == "inspect":
            source = args.input.expanduser().resolve()
            if not source.is_file():
                raise ProviderError("input must be an existing PDF")
            doc = open_document(pymupdf, source, args.password_env)
            try:
                validate_limits(source, doc, args)
                result = document_inventory(doc, source, pymupdf.__version__)
                result["licenseChoice"] = license_choice
            finally:
                doc.close()
            rendered = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True)
            if args.output:
                destination = args.output.expanduser().resolve()
                if destination == source:
                    raise ProviderError("JSON report path must differ from source PDF")
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_text(rendered + "\n", "utf-8")
            else:
                print(rendered)
            return 0
        result = edit(args, pymupdf, license_choice)
        print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        if output_on_failure:
            Path(output_on_failure).expanduser().resolve().unlink(missing_ok=True)
        print(json.dumps({"ok": False, "provider": "pymupdf", "error": str(exc), "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
