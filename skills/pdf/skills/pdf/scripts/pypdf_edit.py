#!/usr/bin/env python3
"""Inspect or make bounded structural/form/annotation edits with explicit pypdf policy."""

from __future__ import annotations

import argparse
from io import BytesIO
import hashlib
import json
import mimetypes
import os
from pathlib import Path
import re
import shutil
import sys
import tempfile
from typing import Any
import unicodedata


MERGE_STAMP_SCHEMA = "open-office-artifact-tool.pdf-merge-stamp.v1"
MERGE_STAMP_RESULT_SCHEMA = "open-office-artifact-tool.pdf-merge-stamp-result.v1"


class ProviderError(RuntimeError):
    pass


def require_pypdf():
    try:
        import pypdf
        from pypdf import PdfReader, PdfWriter
        from pypdf.annotations import Text
        from pypdf.generic import NameObject
    except ImportError as exc:
        raise ProviderError("pypdf is required in the selected Python environment") from exc
    return pypdf, PdfReader, PdfWriter, Text, NameObject


def file_record(path: Path) -> dict[str, Any]:
    payload = path.read_bytes()
    return {
        "path": str(path),
        "bytes": len(payload),
        "sha256": hashlib.sha256(payload).hexdigest(),
    }


def open_reader(PdfReader, source: Path, password_env: str | None):
    reader = PdfReader(str(source), strict=False)
    if reader.is_encrypted:
        if not password_env:
            raise ProviderError("encrypted input requires --password-env naming an authorized environment variable")
        password = os.environ.get(password_env)
        if password is None:
            raise ProviderError(f"password environment variable {password_env!r} is not set")
        if not reader.decrypt(password):
            raise ProviderError("authorized password did not decrypt the PDF")
    return reader


def plain(value: Any, *, depth: int = 0) -> Any:
    if depth > 5:
        return "<depth-limit>"
    try:
        value = value.get_object()
    except Exception:
        pass
    if value is None or isinstance(value, (str, int, float, bool)):
        return value
    if isinstance(value, bytes):
        return {"bytes": len(value), "sha256": hashlib.sha256(value).hexdigest()}
    if isinstance(value, dict):
        return {str(key): plain(item, depth=depth + 1) for key, item in list(value.items())[:200]}
    if isinstance(value, (list, tuple)):
        return [plain(item, depth=depth + 1) for item in list(value)[:200]]
    return str(value)


def signature_policy(reader) -> dict[str, Any]:
    root = reader.trailer["/Root"].get_object()
    perms = root.get("/Perms")
    try:
        perms = perms.get_object() if perms else None
    except Exception:
        pass
    fields = reader.get_fields() or {}
    signature_fields = []
    field_mdp_fields = []
    for name, field in fields.items():
        field_type = str(field.get("/FT", ""))
        value = field.get("/V")
        if field_type == "/Sig" or value is not None and "/ByteRange" in str(value):
            rendered_value = json.dumps(plain(value), ensure_ascii=False, sort_keys=True, default=str) if value is not None else ""
            signature_fields.append({
                "name": str(name),
                "fieldType": field_type or None,
                "signed": value is not None,
            })
            if "/FieldMDP" in rendered_value:
                field_mdp_fields.append(str(name))
    docmdp = None
    if isinstance(perms, dict) and perms.get("/DocMDP") is not None:
        docmdp = plain(perms.get("/DocMDP"))
    return {
        "hasSignatureFields": bool(signature_fields),
        "hasSignedSignatures": any(field["signed"] for field in signature_fields),
        "signatureFields": signature_fields,
        "hasPerms": perms is not None,
        "hasDocMDP": docmdp is not None,
        "docMDP": docmdp,
        "hasFieldMDP": bool(field_mdp_fields),
        "fieldMDPFields": field_mdp_fields,
    }


def inspect_reader(reader, source: Path, version: str) -> dict[str, Any]:
    annotations = []
    widgets = 0
    page_attachment_names = []
    for page_number, page in enumerate(reader.pages, 1):
        page_annots = page.get("/Annots") or []
        for reference in page_annots:
            try:
                annotation = reference.get_object()
                subtype = str(annotation.get("/Subtype", ""))
                if subtype == "/Widget":
                    widgets += 1
                if subtype == "/FileAttachment":
                    filespec = resolve(annotation.get("/FS"))
                    page_attachment_names.append(str(resolve(filespec.get("/UF") or filespec.get("/F") or "")))
                annotations.append({
                    "page": page_number,
                    "subtype": subtype or None,
                    "rect": plain(annotation.get("/Rect")),
                    "field": plain(annotation.get("/T")),
                    "contents": plain(annotation.get("/Contents")),
                })
            except Exception as exc:
                annotations.append({"page": page_number, "error": str(exc)})
    fields = reader.get_fields() or {}
    try:
        document_attachment_names = [str(attachment.name) for attachment in reader.attachment_list]
    except Exception:
        document_attachment_names = []
    attachment_names = document_attachment_names + page_attachment_names
    metadata = plain(reader.metadata or {})
    return {
        "provider": "pypdf",
        "providerVersion": version,
        "strategy": "read-only",
        "silentFallback": False,
        "source": file_record(source),
        "summary": {
            "pages": len(reader.pages),
            "encrypted": bool(reader.is_encrypted),
            "fields": len(fields),
            "widgets": widgets,
            "annotations": len(annotations),
            "attachments": len(attachment_names),
            "documentAttachments": len(document_attachment_names),
            "pageAttachments": len(page_attachment_names),
        },
        "metadata": metadata,
        "fields": {str(name): plain(field) for name, field in fields.items()},
        "annotations": annotations,
        "attachments": attachment_names,
        "signaturePolicy": signature_policy(reader),
    }


def validate_signed_mutation(policy: dict[str, Any], strategy: str, args: argparse.Namespace) -> None:
    signed = policy["hasSignatureFields"] or policy["hasDocMDP"] or policy["hasFieldMDP"] or policy["hasPerms"]
    if not signed:
        return
    if strategy == "incremental" and not args.allow_signed:
        raise ProviderError(
            "signed/signature-constrained input requires --allow-signed after independent pyHanko and DocMDP review"
        )
    if strategy == "rewrite" and not args.invalidate_signatures:
        raise ProviderError("rewriting signed/signature-constrained input requires --invalidate-signatures")


def parse_field(values: list[str]) -> dict[str, str]:
    fields: dict[str, str] = {}
    for entry in values:
        if "=" not in entry:
            raise ProviderError(f"invalid --field {entry!r}; expected NAME=VALUE")
        name, value = entry.split("=", 1)
        if not name:
            raise ProviderError("form field name must not be empty")
        fields[name] = value
    if not fields:
        raise ProviderError("at least one --field NAME=VALUE is required")
    return fields


def resolve(value: Any) -> Any:
    try:
        return value.get_object()
    except Exception:
        return value


def pdf_name_text(value: Any) -> str | None:
    if value is None:
        return None
    raw = str(resolve(value))
    if raw.startswith("/"):
        raw = raw[1:]
    return re.sub(r"#([0-9A-Fa-f]{2})", lambda match: chr(int(match.group(1), 16)), raw) or None


def attachment_mime(declared: Any, display_name: str) -> tuple[str, str]:
    decoded = pdf_name_text(declared)
    if decoded and "/" in decoded:
        return decoded, "declared"
    guessed = mimetypes.guess_type(display_name, strict=False)[0]
    return guessed or "application/octet-stream", "inferred" if guessed else "default"


def indirect_identity(value: Any) -> str | None:
    reference = getattr(resolve(value), "indirect_reference", None)
    if reference is None:
        return None
    return f"{int(reference.idnum)} {int(reference.generation)} R"


def filespec_payload(filespec: Any) -> tuple[bytes, Any, str, str]:
    filespec = resolve(filespec)
    if not isinstance(filespec, dict):
        raise ProviderError("file attachment has no readable FileSpec dictionary")
    display_name = str(resolve(filespec.get("/UF") or filespec.get("/F") or "attachment"))
    internal_key = str(resolve(filespec.get("/F") or filespec.get("/UF") or display_name))
    embedded = resolve(filespec.get("/EF"))
    if not isinstance(embedded, dict):
        raise ProviderError(f"attachment {display_name!r} has no embedded-file dictionary")
    stream = resolve(embedded.get("/UF") or embedded.get("/F"))
    if stream is None or not hasattr(stream, "get_data"):
        raise ProviderError(f"attachment {display_name!r} has no readable embedded-file stream")
    try:
        payload = bytes(stream.get_data())
    except Exception as exc:
        raise ProviderError(f"attachment {display_name!r} stream could not be decoded: {exc}") from exc
    return payload, stream.get("/Subtype"), display_name, internal_key


def attachment_sources(reader) -> list[dict[str, Any]]:
    sources: list[dict[str, Any]] = []
    try:
        document_attachments = list(reader.attachment_list)
    except Exception as exc:
        raise ProviderError(f"document attachment name tree could not be read: {exc}") from exc
    for ordinal, attachment in enumerate(document_attachments, 1):
        try:
            payload = bytes(attachment.content)
        except Exception as exc:
            raise ProviderError(f"document attachment {attachment.name!r} could not be decoded: {exc}") from exc
        display_name = str(attachment.alternative_name or attachment.name or "attachment")
        sources.append({
            "scope": "document",
            "page": None,
            "annotationIndex": None,
            "internalKey": str(attachment.name),
            "displayName": display_name,
            "declaredMime": attachment.subtype,
            "sourceIdentity": indirect_identity(attachment.pdf_object) or f"document:{ordinal}",
            "payload": payload,
        })
    for page_number, page in enumerate(reader.pages, 1):
        for annotation_index, reference in enumerate(page.get("/Annots", []) or []):
            annotation = resolve(reference)
            if not isinstance(annotation, dict) or str(annotation.get("/Subtype", "")) != "/FileAttachment":
                continue
            payload, declared_mime, display_name, internal_key = filespec_payload(annotation.get("/FS"))
            sources.append({
                "scope": "page",
                "page": page_number,
                "annotationIndex": annotation_index,
                "internalKey": internal_key,
                "displayName": display_name,
                "declaredMime": declared_mime,
                "sourceIdentity": indirect_identity(annotation) or f"page:{page_number}:annotation:{annotation_index}",
                "payload": payload,
            })
    return sources


WINDOWS_RESERVED_NAMES = {
    "con", "prn", "aux", "nul",
    *(f"com{index}" for index in range(1, 10)),
    *(f"lpt{index}" for index in range(1, 10)),
}


def safe_attachment_name(display_name: str, ordinal: int, used: set[str]) -> tuple[str, bool]:
    normalized = unicodedata.normalize("NFC", display_name).replace("\\", "/")
    leaf = normalized.rsplit("/", 1)[-1]
    leaf = "".join("_" if unicodedata.category(character).startswith("C") else character for character in leaf)
    leaf = re.sub(r"[^\w.() -]", "_", leaf, flags=re.UNICODE)
    leaf = re.sub(r"\s+", "_", leaf).strip(" ._")
    if not leaf or leaf in {".", ".."}:
        leaf = f"attachment-{ordinal}"
    stem = Path(leaf).stem[:96].strip(" ._") or f"attachment-{ordinal}"
    suffix = Path(leaf).suffix[:24]
    if stem.casefold() in WINDOWS_RESERVED_NAMES:
        stem = f"_{stem}"
    candidate = f"{stem}{suffix}"
    duplicate = 1
    while candidate.casefold() in used:
        duplicate += 1
        candidate = f"{stem}__{duplicate}{suffix}"
    used.add(candidate.casefold())
    return candidate, candidate != display_name


def extract_attachments(args: argparse.Namespace, reader, version: str) -> dict[str, Any]:
    source = args.input.expanduser().resolve()
    output_directory = args.output_directory.expanduser().resolve()
    manifest = args.manifest.expanduser().resolve() if args.manifest else None
    if output_directory == source or manifest == source:
        raise ProviderError("source PDF, quarantine directory, and manifest must be distinct")
    if output_directory.exists():
        raise ProviderError("quarantine output directory must not already exist")
    if manifest is not None and manifest.exists():
        raise ProviderError("attachment manifest must not already exist")
    if manifest is not None and manifest == output_directory or output_directory in manifest.parents:
        raise ProviderError("attachment manifest must be outside the quarantine directory")
    for value, label in [
        (args.max_attachments, "max-attachments"),
        (args.max_total_bytes, "max-total-bytes"),
        (args.max_attachment_bytes, "max-attachment-bytes"),
    ]:
        if value < 1:
            raise ProviderError(f"--{label} must be positive")

    source_before = file_record(source)
    sources = attachment_sources(reader)
    if len(sources) > args.max_attachments:
        raise ProviderError(f"PDF has {len(sources)} attachments; max-attachments is {args.max_attachments}")
    total_bytes = sum(len(entry["payload"]) for entry in sources)
    if total_bytes > args.max_total_bytes:
        raise ProviderError(f"decoded attachments total {total_bytes} bytes; max-total-bytes is {args.max_total_bytes}")
    oversized = [entry["displayName"] for entry in sources if len(entry["payload"]) > args.max_attachment_bytes]
    if oversized:
        raise ProviderError(
            f"attachment {oversized[0]!r} exceeds max-attachment-bytes {args.max_attachment_bytes}"
        )

    output_directory.parent.mkdir(parents=True, exist_ok=True)
    if manifest is not None:
        manifest.parent.mkdir(parents=True, exist_ok=True)
    temporary_directory = Path(tempfile.mkdtemp(prefix=f".{output_directory.name}.", dir=output_directory.parent))
    temporary_manifest: Path | None = None
    promoted_directory = False
    used_names: set[str] = set()
    records: list[dict[str, Any]] = []
    try:
        for ordinal, entry in enumerate(sources, 1):
            saved_name, sanitized = safe_attachment_name(entry["displayName"], ordinal, used_names)
            destination = (temporary_directory / saved_name).resolve()
            if temporary_directory.resolve() not in destination.parents:
                raise ProviderError(f"attachment {entry['displayName']!r} escaped the quarantine directory")
            destination.write_bytes(entry["payload"])
            written = destination.read_bytes()
            if written != entry["payload"]:
                raise ProviderError(f"attachment {entry['displayName']!r} failed byte-for-byte verification")
            mime, mime_source = attachment_mime(entry["declaredMime"], entry["displayName"])
            records.append({
                "index": ordinal,
                "scope": entry["scope"],
                "page": entry["page"],
                "annotationIndex": entry["annotationIndex"],
                "internalKey": entry["internalKey"],
                "displayName": entry["displayName"],
                "sourceIdentity": entry["sourceIdentity"],
                "mime": mime,
                "mimeSource": mime_source,
                "bytes": len(written),
                "sha256": hashlib.sha256(written).hexdigest(),
                "savedName": saved_name,
                "savedPath": Path(os.path.relpath(
                    output_directory / saved_name,
                    manifest.parent if manifest else output_directory.parent,
                )).as_posix(),
                "nameSanitized": sanitized,
            })
        source_after = file_record(source)
        if source_after != source_before:
            raise ProviderError("source PDF changed during read-only attachment extraction")
        result = {
            "schema": "open-office-artifact-tool.pdf-attachments.v1",
            "provider": "pypdf",
            "providerVersion": version,
            "strategy": "read-only",
            "silentFallback": False,
            "source": source_before,
            "outputDirectory": str(output_directory),
            "budgets": {
                "maxAttachments": args.max_attachments,
                "maxTotalBytes": args.max_total_bytes,
                "maxAttachmentBytes": args.max_attachment_bytes,
            },
            "attachments": records,
            "validation": {
                "sourceUnchanged": True,
                "attachmentCount": len(records),
                "documentAttachments": sum(record["scope"] == "document" for record in records),
                "pageAttachments": sum(record["scope"] == "page" for record in records),
                "totalBytes": total_bytes,
                "allHashesVerified": True,
                "allPathsContained": True,
                "duplicateNamesSeparated": len({record["savedName"].casefold() for record in records}) == len(records),
                "attachmentsOpenedOrExecuted": False,
            },
        }
        rendered = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True) + "\n"
        if manifest is not None:
            with tempfile.NamedTemporaryFile(
                prefix=f".{manifest.name}.", suffix=".tmp", dir=manifest.parent, delete=False
            ) as stream:
                temporary_manifest = Path(stream.name)
                stream.write(rendered.encode("utf-8"))
        temporary_directory.replace(output_directory)
        promoted_directory = True
        if manifest is not None:
            temporary_manifest.replace(manifest)
            temporary_manifest = None
        return result
    except Exception:
        if promoted_directory:
            shutil.rmtree(output_directory, ignore_errors=True)
        raise
    finally:
        shutil.rmtree(temporary_directory, ignore_errors=True)
        if temporary_manifest is not None:
            temporary_manifest.unlink(missing_ok=True)


def button_appearance_states(field: Any, reader=None, name: str | None = None) -> set[str]:
    states: set[str] = set()
    widgets = list(field.get("/Kids", []) or []) or [field]
    if reader is not None and name is not None and not field.get("/Kids"):
        widgets = field_widgets(reader, name)
    for widget_reference in widgets:
        widget = resolve(widget_reference)
        appearance = resolve(widget.get("/AP")) if isinstance(widget, dict) else None
        normal = resolve(appearance.get("/N")) if isinstance(appearance, dict) else None
        if isinstance(normal, dict):
            states.update(str(state) for state in normal.keys())
    return states


def qualified_widget_name(widget: Any) -> str:
    parts: list[str] = []
    current = resolve(widget)
    visited: set[tuple[int, int] | int] = set()
    while isinstance(current, dict):
        reference = getattr(current, "indirect_reference", None)
        identity: tuple[int, int] | int = (
            (int(reference.idnum), int(reference.generation)) if reference is not None else id(current)
        )
        if identity in visited:
            raise ProviderError("cyclic AcroForm parent chain")
        visited.add(identity)
        partial_name = current.get("/T")
        if partial_name is not None:
            parts.append(str(resolve(partial_name)))
        parent = current.get("/Parent")
        if parent is None:
            break
        current = resolve(parent)
    return ".".join(reversed(parts))


def field_widgets(reader, name: str) -> list[Any]:
    widgets: list[Any] = []
    for page in reader.pages:
        for reference in page.get("/Annots", []) or []:
            widget = resolve(reference)
            if isinstance(widget, dict) and str(widget.get("/Subtype", "")) == "/Widget":
                if qualified_widget_name(widget) == name:
                    widgets.append(widget)
    return widgets


def prepare_field_values(reader, requested: dict[str, str], NameObject) -> tuple[dict[str, Any], dict[str, Any]]:
    known_fields = reader.get_fields() or {}
    missing = sorted(set(requested) - set(known_fields))
    if missing:
        raise ProviderError(f"form field(s) not found: {', '.join(missing)}")

    prepared: dict[str, Any] = {}
    evidence: dict[str, Any] = {}
    for name, value in requested.items():
        field = known_fields[name]
        field_type = str(field.get("/FT", ""))
        flags = int(field.get("/Ff", 0) or 0)
        if flags & 1:
            raise ProviderError(f"form field {name!r} is read-only")
        if field_type == "/Sig":
            raise ProviderError(f"signature field {name!r} cannot be filled by the basic form adapter")
        if field_type == "/Btn":
            if flags & (1 << 16):
                raise ProviderError(f"push button {name!r} has no fillable value")
            available = sorted(button_appearance_states(field, reader, name))
            target = "/Off" if value.strip().casefold() in {"", "off", "/off"} else f"/{value.lstrip('/')}"
            if target != "/Off" and target not in available:
                choices = ", ".join(state.lstrip("/") for state in available if state != "/Off") or "<none>"
                raise ProviderError(f"button field {name!r} has no appearance state {value!r}; available: {choices}")
            prepared[name] = NameObject(target)
            evidence[name] = {"fieldType": field_type, "appearanceState": target, "availableStates": available}
        elif field_type in {"/Tx", "/Ch"}:
            prepared[name] = value
            evidence[name] = {"fieldType": field_type}
        else:
            raise ProviderError(f"form field {name!r} has unsupported type {field_type or '<missing>'}")
    return prepared, evidence


def validate_filled_fields(reader, expected: dict[str, Any], evidence: dict[str, Any]) -> None:
    fields = reader.get_fields() or {}
    for name, expected_value in expected.items():
        field = fields.get(name)
        if field is None:
            raise ProviderError(f"filled field {name!r} is missing from provider output")
        field_type = evidence[name]["fieldType"]
        actual_value = str(resolve(field.get("/V", "")))
        widgets = field_widgets(reader, name)
        if not widgets:
            raise ProviderError(f"filled field {name!r} has no output widget")
        for widget in widgets:
            appearance = resolve(widget.get("/AP"))
            normal = resolve(appearance.get("/N")) if isinstance(appearance, dict) else None
            if normal is None:
                raise ProviderError(f"filled field {name!r} output widget has no normal appearance")
        if field_type == "/Btn":
            target = str(expected_value)
            if actual_value != target:
                raise ProviderError(f"button field {name!r} value is {actual_value!r}, expected {target!r}")
            selected = []
            for widget in widgets:
                appearance = resolve(widget.get("/AP")) if isinstance(widget, dict) else None
                normal = resolve(appearance.get("/N")) if isinstance(appearance, dict) else None
                if not isinstance(normal, dict) or "/Off" not in {str(state) for state in normal.keys()}:
                    raise ProviderError(f"button field {name!r} output widget has no normal appearance states")
                state = str(resolve(widget.get("/AS", "/Off")))
                if state != "/Off":
                    selected.append(state)
            expected_selected = [] if target == "/Off" else [target]
            if selected != expected_selected:
                raise ProviderError(
                    f"button field {name!r} appearance state is {selected!r}, expected {expected_selected!r}"
                )
        elif actual_value != str(expected_value):
            raise ProviderError(f"form field {name!r} value is {actual_value!r}, expected {str(expected_value)!r}")


def parse_rect(value: str) -> tuple[float, float, float, float]:
    try:
        values = tuple(float(item.strip()) for item in value.split(","))
    except ValueError as exc:
        raise ProviderError("--rect must contain four comma-separated numbers") from exc
    if len(values) != 4 or values[2] <= values[0] or values[3] <= values[1]:
        raise ProviderError("--rect must be x0,y0,x1,y1 with positive width and height")
    return values


def page_reference_key(page: Any) -> tuple[int, int]:
    reference = getattr(page, "indirect_reference", None)
    if reference is None:
        raise ProviderError("page has no indirect object identity")
    return int(reference.idnum), int(reference.generation)


def page_geometry(page: Any) -> dict[str, Any]:
    boxes: dict[str, list[float]] = {}
    for name in ("mediabox", "cropbox", "trimbox", "bleedbox", "artbox"):
        box = getattr(page, name)
        boxes[name] = [float(box.left), float(box.bottom), float(box.right), float(box.top)]
    return {"boxes": boxes, "rotation": int(page.get("/Rotate", 0) or 0)}


def destination_page_number(reader: Any, destination: Any) -> int | None:
    destination = resolve(destination)
    if destination is None:
        return None
    if isinstance(destination, str):
        named = reader.named_destinations.get(destination)
        return reader.get_destination_page_number(named) + 1 if named is not None else None
    try:
        number = reader.get_destination_page_number(destination)
        if number >= 0:
            return number + 1
    except Exception:
        pass
    if isinstance(destination, (list, tuple)) and destination:
        target = destination[0]
        target = resolve(target)
        target_reference = getattr(target, "indirect_reference", None)
        for page_number, page in enumerate(reader.pages, 1):
            reference = getattr(page, "indirect_reference", None)
            if reference is not None and target_reference is not None:
                if (int(reference.idnum), int(reference.generation)) == (
                    int(target_reference.idnum), int(target_reference.generation)
                ):
                    return page_number
            if target is page:
                return page_number
    return None


def outline_records(reader: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []

    def walk(items: list[Any], parents: tuple[str, ...] = ()) -> None:
        previous_title: str | None = None
        for item in items:
            if isinstance(item, list):
                walk(item, parents + ((previous_title,) if previous_title else ()))
                continue
            title = str(getattr(item, "title", item))
            page = destination_page_number(reader, item)
            if page is None:
                raise ProviderError(f"outline item {title!r} has an unresolved destination")
            records.append({"title": title, "page": page, "parentPath": list(parents)})
            previous_title = title

    walk(list(reader.outline or []))
    return records


def named_destination_records(reader: Any) -> dict[str, int]:
    records: dict[str, int] = {}
    for name, destination in reader.named_destinations.items():
        page = destination_page_number(reader, destination)
        if page is None:
            raise ProviderError(f"named destination {name!r} has an unresolved page")
        records[str(name)] = page
    return records


def internal_link_records(reader: Any) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for source_page, page in enumerate(reader.pages, 1):
        for reference in page.get("/Annots", []) or []:
            annotation = resolve(reference)
            if not isinstance(annotation, dict) or str(annotation.get("/Subtype", "")) != "/Link":
                continue
            destination = annotation.get("/Dest")
            action = resolve(annotation.get("/A")) if annotation.get("/A") is not None else None
            if destination is None and isinstance(action, dict) and str(action.get("/S", "")) == "/GoTo":
                destination = action.get("/D")
            if destination is None:
                continue
            target_page = destination_page_number(reader, destination)
            if target_page is None:
                raise ProviderError(f"internal link on page {source_page} has an unresolved destination")
            records.append({
                "page": source_page,
                "targetPage": target_page,
                "rect": [float(value) for value in annotation.get("/Rect", [])],
            })
    return records


def navigation_evidence(reader: Any) -> dict[str, Any]:
    return {
        "outlines": outline_records(reader),
        "namedDestinations": named_destination_records(reader),
        "internalLinks": internal_link_records(reader),
    }


def render_watermark(PdfReader: Any, width: float, height: float, watermark: dict[str, Any]) -> Any:
    try:
        from reportlab.pdfgen import canvas
    except ImportError as exc:
        raise ProviderError("reportlab is required to render the pypdf watermark overlay") from exc
    stream = BytesIO()
    document = canvas.Canvas(stream, pagesize=(width, height), invariant=1, pageCompression=1)
    document.saveState()
    document.setFillAlpha(watermark["opacity"])
    document.setFillColorRGB(0.35, 0.35, 0.35)
    document.setFont("Helvetica-Bold", watermark["fontSize"])
    document.translate(width / 2, height / 2)
    document.rotate(watermark["angle"])
    document.drawCentredString(0, -watermark["fontSize"] / 3, watermark["text"])
    document.restoreState()
    document.save()
    stream.seek(0)
    return PdfReader(stream, strict=True).pages[0]


def normalized_merge_manifest(args: argparse.Namespace, PdfReader: Any) -> dict[str, Any]:
    manifest_path = args.manifest.expanduser().resolve()
    if not manifest_path.is_file():
        raise ProviderError("merge-stamp manifest must be an existing JSON file")
    if manifest_path.stat().st_size > args.max_manifest_bytes:
        raise ProviderError(f"merge-stamp manifest exceeds max-manifest-bytes {args.max_manifest_bytes}")
    try:
        raw = json.loads(manifest_path.read_text("utf-8"))
    except (OSError, json.JSONDecodeError) as exc:
        raise ProviderError(f"merge-stamp manifest is not valid UTF-8 JSON: {exc}") from exc
    if not isinstance(raw, dict) or raw.get("schema") != MERGE_STAMP_SCHEMA:
        raise ProviderError(f"merge-stamp manifest schema must be {MERGE_STAMP_SCHEMA!r}")
    source_specs = raw.get("sources")
    if not isinstance(source_specs, list) or not (2 <= len(source_specs) <= args.max_inputs):
        raise ProviderError(f"merge-stamp sources must contain between 2 and {args.max_inputs} entries")

    sources: list[dict[str, Any]] = []
    source_by_id: dict[str, dict[str, Any]] = {}
    source_paths: set[Path] = set()
    total_bytes = 0
    total_pages = 0
    all_named_destinations: dict[str, str] = {}
    for spec in source_specs:
        if not isinstance(spec, dict):
            raise ProviderError("every merge-stamp source must be an object")
        source_id = spec.get("id")
        source_path_value = spec.get("path")
        if not isinstance(source_id, str) or not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9._-]{0,63}", source_id):
            raise ProviderError("source id must be 1-64 portable characters")
        if source_id in source_by_id:
            raise ProviderError(f"duplicate merge-stamp source id {source_id!r}")
        if not isinstance(source_path_value, str) or not source_path_value.strip():
            raise ProviderError(f"source {source_id!r} path must be a non-empty string")
        source_path = Path(source_path_value).expanduser()
        if not source_path.is_absolute():
            source_path = manifest_path.parent / source_path
        source_path = source_path.resolve()
        if source_path in source_paths:
            raise ProviderError(f"source path appears more than once: {source_path}")
        if not source_path.is_file():
            raise ProviderError(f"source {source_id!r} is not an existing PDF: {source_path}")
        if source_path in {manifest_path, args.output.expanduser().resolve()}:
            raise ProviderError("manifest, source PDFs, and output PDF must all be distinct")
        record = file_record(source_path)
        total_bytes += int(record["bytes"])
        if total_bytes > args.max_total_bytes:
            raise ProviderError(f"source PDFs exceed max-total-bytes {args.max_total_bytes}")
        reader = PdfReader(str(source_path), strict=False)
        if reader.is_encrypted:
            raise ProviderError(
                f"source {source_id!r} is encrypted; merge-stamp has no explicit output-encryption policy"
            )
        page_count = len(reader.pages)
        if page_count < 1:
            raise ProviderError(f"source {source_id!r} has no pages")
        total_pages += page_count
        if total_pages > args.max_pages:
            raise ProviderError(f"source PDFs exceed max-pages {args.max_pages}")
        signature = signature_policy(reader)
        signed = signature["hasSignatureFields"] or signature["hasDocMDP"] or signature["hasFieldMDP"] or signature["hasPerms"]
        if signed and not args.invalidate_signatures:
            raise ProviderError(
                f"source {source_id!r} is signed/signature-constrained; rewrite requires --invalidate-signatures"
            )
        navigation = navigation_evidence(reader)
        for name in navigation["namedDestinations"]:
            if name in all_named_destinations:
                raise ProviderError(
                    f"named destination {name!r} collides between sources "
                    f"{all_named_destinations[name]!r} and {source_id!r}"
                )
            all_named_destinations[name] = source_id
        source = {
            "id": source_id,
            "path": source_path,
            "reader": reader,
            "record": record,
            "pageCount": page_count,
            "signaturePolicy": signature,
            "navigation": navigation,
        }
        sources.append(source)
        source_by_id[source_id] = source
        source_paths.add(source_path)

    sequence_specs = raw.get("sequence")
    if not isinstance(sequence_specs, list) or not sequence_specs:
        raise ProviderError("merge-stamp sequence must be a non-empty array")
    sequence: list[dict[str, Any]] = []
    selected: list[tuple[str, int]] = []
    for segment in sequence_specs:
        if not isinstance(segment, dict) or segment.get("source") not in source_by_id:
            raise ProviderError("every sequence segment must reference a declared source")
        source_id = str(segment["source"])
        page_spec = segment.get("pages")
        if page_spec == "all":
            pages = list(range(1, source_by_id[source_id]["pageCount"] + 1))
        elif isinstance(page_spec, list) and page_spec and all(isinstance(page, int) and not isinstance(page, bool) for page in page_spec):
            pages = list(page_spec)
        else:
            raise ProviderError("sequence pages must be 'all' or a non-empty array of one-based integers")
        for page in pages:
            if page < 1 or page > source_by_id[source_id]["pageCount"]:
                raise ProviderError(f"sequence page {page} is outside source {source_id!r}")
            selected.append((source_id, page))
        sequence.append({"source": source_id, "pages": pages})
    expected = [(source["id"], page) for source in sources for page in range(1, source["pageCount"] + 1)]
    if len(selected) != len(set(selected)):
        raise ProviderError("merge-stamp sequence must not duplicate a source page")
    if set(selected) != set(expected):
        missing = sorted(set(expected) - set(selected))
        extra = sorted(set(selected) - set(expected))
        raise ProviderError(f"merge-stamp sequence must select every source page exactly once; missing={missing}, extra={extra}")

    watermark_specs = raw.get("watermarks")
    if not isinstance(watermark_specs, list) or not watermark_specs:
        raise ProviderError("merge-stamp watermarks must be a non-empty array")
    watermarks: list[dict[str, Any]] = []
    watermarked_sources: set[str] = set()
    for watermark in watermark_specs:
        if not isinstance(watermark, dict) or watermark.get("source") not in source_by_id:
            raise ProviderError("every watermark must reference a declared source")
        source_id = str(watermark["source"])
        if source_id in watermarked_sources:
            raise ProviderError(f"source {source_id!r} has more than one watermark definition")
        text = watermark.get("text")
        opacity = watermark.get("opacity")
        angle = watermark.get("angle", 45)
        font_size = watermark.get("fontSize", 48)
        if not isinstance(text, str) or not text.strip() or len(text) > 128:
            raise ProviderError("watermark text must contain 1-128 characters")
        if not isinstance(opacity, (int, float)) or isinstance(opacity, bool) or not (0 < float(opacity) <= 1):
            raise ProviderError("watermark opacity must be greater than 0 and at most 1")
        if not isinstance(angle, (int, float)) or isinstance(angle, bool) or not (-360 <= float(angle) <= 360):
            raise ProviderError("watermark angle must be between -360 and 360 degrees")
        if not isinstance(font_size, (int, float)) or isinstance(font_size, bool) or not (6 <= float(font_size) <= 240):
            raise ProviderError("watermark fontSize must be between 6 and 240 points")
        rotated = [
            page_number
            for page_number, page in enumerate(source_by_id[source_id]["reader"].pages, 1)
            if int(page.get("/Rotate", 0) or 0) % 360 != 0
        ]
        if rotated:
            raise ProviderError(
                f"watermark source {source_id!r} has rotated page dictionaries {rotated}; "
                "this bounded primitive preserves them but cannot guarantee upright stamp placement"
            )
        watermarks.append({
            "source": source_id,
            "text": text,
            "opacity": float(opacity),
            "angle": float(angle),
            "fontSize": float(font_size),
        })
        watermarked_sources.add(source_id)

    return {
        "manifestPath": manifest_path,
        "manifestRecord": file_record(manifest_path),
        "sources": sources,
        "sequence": sequence,
        "selected": selected,
        "watermarks": watermarks,
        "totalBytes": total_bytes,
        "totalPages": total_pages,
    }


def rewrite_page_tree(writer: Any, ordered_references: list[Any], NameObject: Any, ArrayObject: Any, NumberObject: Any) -> None:
    pages_root = resolve(writer.root_object["/Pages"])
    if not isinstance(pages_root, dict):
        raise ProviderError("writer page tree is unavailable")
    existing = list(pages_root.get("/Kids", []) or [])
    if len(existing) != len(ordered_references):
        raise ProviderError("writer page tree does not match the selected source-page count")
    pages_root[NameObject("/Kids")] = ArrayObject(ordered_references)
    pages_root[NameObject("/Count")] = NumberObject(len(ordered_references))


def page_opacities(page: Any) -> list[float]:
    resources = resolve(page.get("/Resources"))
    states = resolve(resources.get("/ExtGState")) if isinstance(resources, dict) else None
    values: list[float] = []
    if isinstance(states, dict):
        for state in states.values():
            state = resolve(state)
            if isinstance(state, dict) and state.get("/ca") is not None:
                values.append(float(state["/ca"]))
    return sorted(values)


def merge_stamp(args: argparse.Namespace, pypdf: Any, PdfReader: Any, PdfWriter: Any, version: str) -> dict[str, Any]:
    from pypdf import Transformation
    from pypdf.generic import ArrayObject, NameObject, NumberObject

    output = args.output.expanduser().resolve()
    manifest = normalized_merge_manifest(args, PdfReader)
    if output == manifest["manifestPath"]:
        raise ProviderError("merge-stamp manifest and output must differ")
    output.parent.mkdir(parents=True, exist_ok=True)
    writer = PdfWriter()
    page_references: dict[tuple[str, int], Any] = {}
    for source in manifest["sources"]:
        offset = len(writer.pages)
        writer.append(source["reader"], import_outline=True)
        for page_number in range(1, source["pageCount"] + 1):
            page = writer.pages[offset + page_number - 1]
            if page.indirect_reference is None:
                raise ProviderError("pypdf did not assign an indirect reference to an imported page")
            page_references[(source["id"], page_number)] = page.indirect_reference

    ordered_references = [page_references[identity] for identity in manifest["selected"]]
    rewrite_page_tree(writer, ordered_references, NameObject, ArrayObject, NumberObject)
    output_page_map = {identity: index for index, identity in enumerate(manifest["selected"], 1)}
    watermark_records: list[dict[str, Any]] = []
    for watermark in manifest["watermarks"]:
        target_pages = [
            output_page_map[(watermark["source"], source_page)]
            for source_page in range(1, next(source["pageCount"] for source in manifest["sources"] if source["id"] == watermark["source"]) + 1)
        ]
        for output_page in target_pages:
            page = ordered_references[output_page - 1].get_object()
            box = page.mediabox
            width = float(box.width)
            height = float(box.height)
            overlay = render_watermark(PdfReader, width, height, watermark)
            page.merge_transformed_page(
                overlay,
                Transformation().translate(tx=float(box.left), ty=float(box.bottom)),
                over=True,
                expand=False,
            )
        watermark_records.append({**watermark, "outputPages": target_pages})

    source_records_before = {source["id"]: source["record"] for source in manifest["sources"]}
    temporary_name: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent, delete=False) as stream:
            temporary_name = Path(stream.name)
        writer.write(str(temporary_name))
        if not temporary_name.read_bytes().startswith(b"%PDF-"):
            raise ProviderError("provider output is not a PDF")
        result_reader = PdfReader(str(temporary_name), strict=True)
        if len(result_reader.pages) != manifest["totalPages"]:
            raise ProviderError("provider output page count does not match the merge manifest")

        geometry_records: list[dict[str, Any]] = []
        for output_page, (source_id, source_page) in enumerate(manifest["selected"], 1):
            source = next(item for item in manifest["sources"] if item["id"] == source_id)
            expected_geometry = page_geometry(source["reader"].pages[source_page - 1])
            actual_geometry = page_geometry(result_reader.pages[output_page - 1])
            if actual_geometry != expected_geometry:
                raise ProviderError(
                    f"page geometry changed for {source_id} page {source_page}: "
                    f"expected {expected_geometry}, got {actual_geometry}"
                )
            geometry_records.append({
                "outputPage": output_page,
                "source": source_id,
                "sourcePage": source_page,
                **actual_geometry,
            })

        expected_outlines: list[dict[str, Any]] = []
        expected_named: dict[str, int] = {}
        expected_links: list[dict[str, Any]] = []
        for source in manifest["sources"]:
            for record in source["navigation"]["outlines"]:
                expected_outlines.append({**record, "page": output_page_map[(source["id"], record["page"])]})
            for name, page in source["navigation"]["namedDestinations"].items():
                expected_named[name] = output_page_map[(source["id"], page)]
            for record in source["navigation"]["internalLinks"]:
                expected_links.append({
                    **record,
                    "page": output_page_map[(source["id"], record["page"])],
                    "targetPage": output_page_map[(source["id"], record["targetPage"])],
                })
        actual_navigation = navigation_evidence(result_reader)
        sort_records = lambda records: sorted(records, key=lambda record: json.dumps(record, sort_keys=True))
        if sort_records(actual_navigation["outlines"]) != sort_records(expected_outlines):
            raise ProviderError("output outline destinations do not match the reordered source pages")
        if actual_navigation["namedDestinations"] != expected_named:
            raise ProviderError("output named destinations do not match the reordered source pages")
        if sort_records(actual_navigation["internalLinks"]) != sort_records(expected_links):
            raise ProviderError("output internal links do not match the reordered source pages")

        watermark_validation: list[dict[str, Any]] = []
        all_watermark_texts = {watermark["text"] for watermark in watermark_records}
        for watermark in watermark_records:
            target_pages = set(watermark["outputPages"])
            page_counts = []
            for page_number, page in enumerate(result_reader.pages, 1):
                count = (page.extract_text() or "").count(watermark["text"])
                expected_count = 1 if page_number in target_pages else 0
                if count != expected_count:
                    raise ProviderError(
                        f"watermark {watermark['text']!r} count on page {page_number} is {count}, expected {expected_count}"
                    )
                opacities = page_opacities(page)
                if page_number in target_pages and not any(abs(value - watermark["opacity"]) <= 0.001 for value in opacities):
                    raise ProviderError(
                        f"watermark opacity {watermark['opacity']} is not present on output page {page_number}"
                    )
                page_counts.append({"page": page_number, "count": count, "opacities": opacities})
            watermark_validation.append({**watermark, "pages": page_counts})
        if not all_watermark_texts:
            raise ProviderError("merge-stamp produced no watermark validation terms")

        source_records_after = {source["id"]: file_record(source["path"]) for source in manifest["sources"]}
        if source_records_after != source_records_before:
            raise ProviderError("one or more source PDFs changed during merge-stamp")
        temporary_name.replace(output)
        temporary_name = None
    finally:
        if temporary_name is not None:
            temporary_name.unlink(missing_ok=True)

    return {
        "schema": MERGE_STAMP_RESULT_SCHEMA,
        "provider": "pypdf",
        "providerVersion": version,
        "helperProvider": "reportlab",
        "strategy": "rewrite",
        "silentFallback": False,
        "manifest": manifest["manifestRecord"],
        "sources": [
            {
                "id": source["id"],
                **source["record"],
                "pages": source["pageCount"],
                "signaturePolicyBefore": source["signaturePolicy"],
            }
            for source in manifest["sources"]
        ],
        "output": file_record(output),
        "operation": {
            "type": "merge-stamp",
            "sequence": manifest["sequence"],
            "watermarks": watermark_records,
        },
        "validation": {
            "allSourcesUnchanged": True,
            "pageCount": manifest["totalPages"],
            "pageMap": geometry_records,
            "pageGeometryPreserved": True,
            "navigation": {
                "outlines": expected_outlines,
                "namedDestinations": expected_named,
                "internalLinks": expected_links,
                "preserved": True,
            },
            "watermarks": watermark_validation,
        },
    }


def write_mutation(args: argparse.Namespace, reader, PdfWriter, Text, NameObject, version: str) -> dict[str, Any]:
    source = args.input.expanduser().resolve()
    output = args.output.expanduser().resolve()
    if source == output:
        raise ProviderError("input and output must differ; in-place source overwrite is forbidden")
    if not source.is_file():
        raise ProviderError("input must be an existing PDF")
    if reader.is_encrypted:
        raise ProviderError(
            "the shipped pypdf mutation adapter refuses encrypted input because it has no explicit output-encryption policy; "
            "select and probe a provider with the required encryption contract"
        )
    output.parent.mkdir(parents=True, exist_ok=True)
    policy = signature_policy(reader)
    validate_signed_mutation(policy, args.strategy, args)
    if args.command == "fill-form" and args.flatten and args.strategy != "rewrite":
        raise ProviderError("--flatten requires rewrite; incremental flattening is not supported")

    writer = PdfWriter(reader, incremental=True) if args.strategy == "incremental" else PdfWriter(clone_from=reader)
    operation: dict[str, Any]
    if args.command == "fill-form":
        fields = parse_field(args.field)
        prepared_fields, field_evidence = prepare_field_values(reader, fields, NameObject)
        writer.update_page_form_field_values(
            None,
            prepared_fields,
            auto_regenerate=False,
            flatten=args.flatten,
        )
        operation = {
            "type": "fill-form",
            "fields": sorted(fields),
            "fieldEvidence": field_evidence,
            "flatten": bool(args.flatten),
        }
    else:
        if args.page < 1 or args.page > len(writer.pages):
            raise ProviderError(f"--page must be between 1 and {len(writer.pages)}")
        annotation = Text(rect=parse_rect(args.rect), text=args.text, open=args.open)
        writer.add_annotation(args.page - 1, annotation)
        operation = {"type": "add-note", "page": args.page, "rect": list(parse_rect(args.rect)), "open": bool(args.open)}

    source_bytes = source.read_bytes()
    temporary_name = None
    try:
        with tempfile.NamedTemporaryFile(prefix=f".{output.name}.", suffix=".tmp", dir=output.parent, delete=False) as stream:
            temporary_name = Path(stream.name)
        writer.write(str(temporary_name))
        result_bytes = temporary_name.read_bytes()
        if not result_bytes.startswith(b"%PDF-"):
            raise ProviderError("provider output is not a PDF")
        if args.strategy == "incremental" and not result_bytes.startswith(source_bytes):
            raise ProviderError("incremental output does not preserve the exact original byte prefix")
        result_reader = type(reader)(str(temporary_name), strict=False)
        if args.command == "fill-form":
            validate_filled_fields(result_reader, prepared_fields, field_evidence)
        temporary_name.replace(output)
        temporary_name = None
    finally:
        if temporary_name is not None:
            temporary_name.unlink(missing_ok=True)

    result_reader = type(reader)(str(output), strict=False)
    result = {
        "provider": "pypdf",
        "providerVersion": version,
        "strategy": args.strategy,
        "silentFallback": False,
        "source": file_record(source),
        "output": file_record(output),
        "originalPrefixPreserved": output.read_bytes().startswith(source_bytes),
        "signaturePolicyBefore": policy,
        "signaturePolicyAfter": signature_policy(result_reader),
        "operation": operation,
    }
    return result


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    subparsers = root.add_subparsers(dest="command", required=True)

    inspect = subparsers.add_parser("inspect", help="report forms, annotations, attachments, and signature constraints")
    inspect.add_argument("input", type=Path)
    inspect.add_argument("--output", type=Path, help="optional JSON report path")
    inspect.add_argument("--password-env")
    inspect.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    inspect.add_argument("--max-pages", type=int, default=2_000)

    extract = subparsers.add_parser(
        "extract-attachments",
        help="extract document-level and page-level attachments into a path-safe quarantine directory",
    )
    extract.add_argument("input", type=Path)
    extract.add_argument("output_directory", type=Path)
    extract.add_argument("--manifest", type=Path, help="optional JSON inventory path outside the quarantine directory")
    extract.add_argument("--password-env")
    extract.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    extract.add_argument("--max-pages", type=int, default=2_000)
    extract.add_argument("--max-attachments", type=int, default=1_000)
    extract.add_argument("--max-total-bytes", type=int, default=1024 * 1024 * 1024)
    extract.add_argument("--max-attachment-bytes", type=int, default=512 * 1024 * 1024)

    def mutation(name: str):
        command = subparsers.add_parser(name)
        command.add_argument("input", type=Path)
        command.add_argument("output", type=Path)
        command.add_argument("--strategy", choices=["rewrite", "incremental"], required=True)
        command.add_argument("--password-env")
        command.add_argument("--allow-signed", action="store_true")
        command.add_argument("--invalidate-signatures", action="store_true")
        command.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
        command.add_argument("--max-pages", type=int, default=2_000)
        return command

    fill = mutation("fill-form")
    fill.add_argument("--field", action="append", default=[])
    fill.add_argument("--flatten", action="store_true")

    note = mutation("add-note")
    note.add_argument("--page", type=int, required=True)
    note.add_argument("--rect", required=True)
    note.add_argument("--text", required=True)
    note.add_argument("--open", action="store_true")

    merge = subparsers.add_parser(
        "merge-stamp",
        help="merge and reorder complete source PDFs, preserve navigation, and selectively watermark by source id",
    )
    merge.add_argument("manifest", type=Path)
    merge.add_argument("output", type=Path)
    merge.add_argument("--strategy", choices=["rewrite"], required=True)
    merge.add_argument("--invalidate-signatures", action="store_true")
    merge.add_argument("--max-manifest-bytes", type=int, default=1024 * 1024)
    merge.add_argument("--max-inputs", type=int, default=100)
    merge.add_argument("--max-total-bytes", type=int, default=2 * 1024 * 1024 * 1024)
    merge.add_argument("--max-pages", type=int, default=10_000)
    return root


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parser().parse_args()
    output_on_failure = getattr(args, "output", None) if args.command != "inspect" else None
    try:
        pypdf, PdfReader, PdfWriter, Text, NameObject = require_pypdf()
        if args.command == "merge-stamp":
            if args.max_manifest_bytes < 1 or args.max_inputs < 2 or args.max_total_bytes < 1 or args.max_pages < 1:
                raise ProviderError("merge-stamp budgets must be positive and max-inputs must be at least 2")
            result = merge_stamp(args, pypdf, PdfReader, PdfWriter, pypdf.__version__)
            print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
            return 0
        source = args.input.expanduser().resolve()
        if not source.is_file():
            raise ProviderError("input must be an existing PDF")
        if args.max_bytes < 1 or args.max_pages < 1:
            raise ProviderError("max-bytes and max-pages must be positive")
        if source.stat().st_size > args.max_bytes:
            raise ProviderError(f"PDF is {source.stat().st_size} bytes; max-bytes is {args.max_bytes}")
        reader = open_reader(PdfReader, source, getattr(args, "password_env", None))
        if len(reader.pages) > args.max_pages:
            raise ProviderError(f"PDF has {len(reader.pages)} pages; max-pages is {args.max_pages}")
        if args.command == "inspect":
            result = inspect_reader(reader, source, pypdf.__version__)
            rendered = json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True)
            if args.output:
                destination = args.output.expanduser().resolve()
                if destination == source:
                    raise ProviderError("JSON report path must differ from the source PDF")
                destination.parent.mkdir(parents=True, exist_ok=True)
                destination.write_text(rendered + "\n", "utf-8")
            else:
                print(rendered)
        elif args.command == "extract-attachments":
            result = extract_attachments(args, reader, pypdf.__version__)
            print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        else:
            result = write_mutation(args, reader, PdfWriter, Text, NameObject, pypdf.__version__)
            print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        if output_on_failure:
            Path(output_on_failure).expanduser().resolve().unlink(missing_ok=True)
        print(json.dumps({"ok": False, "provider": "pypdf", "error": str(exc), "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
