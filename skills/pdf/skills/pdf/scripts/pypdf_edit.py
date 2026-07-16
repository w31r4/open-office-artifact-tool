#!/usr/bin/env python3
"""Inspect or make bounded form/annotation edits with an explicit pypdf save policy."""

from __future__ import annotations

import argparse
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
    return root


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parser().parse_args()
    output_on_failure = getattr(args, "output", None) if args.command != "inspect" else None
    try:
        pypdf, PdfReader, PdfWriter, Text, NameObject = require_pypdf()
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
