#!/usr/bin/env python3
"""Inspect or make bounded form/annotation edits with an explicit pypdf save policy."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import sys
import tempfile
from typing import Any


class ProviderError(RuntimeError):
    pass


def require_pypdf():
    try:
        import pypdf
        from pypdf import PdfReader, PdfWriter
        from pypdf.annotations import Text
    except ImportError as exc:
        raise ProviderError("pypdf is required in the selected Python environment") from exc
    return pypdf, PdfReader, PdfWriter, Text


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
    for page_number, page in enumerate(reader.pages, 1):
        page_annots = page.get("/Annots") or []
        for reference in page_annots:
            try:
                annotation = reference.get_object()
                subtype = str(annotation.get("/Subtype", ""))
                if subtype == "/Widget":
                    widgets += 1
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
        attachment_names = sorted(str(name) for name in reader.attachments.keys())
    except Exception:
        attachment_names = []
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


def parse_rect(value: str) -> tuple[float, float, float, float]:
    try:
        values = tuple(float(item.strip()) for item in value.split(","))
    except ValueError as exc:
        raise ProviderError("--rect must contain four comma-separated numbers") from exc
    if len(values) != 4 or values[2] <= values[0] or values[3] <= values[1]:
        raise ProviderError("--rect must be x0,y0,x1,y1 with positive width and height")
    return values


def write_mutation(args: argparse.Namespace, reader, PdfWriter, Text, version: str) -> dict[str, Any]:
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
        known_fields = set((reader.get_fields() or {}).keys())
        missing = sorted(set(fields) - known_fields)
        if missing:
            raise ProviderError(f"form field(s) not found: {', '.join(missing)}")
        writer.update_page_form_field_values(
            None,
            fields,
            auto_regenerate=False,
            flatten=args.flatten,
        )
        operation = {"type": "fill-form", "fields": sorted(fields), "flatten": bool(args.flatten)}
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
    args = parser().parse_args()
    output_on_failure = getattr(args, "output", None) if args.command != "inspect" else None
    try:
        pypdf, PdfReader, PdfWriter, Text = require_pypdf()
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
        else:
            result = write_mutation(args, reader, PdfWriter, Text, pypdf.__version__)
            print(json.dumps(result, ensure_ascii=False, indent=2, sort_keys=True))
        return 0
    except Exception as exc:
        if output_on_failure:
            Path(output_on_failure).expanduser().resolve().unlink(missing_ok=True)
        print(json.dumps({"ok": False, "provider": "pypdf", "error": str(exc), "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
