#!/usr/bin/env python3
"""Run bounded qpdf structure inspection and source-bound transactional rewrites."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from typing import Any, Iterator


SCHEMA_INSPECT = "open-office-artifact-tool.qpdf-inspect.v1"
SCHEMA_REWRITE = "open-office-artifact-tool.qpdf-rewrite.v1"
DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_JSON_BYTES = 128 * 1024 * 1024
DEFAULT_MAX_CHECK_BYTES = 2 * 1024 * 1024
DEFAULT_TIMEOUT_SECONDS = 120
MAX_CHECK_LINES = 500
MAX_JSON_NODES = 2_000_000
MAX_EVIDENCE_TEXT_CHARS = 4_096


class ProviderError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def qpdf_path() -> Path:
    configured = os.environ.get("OPEN_OFFICE_PDF_QPDF", "").strip()
    resolved = Path(configured).expanduser() if configured else Path(shutil.which("qpdf") or "")
    if not str(resolved) or not resolved.is_file() or not os.access(resolved, os.X_OK):
        detail = f" configured by OPEN_OFFICE_PDF_QPDF={configured!r}" if configured else " on PATH"
        raise ProviderError(f"qpdf executable is unavailable{detail}")
    return resolved.resolve()


def run_qpdf(
    executable: Path,
    arguments: list[str],
    *,
    timeout_seconds: int,
    stdout_path: Path | None = None,
    max_stdout_bytes: int = DEFAULT_MAX_CHECK_BYTES,
    max_stderr_bytes: int = DEFAULT_MAX_CHECK_BYTES,
) -> subprocess.CompletedProcess:
    command = [str(executable), *arguments]
    stdout_buffer = bytearray()
    stderr_buffer = bytearray()
    violations: list[str] = []
    violation_lock = threading.Lock()
    output_stream = stdout_path.open("wb", buffering=0) if stdout_path is not None else None
    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
        )
    except OSError as exc:
        if output_stream is not None:
            output_stream.close()
        raise ProviderError(f"qpdf could not start: {exc}") from exc

    def record_violation(message: str) -> None:
        with violation_lock:
            if not violations:
                violations.append(message)
        try:
            process.kill()
        except OSError:
            pass

    def pump(stream: Any, buffer: bytearray, limit: int, label: str, sink: Any = None) -> None:
        total = 0
        try:
            while True:
                chunk = stream.read(64 * 1024)
                if not chunk:
                    break
                if total + len(chunk) > limit:
                    record_violation(f"qpdf {label} exceeded the {limit} byte budget")
                    break
                total += len(chunk)
                if sink is None:
                    buffer.extend(chunk)
                else:
                    sink.write(chunk)
        except OSError as exc:
            record_violation(f"qpdf {label} capture failed: {exc}")
        finally:
            stream.close()

    stdout_thread = threading.Thread(
        target=pump,
        args=(process.stdout, stdout_buffer, max_stdout_bytes, "stdout", output_stream),
        daemon=True,
    )
    stderr_thread = threading.Thread(
        target=pump,
        args=(process.stderr, stderr_buffer, max_stderr_bytes, "stderr"),
        daemon=True,
    )
    stdout_thread.start()
    stderr_thread.start()
    timed_out = False
    try:
        process.wait(timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        timed_out = True
        process.kill()
        process.wait()
    stdout_thread.join()
    stderr_thread.join()
    if output_stream is not None:
        output_stream.flush()
        output_stream.close()
    if timed_out:
        raise ProviderError(f"qpdf timed out after {timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    return subprocess.CompletedProcess(command, process.returncode, bytes(stdout_buffer), bytes(stderr_buffer))


def decoded(value: bytes | str | None) -> str:
    if value is None:
        return ""
    if isinstance(value, bytes):
        return value.decode("utf-8", "replace")
    return value


def qpdf_version(executable: Path, timeout_seconds: int) -> str:
    result = run_qpdf(executable, ["--version"], timeout_seconds=timeout_seconds)
    output = (decoded(result.stdout) or decoded(result.stderr)).strip().splitlines()
    if result.returncode != 0 or not output:
        raise ProviderError(f"qpdf --version failed: {decoded(result.stderr).strip()}")
    match = re.search(r"\b(\d+)\.(\d+)(?:\.(\d+))?\b", output[0])
    if not match or int(match.group(1)) < 11:
        raise ProviderError(f"qpdf 11 or newer with JSON v2 is required; received {output[0]!r}")
    return match.group(0)


def input_path(value: str, max_input_bytes: int) -> Path:
    target = Path(value).expanduser().resolve()
    if not target.is_file():
        raise ProviderError(f"input PDF does not exist or is not a regular file: {target}")
    size = target.stat().st_size
    if size < 5 or size > max_input_bytes:
        raise ProviderError(f"input PDF size {size} is outside the 5..{max_input_bytes} byte budget")
    with target.open("rb") as stream:
        if stream.read(5) != b"%PDF-":
            raise ProviderError("input does not begin with a PDF header")
    return target


def expected_hash(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = value.strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise ProviderError("--expected-sha256 must be exactly 64 lowercase or uppercase hexadecimal characters")
    return normalized


def bounded_check_output(result: subprocess.CompletedProcess) -> tuple[list[str], str]:
    stdout = decoded(result.stdout)
    stderr = decoded(result.stderr)
    combined = "\n".join(part for part in (stdout.strip(), stderr.strip()) if part)
    if len(combined.encode("utf-8")) > DEFAULT_MAX_CHECK_BYTES:
        raise ProviderError(f"qpdf diagnostic output exceeded {DEFAULT_MAX_CHECK_BYTES} bytes")
    lines = [line.rstrip() for line in combined.splitlines() if line.strip()]
    if len(lines) > MAX_CHECK_LINES:
        raise ProviderError(f"qpdf diagnostic output exceeded {MAX_CHECK_LINES} lines")
    return lines, combined


def dictionaries(value: Any) -> Iterator[dict[str, Any]]:
    stack = [value]
    visited = 0
    while stack:
        node = stack.pop()
        visited += 1
        if visited > MAX_JSON_NODES:
            raise ProviderError(f"qpdf JSON evidence traversal exceeded {MAX_JSON_NODES} nodes")
        if isinstance(node, dict):
            yield node
            stack.extend(node.values())
        elif isinstance(node, list):
            stack.extend(node)


def bounded_evidence_text(value: Any) -> str | None:
    if value is None:
        return None
    if isinstance(value, str):
        text = value
    elif isinstance(value, (bool, int, float)):
        text = str(value)
    else:
        return f"<{type(value).__name__}>"
    if len(text) <= MAX_EVIDENCE_TEXT_CHARS:
        return text
    return text[:MAX_EVIDENCE_TEXT_CHARS] + "…"


def signature_policy(document: dict[str, Any]) -> dict[str, Any]:
    fields = document.get("acroform", {}).get("fields") or []
    signature_fields = [
        {
            "object": bounded_evidence_text(field.get("object")),
            "name": bounded_evidence_text(field.get("fullname") or field.get("partialname")),
            "hasValue": field.get("value") is not None,
        }
        for field in fields
        if isinstance(field.get("fieldtype"), str) and field["fieldtype"].lstrip("/") == "Sig"
    ]
    qpdf = document.get("qpdf")
    objects = qpdf[1] if isinstance(qpdf, list) and len(qpdf) > 1 and isinstance(qpdf[1], dict) else {}
    evidence: list[dict[str, Any]] = []
    found: set[str] = set()
    for object_id, record in objects.items():
        kinds: set[str] = set()
        for dictionary in dictionaries(record):
            for key, value in dictionary.items():
                if key == "/ByteRange":
                    kinds.add("ByteRange")
                if key == "/Perms":
                    kinds.add("Perms")
                if key == "/DocMDP" or key == "/TransformMethod" and value == "/DocMDP":
                    kinds.add("DocMDP")
                if key == "/FieldMDP" or key == "/TransformMethod" and value == "/FieldMDP":
                    kinds.add("FieldMDP")
                if key in {"/FT", "/Type"} and value == "/Sig":
                    kinds.add("SignatureDictionary")
        if kinds:
            found.update(kinds)
            if len(evidence) < 100:
                evidence.append({"object": bounded_evidence_text(object_id), "kinds": sorted(kinds)})
    if signature_fields:
        found.add("SignatureField")
    return {
        "hasSignatureEvidence": bool(found),
        "hasSignatureFields": bool(signature_fields or "SignatureDictionary" in found),
        "hasByteRange": "ByteRange" in found,
        "hasPerms": "Perms" in found,
        "hasDocMDP": "DocMDP" in found,
        "hasFieldMDP": "FieldMDP" in found,
        "signatureFields": signature_fields[:100],
        "evidenceObjects": evidence,
        "trust": "unknown",
        "requiredProvider": "pyHanko",
    }


def structure_summary(document: dict[str, Any], check_text: str) -> dict[str, Any]:
    qpdf = document.get("qpdf")
    objects = qpdf[1] if isinstance(qpdf, list) and len(qpdf) > 1 and isinstance(qpdf[1], dict) else {}
    pages = document.get("pages") or []
    fields = document.get("acroform", {}).get("fields") or []
    attachments = document.get("attachments") or {}
    outlines = document.get("outlines") or []
    encryption = document.get("encrypt") or {}
    has_struct_tree_root = False
    mark_info_marked = False
    annotation_references: set[str] = set()

    def indirect_reference(value: Any) -> str | None:
        if not isinstance(value, str):
            return None
        match = re.fullmatch(r"(?:obj:)?(\d+\s+\d+\s+R)", value)
        return match.group(1) if match else None

    for record in objects.values():
        for dictionary in dictionaries(record):
            annotations = dictionary.get("/Annots")
            if isinstance(annotations, list):
                annotation_references.update(
                    reference for value in annotations if (reference := indirect_reference(value)) is not None
                )

    annotation_objects = set(annotation_references)
    for object_id, record in objects.items():
        record_is_annotation = False
        for dictionary in dictionaries(record):
            if "/StructTreeRoot" in dictionary:
                has_struct_tree_root = True
            if dictionary.get("/Marked") is True:
                mark_info_marked = True
            if dictionary.get("/Type") == "/Annot":
                record_is_annotation = True
        if record_is_annotation:
            reference = indirect_reference(object_id)
            if reference is not None:
                annotation_objects.add(reference)
    version = re.search(r"PDF Version:\s*([^\s]+)", check_text)
    if "File is not linearized" in check_text:
        linearized = False
    elif "File is linearized" in check_text:
        linearized = True
    else:
        linearized = None
    return {
        "pdfVersion": version.group(1) if version else None,
        "pageCount": len(pages),
        "objectCount": len(objects),
        "formFieldCount": len(fields),
        "attachmentCount": len(attachments),
        "outlineCount": len(outlines),
        "annotationCount": len(annotation_objects),
        "tagged": has_struct_tree_root or mark_info_marked,
        "hasStructTreeRoot": has_struct_tree_root,
        "markInfoMarked": mark_info_marked,
        "encrypted": bool(encryption.get("encrypted")),
        "encryption": {
            "method": encryption.get("parameters", {}).get("method"),
            "bits": encryption.get("parameters", {}).get("bits"),
            "ownerPasswordMatched": bool(encryption.get("ownerpasswordmatched")),
            "userPasswordMatched": bool(encryption.get("userpasswordmatched")),
        },
        "linearized": linearized,
    }


def validate_json_document(document: Any) -> dict[str, Any]:
    if not isinstance(document, dict):
        raise ProviderError("qpdf JSON root must be an object")
    qpdf = document.get("qpdf")
    if not isinstance(qpdf, list) or len(qpdf) < 2 or not isinstance(qpdf[1], dict):
        raise ProviderError("qpdf JSON v2 object table is missing or malformed")
    expected_types = {
        "pages": list,
        "outlines": list,
        "acroform": dict,
        "attachments": dict,
        "encrypt": dict,
    }
    for key, expected_type in expected_types.items():
        if not isinstance(document.get(key), expected_type):
            raise ProviderError(f"qpdf JSON key {key!r} must be a {expected_type.__name__}")
    fields = document["acroform"].get("fields")
    if not isinstance(fields, list) or any(not isinstance(field, dict) for field in fields):
        raise ProviderError("qpdf JSON AcroForm fields must be an array of objects")
    parameters = document["encrypt"].get("parameters")
    if not isinstance(parameters, dict):
        raise ProviderError("qpdf JSON encryption parameters must be an object")
    return document


def inspect_pdf(
    target: Path,
    *,
    executable: Path,
    version: str,
    expected_sha256: str | None,
    max_json_bytes: int,
    timeout_seconds: int,
    display_path: Path | None = None,
) -> dict[str, Any]:
    before_hash = sha256(target)
    if expected_sha256 and before_hash != expected_sha256:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected_sha256}, received {before_hash}")

    check = run_qpdf(executable, ["--check", str(target)], timeout_seconds=timeout_seconds)
    check_lines, check_text = bounded_check_output(check)
    if display_path is not None:
        check_lines = [line.replace(str(target), str(display_path)) for line in check_lines]
    if check.returncode not in {0, 3}:
        raise ProviderError(f"qpdf --check could not process the PDF (exit {check.returncode}): {check_text[-2000:]}")

    json_path: Path | None = None
    try:
        with tempfile.NamedTemporaryFile(prefix="open-office-qpdf-", suffix=".json", delete=False) as stream:
            json_path = Path(stream.name)
        json_result = run_qpdf(executable, [
            "--json=2",
            "--json-key=pages",
            "--json-key=outlines",
            "--json-key=acroform",
            "--json-key=attachments",
            "--json-key=encrypt",
            "--json-key=qpdf",
            "--json-stream-data=none",
            str(target),
        ], timeout_seconds=timeout_seconds, stdout_path=json_path, max_stdout_bytes=max_json_bytes)
        json_lines, json_text = bounded_check_output(json_result)
        if display_path is not None:
            json_lines = [line.replace(str(target), str(display_path)) for line in json_lines]
        if json_result.returncode not in {0, 3}:
            raise ProviderError(f"qpdf JSON inspection failed (exit {json_result.returncode}): {json_text[-2000:]}")
        json_size = json_path.stat().st_size
        if json_size > max_json_bytes:
            raise ProviderError(f"qpdf JSON inspection exceeded the {max_json_bytes} byte budget")
        with json_path.open("r", encoding="utf-8") as stream:
            document = validate_json_document(json.load(stream))
    except (json.JSONDecodeError, UnicodeError, RecursionError) as exc:
        raise ProviderError(f"qpdf returned invalid JSON: {exc}") from exc
    finally:
        if json_path:
            json_path.unlink(missing_ok=True)

    after_hash = sha256(target)
    if before_hash != after_hash:
        raise ProviderError("source PDF changed during read-only qpdf inspection")
    structure = structure_summary(document, check_text)
    return {
        "schema": SCHEMA_INSPECT,
        "ok": True,
        "provider": {"name": "qpdf", "version": version, "executable": str(executable)},
        "silentFallback": False,
        "savePolicy": "read-only",
        "source": {"path": str(target), "bytes": target.stat().st_size, "sha256": before_hash},
        "check": {
            "status": "clean" if check.returncode == 0 else "warnings",
            "exitCode": check.returncode,
            "lines": check_lines,
            "jsonExitCode": json_result.returncode,
            "jsonLines": json_lines,
        },
        "structure": structure,
        "signaturePolicy": signature_policy(document),
    }


def publish_new_file(candidate: Path, destination: Path) -> None:
    try:
        os.link(candidate, destination)
    except FileExistsError as exc:
        raise ProviderError(f"output appeared during the transaction and was not replaced: {destination}") from exc
    except OSError as exc:
        raise ProviderError(f"could not atomically publish output: {exc}") from exc


def output_path(value: str) -> Path:
    requested = Path(value).expanduser()
    if not requested.is_absolute():
        requested = Path.cwd() / requested
    if requested.is_symlink():
        raise ProviderError(f"output path is a symbolic link and will not be followed: {requested}")
    try:
        parent = requested.parent.resolve(strict=True)
    except OSError as exc:
        raise ProviderError(f"output parent directory does not exist or cannot be resolved: {requested.parent}") from exc
    if not parent.is_dir():
        raise ProviderError(f"output parent is not a directory: {parent}")
    destination = parent / requested.name
    if destination.exists() or destination.is_symlink():
        raise ProviderError(f"output already exists and will not be replaced: {destination}")
    return destination


def rewrite_pdf(args: argparse.Namespace) -> dict[str, Any]:
    executable = qpdf_path()
    version = qpdf_version(executable, args.timeout_seconds)
    source = input_path(args.input, args.max_input_bytes)
    destination = output_path(args.output)
    if source == destination:
        raise ProviderError("input and output must be distinct; qpdf never overwrites the source")
    expected = expected_hash(args.expected_sha256)
    if expected is None:
        raise ProviderError("rewrite requires --expected-sha256 from a fresh inspect result")

    before = inspect_pdf(
        source,
        executable=executable,
        version=version,
        expected_sha256=expected,
        max_json_bytes=args.max_json_bytes,
        timeout_seconds=args.timeout_seconds,
    )
    if before["structure"]["encrypted"]:
        raise ProviderError("encrypted input is unsupported by this bounded qpdf rewrite; select an explicit password/encryption workflow")
    signature = before["signaturePolicy"]
    if signature["hasSignatureEvidence"] and not args.invalidate_signatures:
        raise ProviderError("signed or signature-constrained input requires --invalidate-signatures after pyHanko/DocMDP review")

    with tempfile.TemporaryDirectory(prefix=f".{destination.name}.qpdf-", dir=destination.parent) as temporary:
        temporary_root = Path(temporary)
        snapshot = temporary_root / "source.pdf"
        candidate = temporary_root / "candidate.pdf"
        shutil.copyfile(source, snapshot)
        with snapshot.open("rb+") as stream:
            os.fsync(stream.fileno())
        if sha256(snapshot) != expected:
            raise ProviderError("transaction snapshot does not match the inspected source SHA-256")
        arguments = [str(snapshot), str(candidate)]
        if args.mode == "linearize":
            arguments.insert(0, "--linearize")
        write = run_qpdf(executable, arguments, timeout_seconds=args.timeout_seconds)
        write_lines, write_text = bounded_check_output(write)
        write_text = write_text.replace(str(snapshot), str(source)).replace(str(candidate), str(destination))
        write_lines = [
            line.replace(str(snapshot), str(source)).replace(str(candidate), str(destination))
            for line in write_lines
        ]
        if write.returncode not in {0, 3} or not candidate.is_file():
            raise ProviderError(f"qpdf rewrite failed (exit {write.returncode}): {write_text[-2000:]}")
        with candidate.open("rb+") as stream:
            os.fsync(stream.fileno())
        after = inspect_pdf(
            candidate,
            executable=executable,
            version=version,
            expected_sha256=None,
            max_json_bytes=args.max_json_bytes,
            timeout_seconds=args.timeout_seconds,
            display_path=destination,
        )
        if after["check"]["status"] != "clean":
            raise ProviderError("qpdf rewrite output still contains structural warnings")
        for key in ("pageCount", "attachmentCount", "formFieldCount", "annotationCount", "outlineCount", "tagged"):
            if before["structure"][key] != after["structure"][key]:
                raise ProviderError(f"qpdf rewrite changed {key}: {before['structure'][key]} -> {after['structure'][key]}")
        if args.mode == "linearize" and after["structure"]["linearized"] is not True:
            raise ProviderError("qpdf linearize mode did not produce a validated linearized PDF")
        if sha256(source) != expected:
            raise ProviderError("source PDF changed during qpdf rewrite")
        output_hash = sha256(candidate)
        output_bytes = candidate.stat().st_size
        publish_new_file(candidate, destination)

    return {
        "schema": SCHEMA_REWRITE,
        "ok": True,
        "provider": {"name": "qpdf", "version": version, "executable": str(executable)},
        "silentFallback": False,
        "operation": "qpdf-rewrite",
        "mode": args.mode,
        "savePolicy": "rewrite",
        "source": before["source"],
        "output": {"path": str(destination), "bytes": output_bytes, "sha256": output_hash},
        "sourceProtected": sha256(source) == expected,
        "transaction": {"sourceSnapshot": True, "atomicDistinctOutput": True, "outputReplaced": False},
        "signaturePolicyBefore": signature,
        "signaturePolicyAfter": after["signaturePolicy"],
        "signatureInvalidated": bool(signature["hasSignatureEvidence"] and args.invalidate_signatures),
        "checkBefore": before["check"],
        "checkAfter": after["check"],
        "structureBefore": before["structure"],
        "structureAfter": after["structure"],
        "qpdfWrite": {"exitCode": write.returncode, "lines": write_lines},
        "requiredNextGates": ["fresh qpdf inspect", "pdfinfo", "Poppler render every page", "manual visual review"],
    }


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def add_limits(parser: argparse.ArgumentParser) -> None:
    parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    parser.add_argument("--max-json-bytes", type=positive_int, default=DEFAULT_MAX_JSON_BYTES)
    parser.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    probe = subparsers.add_parser("probe", help="require qpdf 11+ and report the exact executable")
    probe.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)

    inspect = subparsers.add_parser("inspect", help="produce bounded structure and signature-policy evidence")
    inspect.add_argument("input")
    inspect.add_argument("--expected-sha256")
    add_limits(inspect)

    rewrite = subparsers.add_parser("rewrite", help="repair or linearize into a new transactionally published PDF")
    rewrite.add_argument("input")
    rewrite.add_argument("output")
    rewrite.add_argument("--mode", choices=["repair", "linearize"], required=True)
    rewrite.add_argument("--expected-sha256", required=True)
    rewrite.add_argument("--invalidate-signatures", action="store_true")
    add_limits(rewrite)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = build_parser().parse_args()
    try:
        if args.command == "probe":
            executable = qpdf_path()
            result = {
                "ok": True,
                "provider": "qpdf",
                "providerVersion": qpdf_version(executable, args.timeout_seconds),
                "executable": str(executable),
                "jsonVersion": 2,
                "integration": "shipped-thin-script-external-cli",
                "silentFallback": False,
            }
        elif args.command == "inspect":
            executable = qpdf_path()
            result = inspect_pdf(
                input_path(args.input, args.max_input_bytes),
                executable=executable,
                version=qpdf_version(executable, args.timeout_seconds),
                expected_sha256=expected_hash(args.expected_sha256),
                max_json_bytes=args.max_json_bytes,
                timeout_seconds=args.timeout_seconds,
            )
        else:
            result = rewrite_pdf(args)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (ProviderError, OSError) as exc:
        print(json.dumps({
            "ok": False,
            "provider": "qpdf",
            "error": str(exc),
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
