#!/usr/bin/env python3
"""Inspect and remove bounded active or auxiliary PDF structures with pikepdf."""

from __future__ import annotations

import argparse
import errno
import hashlib
from importlib import metadata
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from typing import Any


SCHEMA_INSPECT = "open-office-artifact-tool.pikepdf-inspect.v1"
SCHEMA_CLEAN = "open-office-artifact-tool.pikepdf-structure-clean.v1"
SUPPORTED_MIN = (10, 10, 0)
SUPPORTED_MAX_EXCLUSIVE = (10, 11, 0)
DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_OUTPUT_BYTES = 1024 * 1024 * 1024
DEFAULT_MAX_PAGES = 10_000
DEFAULT_MAX_OBJECTS = 2_000_000
DEFAULT_TIMEOUT_SECONDS = 120
DEFAULT_MAX_STDOUT_BYTES = 4 * 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 512 * 1024
MAX_WORKER_CONFIG_BYTES = 128 * 1024
MAX_WARNING_COUNT = 100
MAX_WARNING_CHARS = 2_048

PROFILES = {
    "active-content": (
        "remove_javascript",
        "remove_external_access",
        "remove_multimedia",
    ),
    "active-and-auxiliary": (
        "remove_javascript",
        "remove_external_access",
        "remove_multimedia",
        "remove_attachments",
        "remove_thumbnails",
        "remove_search_index",
        "remove_web_capture",
        "remove_private_app_data",
        "remove_collection",
    ),
}

ACTION_CATEGORIES = {
    "/JavaScript": "javascriptActions",
    "/URI": "externalAccessActions",
    "/Launch": "externalAccessActions",
    "/GoToR": "externalAccessActions",
    "/GoToE": "externalAccessActions",
    "/SubmitForm": "externalAccessActions",
    "/ImportData": "externalAccessActions",
    "/Rendition": "multimediaActions",
    "/Movie": "multimediaActions",
    "/Sound": "multimediaActions",
    "/RichMediaExecute": "multimediaActions",
}

KEY_CATEGORIES = {
    "/JavaScript": "javascriptNameTrees",
    "/EmbeddedFiles": "embeddedFileNameTrees",
    "/AF": "associatedFileReferences",
    "/EF": "embeddedFileSpecifications",
    "/Thumb": "thumbnails",
    "/Renditions": "renditionNameTrees",
    "/RichMediaContent": "multimediaPayloadReferences",
    "/RichMediaSettings": "multimediaPayloadReferences",
    "/3DD": "multimediaPayloadReferences",
    "/Movie": "multimediaPayloadReferences",
    "/Sound": "multimediaPayloadReferences",
    "/SpiderInfo": "webCaptureDictionaries",
    "/PieceInfo": "privateApplicationData",
    "/SearchIndex": "searchIndexes",
    "/Collection": "collectionDictionaries",
}

PROFILE_ZERO_CATEGORIES = {
    "active-content": {
        "javascriptActions",
        "javascriptNameTrees",
        "externalAccessActions",
        "multimediaActions",
        "renditionNameTrees",
        "multimediaPayloadReferences",
    },
    "active-and-auxiliary": set(KEY_CATEGORIES.values())
    | {
        "javascriptActions",
        "externalAccessActions",
        "multimediaActions",
        "attachmentFileSpecifications",
    },
}

STABLE_STRUCTURE_KEYS = (
    "pageCount",
    "annotationCount",
    "formFieldCount",
    "hasXfa",
    "hasMetadata",
    "hasStructTreeRoot",
    "hasOutlines",
)


class ProviderError(RuntimeError):
    pass


def bounded_text(value: Any, limit: int = MAX_WARNING_CHARS) -> str:
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parsed_version(value: str) -> tuple[int, int, int]:
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", value)
    if not match:
        raise ProviderError(f"cannot parse pikepdf version {value!r}")
    return tuple(int(part) for part in match.groups())


def provider_version() -> str:
    try:
        value = metadata.version("pikepdf")
    except metadata.PackageNotFoundError as exc:
        raise ProviderError(
            f"pikepdf is unavailable in {sys.executable}; install pikepdf >=10.10.0,<10.11.0 in the explicitly selected PDF provider environment"
        ) from exc
    parsed = parsed_version(value)
    if not (SUPPORTED_MIN <= parsed < SUPPORTED_MAX_EXCLUSIVE):
        raise ProviderError(
            "pikepdf >=10.10.0,<10.11.0 is required by this validated adapter; "
            f"received {value!r}"
        )
    return value


def require_pikepdf() -> tuple[Any, str]:
    version = provider_version()
    try:
        import pikepdf
        import pikepdf.sanitize as sanitize
    except Exception as exc:
        raise ProviderError(f"pikepdf structure-clean APIs could not be imported: {bounded_text(exc)}") from exc
    for name in sorted({operation for operations in PROFILES.values() for operation in operations}):
        if not callable(getattr(sanitize.Sanitizer, name, None)):
            raise ProviderError(f"pikepdf {version} is missing required Sanitizer.{name}()")
    return pikepdf, version


def probe() -> dict[str, Any]:
    _, version = require_pikepdf()
    return {
        "ok": True,
        "provider": "pikepdf",
        "providerVersion": version,
        "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        "integration": "shipped-thin-script-external-python",
        "operation": "structure-clean",
        "profiles": {name: list(operations) for name, operations in PROFILES.items()},
        "savePolicies": ["read-only", "rewrite"],
        "incrementalSave": False,
        "passwordInput": False,
        "arbitraryProviderFlags": False,
        "silentFallback": False,
        "providerIsRedactor": False,
        "providerIsCompleteSanitizer": False,
        "providerIsSandbox": False,
    }


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def expected_hash(value: str) -> str:
    normalized = value.strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise ProviderError("--expected-sha256 must be exactly 64 hexadecimal characters")
    return normalized


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


def destination_path(value: str, source: Path) -> Path:
    target = Path(value).expanduser()
    if not target.is_absolute():
        target = Path.cwd() / target
    target = Path(os.path.abspath(target))
    if os.path.lexists(target):
        if target.is_symlink():
            raise ProviderError(f"output path is a symbolic link and will not be followed: {target}")
        raise ProviderError(f"output path already exists and will not be replaced: {target}")
    if not target.parent.is_dir():
        raise ProviderError(f"output parent directory does not exist: {target.parent}")
    if target.resolve(strict=False) == source.resolve():
        raise ProviderError("input and output must be distinct; the source PDF is never overwritten")
    return target


def validate_limits(args: argparse.Namespace) -> None:
    for label, value, maximum in (
        ("--max-input-bytes", args.max_input_bytes, DEFAULT_MAX_INPUT_BYTES),
        ("--max-output-bytes", getattr(args, "max_output_bytes", DEFAULT_MAX_OUTPUT_BYTES), DEFAULT_MAX_OUTPUT_BYTES),
        ("--max-pages", args.max_pages, DEFAULT_MAX_PAGES),
        ("--max-objects", args.max_objects, DEFAULT_MAX_OBJECTS),
        ("--timeout-seconds", args.timeout_seconds, DEFAULT_TIMEOUT_SECONDS),
        ("--max-stdout-bytes", args.max_stdout_bytes, DEFAULT_MAX_STDOUT_BYTES),
        ("--max-stderr-bytes", args.max_stderr_bytes, DEFAULT_MAX_STDERR_BYTES),
    ):
        if value > maximum:
            raise ProviderError(f"{label} cannot exceed the hard maximum {maximum}")


def input_trust(args: argparse.Namespace) -> str:
    if getattr(args, "trusted_input", False):
        return "trusted-input"
    if getattr(args, "caller_isolated", False):
        return "caller-isolated"
    raise ProviderError("select exactly one of --trusted-input or --caller-isolated")


def object_type(value: Any) -> Any:
    return getattr(value, "_type_code", None)


def object_ref(value: Any) -> str:
    objgen = getattr(value, "objgen", (0, 0))
    if objgen != (0, 0):
        return f"{objgen[0]} {objgen[1]} R"
    return "direct"


def scan_pdf(pikepdf: Any, pdf: Any, max_pages: int, max_objects: int) -> dict[str, Any]:
    page_count = len(pdf.pages)
    object_count = len(pdf.objects)
    if page_count > max_pages:
        raise ProviderError(f"PDF page count {page_count} exceeds the {max_pages} page budget")
    if object_count > max_objects:
        raise ProviderError(f"PDF object count {object_count} exceeds the {max_objects} object budget")

    counts = {category: 0 for category in set(KEY_CATEGORIES.values()) | set(ACTION_CATEGORIES.values())}
    counts["attachmentFileSpecifications"] = 0
    signature = {
        "signatureFields": 0,
        "signatureObjects": 0,
        "byteRanges": 0,
        "docMDPReferences": 0,
        "fieldMDPReferences": 0,
        "hasPerms": False,
    }
    annotation_count = 0
    form_field_count = 0
    has_xfa = False
    has_metadata = "/Metadata" in pdf.Root or "/Info" in pdf.trailer
    has_struct_tree_root = "/StructTreeRoot" in pdf.Root
    has_outlines = "/Outlines" in pdf.Root
    evidence: dict[str, list[str]] = {category: [] for category in counts}

    def hit(category: str, owner: Any, key: str) -> None:
        counts[category] += 1
        if len(evidence[category]) < 32:
            evidence[category].append(f"{object_ref(owner)}:{key}")

    def visit(value: Any, owner: Any, depth: int) -> None:
        nonlocal annotation_count, form_field_count, has_xfa
        if depth > 64:
            raise ProviderError("PDF contains a direct-object nesting depth greater than 64")
        type_code = object_type(value)
        if type_code in {pikepdf.ObjectType.dictionary, pikepdf.ObjectType.stream}:
            current_owner = value if getattr(value, "objgen", (0, 0)) != (0, 0) else owner
            try:
                items = list(value.items())
            except Exception as exc:
                raise ProviderError(f"cannot inspect PDF dictionary {object_ref(current_owner)}: {bounded_text(exc)}") from exc
            subtype = str(value.get("/Subtype", ""))
            object_pdf_type = str(value.get("/Type", ""))
            field_type = str(value.get("/FT", ""))
            if object_pdf_type == "/Annot" or subtype in {
                "/Text", "/Link", "/FreeText", "/Line", "/Square", "/Circle", "/Polygon", "/PolyLine",
                "/Highlight", "/Underline", "/Squiggly", "/StrikeOut", "/Stamp", "/Caret", "/Ink",
                "/Popup", "/FileAttachment", "/Sound", "/Movie", "/Widget", "/Screen", "/PrinterMark",
                "/TrapNet", "/Watermark", "/3D", "/Redact", "/RichMedia",
            }:
                annotation_count += 1
            if field_type:
                form_field_count += 1
            if field_type == "/Sig":
                signature["signatureFields"] += 1
            if object_pdf_type == "/Sig":
                signature["signatureObjects"] += 1
            for raw_key, child in items:
                key = str(raw_key)
                category = KEY_CATEGORIES.get(key)
                if category:
                    hit(category, current_owner, key)
                if key == "/FS" and subtype == "/FileAttachment":
                    hit("attachmentFileSpecifications", current_owner, key)
                if key == "/XFA":
                    has_xfa = True
                if key == "/ByteRange":
                    signature["byteRanges"] += 1
                elif key == "/Perms":
                    signature["hasPerms"] = True
                elif key == "/DocMDP" or str(child) == "/DocMDP":
                    signature["docMDPReferences"] += 1
                elif key == "/FieldMDP" or str(child) == "/FieldMDP":
                    signature["fieldMDPReferences"] += 1
                if key == "/S":
                    action_category = ACTION_CATEGORIES.get(str(child))
                    if action_category:
                        hit(action_category, current_owner, f"/S={str(child)}")
                child_objgen = getattr(child, "objgen", (0, 0))
                if child_objgen == (0, 0):
                    visit(child, current_owner, depth + 1)
        elif type_code == pikepdf.ObjectType.array:
            for child in value:
                if getattr(child, "objgen", (0, 0)) == (0, 0):
                    visit(child, owner, depth + 1)

    for indirect in pdf.objects:
        visit(indirect, indirect, 0)

    try:
        attachment_count = len(pdf.attachments)
    except Exception as exc:
        raise ProviderError(f"cannot inspect PDF attachment name tree: {bounded_text(exc)}") from exc
    signature["hasEvidence"] = bool(
        signature["signatureFields"]
        or signature["signatureObjects"]
        or signature["byteRanges"]
        or signature["docMDPReferences"]
        or signature["fieldMDPReferences"]
        or signature["hasPerms"]
    )
    return {
        "pageCount": page_count,
        "objectCount": object_count,
        "annotationCount": annotation_count,
        "formFieldCount": form_field_count,
        "attachmentCount": attachment_count,
        "hasXfa": has_xfa,
        "hasMetadata": has_metadata,
        "hasStructTreeRoot": has_struct_tree_root,
        "hasOutlines": has_outlines,
        "featureCounts": counts,
        "featureEvidence": {key: values for key, values in evidence.items() if values},
        "signatureEvidence": signature,
    }


def bounded_warnings(pdf: Any) -> list[str]:
    warnings = [bounded_text(value) for value in pdf.get_warnings()]
    if len(warnings) > MAX_WARNING_COUNT:
        raise ProviderError(f"pikepdf emitted more than {MAX_WARNING_COUNT} parser warnings")
    return warnings


def worker_main() -> int:
    try:
        raw_config = sys.stdin.buffer.read(MAX_WORKER_CONFIG_BYTES + 1)
        if len(raw_config) > MAX_WORKER_CONFIG_BYTES:
            raise ProviderError(f"worker configuration exceeds {MAX_WORKER_CONFIG_BYTES} bytes")
        config = json.loads(raw_config)
        if not isinstance(config, dict):
            raise ProviderError("worker configuration must be a JSON object")
        pikepdf, version = require_pikepdf()
        input_file = Path(config["input"])
        operation = config["operation"]
        with pikepdf.open(
            input_file,
            suppress_warnings=True,
            attempt_recovery=False,
            inherit_page_attributes=False,
            allow_overwriting_input=False,
        ) as pdf:
            if pdf.is_encrypted:
                raise ProviderError("encrypted PDFs are unsupported; this adapter accepts no password or decryption policy")
            before = scan_pdf(pikepdf, pdf, config["maxPages"], config["maxObjects"])
            warnings_before = bounded_warnings(pdf)
            if warnings_before:
                raise ProviderError(f"pikepdf reported parser warnings; route damaged files to qpdf repair first: {warnings_before[0]}")
            if operation == "inspect":
                print(json.dumps({
                    "workerSchema": 1,
                    "providerVersion": version,
                    "structure": before,
                    "warnings": warnings_before,
                }, separators=(",", ":"), sort_keys=True))
                return 0
            if operation != "clean":
                raise ProviderError(f"unsupported worker operation: {operation!r}")
            profile = config["profile"]
            if profile not in PROFILES:
                raise ProviderError(f"unsupported structure-clean profile: {profile!r}")
            sanitizer = pikepdf.sanitize.Sanitizer()
            for method in PROFILES[profile]:
                getattr(sanitizer, method)()
            sanitizer.apply(pdf)
            output_file = Path(config["output"])
            pdf.save(
                output_file,
                preserve_pdfa=True,
                fix_metadata_version=False,
                object_stream_mode=pikepdf.ObjectStreamMode.preserve,
                deterministic_id=True,
            )

        with pikepdf.open(
            output_file,
            suppress_warnings=True,
            attempt_recovery=False,
            inherit_page_attributes=False,
            allow_overwriting_input=False,
        ) as output_pdf:
            if output_pdf.is_encrypted:
                raise ProviderError("structure-clean output is unexpectedly encrypted")
            after = scan_pdf(pikepdf, output_pdf, config["maxPages"], config["maxObjects"])
            warnings_after = bounded_warnings(output_pdf)
            if warnings_after:
                raise ProviderError(f"pikepdf reported output parser warnings: {warnings_after[0]}")
        remaining = {
            category: after["featureCounts"].get(category, 0)
            for category in sorted(PROFILE_ZERO_CATEGORIES[profile])
            if after["featureCounts"].get(category, 0)
        }
        if after["attachmentCount"] and "remove_attachments" in PROFILES[profile]:
            remaining["attachmentCount"] = after["attachmentCount"]
        if remaining:
            raise ProviderError(f"structure-clean postcondition failed; forbidden features remain: {remaining}")
        changed_topology = {
            key: {"before": before[key], "after": after[key]}
            for key in STABLE_STRUCTURE_KEYS
            if before[key] != after[key]
        }
        if changed_topology:
            raise ProviderError(f"structure-clean changed protected document topology: {changed_topology}")
        print(json.dumps({
            "workerSchema": 1,
            "providerVersion": version,
            "profile": profile,
            "appliedOperations": list(PROFILES[profile]),
            "structureBefore": before,
            "structureAfter": after,
            "warningsBefore": warnings_before,
            "warningsAfter": warnings_after,
            "stableTopology": {key: after[key] for key in STABLE_STRUCTURE_KEYS},
        }, separators=(",", ":"), sort_keys=True))
        return 0
    except Exception as exc:
        print(bounded_text(f"{type(exc).__name__}: {exc}"), file=sys.stderr)
        return 2


def run_worker(
    config: dict[str, Any], timeout_seconds: int, max_stdout_bytes: int, max_stderr_bytes: int
) -> dict[str, Any]:
    process = subprocess.Popen(
        [sys.executable, str(Path(__file__).resolve()), "_worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        shell=False,
        env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        start_new_session=os.name != "nt",
    )
    stdout = bytearray()
    stderr = bytearray()
    violations: list[str] = []
    lock = threading.Lock()

    def stop_process(message: str) -> None:
        with lock:
            if not violations:
                violations.append(message)
        try:
            if os.name != "nt":
                os.killpg(process.pid, 9)
            else:
                process.kill()
        except OSError:
            pass

    def pump(stream: Any, buffer: bytearray, limit: int, label: str) -> None:
        try:
            while True:
                chunk = stream.read(64 * 1024)
                if not chunk:
                    break
                if len(buffer) + len(chunk) > limit:
                    stop_process(f"pikepdf worker {label} exceeded the {limit} byte budget")
                    break
                buffer.extend(chunk)
        finally:
            stream.close()

    stdout_thread = threading.Thread(target=pump, args=(process.stdout, stdout, max_stdout_bytes, "stdout"), daemon=True)
    stderr_thread = threading.Thread(target=pump, args=(process.stderr, stderr, max_stderr_bytes, "stderr"), daemon=True)
    stdout_thread.start()
    stderr_thread.start()
    assert process.stdin is not None
    process.stdin.write(json.dumps(config, separators=(",", ":")).encode("utf-8"))
    process.stdin.close()
    timed_out = False
    try:
        process.wait(timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        timed_out = True
        stop_process(f"pikepdf operation timed out after {timeout_seconds} seconds")
        process.wait()
    stdout_thread.join()
    stderr_thread.join()
    if timed_out:
        raise ProviderError(f"pikepdf operation timed out after {timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    stderr_text = stderr.decode("utf-8", "replace").strip()
    if process.returncode != 0:
        raise ProviderError(f"pikepdf worker failed (exit {process.returncode}): {bounded_text(stderr_text)}")
    if stderr_text:
        raise ProviderError(f"pikepdf worker emitted unexpected diagnostics: {bounded_text(stderr_text)}")
    try:
        result = json.loads(stdout.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("pikepdf worker did not return one valid UTF-8 JSON document") from exc
    if not isinstance(result, dict) or result.get("workerSchema") != 1:
        raise ProviderError("pikepdf worker returned an unsupported result schema")
    return result


def snapshot_source(source: Path, expected: str, temporary_root: Path) -> Path:
    snapshot = temporary_root / "source.pdf"
    shutil.copyfile(source, snapshot)
    with snapshot.open("rb+") as stream:
        os.fsync(stream.fileno())
    snapshot.chmod(0o400)
    if sha256(snapshot) != expected:
        raise ProviderError("private source snapshot does not match the expected SHA-256")
    return snapshot


def common_worker_config(args: argparse.Namespace, snapshot: Path, operation: str) -> dict[str, Any]:
    return {
        "operation": operation,
        "input": str(snapshot),
        "maxPages": args.max_pages,
        "maxObjects": args.max_objects,
    }


def inspect(args: argparse.Namespace) -> dict[str, Any]:
    validate_limits(args)
    trust = input_trust(args)
    source = input_path(args.input, args.max_input_bytes)
    expected = expected_hash(args.expected_sha256)
    actual = sha256(source)
    if actual != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {actual}")
    _, version = require_pikepdf()
    with tempfile.TemporaryDirectory(prefix="open-office-pikepdf-inspect-") as temporary:
        temporary_root = Path(temporary)
        snapshot = snapshot_source(source, expected, temporary_root)
        worker = run_worker(
            common_worker_config(args, snapshot, "inspect"),
            args.timeout_seconds,
            args.max_stdout_bytes,
            args.max_stderr_bytes,
        )
        if sha256(snapshot) != expected or sha256(source) != expected:
            raise ProviderError("source PDF changed during pikepdf inspection")
    if worker["providerVersion"] != version:
        raise ProviderError("pikepdf controller and worker versions do not match")
    return {
        "schema": SCHEMA_INSPECT,
        "ok": True,
        "operationCompleted": True,
        "provider": {
            "name": "pikepdf",
            "version": version,
            "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        },
        "operation": "inspect-structure-clean-surface",
        "savePolicy": "read-only",
        "silentFallback": False,
        "inputTrust": trust,
        "providerIsSandbox": False,
        "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": expected},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "readOnly": True, "sha256": expected},
        "structure": worker["structure"],
        "warnings": worker["warnings"],
        "profiles": {name: list(operations) for name, operations in PROFILES.items()},
        "limitations": [
            "inspection does not execute JavaScript, attachments, multimedia, XFA, or external actions",
            "this process is not a malware sandbox; attacker-chosen files require caller isolation",
            "feature counts describe the bounded structure-clean surface, not every PDF security property",
        ],
    }


def publish_without_replace(candidate: Path, destination: Path) -> None:
    try:
        os.link(candidate, destination)
    except FileExistsError as exc:
        raise ProviderError(f"output path appeared during the transaction and was not replaced: {destination}") from exc
    except OSError as exc:
        if exc.errno == errno.EXDEV:
            raise ProviderError("transactional output promotion crossed file systems and was refused") from exc
        raise ProviderError(f"cannot atomically publish output without replacement: {bounded_text(exc)}") from exc
    try:
        directory_fd = os.open(destination.parent, os.O_RDONLY)
        try:
            os.fsync(directory_fd)
        finally:
            os.close(directory_fd)
    except OSError:
        pass


def clean(args: argparse.Namespace) -> dict[str, Any]:
    validate_limits(args)
    trust = input_trust(args)
    if not args.invalidate_signatures:
        raise ProviderError("structure-clean is a full rewrite and requires --invalidate-signatures acknowledgement")
    source = input_path(args.input, args.max_input_bytes)
    destination = destination_path(args.output, source)
    expected = expected_hash(args.expected_sha256)
    actual = sha256(source)
    if actual != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {actual}")
    _, version = require_pikepdf()
    with tempfile.TemporaryDirectory(prefix=".open-office-pikepdf-", dir=destination.parent) as temporary:
        temporary_root = Path(temporary)
        temporary_root.chmod(0o700)
        snapshot = snapshot_source(source, expected, temporary_root)
        candidate = temporary_root / "candidate.pdf"
        config = common_worker_config(args, snapshot, "clean")
        config.update({"output": str(candidate), "profile": args.profile})
        worker = run_worker(config, args.timeout_seconds, args.max_stdout_bytes, args.max_stderr_bytes)
        if worker["providerVersion"] != version:
            raise ProviderError("pikepdf controller and worker versions do not match")
        if not candidate.is_file():
            raise ProviderError("pikepdf worker did not produce the transactional output")
        candidate_size = candidate.stat().st_size
        if candidate_size < 5 or candidate_size > args.max_output_bytes:
            raise ProviderError(
                f"output PDF size {candidate_size} is outside the 5..{args.max_output_bytes} byte budget"
            )
        candidate_bytes = candidate.read_bytes()
        if not candidate_bytes.startswith(b"%PDF-"):
            raise ProviderError("structure-clean output does not begin with a PDF header")
        snapshot_bytes = snapshot.read_bytes()
        source_prefix_preserved = candidate_bytes.startswith(snapshot_bytes)
        if source_prefix_preserved:
            raise ProviderError("structure-clean output retains the complete original byte prefix; incremental output is forbidden")
        startxref_count = candidate_bytes.count(b"startxref")
        eof_count = candidate_bytes.count(b"%%EOF")
        if startxref_count != 1 or eof_count != 1:
            raise ProviderError(
                f"structure-clean output must contain one final revision marker; found startxref={startxref_count}, EOF={eof_count}"
            )
        output_hash = hashlib.sha256(candidate_bytes).hexdigest()
        if sha256(snapshot) != expected or sha256(source) != expected:
            raise ProviderError("source PDF changed during pikepdf structure-clean")
        publish_without_replace(candidate, destination)
    if sha256(source) != expected:
        try:
            destination.unlink()
        except OSError:
            pass
        raise ProviderError("source PDF changed before the structure-clean report was finalized")
    if sha256(destination) != output_hash:
        try:
            destination.unlink()
        except OSError:
            pass
        raise ProviderError("published output hash does not match the validated candidate")
    signature_before = worker["structureBefore"]["signatureEvidence"]
    return {
        "schema": SCHEMA_CLEAN,
        "ok": True,
        "operationCompleted": True,
        "provider": {
            "name": "pikepdf",
            "version": version,
            "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        },
        "operation": "structure-clean",
        "profile": args.profile,
        "appliedOperations": worker["appliedOperations"],
        "savePolicy": "rewrite",
        "silentFallback": False,
        "inputTrust": trust,
        "providerIsSandbox": False,
        "providerIsRedactor": False,
        "providerIsCompleteSanitizer": False,
        "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": expected},
        "output": {"path": str(destination), "bytes": destination.stat().st_size, "sha256": output_hash},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "readOnly": True, "sha256": expected},
        "signaturePolicy": {
            "evidenceBefore": signature_before,
            "invalidationAcknowledged": True,
            "signatureInvalidated": bool(signature_before["hasEvidence"]),
            "oldSignatureApprovalClaimed": False,
        },
        "structureBefore": worker["structureBefore"],
        "structureAfter": worker["structureAfter"],
        "validation": {
            "profilePostconditionsPassed": True,
            "stableTopology": worker["stableTopology"],
            "warningsBefore": worker["warningsBefore"],
            "warningsAfter": worker["warningsAfter"],
            "fullRewrite": {
                "sourcePrefixPreserved": False,
                "startxrefCount": startxref_count,
                "eofCount": eof_count,
            },
            "sourceReproved": True,
        },
        "transaction": {
            "atomicDistinctOutput": True,
            "noReplace": True,
            "privateSourceSnapshot": True,
        },
        "notRemoved": [
            "visible page content",
            "DocumentInfo and XMP metadata",
            "AcroForm fields and values",
            "XFA packets",
            "annotations after their selected actions or payload references are defanged",
            "hidden text and OCR layers",
            "signature fields or appearances",
        ],
        "limitations": [
            "structure-clean is not redaction, metadata scrub, form flattening, XFA flattening, or complete sanitize",
            "removing JavaScript can break interactive form validation and removing external actions can disable ordinary links",
            "this process is not a malware sandbox; attacker-chosen files require caller isolation",
            "pikepdf rewrites and coalesces prior revisions; any existing digital signature no longer approves the output bytes",
        ],
        "requiredFollowups": [
            "run qpdf structural inspection on the exact output bytes",
            "render every output page with Poppler and compare it with the source",
            "run the independent residue scanner for the intended delivery policy",
            "run pyHanko when the source contained signature evidence",
            "rerun veraPDF when archival or accessibility conformance matters",
            "bind the exact source and output hashes into the canonical PDF audit",
        ],
    }


def add_limits(parser: argparse.ArgumentParser, include_output: bool = False) -> None:
    parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    if include_output:
        parser.add_argument("--max-output-bytes", type=positive_int, default=DEFAULT_MAX_OUTPUT_BYTES)
    parser.add_argument("--max-pages", type=positive_int, default=DEFAULT_MAX_PAGES)
    parser.add_argument("--max-objects", type=positive_int, default=DEFAULT_MAX_OBJECTS)
    parser.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)
    parser.add_argument("--max-stdout-bytes", type=positive_int, default=DEFAULT_MAX_STDOUT_BYTES)
    parser.add_argument("--max-stderr-bytes", type=positive_int, default=DEFAULT_MAX_STDERR_BYTES)


def add_trust(parser: argparse.ArgumentParser) -> None:
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--trusted-input", action="store_true")
    group.add_argument("--caller-isolated", action="store_true")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("probe", help="require the validated pikepdf 10.10.x API surface")

    inspect_parser = subparsers.add_parser("inspect", help="inspect the bounded structure-clean surface")
    inspect_parser.add_argument("input")
    inspect_parser.add_argument("--expected-sha256", required=True)
    add_trust(inspect_parser)
    add_limits(inspect_parser)

    clean_parser = subparsers.add_parser("clean", help="apply one fixed full-rewrite structure-clean profile")
    clean_parser.add_argument("input")
    clean_parser.add_argument("output")
    clean_parser.add_argument("--profile", choices=PROFILES, required=True)
    clean_parser.add_argument("--expected-sha256", required=True)
    clean_parser.add_argument("--invalidate-signatures", action="store_true")
    add_trust(clean_parser)
    add_limits(clean_parser, include_output=True)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    if len(sys.argv) > 1 and sys.argv[1] == "_worker":
        return worker_main()
    args = build_parser().parse_args()
    try:
        if args.command == "probe":
            result = probe()
        elif args.command == "inspect":
            result = inspect(args)
        else:
            result = clean(args)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except ProviderError as exc:
        print(json.dumps({
            "ok": False,
            "error": str(exc),
            "provider": "pikepdf",
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
