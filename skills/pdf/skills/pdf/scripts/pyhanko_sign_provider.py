#!/usr/bin/env python3
"""Create bounded source-bound PDF signatures with local PKCS#12 credentials."""

from __future__ import annotations

import argparse
import base64
from datetime import datetime, timezone
import errno
import getpass
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import threading
from typing import Any

from pyhanko_provider import provider_versions, sha256


SCHEMA_INSPECT = "open-office-artifact-tool.pyhanko-signing-inspect.v1"
SCHEMA_SIGN = "open-office-artifact-tool.pyhanko-sign.v1"
DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_OUTPUT_BYTES = 1024 * 1024 * 1024
DEFAULT_MAX_CREDENTIAL_BYTES = 16 * 1024 * 1024
DEFAULT_MAX_PAGES = 10_000
DEFAULT_MAX_FIELDS = 10_000
DEFAULT_MAX_SIGNATURES = 64
DEFAULT_TIMEOUT_SECONDS = 120
DEFAULT_MAX_STDOUT_BYTES = 8 * 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 512 * 1024
MAX_WORKER_CONFIG_BYTES = 128 * 1024
MAX_PASSPHRASE_BYTES = 4_096
MAX_METADATA_CHARS = 512
MAX_FIELD_NAME_CHARS = 128

SUBFILTERS = {
    "pades": "PADES",
    "adobe-pkcs7-detached": "ADOBE_PKCS7_DETACHED",
}

DOCMDP_PERMISSIONS = {
    "no-changes": "NO_CHANGES",
    "fill-forms": "FILL_FORMS",
    "annotate": "ANNOTATE",
}


class ProviderError(RuntimeError):
    pass


def bounded_text(value: Any, limit: int = 2_048) -> str:
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def require_signing_runtime() -> dict[str, str]:
    try:
        versions = provider_versions()
    except Exception as exc:
        raise ProviderError(bounded_text(exc)) from exc
    try:
        from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter  # noqa: F401
        from pyhanko.pdf_utils.reader import PdfFileReader  # noqa: F401
        from pyhanko.sign import fields, signers  # noqa: F401
    except Exception as exc:
        raise ProviderError(f"pyHanko signing APIs could not be imported: {bounded_text(exc)}") from exc
    for name in ("SimpleSigner", "PdfSignatureMetadata", "sign_pdf"):
        if not hasattr(signers, name):
            raise ProviderError(f"pyHanko {versions['pyHanko']} is missing signers.{name}")
    for name in ("SigFieldSpec", "SigSeedSubFilter", "MDPPerm", "enumerate_sig_fields"):
        if not hasattr(fields, name):
            raise ProviderError(f"pyHanko {versions['pyHanko']} is missing fields.{name}")
    return versions


def probe() -> dict[str, Any]:
    versions = require_signing_runtime()
    return {
        "ok": True,
        "provider": "pyhanko",
        "providerVersion": versions["pyHanko"],
        "certvalidatorVersion": versions["pyhankoCertvalidator"],
        "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        "integration": "shipped-thin-script-external-python",
        "operation": "sign-local-pkcs12",
        "credentialTypes": ["pkcs12"],
        "fieldModes": ["existing", "create-invisible", "create-visible"],
        "signatureKinds": ["approval", "certification"],
        "subfilters": list(SUBFILTERS),
        "savePolicies": ["read-only", "incremental"],
        "passphraseChannels": ["stdin-pipe-or-hidden-tty-prompt", "none"],
        "networkAllowed": False,
        "timestampAuthoritySupported": False,
        "ltvEmbeddingSupported": False,
        "pkcs11Supported": False,
        "arbitraryProviderFlags": False,
        "silentFallback": False,
        "padesProfileConformanceClaimed": False,
    }


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def nonnegative_int(value: str) -> int:
    parsed = int(value)
    if parsed < 0:
        raise argparse.ArgumentTypeError("must be a nonnegative integer")
    return parsed


def expected_hash(value: str) -> str:
    normalized = value.strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise ProviderError("expected SHA-256 values must contain exactly 64 hexadecimal characters")
    return normalized


def trusted_input(args: argparse.Namespace) -> str:
    if getattr(args, "trusted_input", False):
        return "trusted-input"
    if getattr(args, "caller_isolated", False):
        return "caller-isolated"
    raise ProviderError("select exactly one of --trusted-input or --caller-isolated")


def regular_input_path(value: str, label: str, maximum_bytes: int, *, pdf: bool = False) -> Path:
    raw = Path(value).expanduser()
    absolute = raw if raw.is_absolute() else Path.cwd() / raw
    absolute = Path(os.path.abspath(absolute))
    try:
        metadata = absolute.lstat()
    except FileNotFoundError as exc:
        raise ProviderError(f"{label} does not exist: {absolute}") from exc
    if stat.S_ISLNK(metadata.st_mode):
        raise ProviderError(f"{label} is a symbolic link and will not be followed: {absolute}")
    if not stat.S_ISREG(metadata.st_mode):
        raise ProviderError(f"{label} is not a regular file: {absolute}")
    if metadata.st_size < 1 or metadata.st_size > maximum_bytes:
        raise ProviderError(f"{label} size {metadata.st_size} is outside the 1..{maximum_bytes} byte budget")
    if pdf:
        with absolute.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ProviderError(f"{label} does not begin with a PDF header")
    return absolute


def destination_path(value: str, source: Path) -> Path:
    raw = Path(value).expanduser()
    target = raw if raw.is_absolute() else Path.cwd() / raw
    target = Path(os.path.abspath(target))
    if os.path.lexists(target):
        if target.is_symlink():
            raise ProviderError(f"output path is a symbolic link and will not be followed: {target}")
        raise ProviderError(f"output path already exists and will not be replaced: {target}")
    if not target.parent.is_dir():
        raise ProviderError(f"output parent directory does not exist: {target.parent}")
    if target == source:
        raise ProviderError("input and output must be distinct; the source PDF is never overwritten")
    return target


def field_name(value: str) -> str:
    if not (1 <= len(value) <= MAX_FIELD_NAME_CHARS):
        raise ProviderError(f"--field-name must contain 1..{MAX_FIELD_NAME_CHARS} characters")
    if not re.fullmatch(r"[A-Za-z0-9_.-]+", value):
        raise ProviderError("--field-name may contain only ASCII letters, digits, underscore, hyphen, and hierarchy dots")
    if value.startswith(".") or value.endswith(".") or ".." in value:
        raise ProviderError("--field-name contains an empty hierarchy segment")
    return value


def bounded_metadata(value: str | None, label: str) -> str | None:
    if value is None:
        return None
    if not value.strip():
        raise ProviderError(f"{label} cannot be empty or whitespace")
    if len(value) > MAX_METADATA_CHARS:
        raise ProviderError(f"{label} cannot exceed {MAX_METADATA_CHARS} characters")
    if any(ord(character) < 0x20 and character not in "\t\n\r" for character in value):
        raise ProviderError(f"{label} contains unsupported control characters")
    return value


def parse_box(value: str) -> tuple[int, int, int, int]:
    parts = value.split(",")
    if len(parts) != 4:
        raise argparse.ArgumentTypeError("must be four comma-separated integer PDF coordinates: x1,y1,x2,y2")
    try:
        box = tuple(int(part.strip()) for part in parts)
    except ValueError as exc:
        raise argparse.ArgumentTypeError("box coordinates must be integers") from exc
    if not (box[0] < box[2] and box[1] < box[3]):
        raise argparse.ArgumentTypeError("box must satisfy x1 < x2 and y1 < y2")
    return box  # type: ignore[return-value]


def validate_limits(args: argparse.Namespace) -> None:
    for label, value, maximum in (
        ("--max-input-bytes", args.max_input_bytes, DEFAULT_MAX_INPUT_BYTES),
        ("--max-output-bytes", getattr(args, "max_output_bytes", DEFAULT_MAX_OUTPUT_BYTES), DEFAULT_MAX_OUTPUT_BYTES),
        ("--max-credential-bytes", getattr(args, "max_credential_bytes", DEFAULT_MAX_CREDENTIAL_BYTES), DEFAULT_MAX_CREDENTIAL_BYTES),
        ("--max-pages", args.max_pages, DEFAULT_MAX_PAGES),
        ("--max-fields", args.max_fields, DEFAULT_MAX_FIELDS),
        ("--max-signatures", args.max_signatures, DEFAULT_MAX_SIGNATURES),
        ("--timeout-seconds", args.timeout_seconds, DEFAULT_TIMEOUT_SECONDS),
        ("--max-stdout-bytes", args.max_stdout_bytes, DEFAULT_MAX_STDOUT_BYTES),
        ("--max-stderr-bytes", args.max_stderr_bytes, DEFAULT_MAX_STDERR_BYTES),
    ):
        if value > maximum:
            raise ProviderError(f"{label} cannot exceed the hard maximum {maximum}")


def read_passphrase(args: argparse.Namespace) -> bytearray | None:
    if args.no_passphrase:
        return None
    if sys.stdin.isatty():
        try:
            entered = getpass.getpass("PKCS#12 passphrase: ")
            raw = bytearray(entered.encode("utf-8"))
            entered = ""
        except (EOFError, KeyboardInterrupt) as exc:
            raise ProviderError("PKCS#12 passphrase input was cancelled") from exc
    else:
        raw = bytearray(sys.stdin.buffer.read(MAX_PASSPHRASE_BYTES + 1))
    if len(raw) > MAX_PASSPHRASE_BYTES:
        for index in range(len(raw)):
            raw[index] = 0
        raise ProviderError(f"PKCS#12 passphrase exceeds the {MAX_PASSPHRASE_BYTES} byte budget")
    if raw.endswith(b"\n"):
        del raw[-1:]
        if raw.endswith(b"\r"):
            del raw[-1:]
    if not raw:
        raise ProviderError("empty PKCS#12 passphrase input is refused; use --no-passphrase deliberately")
    return raw


def snapshot_file(source: Path, destination: Path, expected: str, mode: int = 0o400) -> None:
    shutil.copyfile(source, destination)
    with destination.open("rb+") as stream:
        os.fsync(stream.fileno())
    destination.chmod(mode)
    if sha256(destination) != expected:
        raise ProviderError(f"private snapshot does not match expected SHA-256 for {source}")


def has_exact_prefix(source: Path, candidate: Path, chunk_bytes: int = 1024 * 1024) -> bool:
    if candidate.stat().st_size <= source.stat().st_size:
        return False
    with source.open("rb") as source_stream, candidate.open("rb") as candidate_stream:
        while True:
            source_chunk = source_stream.read(chunk_bytes)
            if not source_chunk:
                return True
            if candidate_stream.read(len(source_chunk)) != source_chunk:
                return False


def object_reference(value: Any) -> str:
    reference = getattr(value, "reference", None)
    if reference is None:
        return "direct"
    return f"{reference.idnum} {reference.generation} R"


def number(value: Any) -> float:
    try:
        return float(value)
    except (TypeError, ValueError) as exc:
        raise ProviderError(f"page geometry contains a nonnumeric value: {bounded_text(value)}") from exc


def page_snapshot(handler: Any, page_index: int) -> dict[str, Any]:
    from pyhanko.pdf_utils.rw_common import find_inherited_value_in_tree

    page_ref, _ = handler.find_page_for_modification(page_index)
    page = page_ref.get_object()
    media_raw = find_inherited_value_in_tree(page, "/MediaBox", "/Parent")
    if media_raw is None or len(media_raw) != 4:
        raise ProviderError(f"page {page_index} has no valid inherited MediaBox")
    crop_raw = find_inherited_value_in_tree(page, "/CropBox", "/Parent") or media_raw
    if len(crop_raw) != 4:
        raise ProviderError(f"page {page_index} has no valid inherited CropBox")
    rotation_raw = find_inherited_value_in_tree(page, "/Rotate", "/Parent")
    rotation = int(rotation_raw or 0) % 360
    if rotation not in {0, 90, 180, 270}:
        raise ProviderError(f"page {page_index} has unsupported rotation {rotation}")
    return {
        "pageIndex": page_index,
        "pageReference": object_reference(page_ref),
        "mediaBox": [number(value) for value in media_raw],
        "cropBox": [number(value) for value in crop_raw],
        "rotation": rotation,
    }


def signature_inventory(reader: Any, max_fields: int, max_signatures: int) -> dict[str, Any]:
    from pyhanko.sign.fields import enumerate_sig_fields

    records = []
    for name, value, reference in enumerate_sig_fields(reader):
        if len(records) >= max_fields:
            raise ProviderError(f"PDF contains more than {max_fields} signature fields")
        field = reference.get_object()
        rectangle = field.get("/Rect")
        records.append({
            "name": str(name),
            "filled": bool(value),
            "reference": object_reference(reference),
            "hasWidget": str(field.get("/Subtype", "")) == "/Widget" or rectangle is not None,
            "rectangle": [number(item) for item in rectangle] if rectangle is not None and len(rectangle) == 4 else None,
        })
    signatures = list(reader.embedded_signatures)
    if len(signatures) > max_signatures:
        raise ProviderError(f"PDF contains more than {max_signatures} embedded signatures")
    return {
        "signatureFields": records,
        "fieldCount": len(records),
        "emptyFieldCount": sum(not record["filled"] for record in records),
        "filledFieldCount": sum(record["filled"] for record in records),
        "signatureCount": len(signatures),
        "signatureNames": [str(signature.field_name) for signature in signatures],
        "hasCertificationSignature": any(signature.docmdp_level is not None for signature in signatures),
        "revisionCount": int(reader.total_revisions),
    }


def certificate_identity(certificate: Any) -> dict[str, Any]:
    validity = certificate["tbs_certificate"]["validity"]
    not_before = validity["not_before"].native
    not_after = validity["not_after"].native
    try:
        key_usage = sorted(certificate.key_usage_value.native)
    except (KeyError, ValueError, TypeError):
        key_usage = []
    fingerprint = hashlib.sha256(certificate.dump()).hexdigest()
    return {
        "subject": bounded_text(certificate.subject.human_friendly, 2_048),
        "issuer": bounded_text(certificate.issuer.human_friendly, 2_048),
        "serialNumber": str(certificate.serial_number),
        "sha256Fingerprint": fingerprint,
        "notValidBefore": not_before.isoformat(),
        "notValidAfter": not_after.isoformat(),
        "keyUsage": key_usage,
        "selfIssued": certificate.subject == certificate.issuer,
    }


def validate_credential(signer: Any) -> dict[str, Any]:
    if signer is None or signer.signing_cert is None or signer.signing_key is None:
        raise ProviderError("PKCS#12 credential did not yield one signing certificate and private key")
    identity = certificate_identity(signer.signing_cert)
    now = datetime.now(timezone.utc)
    validity = signer.signing_cert["tbs_certificate"]["validity"]
    not_before = validity["not_before"].native
    not_after = validity["not_after"].native
    if now < not_before or now > not_after:
        raise ProviderError("PKCS#12 signing certificate is not valid at the current system time")
    usages = set(identity["keyUsage"])
    if usages and not usages.intersection({"digital_signature", "non_repudiation", "content_commitment"}):
        raise ProviderError("PKCS#12 signing certificate key usage does not permit PDF signing")
    return identity


def inspect_worker(config: dict[str, Any]) -> dict[str, Any]:
    from pyhanko.pdf_utils.reader import PdfFileReader

    with Path(config["input"]).open("rb") as stream:
        reader = PdfFileReader(stream, strict=True)
        if reader.encrypted:
            raise ProviderError("encrypted PDFs are unsupported; no password or decryption workflow is exposed")
        page_count = int(reader.root["/Pages"]["/Count"])
        if page_count < 1 or page_count > config["maxPages"]:
            raise ProviderError(f"PDF page count {page_count} is outside the 1..{config['maxPages']} page budget")
        inventory = signature_inventory(reader, config["maxFields"], config["maxSignatures"])
        selected_page = None
        if config.get("pageIndex") is not None:
            if config["pageIndex"] >= page_count:
                raise ProviderError(f"page index {config['pageIndex']} is outside the 0..{page_count - 1} range")
            selected_page = page_snapshot(reader, config["pageIndex"])
    return {"pageCount": page_count, "selectedPage": selected_page, **inventory}


def sign_worker(config: dict[str, Any]) -> dict[str, Any]:
    from pyhanko.pdf_utils.incremental_writer import IncrementalPdfFileWriter
    from pyhanko.sign import fields, signers

    encoded_passphrase = config.pop("passphraseBase64", None)
    passphrase = (
        bytearray(base64.b64decode(encoded_passphrase, validate=True))
        if encoded_passphrase is not None else None
    )
    encoded_passphrase = None
    try:
        signer = signers.SimpleSigner.load_pkcs12(
            config["credential"],
            passphrase=bytes(passphrase) if passphrase is not None else None,
            prefer_pss=False,
        )
    finally:
        if passphrase is not None:
            for index in range(len(passphrase)):
                passphrase[index] = 0
    credential_identity = validate_credential(signer)

    with Path(config["input"]).open("rb") as source:
        writer = IncrementalPdfFileWriter(source, strict=True)
        if writer.prev.encrypted:
            raise ProviderError("encrypted PDFs are unsupported; no password or decryption workflow is exposed")
        page_count = int(writer.root["/Pages"]["/Count"])
        if page_count < 1 or page_count > config["maxPages"]:
            raise ProviderError(f"PDF page count {page_count} is outside the 1..{config['maxPages']} page budget")
        before = signature_inventory(writer.prev, config["maxFields"], config["maxSignatures"])
        if before["signatureCount"] != config["expectedSignatureCount"]:
            raise ProviderError(
                f"signature count changed from expected {config['expectedSignatureCount']} to {before['signatureCount']}"
            )
        if before["signatureCount"] and not config["allowExistingSignatures"]:
            raise ProviderError("existing signatures require --allow-existing-signatures and post-sign validation")
        if config["signatureKind"] == "certification" and before["signatureCount"]:
            raise ProviderError("a certification signature must be the first signature in the document")

        matching = [record for record in before["signatureFields"] if record["name"] == config["fieldName"]]
        if config["fieldMode"] == "existing":
            if len(matching) != 1 or matching[0]["filled"]:
                raise ProviderError("existing field mode requires exactly one empty signature field with the requested name")
        elif matching:
            raise ProviderError("field creation requires that the requested signature field name does not already exist")

        new_field_spec = None
        selected_page = None
        if config["fieldMode"] == "create-invisible":
            new_field_spec = fields.SigFieldSpec(sig_field_name=config["fieldName"], on_page=0, box=None)
        elif config["fieldMode"] == "create-visible":
            page_index = config["pageIndex"]
            if page_index >= page_count:
                raise ProviderError(f"page index {page_index} is outside the 0..{page_count - 1} range")
            selected_page = page_snapshot(writer, page_index)
            if selected_page["rotation"] != 0:
                raise ProviderError("new visible signature fields require an unrotated page")
            crop = selected_page["cropBox"]
            box = config["box"]
            if not (crop[0] <= box[0] < box[2] <= crop[2] and crop[1] <= box[1] < box[3] <= crop[3]):
                raise ProviderError("visible signature field box must fit wholly inside the selected page CropBox")
            new_field_spec = fields.SigFieldSpec(
                sig_field_name=config["fieldName"], on_page=page_index, box=tuple(box)
            )

        metadata = signers.PdfSignatureMetadata(
            field_name=config["fieldName"],
            md_algorithm="sha256",
            reason=config.get("reason"),
            location=config.get("location"),
            contact_info=config.get("contactInfo"),
            name=config.get("signerName"),
            certify=config["signatureKind"] == "certification",
            subfilter=getattr(fields.SigSeedSubFilter, SUBFILTERS[config["subfilter"]]),
            embed_validation_info=False,
            use_pades_lta=False,
            docmdp_permissions=getattr(fields.MDPPerm, DOCMDP_PERMISSIONS[config["docMDPPermission"]]),
        )
        with Path(config["output"]).open("wb") as output:
            signers.sign_pdf(
                writer,
                metadata,
                signer=signer,
                new_field_spec=new_field_spec,
                existing_fields_only=config["fieldMode"] == "existing",
                in_place=False,
                output=output,
            )
            output.flush()
            os.fsync(output.fileno())

    with Path(config["output"]).open("rb") as stream:
        from pyhanko.pdf_utils.reader import PdfFileReader

        output_reader = PdfFileReader(stream, strict=True)
        after = signature_inventory(output_reader, config["maxFields"], config["maxSignatures"])
        if int(output_reader.root["/Pages"]["/Count"]) != page_count:
            raise ProviderError("signing changed the PDF page count")
    if after["signatureCount"] != before["signatureCount"] + 1:
        raise ProviderError("signing did not add exactly one embedded signature")
    if after["signatureNames"][-1] != config["fieldName"]:
        raise ProviderError("the new embedded signature does not use the requested field name")
    return {
        "workerSchema": 1,
        "credential": credential_identity,
        "pageCount": page_count,
        "selectedPage": selected_page,
        "before": before,
        "after": after,
    }


def worker_main() -> int:
    logging.disable(logging.CRITICAL)
    try:
        raw_config = bytearray(sys.stdin.buffer.read(MAX_WORKER_CONFIG_BYTES + 1))
        if len(raw_config) > MAX_WORKER_CONFIG_BYTES:
            raise ProviderError(f"worker configuration exceeds {MAX_WORKER_CONFIG_BYTES} bytes")
        try:
            config = json.loads(raw_config)
        finally:
            for index in range(len(raw_config)):
                raw_config[index] = 0
        if not isinstance(config, dict):
            raise ProviderError("worker configuration must be a JSON object")
        require_signing_runtime()
        operation = config.get("operation")
        result = inspect_worker(config) if operation == "inspect" else sign_worker(config) if operation == "sign" else None
        if result is None:
            raise ProviderError(f"unsupported worker operation: {operation!r}")
        if "workerSchema" not in result:
            result["workerSchema"] = 1
        print(json.dumps(result, separators=(",", ":"), sort_keys=True))
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
                    stop_process(f"pyHanko signing worker {label} exceeded the {limit} byte budget")
                    break
                buffer.extend(chunk)
        finally:
            stream.close()

    stdout_thread = threading.Thread(target=pump, args=(process.stdout, stdout, max_stdout_bytes, "stdout"), daemon=True)
    stderr_thread = threading.Thread(target=pump, args=(process.stderr, stderr, max_stderr_bytes, "stderr"), daemon=True)
    stdout_thread.start()
    stderr_thread.start()
    assert process.stdin is not None
    encoded_config = bytearray(json.dumps(config, separators=(",", ":")).encode("utf-8"))
    try:
        process.stdin.write(encoded_config)
        process.stdin.close()
    finally:
        for index in range(len(encoded_config)):
            encoded_config[index] = 0
        config["passphraseBase64"] = None
    timed_out = False
    try:
        process.wait(timeout=timeout_seconds)
    except subprocess.TimeoutExpired:
        timed_out = True
        stop_process(f"pyHanko signing operation timed out after {timeout_seconds} seconds")
        process.wait()
    stdout_thread.join()
    stderr_thread.join()
    if timed_out:
        raise ProviderError(f"pyHanko signing operation timed out after {timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    stderr_text = stderr.decode("utf-8", "replace").strip()
    if process.returncode != 0:
        raise ProviderError(f"pyHanko signing worker failed (exit {process.returncode}): {bounded_text(stderr_text)}")
    if stderr_text:
        raise ProviderError(f"pyHanko signing worker emitted unexpected diagnostics: {bounded_text(stderr_text)}")
    try:
        result = json.loads(stdout.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("pyHanko signing worker did not return one valid UTF-8 JSON document") from exc
    if not isinstance(result, dict) or result.get("workerSchema") != 1:
        raise ProviderError("pyHanko signing worker returned an unsupported result schema")
    return result


def run_validator(
    candidate: Path, output_hash: str, timeout_seconds: int, max_stdout_bytes: int, max_stderr_bytes: int
) -> dict[str, Any]:
    validator = Path(__file__).with_name("pyhanko_provider.py")
    validator_stdout_bytes = min(max_stdout_bytes, 2 * 1024 * 1024)
    validator_stderr_bytes = min(max_stderr_bytes, 512 * 1024)
    command = [
        sys.executable,
        str(validator),
        "verify",
        str(candidate),
        "--expected-sha256",
        output_hash,
        "--revocation-policy",
        "none",
        "--require-signature",
        "--require-all-integrity-valid",
        "--require-docmdp-compliant",
        "--timeout-seconds",
        str(timeout_seconds),
        "--max-stdout-bytes",
        str(validator_stdout_bytes),
        "--max-stderr-bytes",
        str(validator_stderr_bytes),
    ]
    try:
        result = subprocess.run(
            command,
            check=False,
            capture_output=True,
            timeout=timeout_seconds + 5,
            env={**os.environ, "PYTHONDONTWRITEBYTECODE": "1"},
        )
    except subprocess.TimeoutExpired as exc:
        raise ProviderError("post-signature pyHanko validation timed out") from exc
    if len(result.stdout) > validator_stdout_bytes or len(result.stderr) > validator_stderr_bytes:
        raise ProviderError("post-signature pyHanko validation exceeded its output budget")
    payload = result.stdout if result.returncode == 0 else result.stderr
    try:
        report = json.loads(payload.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("post-signature pyHanko validation did not return valid bounded JSON") from exc
    if result.returncode != 0:
        failures = report.get("policyGates", {}).get("failures", []) if isinstance(report, dict) else []
        detail = failures[0] if failures else report.get("error", "validation failed") if isinstance(report, dict) else "validation failed"
        raise ProviderError(f"post-signature validation failed: {bounded_text(detail)}")
    if report.get("schema") != "open-office-artifact-tool.pyhanko-verify.v1":
        raise ProviderError("post-signature validator returned an unsupported report schema")
    return report


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


def common_config(args: argparse.Namespace, source_snapshot: Path, operation: str) -> dict[str, Any]:
    return {
        "operation": operation,
        "input": str(source_snapshot),
        "maxPages": args.max_pages,
        "maxFields": args.max_fields,
        "maxSignatures": args.max_signatures,
    }


def inspect(args: argparse.Namespace) -> dict[str, Any]:
    validate_limits(args)
    trust = trusted_input(args)
    source = regular_input_path(args.input, "input PDF", args.max_input_bytes, pdf=True)
    expected = expected_hash(args.expected_sha256)
    if sha256(source) != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {sha256(source)}")
    versions = require_signing_runtime()
    with tempfile.TemporaryDirectory(prefix="open-office-pyhanko-sign-inspect-") as temporary:
        snapshot = Path(temporary) / "source.pdf"
        snapshot_file(source, snapshot, expected)
        config = common_config(args, snapshot, "inspect")
        config["pageIndex"] = args.page_index
        result = run_worker(config, args.timeout_seconds, args.max_stdout_bytes, args.max_stderr_bytes)
        if sha256(snapshot) != expected or sha256(source) != expected:
            raise ProviderError("source PDF changed during signing inspection")
    return {
        "schema": SCHEMA_INSPECT,
        "ok": True,
        "operationCompleted": True,
        "provider": {"name": "pyhanko", "version": versions["pyHanko"]},
        "operation": "inspect-signing-surface",
        "savePolicy": "read-only",
        "silentFallback": False,
        "inputTrust": trust,
        "networkAllowed": False,
        "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": expected},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "readOnly": True, "sha256": expected},
        "pageCount": result["pageCount"],
        "selectedPage": result["selectedPage"],
        "signatureFields": result["signatureFields"],
        "summary": {
            "fieldCount": result["fieldCount"],
            "emptyFieldCount": result["emptyFieldCount"],
            "filledFieldCount": result["filledFieldCount"],
            "signatureCount": result["signatureCount"],
            "signatureNames": result["signatureNames"],
            "hasCertificationSignature": result["hasCertificationSignature"],
            "revisionCount": result["revisionCount"],
        },
        "limitations": [
            "inspection does not validate signature integrity or certificate trust; use pyhanko_provider.py verify",
            "the signing adapter is not a malware sandbox; attacker-chosen files require caller isolation",
            "timestamp, LTV, PKCS#11, and remote-signing capabilities are not exposed",
        ],
    }


def validate_sign_args(args: argparse.Namespace) -> None:
    field_name(args.field_name)
    for value, label in (
        (args.reason, "--reason"),
        (args.location, "--location"),
        (args.contact_info, "--contact-info"),
        (args.signer_name, "--signer-name"),
    ):
        bounded_metadata(value, label)
    if args.signature_kind == "certification" and args.docmdp_permission is None:
        raise ProviderError("certification signatures require an explicit --docmdp-permission")
    if args.signature_kind == "approval" and args.docmdp_permission is not None:
        raise ProviderError("--docmdp-permission is accepted only for certification signatures")
    if args.field_mode == "create-visible":
        if args.page_index is None or args.box is None:
            raise ProviderError("create-visible requires both --page-index and --box")
    elif args.page_index is not None or args.box is not None:
        raise ProviderError("--page-index and --box are accepted only with --field-mode create-visible")
    if args.expected_signature_count > args.max_signatures:
        raise ProviderError("--expected-signature-count exceeds --max-signatures")


def sign(args: argparse.Namespace) -> dict[str, Any]:
    validate_limits(args)
    validate_sign_args(args)
    trust = trusted_input(args)
    source = regular_input_path(args.input, "input PDF", args.max_input_bytes, pdf=True)
    destination = destination_path(args.output, source)
    credential = regular_input_path(
        args.credential, "PKCS#12 credential", args.max_credential_bytes, pdf=False
    )
    source_expected = expected_hash(args.expected_sha256)
    credential_expected = expected_hash(args.credential_sha256)
    source_actual = sha256(source)
    credential_actual = sha256(credential)
    if source_actual != source_expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {source_expected}, received {source_actual}")
    if credential_actual != credential_expected:
        raise ProviderError(
            f"PKCS#12 credential SHA-256 mismatch: expected {credential_expected}, received {credential_actual}"
        )
    versions = require_signing_runtime()
    passphrase = read_passphrase(args)
    output_hash = ""
    validator_report: dict[str, Any] = {}
    preflight_validator_report: dict[str, Any] | None = None
    worker: dict[str, Any] = {}
    try:
        with tempfile.TemporaryDirectory(prefix=".open-office-pyhanko-sign-", dir=destination.parent) as temporary:
            root = Path(temporary)
            root.chmod(0o700)
            source_snapshot = root / "source.pdf"
            credential_snapshot = root / "credential.p12"
            candidate = root / "candidate.pdf"
            snapshot_file(source, source_snapshot, source_expected)
            snapshot_file(credential, credential_snapshot, credential_expected)
            if args.expected_signature_count:
                preflight_validator_report = run_validator(
                    source_snapshot,
                    source_expected,
                    args.timeout_seconds,
                    args.max_stdout_bytes,
                    args.max_stderr_bytes,
                )
                if len(preflight_validator_report["signatures"]) != args.expected_signature_count:
                    raise ProviderError("pre-signature validator observed an unexpected signature count")
            config = common_config(args, source_snapshot, "sign")
            config.update({
                "output": str(candidate),
                "credential": str(credential_snapshot),
                "passphraseBase64": base64.b64encode(bytes(passphrase)).decode("ascii") if passphrase is not None else None,
                "expectedSignatureCount": args.expected_signature_count,
                "allowExistingSignatures": bool(args.allow_existing_signatures),
                "fieldName": args.field_name,
                "fieldMode": args.field_mode,
                "signatureKind": args.signature_kind,
                "docMDPPermission": args.docmdp_permission or "fill-forms",
                "subfilter": args.subfilter,
                "pageIndex": args.page_index,
                "box": list(args.box) if args.box is not None else None,
                "reason": args.reason,
                "location": args.location,
                "contactInfo": args.contact_info,
                "signerName": args.signer_name,
            })
            worker = run_worker(config, args.timeout_seconds, args.max_stdout_bytes, args.max_stderr_bytes)
            if not candidate.is_file():
                raise ProviderError("pyHanko signing worker did not produce a transactional output")
            candidate_size = candidate.stat().st_size
            if candidate_size < 5 or candidate_size > args.max_output_bytes:
                raise ProviderError(
                    f"signed PDF size {candidate_size} is outside the 5..{args.max_output_bytes} byte budget"
                )
            if not has_exact_prefix(source_snapshot, candidate):
                raise ProviderError("signed output does not preserve the complete exact source byte prefix")
            output_hash = sha256(candidate)
            if sha256(source_snapshot) != source_expected or sha256(source) != source_expected:
                raise ProviderError("source PDF changed during signing")
            if sha256(credential_snapshot) != credential_expected or sha256(credential) != credential_expected:
                raise ProviderError("PKCS#12 credential changed during signing")
            candidate.chmod(0o400)
            validator_report = run_validator(
                candidate, output_hash, args.timeout_seconds, args.max_stdout_bytes, args.max_stderr_bytes
            )
            signatures = validator_report["signatures"]
            if len(signatures) != args.expected_signature_count + 1:
                raise ProviderError("post-signature validator did not observe exactly one added signature")
            if signatures[-1].get("fieldName") != args.field_name:
                raise ProviderError("post-signature validator resolved a different final signature field")
            if signatures[-1].get("coverage") != "entire-file":
                raise ProviderError("new signature does not cover the entire final file")
            if sha256(source_snapshot) != source_expected or sha256(source) != source_expected:
                raise ProviderError("source PDF changed during post-signature validation")
            if sha256(credential_snapshot) != credential_expected or sha256(credential) != credential_expected:
                raise ProviderError("PKCS#12 credential changed during post-signature validation")
            candidate.chmod(0o600)
            with candidate.open("rb+") as stream:
                os.fsync(stream.fileno())
            publish_without_replace(candidate, destination)
    finally:
        if passphrase is not None:
            for index in range(len(passphrase)):
                passphrase[index] = 0

    for path, expected, label in (
        (source, source_expected, "source PDF"),
        (credential, credential_expected, "PKCS#12 credential"),
        (destination, output_hash, "published signed PDF"),
    ):
        if sha256(path) != expected:
            try:
                destination.unlink()
            except OSError:
                pass
            raise ProviderError(f"{label} hash changed before the signing report was finalized")

    new_signature = validator_report["signatures"][-1]
    return {
        "schema": SCHEMA_SIGN,
        "ok": True,
        "operationCompleted": True,
        "provider": {
            "name": "pyhanko",
            "version": versions["pyHanko"],
            "certvalidatorVersion": versions["pyhankoCertvalidator"],
            "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        },
        "operation": "sign-local-pkcs12",
        "savePolicy": "incremental",
        "silentFallback": False,
        "networkAllowed": False,
        "inputTrust": trust,
        "source": {"path": str(source), "bytes": source.stat().st_size, "sha256": source_expected},
        "output": {"path": str(destination), "bytes": destination.stat().st_size, "sha256": output_hash},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "readOnly": True, "sha256": source_expected},
        "credential": {
            "path": str(credential),
            "bytes": credential.stat().st_size,
            "sha256": credential_expected,
            "privateSnapshot": True,
            "passphraseChannel": "none" if args.no_passphrase else "stdin",
            "secretLogged": False,
            "certificate": worker["credential"],
            "certificateTrustValidated": False,
        },
        "signature": {
            "fieldName": args.field_name,
            "fieldMode": args.field_mode,
            "signatureKind": args.signature_kind,
            "docMDPPermission": args.docmdp_permission,
            "subfilter": args.subfilter,
            "digestAlgorithm": "sha256",
            "page": worker["selectedPage"],
            "box": list(args.box) if args.box is not None else None,
            "metadata": {
                "reason": args.reason,
                "location": args.location,
                "contactInfo": args.contact_info,
                "signerName": args.signer_name,
            },
            "timestampAuthorityUsed": False,
            "validationInfoEmbedded": False,
            "ltvEnabled": False,
            "padesProfileConformanceClaimed": False,
        },
        "existingSignatures": {
            "expected": args.expected_signature_count,
            "acknowledged": bool(args.allow_existing_signatures),
            "preflightValidated": preflight_validator_report is not None,
            "preflightAllIntegrityValid": (
                preflight_validator_report["summary"]["allIntegrityValid"]
                if preflight_validator_report is not None else None
            ),
            "preflightAllDocMDPCompliant": (
                preflight_validator_report["summary"]["allDocMDPCompliant"]
                if preflight_validator_report is not None else None
            ),
            "before": worker["before"],
            "after": worker["after"],
            "oldSignerApprovalOfNewRevisionClaimed": False,
        },
        "validation": {
            "sourcePrefixPreserved": True,
            "sourceReproved": True,
            "credentialReproved": True,
            "signatureCountDelta": 1,
            "newSignature": new_signature,
            "allIntegrityValid": validator_report["summary"]["allIntegrityValid"],
            "allDocMDPCompliant": validator_report["summary"]["allDocMDPCompliant"],
            "postValidationSchema": validator_report["schema"],
        },
        "transaction": {
            "atomicDistinctOutput": True,
            "noReplace": True,
            "privateSourceSnapshot": True,
            "privateCredentialSnapshot": True,
        },
        "requiredFollowups": [
            "run pyhanko_provider.py verify with explicit trust roots and the exact output hash",
            "run qpdf structural inspection on the exact output bytes",
            "render every output page with Poppler and review the signature appearance when visible",
            "bind source, credential, output, field, signer identity, and validation evidence into the canonical PDF audit",
        ],
        "limitations": [
            "the adapter signs only with a local PKCS#12 credential supplied by the caller",
            "certificate trust is not established during signing and must be validated separately",
            "timestamp authorities, LTV/DSS embedding, PKCS#11, remote signing, and complete PAdES conformance are not implemented",
            "an earlier valid signature covers its own revision; it does not imply approval of this new revision",
            "the adapter is not a malware sandbox; attacker-chosen input requires caller isolation",
        ],
    }


def add_limits(parser: argparse.ArgumentParser, *, output: bool = False, credential: bool = False) -> None:
    parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    if output:
        parser.add_argument("--max-output-bytes", type=positive_int, default=DEFAULT_MAX_OUTPUT_BYTES)
    if credential:
        parser.add_argument("--max-credential-bytes", type=positive_int, default=DEFAULT_MAX_CREDENTIAL_BYTES)
    parser.add_argument("--max-pages", type=positive_int, default=DEFAULT_MAX_PAGES)
    parser.add_argument("--max-fields", type=positive_int, default=DEFAULT_MAX_FIELDS)
    parser.add_argument("--max-signatures", type=positive_int, default=DEFAULT_MAX_SIGNATURES)
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
    subparsers.add_parser("probe", help="require the validated local-PKCS#12 pyHanko signing surface")

    inspect_parser = subparsers.add_parser("inspect", help="inspect signature fields and selected page geometry")
    inspect_parser.add_argument("input")
    inspect_parser.add_argument("--expected-sha256", required=True)
    inspect_parser.add_argument("--page-index", type=nonnegative_int)
    add_trust(inspect_parser)
    add_limits(inspect_parser)

    sign_parser = subparsers.add_parser("sign", help="add one bounded incremental PDF signature")
    sign_parser.add_argument("input")
    sign_parser.add_argument("output")
    sign_parser.add_argument("--expected-sha256", required=True)
    sign_parser.add_argument("--credential", required=True)
    sign_parser.add_argument("--credential-sha256", required=True)
    passphrase = sign_parser.add_mutually_exclusive_group(required=True)
    passphrase.add_argument("--passphrase-stdin", action="store_true")
    passphrase.add_argument("--no-passphrase", action="store_true")
    sign_parser.add_argument("--field-name", required=True)
    sign_parser.add_argument("--field-mode", choices=["existing", "create-invisible", "create-visible"], required=True)
    sign_parser.add_argument("--page-index", type=nonnegative_int)
    sign_parser.add_argument("--box", type=parse_box)
    sign_parser.add_argument("--signature-kind", choices=["approval", "certification"], required=True)
    sign_parser.add_argument("--docmdp-permission", choices=DOCMDP_PERMISSIONS)
    sign_parser.add_argument("--subfilter", choices=SUBFILTERS, default="pades")
    sign_parser.add_argument("--expected-signature-count", type=nonnegative_int, required=True)
    sign_parser.add_argument("--allow-existing-signatures", action="store_true")
    sign_parser.add_argument("--reason")
    sign_parser.add_argument("--location")
    sign_parser.add_argument("--contact-info")
    sign_parser.add_argument("--signer-name")
    add_trust(sign_parser)
    add_limits(sign_parser, output=True, credential=True)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python

    reexec_configured_provider_python()
    if len(sys.argv) > 1 and sys.argv[1] == "_worker":
        return worker_main()
    args = build_parser().parse_args()
    try:
        result = probe() if args.command == "probe" else inspect(args) if args.command == "inspect" else sign(args)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except ProviderError as exc:
        print(json.dumps({
            "ok": False,
            "error": str(exc),
            "provider": "pyhanko",
            "operation": "sign-local-pkcs12" if getattr(args, "command", None) == "sign" else getattr(args, "command", None),
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
