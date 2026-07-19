#!/usr/bin/env python3
"""Validate PDF signatures with bounded, source-bound pyHanko evidence."""

from __future__ import annotations

import argparse
from datetime import datetime
import hashlib
from importlib import metadata
import json
import logging
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import threading
from typing import Any


SCHEMA = "open-office-artifact-tool.pyhanko-verify.v1"
SUPPORTED_PYHANKO_MIN = (0, 35, 0)
SUPPORTED_PYHANKO_MAX_EXCLUSIVE = (0, 36, 0)
SUPPORTED_CERTVALIDATOR_MIN = (0, 31, 0)
SUPPORTED_CERTVALIDATOR_MAX_EXCLUSIVE = (0, 32, 0)
DEFAULT_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_TIMEOUT_SECONDS = 120
DEFAULT_MAX_STDOUT_BYTES = 2 * 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 512 * 1024
MAX_WORKER_CONFIG_BYTES = 256 * 1024
MAX_CERTIFICATE_BYTES = 4 * 1024 * 1024
MAX_CERTIFICATE_TOTAL_BYTES = 64 * 1024 * 1024
MAX_CERTIFICATES_PER_ROLE = 32
MAX_SIGNATURES = 64
MAX_TEXT_CHARS = 4_096


class ProviderError(RuntimeError):
    pass


def bounded_text(value: Any, limit: int = MAX_TEXT_CHARS) -> str | None:
    if value is None:
        return None
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def package_version(name: str) -> str:
    try:
        return metadata.version(name)
    except metadata.PackageNotFoundError as exc:
        raise ProviderError(
            f"{name} is unavailable in {sys.executable}; install pyHanko 0.35.x in the explicitly selected PDF provider environment"
        ) from exc


def parsed_version(value: str) -> tuple[int, int, int]:
    match = re.match(r"^(\d+)\.(\d+)\.(\d+)", value)
    if not match:
        raise ProviderError(f"cannot parse pyHanko version {value!r}")
    return tuple(int(part) for part in match.groups())


def provider_versions() -> dict[str, str]:
    pyhanko = package_version("pyHanko")
    parsed = parsed_version(pyhanko)
    if not (SUPPORTED_PYHANKO_MIN <= parsed < SUPPORTED_PYHANKO_MAX_EXCLUSIVE):
        raise ProviderError(
            "pyHanko >=0.35.0,<0.36.0 is required by this validated adapter; "
            f"received {pyhanko!r}"
        )
    certvalidator = package_version("pyhanko-certvalidator")
    parsed_certvalidator = parsed_version(certvalidator)
    if not (
        SUPPORTED_CERTVALIDATOR_MIN
        <= parsed_certvalidator
        < SUPPORTED_CERTVALIDATOR_MAX_EXCLUSIVE
    ):
        raise ProviderError(
            "pyhanko-certvalidator >=0.31.0,<0.32.0 is required by this validated adapter; "
            f"received {certvalidator!r}"
        )
    return {"pyHanko": pyhanko, "pyhankoCertvalidator": certvalidator}


def probe() -> dict[str, Any]:
    versions = provider_versions()
    try:
        from pyhanko.pdf_utils.reader import PdfFileReader  # noqa: F401
        from pyhanko.sign.validation import validate_pdf_signature  # noqa: F401
        from pyhanko_certvalidator import ValidationContext  # noqa: F401
    except Exception as exc:
        raise ProviderError(f"pyHanko validation APIs could not be imported: {bounded_text(exc)}") from exc
    return {
        "ok": True,
        "provider": "pyhanko",
        "providerVersion": versions["pyHanko"],
        "certvalidatorVersion": versions["pyhankoCertvalidator"],
        "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        "integration": "shipped-thin-script-external-python",
        "capabilities": ["read-only-signature-validation", "difference-analysis", "DocMDP", "FieldMDP"],
        "trustRoots": "explicit-only",
        "networkAllowed": False,
        "silentFallback": False,
        "padesProfileConformanceClaimed": False,
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


def parse_moment(value: str | None) -> str | None:
    if value is None:
        return None
    normalized = value.strip()
    if normalized.endswith("Z"):
        normalized = normalized[:-1] + "+00:00"
    try:
        parsed = datetime.fromisoformat(normalized)
    except ValueError as exc:
        raise ProviderError("--moment must be an ISO 8601 timestamp") from exc
    if parsed.tzinfo is None or parsed.utcoffset() is None:
        raise ProviderError("--moment must include an explicit UTC offset")
    return parsed.isoformat()


def certificate_path(value: str) -> Path:
    target = Path(value).expanduser().resolve()
    if not target.is_file():
        raise ProviderError(f"certificate does not exist or is not a regular file: {target}")
    size = target.stat().st_size
    if size < 1 or size > MAX_CERTIFICATE_BYTES:
        raise ProviderError(
            f"certificate {target} size {size} is outside the 1..{MAX_CERTIFICATE_BYTES} byte budget"
        )
    return target


def snapshot_certificates(
    values: list[str], role: str, temporary_root: Path
) -> tuple[list[str], list[dict[str, Any]]]:
    if len(values) > MAX_CERTIFICATES_PER_ROLE:
        raise ProviderError(f"at most {MAX_CERTIFICATES_PER_ROLE} {role} certificate files are accepted")
    snapshots: list[str] = []
    evidence: list[dict[str, Any]] = []
    total = 0
    for index, value in enumerate(values):
        source = certificate_path(value)
        total += source.stat().st_size
        if total > MAX_CERTIFICATE_TOTAL_BYTES:
            raise ProviderError(f"{role} certificates exceed the {MAX_CERTIFICATE_TOTAL_BYTES} byte aggregate budget")
        source_hash = sha256(source)
        snapshot = temporary_root / f"{role}-{index}.cert"
        shutil.copyfile(source, snapshot)
        with snapshot.open("rb+") as stream:
            os.fsync(stream.fileno())
        if sha256(snapshot) != source_hash:
            raise ProviderError(f"{role} certificate snapshot hash mismatch: {source}")
        snapshots.append(str(snapshot))
        evidence.append({
            "path": str(source),
            "bytes": source.stat().st_size,
            "sha256": source_hash,
        })
    return snapshots, evidence


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
    )
    stdout = bytearray()
    stderr = bytearray()
    violations: list[str] = []
    lock = threading.Lock()

    def record_violation(message: str) -> None:
        with lock:
            if not violations:
                violations.append(message)
        try:
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
                    record_violation(f"pyHanko worker {label} exceeded the {limit} byte budget")
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
        process.kill()
        process.wait()
    stdout_thread.join()
    stderr_thread.join()
    if timed_out:
        raise ProviderError(f"pyHanko validation timed out after {timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    stderr_text = stderr.decode("utf-8", "replace").strip()
    if process.returncode != 0:
        raise ProviderError(f"pyHanko worker failed (exit {process.returncode}): {bounded_text(stderr_text)}")
    if stderr_text:
        raise ProviderError(f"pyHanko worker emitted unexpected diagnostics: {bounded_text(stderr_text)}")
    try:
        result = json.loads(stdout.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("pyHanko worker did not return one valid UTF-8 JSON document") from exc
    if not isinstance(result, dict) or result.get("workerSchema") != 1:
        raise ProviderError("pyHanko worker returned an unsupported result schema")
    return result


def iso_datetime(value: Any) -> str | None:
    return value.isoformat() if isinstance(value, datetime) else None


def enum_name(value: Any) -> str | None:
    if value is None:
        return None
    name = getattr(value, "name", None)
    return str(name).lower().replace("_", "-") if name else bounded_text(value)


def certificate_identity(certificate: Any) -> dict[str, Any]:
    fingerprint = str(certificate.sha256_fingerprint).replace(" ", "").lower()
    return {
        "subject": bounded_text(certificate.subject.human_friendly, 2_048),
        "issuer": bounded_text(certificate.issuer.human_friendly, 2_048),
        "serialNumber": str(certificate.serial_number),
        "sha256Fingerprint": fingerprint,
        "notValidBefore": iso_datetime(certificate.not_valid_before),
        "notValidAfter": iso_datetime(certificate.not_valid_after),
    }


def timestamp_status(status: Any) -> dict[str, Any] | None:
    if status is None:
        return None
    return {
        "intact": bool(status.intact),
        "cryptographicallyValid": bool(status.valid),
        "trusted": bool(status.trusted),
        "bottomLine": bool(status.bottom_line),
        "validationTime": iso_datetime(getattr(status, "validation_time", None)),
        "summary": bounded_text(status.summary(), 1_024),
    }


def signature_base(signature: Any, index: int) -> dict[str, Any]:
    byte_range = [int(value) for value in signature.byte_range]
    docmdp = signature.docmdp_level
    fieldmdp = signature.fieldmdp
    return {
        "index": index,
        "fieldName": bounded_text(signature.field_name, 2_048),
        "signedRevision": int(signature.signed_revision),
        "byteRange": byte_range,
        "signedByteCount": sum(byte_range[1::2]),
        "filter": bounded_text(signature.sig_object.get("/Filter"), 256),
        "subFilter": bounded_text(signature.sig_object.get("/SubFilter"), 256),
        "docMDP": {
            "present": docmdp is not None,
            "permission": enum_name(docmdp),
            "permissionCode": int(docmdp.value) if docmdp is not None else None,
        },
        "fieldMDP": {
            "present": fieldmdp is not None,
            "action": enum_name(fieldmdp.action) if fieldmdp is not None else None,
            "fields": [bounded_text(value, 2_048) for value in (fieldmdp.fields or [])][:1_024]
            if fieldmdp is not None else [],
        },
        "signerCertificate": certificate_identity(signature.signer_cert),
        "embeddedCertificateCount": 1 + len(signature.other_embedded_certs),
    }


def validated_signature(signature: Any, index: int, signer_context: Any, timestamp_context: Any) -> dict[str, Any]:
    from pyhanko.sign.validation import validate_pdf_signature

    result = signature_base(signature, index)
    try:
        status = validate_pdf_signature(
            signature,
            signer_validation_context=signer_context,
            ts_validation_context=timestamp_context,
        )
        changed_fields = []
        if status.diff_result is not None:
            changed_fields = sorted(bounded_text(value, 2_048) for value in status.diff_result.changed_form_fields)[:1_024]
        result.update({
            "validationCompleted": True,
            "intact": bool(status.intact),
            "cryptographicallyValid": bool(status.valid),
            "trusted": bool(status.trusted),
            "bottomLine": bool(status.bottom_line),
            "summary": bounded_text(status.summary(), 1_024),
            "coverage": enum_name(status.coverage),
            "modificationLevel": enum_name(status.modification_level),
            "changedFormFields": changed_fields,
            "docMDPCompliant": bool(status.docmdp_ok),
            "hasSeedValues": bool(status.has_seed_values),
            "seedValueCompliant": bool(status.seed_value_ok) if status.has_seed_values else None,
            "seedValueError": bounded_text(status.seed_value_constraint_error),
            "digestAlgorithm": bounded_text(status.md_algorithm, 256),
            "signatureMechanism": bounded_text(status.pkcs7_signature_mechanism, 256),
            "validationTime": iso_datetime(status.validation_time),
            "signerReportedTime": iso_datetime(status.signer_reported_dt),
            "revoked": bool(status.revoked),
            "revocationEvidencePresent": status.revocation_details is not None,
            "trustProblemIndicator": enum_name(status.trust_problem_indic),
            "signatureTimestamp": timestamp_status(status.timestamp_validity),
            "contentTimestamp": timestamp_status(status.content_timestamp_validity),
        })
    except Exception as exc:  # pyHanko reports malformed CMS and unsupported constructs here
        result.update({
            "validationCompleted": False,
            "validationError": bounded_text(f"{type(exc).__name__}: {exc}"),
        })
    return result


def worker_main() -> int:
    logging.disable(logging.CRITICAL)
    try:
        raw_config = sys.stdin.buffer.read(MAX_WORKER_CONFIG_BYTES + 1)
        if len(raw_config) > MAX_WORKER_CONFIG_BYTES:
            raise ProviderError(f"worker configuration exceeds {MAX_WORKER_CONFIG_BYTES} bytes")
        config = json.loads(raw_config)
        if not isinstance(config, dict):
            raise ProviderError("worker configuration must be a JSON object")
        provider_versions()
        from pyhanko.keys import load_cert_from_pemder
        from pyhanko.pdf_utils.reader import PdfFileReader
        from pyhanko_certvalidator import ValidationContext

        roots = [load_cert_from_pemder(path) for path in config["trustRoots"]]
        others = [load_cert_from_pemder(path) for path in config["otherCertificates"]]
        moment = datetime.fromisoformat(config["moment"]) if config.get("moment") else None

        def validation_context() -> Any:
            return ValidationContext(
                trust_roots=roots,
                other_certs=others,
                moment=moment,
                allow_fetching=False,
                revocation_mode=config["revocationPolicy"],
            )

        with Path(config["input"]).open("rb") as stream:
            reader = PdfFileReader(stream, strict=True)
            if reader.security_handler is not None:
                raise ProviderError("encrypted PDFs are unsupported by this validation slice")
            signatures = list(reader.embedded_signatures)
            if len(signatures) > MAX_SIGNATURES:
                raise ProviderError(f"PDF contains more than {MAX_SIGNATURES} embedded signatures")
            records = [
                validated_signature(signature, index, validation_context(), validation_context())
                for index, signature in enumerate(signatures)
            ]
            total_revisions = int(reader.xrefs.total_revisions)
        print(json.dumps({
            "workerSchema": 1,
            "totalRevisions": total_revisions,
            "trustAnchors": [certificate_identity(value) for value in roots],
            "otherCertificates": [certificate_identity(value) for value in others],
            "signatures": records,
        }, separators=(",", ":"), sort_keys=True))
        return 0
    except Exception as exc:
        print(bounded_text(f"{type(exc).__name__}: {exc}"), file=sys.stderr)
        return 2


def aggregate(records: list[dict[str, Any]]) -> dict[str, Any]:
    complete = [record for record in records if record.get("validationCompleted")]
    return {
        "signatureCount": len(records),
        "validationCompletedCount": len(complete),
        "allValidationCompleted": len(complete) == len(records),
        "allIntegrityValid": bool(records) and all(record.get("intact") and record.get("cryptographicallyValid") for record in records),
        "allTrusted": bool(records) and all(record.get("trusted") for record in records),
        "allBottomLine": bool(records) and all(record.get("bottomLine") for record in records),
        "allDocMDPCompliant": bool(records) and all(record.get("docMDPCompliant") for record in records),
        "hasPostSigningChanges": any(record.get("modificationLevel") not in {None, "none"} for record in records),
        "hasSignatureTimestamp": any(record.get("signatureTimestamp") is not None for record in records),
    }


def conclusion(summary: dict[str, Any]) -> str:
    if summary["signatureCount"] == 0:
        return "unsigned"
    if not summary["allValidationCompleted"]:
        return "validation-incomplete"
    if not summary["allIntegrityValid"]:
        return "integrity-failure"
    if not summary["allDocMDPCompliant"]:
        return "post-signing-policy-failure"
    if not summary["allTrusted"]:
        return "cryptographically-valid-untrusted"
    if summary["allBottomLine"]:
        return "valid-under-selected-policy"
    return "policy-indeterminate"


def verify(args: argparse.Namespace) -> dict[str, Any]:
    versions = provider_versions()
    for label, value, maximum in (
        ("--max-input-bytes", args.max_input_bytes, DEFAULT_MAX_INPUT_BYTES),
        ("--timeout-seconds", args.timeout_seconds, DEFAULT_TIMEOUT_SECONDS),
        ("--max-stdout-bytes", args.max_stdout_bytes, DEFAULT_MAX_STDOUT_BYTES),
        ("--max-stderr-bytes", args.max_stderr_bytes, DEFAULT_MAX_STDERR_BYTES),
    ):
        if value > maximum:
            raise ProviderError(f"{label} cannot exceed the hard maximum {maximum}")
    source = input_path(args.input, args.max_input_bytes)
    source_bytes = source.stat().st_size
    expected = expected_hash(args.expected_sha256)
    source_hash = sha256(source)
    if source_hash != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {source_hash}")
    moment = parse_moment(args.moment)
    if args.trust_policy == "explicit-roots" and not args.trust_root:
        raise ProviderError("--trust-policy explicit-roots requires at least one --trust-root certificate")
    if args.trust_policy == "cryptographic-only" and args.trust_root:
        raise ProviderError("--trust-root requires --trust-policy explicit-roots; trust is never inferred silently")
    if args.require_all_trusted and args.trust_policy != "explicit-roots":
        raise ProviderError("--require-all-trusted requires --trust-policy explicit-roots")

    with tempfile.TemporaryDirectory(prefix="open-office-pyhanko-") as temporary:
        temporary_root = Path(temporary)
        snapshot = temporary_root / "source.pdf"
        shutil.copyfile(source, snapshot)
        with snapshot.open("rb+") as stream:
            os.fsync(stream.fileno())
        if sha256(snapshot) != expected:
            raise ProviderError("private validation snapshot does not match the inspected source SHA-256")
        roots, root_evidence = snapshot_certificates(args.trust_root, "trust-root", temporary_root)
        others, other_evidence = snapshot_certificates(args.other_cert, "other-cert", temporary_root)
        if sum(value["bytes"] for value in [*root_evidence, *other_evidence]) > MAX_CERTIFICATE_TOTAL_BYTES:
            raise ProviderError(f"all certificate inputs exceed the {MAX_CERTIFICATE_TOTAL_BYTES} byte aggregate budget")
        worker = run_worker({
            "input": str(snapshot),
            "trustPolicy": args.trust_policy,
            "trustRoots": roots,
            "otherCertificates": others,
            "moment": moment,
            "revocationPolicy": args.revocation_policy,
        }, args.timeout_seconds, args.max_stdout_bytes, args.max_stderr_bytes)
        if len(worker["trustAnchors"]) != len(root_evidence) or len(worker["otherCertificates"]) != len(other_evidence):
            raise ProviderError("pyHanko worker certificate evidence count does not match the snapshotted inputs")
        if sha256(snapshot) != expected or sha256(source) != expected:
            raise ProviderError("source PDF changed during signature validation")
        for evidence in [*root_evidence, *other_evidence]:
            if sha256(Path(evidence["path"])) != evidence["sha256"]:
                raise ProviderError(f"certificate changed during signature validation: {evidence['path']}")

    if sha256(source) != expected:
        raise ProviderError("source PDF changed before the validation report was finalized")
    summary = aggregate(worker["signatures"])
    gates = {
        "requireSignature": bool(args.require_signature),
        "requireAllIntegrityValid": bool(args.require_all_integrity_valid),
        "requireAllTrusted": bool(args.require_all_trusted),
        "requireDocMDPCompliant": bool(args.require_docmdp_compliant),
        "requireAllBottomLine": bool(args.require_all_bottom_line),
    }
    failures = []
    if args.require_signature and summary["signatureCount"] == 0:
        failures.append("document has no embedded signatures")
    if summary["signatureCount"] and not summary["allIntegrityValid"]:
        failures.append("one or more signatures are not intact and cryptographically valid")
    elif args.require_all_integrity_valid and not summary["allIntegrityValid"]:
        failures.append("not every signature is intact and cryptographically valid")
    if args.require_all_trusted and not summary["allTrusted"]:
        failures.append("not every signature chains to the explicit trust roots")
    if summary["signatureCount"] and not summary["allDocMDPCompliant"]:
        failures.append("one or more post-signing changes do not comply with DocMDP")
    elif args.require_docmdp_compliant and not summary["allDocMDPCompliant"]:
        failures.append("not every post-signing change complies with DocMDP")
    if args.require_all_bottom_line and not summary["allBottomLine"]:
        failures.append("not every signature satisfies pyHanko's selected-policy bottom line")
    validation_errors = [record for record in worker["signatures"] if not record.get("validationCompleted")]
    if validation_errors:
        failures.append(f"{len(validation_errors)} signature validation record(s) are incomplete")

    return {
        "schema": SCHEMA,
        "ok": not failures,
        "operationCompleted": True,
        "provider": {
            "name": "pyhanko",
            "version": versions["pyHanko"],
            "certvalidatorVersion": versions["pyhankoCertvalidator"],
            "python": {"executable": sys.executable, "version": sys.version.split()[0]},
        },
        "operation": "verify-signatures",
        "savePolicy": "read-only",
        "silentFallback": False,
        "networkAllowed": False,
        "padesProfileConformanceClaimed": False,
        "source": {"path": str(source), "bytes": source_bytes, "sha256": expected},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "sha256": expected},
        "validationPolicy": {
            "trustPolicy": args.trust_policy,
            "trustRoots": [
                {**evidence, "certificate": identity}
                for evidence, identity in zip(root_evidence, worker["trustAnchors"])
            ],
            "otherCertificates": [
                {**evidence, "certificate": identity}
                for evidence, identity in zip(other_evidence, worker["otherCertificates"])
            ],
            "validationMoment": moment or "provider-current-time",
            "revocationPolicy": args.revocation_policy,
            "networkAllowed": False,
        },
        "revisionCount": worker["totalRevisions"],
        "signatures": worker["signatures"],
        "summary": summary,
        "conclusion": conclusion(summary),
        "policyGates": {"requested": gates, "passed": not failures, "failures": failures},
        "limitations": [
            "certificate trust is only relative to the explicit validation policy in this report",
            "no network fetching is performed; revocation evidence must already be available to pyHanko",
            "this adapter does not claim complete PAdES profile conformance",
        ],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("probe", help="require the validated pyHanko core library range")

    verify_parser = subparsers.add_parser("verify", help="validate embedded signatures without mutating the PDF")
    verify_parser.add_argument("input")
    verify_parser.add_argument("--expected-sha256", required=True)
    verify_parser.add_argument("--trust-policy", choices=["cryptographic-only", "explicit-roots"], default="cryptographic-only")
    verify_parser.add_argument("--trust-root", action="append", default=[])
    verify_parser.add_argument("--other-cert", action="append", default=[])
    verify_parser.add_argument("--moment")
    verify_parser.add_argument("--revocation-policy", choices=["none", "soft-fail", "hard-fail", "require"], default="none")
    verify_parser.add_argument("--require-signature", action="store_true")
    verify_parser.add_argument("--require-all-integrity-valid", action="store_true")
    verify_parser.add_argument("--require-all-trusted", action="store_true")
    verify_parser.add_argument("--require-docmdp-compliant", action="store_true")
    verify_parser.add_argument("--require-all-bottom-line", action="store_true")
    verify_parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    verify_parser.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)
    verify_parser.add_argument("--max-stdout-bytes", type=positive_int, default=DEFAULT_MAX_STDOUT_BYTES)
    verify_parser.add_argument("--max-stderr-bytes", type=positive_int, default=DEFAULT_MAX_STDERR_BYTES)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    if len(sys.argv) > 1 and sys.argv[1] == "_worker":
        return worker_main()
    args = build_parser().parse_args()
    try:
        result = probe() if args.command == "probe" else verify(args)
        stream = sys.stdout if result.get("ok", False) else sys.stderr
        print(json.dumps(result, indent=2, sort_keys=True), file=stream)
        return 0 if result.get("ok", False) else 2
    except (ProviderError, OSError, ValueError) as exc:
        print(json.dumps({
            "ok": False,
            "provider": "pyhanko",
            "error": bounded_text(exc),
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
