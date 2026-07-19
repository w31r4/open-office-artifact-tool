#!/usr/bin/env python3
"""Run bounded, source-bound PDF/A or PDF/UA validation through veraPDF."""

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
from typing import Any


SCHEMA = "open-office-artifact-tool.verapdf-validation.v1"
MINIMUM_VERSION = (1, 30, 0)
MAXIMUM_VERSION_EXCLUSIVE = (1, 31, 0)
HARD_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_INPUT_BYTES = HARD_MAX_INPUT_BYTES
HARD_MAX_TIMEOUT_SECONDS = 180
DEFAULT_TIMEOUT_SECONDS = 90
HARD_MAX_STDOUT_BYTES = 16 * 1024 * 1024
DEFAULT_MAX_STDOUT_BYTES = 8 * 1024 * 1024
HARD_MAX_STDERR_BYTES = 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 256 * 1024
HARD_MAX_FAILURES_DISPLAYED = 20
DEFAULT_MAX_FAILURES_DISPLAYED = 3
MAX_RULE_SUMMARIES = 512
MAX_TAGS = 64
MAX_ERROR_ARGUMENTS = 32

FLAVOURS = {
    "1a": {"family": "PDF/A", "profile": "PDF/A-1a validation profile"},
    "1b": {"family": "PDF/A", "profile": "PDF/A-1b validation profile"},
    "2a": {"family": "PDF/A", "profile": "PDF/A-2a validation profile"},
    "2b": {"family": "PDF/A", "profile": "PDF/A-2b validation profile"},
    "2u": {"family": "PDF/A", "profile": "PDF/A-2u validation profile"},
    "3a": {"family": "PDF/A", "profile": "PDF/A-3a validation profile"},
    "3b": {"family": "PDF/A", "profile": "PDF/A-3b validation profile"},
    "3u": {"family": "PDF/A", "profile": "PDF/A-3u validation profile"},
    "4": {"family": "PDF/A", "profile": "PDF/A-4 validation profile"},
    "4e": {"family": "PDF/A", "profile": "PDF/A-4e validation profile"},
    "4f": {"family": "PDF/A", "profile": "PDF/A-4f validation profile"},
    "ua1": {"family": "PDF/UA", "profile": "PDF/UA-1 validation profile"},
    "ua2": {"family": "PDF/UA", "profile": "PDF/UA-2 + Tagged PDF validation profile"},
}


class ProviderError(RuntimeError):
    pass


def bounded_text(value: Any, limit: int = 2_048) -> str:
    text = str(value)
    return text if len(text) <= limit else text[:limit] + "…"


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def file_identity(path: Path, label: str) -> tuple[int, str]:
    try:
        size = path.stat().st_size
        digest = sha256(path)
    except OSError as exc:
        raise ProviderError(f"{label} became unavailable during validation: {bounded_text(exc)}") from exc
    return size, digest


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


def checked_budget(name: str, value: int, hard_maximum: int) -> int:
    if value > hard_maximum:
        raise ProviderError(f"{name} cannot exceed the hard maximum {hard_maximum}")
    return value


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


def executable_path() -> Path:
    configured = os.environ.get("OPEN_OFFICE_PDF_VERAPDF", "").strip()
    candidate = Path(configured).expanduser() if configured else Path(shutil.which("verapdf") or "")
    if not str(candidate) or not candidate.is_file() or not os.access(candidate, os.X_OK):
        detail = f" configured by OPEN_OFFICE_PDF_VERAPDF={configured!r}" if configured else " on PATH"
        raise ProviderError(f"veraPDF executable is unavailable{detail}")
    return candidate.resolve()


def provider_environment(private_root: Path | None = None) -> dict[str, str]:
    environment = dict(os.environ)
    # veraPDF's launcher honours these Java/classpath injection variables. The
    # adapter exposes no arbitrary JVM or classpath surface.
    for name in (
        "BASH_ENV",
        "CLASSPATH_PREFIX",
        "ENDORSED_DIR",
        "ENV",
        "JAVA_OPTS",
        "JDK_JAVA_OPTIONS",
        "REPO",
        "_JAVA_OPTIONS",
        "JAVA_TOOL_OPTIONS",
    ):
        environment.pop(name, None)
    environment["NO_COLOR"] = "1"
    if private_root is not None:
        environment["HOME"] = str(private_root)
        environment["TMPDIR"] = str(private_root)
        environment["XDG_CACHE_HOME"] = str(private_root / "cache")
        environment["XDG_CONFIG_HOME"] = str(private_root / "config")
    return environment


def run_bounded(
    executable: Path,
    arguments: list[str],
    *,
    timeout_seconds: int,
    max_stdout_bytes: int,
    max_stderr_bytes: int,
    environment: dict[str, str] | None = None,
) -> subprocess.CompletedProcess:
    command = [str(executable), *arguments]
    stdout = bytearray()
    stderr = bytearray()
    violations: list[str] = []
    lock = threading.Lock()
    try:
        process = subprocess.Popen(
            command,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=False,
            env=environment or provider_environment(),
        )
    except OSError as exc:
        raise ProviderError(f"veraPDF could not start: {bounded_text(exc)}") from exc

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
                    record_violation(f"veraPDF {label} exceeded the {limit} byte budget")
                    break
                buffer.extend(chunk)
        except OSError as exc:
            record_violation(f"veraPDF {label} capture failed: {bounded_text(exc)}")
        finally:
            stream.close()

    stdout_thread = threading.Thread(target=pump, args=(process.stdout, stdout, max_stdout_bytes, "stdout"), daemon=True)
    stderr_thread = threading.Thread(target=pump, args=(process.stderr, stderr, max_stderr_bytes, "stderr"), daemon=True)
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
    if timed_out:
        raise ProviderError(f"veraPDF timed out after {timeout_seconds} seconds")
    if violations:
        raise ProviderError(violations[0])
    return subprocess.CompletedProcess(command, process.returncode, bytes(stdout), bytes(stderr))


def decoded(value: bytes) -> str:
    return value.decode("utf-8", "replace")


def parsed_version(value: str) -> tuple[int, int, int] | None:
    match = re.search(r"\bveraPDF\s+(\d+)\.(\d+)\.(\d+)\b", value, re.IGNORECASE)
    return tuple(int(part) for part in match.groups()) if match else None


def version_text(value: tuple[int, int, int]) -> str:
    return ".".join(str(part) for part in value)


def provider_version(executable: Path, timeout_seconds: int) -> str:
    result = run_bounded(
        executable,
        ["--version"],
        timeout_seconds=timeout_seconds,
        max_stdout_bytes=32 * 1024,
        max_stderr_bytes=32 * 1024,
    )
    output = "\n".join(part for part in (decoded(result.stdout).strip(), decoded(result.stderr).strip()) if part)
    version = parsed_version(output)
    if result.returncode != 0 or version is None:
        raise ProviderError(f"veraPDF --version failed or returned an unsupported response: {bounded_text(output)}")
    if not MINIMUM_VERSION <= version < MAXIMUM_VERSION_EXCLUSIVE:
        raise ProviderError(
            f"veraPDF >= {version_text(MINIMUM_VERSION)} and < {version_text(MAXIMUM_VERSION_EXCLUSIVE)} "
            f"is required; received {version_text(version)}"
        )
    return version_text(version)


def supported_profiles(executable: Path, timeout_seconds: int) -> list[str]:
    result = run_bounded(
        executable,
        ["--list"],
        timeout_seconds=timeout_seconds,
        max_stdout_bytes=128 * 1024,
        max_stderr_bytes=32 * 1024,
    )
    output = "\n".join(part for part in (decoded(result.stdout).strip(), decoded(result.stderr).strip()) if part)
    if result.returncode != 0:
        raise ProviderError(f"veraPDF --list failed: {bounded_text(output)}")
    profiles = sorted(set(re.findall(r"^\s*([0-9a-z]+)\s+-\s+", output, re.MULTILINE)))
    missing = sorted(set(FLAVOURS) - set(profiles))
    if missing:
        raise ProviderError(f"veraPDF is missing required built-in profiles: {', '.join(missing)}")
    return profiles


def probe(timeout_seconds: int) -> dict[str, Any]:
    executable = executable_path()
    version = provider_version(executable, timeout_seconds)
    profiles = supported_profiles(executable, timeout_seconds)
    return {
        "ok": True,
        "provider": "verapdf",
        "providerVersion": version,
        "executable": str(executable),
        "integration": "shipped-thin-script-external-cli",
        "capabilities": ["PDF/A-machine-validation", "PDF/UA-machine-validation"],
        "profiles": profiles,
        "customProfilesAccepted": False,
        "passwordsAccepted": False,
        "silentFallback": False,
    }


def nonnegative_int(value: Any, label: str) -> int:
    if isinstance(value, bool) or not isinstance(value, int) or value < 0:
        raise ProviderError(f"veraPDF report field {label} is not a nonnegative integer")
    return value


def bounded_tags(value: Any) -> list[str]:
    if value is None:
        return []
    if not isinstance(value, list) or len(value) > MAX_TAGS:
        raise ProviderError("veraPDF rule tags are malformed or exceed the tag budget")
    return [bounded_text(tag, 128) for tag in value]


def parsed_check(value: Any) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ProviderError("veraPDF failed-check evidence is malformed")
    arguments = value.get("errorArguments") or []
    if not isinstance(arguments, list) or len(arguments) > MAX_ERROR_ARGUMENTS:
        raise ProviderError("veraPDF failed-check arguments are malformed or exceed the argument budget")
    return {
        "status": bounded_text(value.get("status", "unknown"), 64),
        "context": bounded_text(value.get("context", ""), 2_048),
        "errorMessage": bounded_text(value.get("errorMessage", ""), 2_048),
        "errorArguments": [None if item is None else bounded_text(item, 256) for item in arguments],
    }


def parsed_rule(value: Any, max_failures_displayed: int) -> dict[str, Any]:
    if not isinstance(value, dict):
        raise ProviderError("veraPDF rule summary is malformed")
    checks = value.get("checks") or []
    if not isinstance(checks, list) or len(checks) > max_failures_displayed:
        raise ProviderError("veraPDF rule check list exceeds the configured display budget")
    test_number = value.get("testNumber")
    if isinstance(test_number, bool) or not isinstance(test_number, (int, str)):
        raise ProviderError("veraPDF rule test number is malformed")
    return {
        "ruleStatus": bounded_text(value.get("ruleStatus", "unknown"), 64),
        "specification": bounded_text(value.get("specification", ""), 256),
        "clause": bounded_text(value.get("clause", ""), 128),
        "testNumber": test_number if isinstance(test_number, int) else bounded_text(test_number, 64),
        "status": bounded_text(value.get("status", "unknown"), 64),
        "failedChecks": nonnegative_int(value.get("failedChecks"), "rule.failedChecks"),
        "tags": bounded_tags(value.get("tags")),
        "description": bounded_text(value.get("description", ""), 2_048),
        "object": bounded_text(value.get("object", ""), 512),
        "test": bounded_text(value.get("test", ""), 2_048),
        "reportedChecks": [parsed_check(check) for check in checks],
    }


def parsed_release_details(value: Any) -> list[dict[str, str]]:
    if not isinstance(value, list) or not value or len(value) > 32:
        raise ProviderError("veraPDF build release details are missing or malformed")
    details: list[dict[str, str]] = []
    for item in value:
        if not isinstance(item, dict):
            raise ProviderError("veraPDF build release detail is malformed")
        identifier = bounded_text(item.get("id", ""), 128)
        version = bounded_text(item.get("version", ""), 64)
        parsed = tuple(int(part) for part in version.split(".")) if re.fullmatch(r"\d+\.\d+\.\d+", version) else None
        if parsed is None or not MINIMUM_VERSION <= parsed < MAXIMUM_VERSION_EXCLUSIVE:
            raise ProviderError(f"veraPDF report contains an unsupported component version: {identifier}={version}")
        details.append({"id": identifier, "version": version})
    return details


def parse_report(
    raw: bytes,
    *,
    returncode: int,
    flavour: str,
    source_bytes: int,
    max_failures_displayed: int,
) -> dict[str, Any]:
    try:
        payload = json.loads(raw.decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ProviderError("veraPDF did not return one valid UTF-8 JSON document") from exc
    try:
        report = payload["report"]
        release_details = parsed_release_details(report["buildInformation"]["releaseDetails"])
        jobs = report["jobs"]
        batch = report["batchSummary"]
    except (KeyError, TypeError) as exc:
        raise ProviderError("veraPDF JSON report is missing required top-level fields") from exc
    if not isinstance(jobs, list) or len(jobs) != 1 or not isinstance(jobs[0], dict):
        raise ProviderError("veraPDF must return exactly one validation job")
    job = jobs[0]
    item_details = job.get("itemDetails")
    results = job.get("validationResult")
    if not isinstance(item_details, dict) or item_details.get("size") != source_bytes:
        raise ProviderError("veraPDF item size does not match the immutable source snapshot")
    if not isinstance(results, list) or len(results) != 1 or not isinstance(results[0], dict):
        raise ProviderError("veraPDF must return exactly one validation result")
    result = results[0]
    details = result.get("details")
    if not isinstance(details, dict):
        raise ProviderError("veraPDF validation details are missing or malformed")
    rules = details.get("ruleSummaries") or []
    if not isinstance(rules, list) or len(rules) > MAX_RULE_SUMMARIES:
        raise ProviderError(f"veraPDF rule summaries exceed the {MAX_RULE_SUMMARIES} rule budget")
    parsed_rules = [parsed_rule(rule, max_failures_displayed) for rule in rules]
    passed_rules = nonnegative_int(details.get("passedRules"), "passedRules")
    failed_rules = nonnegative_int(details.get("failedRules"), "failedRules")
    passed_checks = nonnegative_int(details.get("passedChecks"), "passedChecks")
    failed_checks = nonnegative_int(details.get("failedChecks"), "failedChecks")
    if failed_rules != len(parsed_rules):
        raise ProviderError("veraPDF did not report one bounded summary for every failed rule")
    compliant = result.get("compliant")
    if not isinstance(compliant, bool):
        raise ProviderError("veraPDF compliance result is not boolean")
    if compliant and (failed_rules or failed_checks):
        raise ProviderError("veraPDF marked a report compliant while also reporting failures")
    if not compliant and failed_rules == 0 and failed_checks == 0:
        raise ProviderError("veraPDF marked a report noncompliant without failure evidence")
    if result.get("jobEndStatus") != "normal":
        raise ProviderError(f"veraPDF validation did not end normally: {bounded_text(result.get('jobEndStatus'))}")
    expected_profile = FLAVOURS[flavour]["profile"]
    if result.get("profileName") != expected_profile:
        raise ProviderError(
            f"veraPDF validated the wrong profile: expected {expected_profile!r}, received {result.get('profileName')!r}"
        )
    expected_returncode = 0 if compliant else 1
    if returncode != expected_returncode:
        raise ProviderError(
            f"veraPDF exit status {returncode} contradicts the compliance result (expected {expected_returncode})"
        )
    if not isinstance(batch, dict):
        raise ProviderError("veraPDF batch summary is malformed")
    for field in ("outOfMemory", "veraExceptions", "failedEncryptedJobs", "failedParsingJobs"):
        if nonnegative_int(batch.get(field), f"batchSummary.{field}") != 0:
            raise ProviderError(f"veraPDF batch summary reports {field}={batch.get(field)}")
    if nonnegative_int(batch.get("totalJobs"), "batchSummary.totalJobs") != 1:
        raise ProviderError("veraPDF batch summary does not describe exactly one job")
    validation_summary = batch.get("validationSummary")
    if not isinstance(validation_summary, dict):
        raise ProviderError("veraPDF validation summary is missing or malformed")
    if (
        nonnegative_int(validation_summary.get("totalJobCount"), "validationSummary.totalJobCount") != 1
        or nonnegative_int(validation_summary.get("successfulJobCount"), "validationSummary.successfulJobCount") != 1
        or nonnegative_int(validation_summary.get("failedJobCount"), "validationSummary.failedJobCount") != 0
    ):
        raise ProviderError("veraPDF validation summary does not prove one successfully processed job")
    return {
        "releaseDetails": release_details,
        "profileName": expected_profile,
        "compliant": compliant,
        "statement": bounded_text(result.get("statement", ""), 1_024),
        "summary": {
            "passedRules": passed_rules,
            "failedRules": failed_rules,
            "passedChecks": passed_checks,
            "failedChecks": failed_checks,
            "tags": bounded_tags(details.get("tags")),
        },
        "failedRuleSummaries": parsed_rules,
        "processingTime": {
            "duration": bounded_text((job.get("processingTime") or {}).get("duration", ""), 64)
        },
    }


def validate(args: argparse.Namespace) -> dict[str, Any]:
    max_input_bytes = checked_budget("--max-input-bytes", args.max_input_bytes, HARD_MAX_INPUT_BYTES)
    timeout_seconds = checked_budget("--timeout-seconds", args.timeout_seconds, HARD_MAX_TIMEOUT_SECONDS)
    max_stdout_bytes = checked_budget("--max-stdout-bytes", args.max_stdout_bytes, HARD_MAX_STDOUT_BYTES)
    max_stderr_bytes = checked_budget("--max-stderr-bytes", args.max_stderr_bytes, HARD_MAX_STDERR_BYTES)
    max_failures_displayed = checked_budget(
        "--max-failures-displayed", args.max_failures_displayed, HARD_MAX_FAILURES_DISPLAYED
    )
    source = input_path(args.input, max_input_bytes)
    expected = expected_hash(args.expected_sha256)
    source_bytes, source_hash = file_identity(source, "source PDF")
    if source_hash != expected:
        raise ProviderError("source SHA-256 mismatch; inspect the exact current bytes before validation")
    executable = executable_path()
    version = provider_version(executable, min(timeout_seconds, 30))

    with tempfile.TemporaryDirectory(prefix="open-office-verapdf-") as temporary:
        private_root = Path(temporary)
        os.chmod(private_root, 0o700)
        snapshot = private_root / "input.pdf"
        try:
            shutil.copyfile(source, snapshot)
            with snapshot.open("rb+") as stream:
                os.fsync(stream.fileno())
            os.chmod(snapshot, 0o400)
        except OSError as exc:
            raise ProviderError(f"could not create the private source snapshot: {bounded_text(exc)}") from exc
        _, snapshot_hash = file_identity(snapshot, "private source snapshot")
        if snapshot_hash != expected:
            raise ProviderError("private source snapshot hash mismatch")
        result = run_bounded(
            executable,
            [
                "--format",
                "json",
                "--loglevel",
                "0",
                "--maxfailuresdisplayed",
                str(max_failures_displayed),
                "--flavour",
                args.flavour,
                str(snapshot),
            ],
            timeout_seconds=timeout_seconds,
            max_stdout_bytes=max_stdout_bytes,
            max_stderr_bytes=max_stderr_bytes,
            environment=provider_environment(private_root),
        )
        diagnostics = decoded(result.stderr).strip()
        if diagnostics:
            raise ProviderError(f"veraPDF emitted unexpected diagnostics: {bounded_text(diagnostics)}")
        _, snapshot_hash = file_identity(snapshot, "private source snapshot")
        if snapshot_hash != expected:
            raise ProviderError("veraPDF changed the private read-only source snapshot")
        parsed = parse_report(
            result.stdout,
            returncode=result.returncode,
            flavour=args.flavour,
            source_bytes=source_bytes,
            max_failures_displayed=max_failures_displayed,
        )
        raw_report = {
            "format": "json",
            "bytes": len(result.stdout),
            "sha256": hashlib.sha256(result.stdout).hexdigest(),
            "retained": False,
        }

    final_source_bytes, final_source_hash = file_identity(source, "source PDF")
    if final_source_bytes != source_bytes or final_source_hash != expected:
        raise ProviderError("source PDF changed during validation")
    human_review_required = FLAVOURS[args.flavour]["family"] == "PDF/UA"
    failures = [] if parsed["compliant"] or not args.require_compliant else [
        f"PDF is not compliant with the explicitly selected {parsed['profileName']} machine rules"
    ]
    return {
        "schema": SCHEMA,
        "ok": not failures,
        "operationCompleted": True,
        "operation": "validate-conformance",
        "provider": {
            "name": "verapdf",
            "version": version,
            "executable": str(executable),
            "releaseDetails": parsed["releaseDetails"],
        },
        "savePolicy": "read-only",
        "silentFallback": False,
        "customProfilesAccepted": False,
        "passwordsAccepted": False,
        "networkFetchingConfigured": False,
        "source": {"path": str(source), "bytes": source_bytes, "sha256": expected},
        "sourceProtected": True,
        "sourceSnapshot": {"privateCopy": True, "sha256": expected},
        "validationPolicy": {
            "flavour": args.flavour,
            "family": FLAVOURS[args.flavour]["family"],
            "profileName": parsed["profileName"],
            "explicitBuiltInProfile": True,
            "maxFailuresDisplayedPerRule": max_failures_displayed,
        },
        "machineRuleCompliant": parsed["compliant"],
        "statement": parsed["statement"],
        "summary": parsed["summary"],
        "failedRuleSummaries": parsed["failedRuleSummaries"],
        "processingTime": parsed["processingTime"],
        "rawProviderReport": raw_report,
        "humanReview": {
            "required": human_review_required,
            "reason": (
                "PDF/UA includes author-intent and usability judgments that machine rules cannot establish"
                if human_review_required
                else "independent visual and semantic QA is still recommended for archival delivery"
            ),
        },
        "policyGates": {
            "requested": {"requireCompliant": bool(args.require_compliant)},
            "passed": not failures,
            "failures": failures,
        },
        "limitations": [
            "the report proves only the selected veraPDF machine-rule result for these exact bytes",
            "PDF/UA machine compliance does not replace human review of reading order, semantics, alternatives, contrast, or assistive-technology usability",
            "this adapter does not repair metadata or structure and does not accept custom profiles, passwords, directories, or network inputs",
        ],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    probe_parser = subparsers.add_parser("probe", help="require the validated veraPDF CLI range and built-in profiles")
    probe_parser.add_argument("--timeout-seconds", type=positive_int, default=30)

    validate_parser = subparsers.add_parser("validate", help="validate one exact PDF against one explicit built-in profile")
    validate_parser.add_argument("input")
    validate_parser.add_argument("--expected-sha256", required=True)
    validate_parser.add_argument("--flavour", choices=sorted(FLAVOURS), required=True)
    validate_parser.add_argument("--require-compliant", action="store_true")
    validate_parser.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    validate_parser.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)
    validate_parser.add_argument("--max-stdout-bytes", type=positive_int, default=DEFAULT_MAX_STDOUT_BYTES)
    validate_parser.add_argument("--max-stderr-bytes", type=positive_int, default=DEFAULT_MAX_STDERR_BYTES)
    validate_parser.add_argument(
        "--max-failures-displayed", type=positive_int, default=DEFAULT_MAX_FAILURES_DISPLAYED
    )
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python

    reexec_configured_provider_python()
    args = build_parser().parse_args()
    try:
        if args.command == "probe":
            timeout = checked_budget("--timeout-seconds", args.timeout_seconds, HARD_MAX_TIMEOUT_SECONDS)
            print(json.dumps(probe(timeout), indent=2, sort_keys=True))
            return 0
        report = validate(args)
        stream = sys.stdout if report["ok"] else sys.stderr
        print(json.dumps(report, indent=2, sort_keys=True), file=stream)
        return 0 if report["ok"] else 2
    except ProviderError as exc:
        print(
            json.dumps(
                {"schema": SCHEMA, "ok": False, "error": str(exc), "silentFallback": False},
                indent=2,
                sort_keys=True,
            ),
            file=sys.stderr,
        )
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
