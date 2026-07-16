#!/usr/bin/env python3
"""Validate the canonical PDF mutation audit envelope against delivered bytes."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys


SCHEMA = "open-office-artifact-tool.pdf-audit.v1"
SHA256 = re.compile(r"^[0-9a-f]{64}$")
SAVE_POLICIES = {"rewrite", "incremental", "sanitize"}


class AuditError(RuntimeError):
    pass


def digest(path: Path) -> dict[str, str | int]:
    payload = path.read_bytes()
    return {"bytes": len(payload), "sha256": hashlib.sha256(payload).hexdigest()}


def require_object(value, label: str) -> dict:
    if not isinstance(value, dict):
        raise AuditError(f"{label} must be an object")
    return value


def validate_file_evidence(evidence, actual_path: Path, label: str) -> None:
    record = require_object(evidence, label)
    if not isinstance(record.get("path"), str) or not record["path"].strip():
        raise AuditError(f"{label}.path must be a non-empty string")
    recorded_path = Path(record["path"]).expanduser()
    if not recorded_path.is_absolute() or recorded_path.resolve() != actual_path.expanduser().resolve():
        raise AuditError(f"{label}.path must be the exact absolute path {actual_path.expanduser().resolve()}")
    if not isinstance(record.get("bytes"), int) or record["bytes"] < 1:
        raise AuditError(f"{label}.bytes must be a positive integer")
    if not isinstance(record.get("sha256"), str) or not SHA256.fullmatch(record["sha256"]):
        raise AuditError(f"{label}.sha256 must be a lowercase SHA-256 digest")
    if not actual_path.is_file():
        raise AuditError(f"{label} file does not exist: {actual_path}")
    actual = digest(actual_path)
    if record["bytes"] != actual["bytes"] or record["sha256"] != actual["sha256"]:
        raise AuditError(f"{label} bytes/hash do not match {actual_path}")


def validate_record(record: dict, source: Path, artifact: Path | None, required_operation: str | None) -> dict:
    record = require_object(record, "audit")
    if record.get("schema") != SCHEMA:
        raise AuditError(f"schema must be {SCHEMA!r}")
    status = record.get("status")
    if status not in {"succeeded", "failed_closed"}:
        raise AuditError("status must be 'succeeded' or 'failed_closed'")
    validate_file_evidence(record.get("source"), source, "source")

    provider = require_object(record.get("provider"), "provider")
    for field in ("actual", "version"):
        if not isinstance(provider.get(field), str) or not provider[field].strip():
            raise AuditError(f"provider.{field} must be a non-empty string")
    if provider.get("silentFallback") is not False:
        raise AuditError("provider.silentFallback must be false")

    policy = require_object(record.get("savePolicy"), "savePolicy")
    if policy.get("strategy") not in SAVE_POLICIES:
        raise AuditError("savePolicy.strategy must be rewrite, incremental, or sanitize")
    preflight = require_object(record.get("preflight"), "preflight")
    if not isinstance(preflight.get("probeCompleted"), bool) or not isinstance(preflight.get("planCompleted"), bool):
        raise AuditError("preflight.probeCompleted and preflight.planCompleted must be booleans")
    operation = require_object(record.get("operation"), "operation")
    if not isinstance(operation.get("type"), str) or not operation["type"].strip():
        raise AuditError("operation.type must be a non-empty string")
    if required_operation and operation["type"] != required_operation:
        raise AuditError(f"operation.type must be {required_operation!r}")
    require_object(record.get("validation"), "validation")

    if status == "succeeded":
        if preflight["probeCompleted"] is not True or preflight["planCompleted"] is not True:
            raise AuditError("a succeeded audit requires completed provider probe and route plan")
        if artifact is None:
            raise AuditError("--artifact is required for a succeeded audit")
        validate_file_evidence(record.get("output"), artifact, "output")
    else:
        if record.get("output") is not None:
            raise AuditError("failed_closed audit output must be null")
        if artifact is not None and artifact.exists():
            raise AuditError("failed_closed audit must not have a partial artifact")
        if not isinstance(record.get("reason"), str) or not record["reason"].strip():
            raise AuditError("failed_closed audit requires a non-empty reason")

    return {
        "ok": True,
        "schema": SCHEMA,
        "status": status,
        "provider": provider["actual"],
        "providerVersion": provider["version"],
        "savePolicy": policy["strategy"],
        "operation": operation["type"],
        "silentFallback": False,
    }


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    subparsers = root.add_subparsers(dest="command", required=True)
    validate = subparsers.add_parser("validate", help="validate a canonical audit and recompute byte evidence")
    validate.add_argument("audit", type=Path)
    validate.add_argument("--source", type=Path, required=True)
    validate.add_argument("--artifact", type=Path)
    validate.add_argument("--require-operation")
    return root


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = parser().parse_args()
    try:
        record = json.loads(args.audit.read_text("utf8"))
        print(json.dumps(validate_record(record, args.source, args.artifact, args.require_operation), indent=2, sort_keys=True))
        return 0
    except (AuditError, OSError, json.JSONDecodeError) as exc:
        print(json.dumps({"ok": False, "error": str(exc)}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
