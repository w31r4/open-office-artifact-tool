#!/usr/bin/env python3
"""Probe and validate explicit PDF provider/save-policy contracts."""

from __future__ import annotations

import argparse
import hashlib
import importlib.util
import json
import os
from pathlib import Path
import shutil
import subprocess
import sys


PROVIDERS = {
    "artifact-tool": {"kind": "command", "commands": ["node"], "role": "greenfield tagged model", "integration": "shipped-js-api"},
    "reportlab": {"kind": "module", "module": "reportlab", "role": "greenfield layout generation", "integration": "shipped-thin-script"},
    "pdfplumber": {"kind": "module", "module": "pdfplumber", "role": "read-only extraction", "integration": "shipped-thin-script"},
    "pypdf": {"kind": "module", "module": "pypdf", "role": "basic structure, attachment quarantine, complete-source merge/reorder/selective stamp, forms, annotations, rewrite/incremental", "integration": "shipped-thin-script"},
    "pymupdf": {"kind": "module", "module": "pymupdf", "role": "advanced imported-PDF editing and sanitize", "license": "agpl-or-commercial", "integration": "shipped-thin-script"},
    "poppler": {"kind": "command", "commands": ["pdfinfo", "pdftoppm"], "role": "native file/render QA", "integration": "shipped-js-adapter-and-cli-workflow"},
    "qpdf": {"kind": "command", "commands": ["qpdf"], "role": "structure inspection, recovery, encryption, and rewrite", "integration": "external-documented"},
    "pikepdf": {"kind": "module", "module": "pikepdf", "role": "Python qpdf structure, attachment, and active-content operations", "integration": "planned-no-shipped-adapter"},
    "pyhanko": {"kind": "command_or_module", "commands": ["pyhanko"], "module": "pyhanko", "role": "signing and signature validation", "integration": "external-documented"},
    "verapdf": {"kind": "command", "commands": ["verapdf"], "role": "PDF/A and PDF/UA validation", "integration": "external-documented"},
    "ocrmypdf": {"kind": "command_or_module", "commands": ["ocrmypdf"], "module": "ocrmypdf", "role": "scanned-PDF OCR and searchable layer generation", "integration": "planned-no-shipped-adapter"},
    "tesseract": {"kind": "command", "commands": ["tesseract"], "role": "OCR engine used by strict image residue checks", "integration": "external-required-for-image-ocr"},
}


TASKS = {
    "create-tagged": {"providers": ["artifact-tool"], "strategies": ["rewrite"], "input": "none"},
    "create-layout": {"providers": ["reportlab"], "strategies": ["rewrite"], "input": "none"},
    "extract": {"providers": ["pdfplumber", "pypdf", "pymupdf"], "strategies": ["read-only"], "input": "existing"},
    "extract-attachments": {"providers": ["pypdf"], "strategies": ["read-only"], "input": "existing"},
    "inspect": {"providers": ["artifact-tool", "pypdf", "pymupdf", "qpdf"], "strategies": ["read-only"], "input": "existing"},
    "edit-content": {"providers": ["pymupdf"], "strategies": ["rewrite", "incremental"], "input": "existing", "mutation": True},
    "fill-form": {"providers": ["pypdf", "pymupdf"], "strategies": ["rewrite", "incremental"], "input": "existing", "mutation": True},
    "annotate": {"providers": ["pypdf", "pymupdf"], "strategies": ["rewrite", "incremental"], "input": "existing", "mutation": True},
    "merge-stamp": {"providers": ["pypdf"], "strategies": ["rewrite"], "input": "existing", "mutation": True},
    "repair": {"providers": ["qpdf", "pikepdf"], "strategies": ["rewrite"], "input": "existing", "mutation": True},
    "structure-clean": {"providers": ["qpdf", "pikepdf"], "strategies": ["rewrite"], "input": "existing", "mutation": True, "invalidate_signatures": True},
    "ocr": {"providers": ["ocrmypdf"], "strategies": ["rewrite"], "input": "existing", "mutation": True},
    "sign": {"providers": ["pyhanko"], "strategies": ["incremental"], "input": "existing", "mutation": True},
    "verify-signature": {"providers": ["pyhanko"], "strategies": ["read-only"], "input": "existing"},
    "validate-conformance": {"providers": ["verapdf"], "strategies": ["read-only"], "input": "existing"},
    "redact": {"providers": ["pymupdf"], "strategies": ["sanitize"], "input": "existing", "mutation": True, "invalidate_signatures": True},
    "sanitize": {"providers": ["pymupdf"], "strategies": ["sanitize"], "input": "existing", "mutation": True, "invalidate_signatures": True},
    "render": {"providers": ["poppler"], "strategies": ["read-only"], "input": "existing"},
}


class ContractError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def module_version(name: str) -> str | None:
    if importlib.util.find_spec(name) is None:
        return None
    try:
        module = __import__(name)
        return str(getattr(module, "__version__", getattr(module, "Version", "available")))
    except Exception as exc:  # pragma: no cover - provider-specific import failures
        return f"import-error: {exc}"


def command_version(command: str) -> str | None:
    resolved = shutil.which(command)
    if not resolved:
        return None
    for args in (["-v"], ["--version"], ["-version"]):
        try:
            result = subprocess.run([resolved, *args], check=False, capture_output=True, text=True, timeout=10)
        except Exception:
            continue
        output = (result.stdout or result.stderr).strip().splitlines()
        if output and result.returncode == 0 and not any("couldn't open file" in line.lower() for line in output):
            return output[0][:300]
    return resolved


def probe_provider(name: str) -> dict:
    config = PROVIDERS[name]
    result = {"provider": name, "role": config["role"], "kind": config["kind"], "integration": config["integration"], "available": False, "evidence": {}}
    if config["kind"] == "module":
        version = module_version(config["module"])
        result["available"] = version is not None and not version.startswith("import-error:")
        result["evidence"]["module"] = config["module"]
        result["evidence"]["version"] = version
    elif config["kind"] == "command":
        versions = {command: command_version(command) for command in config["commands"]}
        result["available"] = all(versions.values())
        result["evidence"]["commands"] = versions
    else:
        versions = {command: command_version(command) for command in config["commands"]}
        module = module_version(config["module"])
        result["available"] = any(versions.values()) or (module is not None and not module.startswith("import-error:"))
        result["evidence"].update({"commands": versions, "module": config["module"], "version": module})
    if config.get("license"):
        result["license"] = config["license"]
        result["licenseAccepted"] = str(os.environ.get("OPEN_OFFICE_PDF_PYMUPDF_LICENSE", "")).lower() in {"agpl", "commercial"}
    return result


def accepted_pymupdf_license(value: str | None) -> str | None:
    selected = str(value or os.environ.get("OPEN_OFFICE_PDF_PYMUPDF_LICENSE", "")).strip().lower()
    return selected if selected in {"agpl", "commercial"} else None


def validate_plan(args: argparse.Namespace) -> dict:
    capability = TASKS[args.task]
    if args.provider not in capability["providers"]:
        raise ContractError(
            f"provider {args.provider!r} cannot perform task {args.task!r}; "
            f"allowed providers: {', '.join(capability['providers'])}"
        )
    if args.strategy not in capability["strategies"]:
        raise ContractError(
            f"strategy {args.strategy!r} is invalid for task {args.task!r}; "
            f"allowed strategies: {', '.join(capability['strategies'])}"
        )
    if args.strategy == "sanitize" and args.provider != "pymupdf":
        raise ContractError("sanitize requires the explicit pymupdf provider")
    license_choice = accepted_pymupdf_license(args.accept_license) if args.provider == "pymupdf" else None
    if args.provider == "pymupdf" and not license_choice:
        raise ContractError("pymupdf requires --accept-license agpl|commercial or OPEN_OFFICE_PDF_PYMUPDF_LICENSE")

    input_path = Path(args.input).expanduser().resolve() if args.input else None
    output_path = Path(args.output).expanduser().resolve() if args.output else None
    if capability["input"] == "existing":
        if not input_path or not input_path.is_file():
            raise ContractError("this task requires an existing --input PDF")
    elif input_path:
        raise ContractError("greenfield creation does not accept --input")
    if capability.get("mutation"):
        if not output_path:
            raise ContractError("mutating tasks require --output")
        if input_path == output_path:
            raise ContractError("input and output must be different; never overwrite the source PDF in place")
    elif output_path and args.strategy == "read-only":
        raise ContractError("read-only tasks do not write a PDF --output")
    if capability.get("invalidate_signatures") and not args.invalidate_signatures:
        raise ContractError("this destructive rewrite requires explicit --invalidate-signatures acknowledgement")

    probe = probe_provider(args.provider)
    if args.require_provider and not probe["available"]:
        raise ContractError(f"required provider {args.provider!r} is unavailable: {probe['evidence']}")
    result = {
        "task": args.task,
        "provider": args.provider,
        "strategy": args.strategy,
        "inputMode": capability["input"],
        "mutation": bool(capability.get("mutation")),
        "silentFallback": False,
        "providerProbe": probe,
    }
    if input_path:
        result["input"] = {"path": str(input_path), "bytes": input_path.stat().st_size, "sha256": sha256(input_path)}
    if output_path:
        result["output"] = {"path": str(output_path)}
    if license_choice:
        result["licenseChoice"] = license_choice
    if capability.get("invalidate_signatures"):
        result["invalidateSignatures"] = True
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    check = subparsers.add_parser("check", help="probe one or every provider")
    check.add_argument("--provider", choices=["all", *PROVIDERS], required=True)
    check.add_argument("--require", action="store_true", help="fail if the selected provider is unavailable")

    plan = subparsers.add_parser("plan", help="validate an explicit task/provider/save-policy contract")
    plan.add_argument("--task", choices=TASKS, required=True)
    plan.add_argument("--provider", choices=PROVIDERS, required=True)
    plan.add_argument("--strategy", choices=["read-only", "rewrite", "incremental", "sanitize"], required=True)
    plan.add_argument("--input")
    plan.add_argument("--output")
    plan.add_argument("--accept-license", choices=["agpl", "commercial"])
    plan.add_argument("--invalidate-signatures", action="store_true")
    plan.add_argument("--require-provider", action="store_true")
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = build_parser().parse_args()
    try:
        if args.command == "check":
            names = list(PROVIDERS) if args.provider == "all" else [args.provider]
            probes = [probe_provider(name) for name in names]
            print(json.dumps({"providers": probes}, indent=2, sort_keys=True))
            if args.require and not all(probe["available"] for probe in probes):
                return 2
            return 0
        print(json.dumps(validate_plan(args), indent=2, sort_keys=True))
        return 0
    except ContractError as exc:
        print(json.dumps({"ok": False, "error": str(exc), "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
