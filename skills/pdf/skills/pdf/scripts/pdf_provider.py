#!/usr/bin/env python3
"""Probe and validate explicit PDF provider/save-policy contracts."""

from __future__ import annotations

import argparse
import hashlib
from importlib import metadata
import importlib.util
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys


CATALOG_SCHEMA = "open-office-artifact-tool.pdf-provider-catalog.v1"


def load_catalog() -> dict:
    """Read the canonical Node export without maintaining a Python mirror.

    A packaged Skill may run in the package itself or be copied into an Agent
    workspace. In both cases Node module resolution finds the installed public
    `open-office-artifact-tool/pdf/providers` subpath. This command reads data
    only; it never initializes MuPDF or a specialist provider.
    """

    node = shutil.which("node")
    if not node:
        raise RuntimeError("Node.js is required to read the canonical PDF provider catalog")
    program = (
        "import { PDF_PROVIDER_CATALOG } from 'open-office-artifact-tool/pdf/providers';"
        "process.stdout.write(JSON.stringify(PDF_PROVIDER_CATALOG));"
    )
    candidates = [Path.cwd(), *[parent for parent in Path(__file__).resolve().parents if (parent / "package.json").is_file()]]
    errors = []
    for cwd in dict.fromkeys(candidates):
        result = subprocess.run(
            [node, "--input-type=module", "--eval", program],
            cwd=str(cwd),
            check=False,
            capture_output=True,
            text=True,
            timeout=10,
        )
        if result.returncode == 0:
            try:
                return json.loads(result.stdout)
            except json.JSONDecodeError as exc:
                raise RuntimeError("canonical PDF provider catalog emitted invalid JSON") from exc
        errors.append((result.stderr or result.stdout).strip())
    detail = next((error for error in errors if error), "module was not resolvable")
    raise RuntimeError(f"unable to load canonical PDF provider catalog: {detail[:300]}")


CATALOG = load_catalog()
if CATALOG.get("schema") != CATALOG_SCHEMA or CATALOG.get("schemaVersion") != 1:
    raise RuntimeError("PDF provider catalog has an unsupported schema")
if not isinstance(CATALOG.get("providers"), dict) or not isinstance(CATALOG.get("tasks"), dict):
    raise RuntimeError("PDF provider catalog is missing providers or tasks")
PROVIDERS = CATALOG["providers"]
TASKS = CATALOG["tasks"]


class ContractError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def module_version(name: str, distribution: str | None = None) -> str | None:
    if importlib.util.find_spec(name) is None:
        return None
    try:
        module = __import__(name)
        if distribution:
            return metadata.version(distribution)
        return str(getattr(module, "__version__", getattr(module, "Version", "available")))
    except Exception as exc:  # pragma: no cover - provider-specific import failures
        return f"import-error: {exc}"


def command_version(command: str, environment: str | None = None, require_version_output: bool = False) -> str | None:
    configured = str(os.environ.get(environment, "")).strip() if environment else ""
    if configured:
        candidate = Path(configured).expanduser()
        if not candidate.is_file() or not os.access(candidate, os.X_OK):
            return None
        resolved = str(candidate.resolve())
    else:
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
    return None if require_version_output else resolved


def command_major(version: str | None) -> int | None:
    if not version:
        return None
    match = re.search(r"(?<!\d)(\d+)(?:\.\d+)", version)
    return int(match.group(1)) if match else None


def semantic_version(value: str | None) -> tuple[int, int, int] | None:
    if not value:
        return None
    match = re.search(r"(?<!\d)(\d+)\.(\d+)\.(\d+)", value)
    return tuple(int(part) for part in match.groups()) if match else None


def version_in_range(value: str | None, config: dict, prefix: str = "") -> bool:
    parsed = semantic_version(value)
    if parsed is None:
        return False
    minimum = config.get(f"{prefix}minimumVersion")
    maximum = config.get(f"{prefix}maximumVersionExclusive")
    if minimum and parsed < semantic_version(minimum):
        return False
    if maximum and parsed >= semantic_version(maximum):
        return False
    exact = config.get(f"{prefix}exactVersion")
    return not exact or parsed == semantic_version(exact)


def node_package_version(package: str) -> str | None:
    node = shutil.which("node")
    if not node:
        return None
    program = "const fs=require('fs'),path=require('path');let p=path.dirname(require.resolve(process.argv[1]));for(;;){const c=path.join(p,'package.json');try{const m=JSON.parse(fs.readFileSync(c));if(m.name===process.argv[1]){process.stdout.write(m.version);break}}catch{}const n=path.dirname(p);if(n===p)process.exit(2);p=n}"
    try:
        result = subprocess.run([node, "-e", program, package], check=False, capture_output=True, text=True, timeout=10)
    except Exception:
        return None
    return result.stdout.strip() if result.returncode == 0 and result.stdout.strip() else None


def probe_provider(name: str) -> dict:
    config = PROVIDERS[name]
    result = {"provider": name, "role": config["role"], "kind": config["kind"], "integration": config["integration"], "available": False, "evidence": {}}
    if config["kind"] == "core":
        result["available"] = bool(shutil.which("node"))
        result["evidence"] = {"runtime": "open-office-artifact-tool", "node": shutil.which("node")}
    elif config["kind"] == "node-package":
        version = node_package_version(config["package"])
        result["available"] = version_in_range(version, config)
        result["evidence"] = {"package": config["package"], "version": version, "expectedVersion": config["exactVersion"]}
    elif config["kind"] == "python-module":
        version = module_version(config["module"], config.get("distribution"))
        result["available"] = version is not None and not version.startswith("import-error:")
        result["evidence"]["module"] = config["module"]
        result["evidence"]["version"] = version
        if config.get("minimumVersion") is not None:
            result["evidence"].update({
                "minimumVersion": config["minimumVersion"],
                "maximumVersionExclusive": config["maximumVersionExclusive"],
            })
            result["available"] = bool(result["available"] and version_in_range(version, config))
        if config.get("companionModule"):
            companion = module_version(config["companionModule"], config["companionDistribution"])
            result["evidence"].update({
                "companionModule": config["companionModule"],
                "companionVersion": companion,
                "companionMinimumVersion": config["companionMinimumVersion"],
                "companionMaximumVersionExclusive": config["companionMaximumVersionExclusive"],
            })
            companion_config = {
                "minimumVersion": config["companionMinimumVersion"],
                "maximumVersionExclusive": config["companionMaximumVersionExclusive"],
            }
            result["available"] = bool(result["available"] and version_in_range(companion, companion_config))
    elif config["kind"] == "command":
        versions = {
            command: command_version(command, config.get("environment"), config.get("requireVersionOutput", False))
            for command in config["commands"]
        }
        result["available"] = all(versions.values())
        result["evidence"]["commands"] = versions
        if config.get("minimumMajor") is not None:
            majors = {command: command_major(version) for command, version in versions.items()}
            result["evidence"].update({"majorVersions": majors, "minimumMajor": config["minimumMajor"]})
            result["available"] = all(
                major is not None and major >= config["minimumMajor"]
                for major in majors.values()
            )
        if config.get("minimumVersion") is not None:
            result["evidence"].update({
                "semanticVersions": {command: ".".join(str(part) for part in semantic_version(version)) if semantic_version(version) else None for command, version in versions.items()},
                "minimumVersion": config["minimumVersion"],
                "maximumVersionExclusive": config["maximumVersionExclusive"],
            })
            result["available"] = bool(result["available"] and all(version_in_range(version, config) for version in versions.values()))
    if config.get("license", {}).get("requiresAcknowledgement"):
        result["license"] = config["license"].get("id", config["license"]["expression"])
        result["licenseAccepted"] = str(os.environ.get("OPEN_OFFICE_PDF_PYMUPDF_LICENSE", "")).lower() in set(config["license"].get("acceptedValues", []))
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
    if capability.get("invalidateSignatures") and not args.invalidate_signatures:
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
        "integration": capability.get("integration", probe["integration"]),
        "providerProbe": probe,
    }
    if input_path:
        result["input"] = {"path": str(input_path), "bytes": input_path.stat().st_size, "sha256": sha256(input_path)}
    if output_path:
        result["output"] = {"path": str(output_path)}
    if license_choice:
        result["licenseChoice"] = license_choice
    if capability.get("invalidateSignatures"):
        result["invalidateSignatures"] = True
    return result


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    check = subparsers.add_parser("check", help="probe one or every provider")
    check.add_argument("--provider", choices=["all", *PROVIDERS], required=True)
    check.add_argument("--require", action="store_true", help="fail if the selected provider is unavailable")

    catalog = subparsers.add_parser("catalog", help="print the canonical provider catalog without probing a provider runtime")
    catalog.add_argument("--compact", action="store_true", help="omit indentation")

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
    args = build_parser().parse_args()
    try:
        if args.command == "catalog":
            print(json.dumps(CATALOG, indent=None if args.compact else 2, sort_keys=True))
            return 0
        from python_runtime import reexec_configured_provider_python
        reexec_configured_provider_python()
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
