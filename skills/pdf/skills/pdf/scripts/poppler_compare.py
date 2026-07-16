#!/usr/bin/env python3
"""Render and compare a manifest-driven PDF merge with objective Poppler evidence."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import shutil
import subprocess
import sys


MANIFEST_SCHEMA = "open-office-artifact-tool.pdf-merge-stamp.v1"
REPORT_SCHEMA = "open-office-artifact-tool.pdf-poppler-compare.v1"


class CompareError(RuntimeError):
    pass


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def render_pdf(source: Path, directory: Path, label: str, dpi: int, timeout: int) -> list[Path]:
    command = shutil.which("pdftoppm")
    if not command:
        raise CompareError("pdftoppm is unavailable; Poppler comparison cannot run")
    directory.mkdir(parents=True, exist_ok=True)
    prefix = directory / label
    for stale in directory.glob(f"{label}-*.png"):
        stale.unlink()
    result = subprocess.run(
        [command, "-png", "-r", str(dpi), str(source), str(prefix)],
        check=False,
        capture_output=True,
        text=True,
        timeout=timeout,
    )
    if result.returncode != 0:
        raise CompareError(f"pdftoppm failed for {source}: {(result.stderr or result.stdout).strip()[:500]}")
    pages = sorted(
        directory.glob(f"{label}-*.png"),
        key=lambda value: int(re.search(r"-(\d+)\.png$", value.name).group(1)),
    )
    if not pages:
        raise CompareError(f"pdftoppm produced no pages for {source}")
    return pages


def image_evidence(source_path: Path, output_path: Path) -> dict:
    try:
        from PIL import Image, ImageChops
    except ImportError as exc:
        raise CompareError("Pillow is required for Poppler pixel comparison") from exc
    with Image.open(source_path) as source_image, Image.open(output_path) as output_image:
        source = source_image.convert("RGB")
        output = output_image.convert("RGB")
        same_dimensions = source.size == output.size
        if same_dimensions:
            difference = ImageChops.difference(source, output)
            difference_box = difference.getbbox()
            changed_bbox = list(difference_box) if difference_box else None
        else:
            changed_bbox = None
        source_histogram = source.convert("L").histogram()
        output_histogram = output.convert("L").histogram()
        source_pixels = source.width * source.height
        output_pixels = output.width * output.height
        source_dark_ratio = sum(source_histogram[:10]) / source_pixels
        output_dark_ratio = sum(output_histogram[:10]) / output_pixels
        return {
            "sourcePixels": [source.width, source.height],
            "outputPixels": [output.width, output.height],
            "sameDimensions": same_dimensions,
            "sourceNonBlank": sum(source_histogram[:250]) > 0,
            "outputNonBlank": sum(output_histogram[:250]) > 0,
            "pixelStable": same_dimensions and changed_bbox is None,
            "changedPixelsBBox": changed_bbox,
            "sourceDarkRatio": round(source_dark_ratio, 8),
            "outputDarkRatio": round(output_dark_ratio, 8),
            "darkRatioDelta": round(output_dark_ratio - source_dark_ratio, 8),
        }


def load_manifest(path: Path, max_manifest_bytes: int) -> dict:
    if not path.is_file():
        raise CompareError("manifest must be an existing JSON file")
    if path.stat().st_size > max_manifest_bytes:
        raise CompareError(f"manifest exceeds max-manifest-bytes {max_manifest_bytes}")
    try:
        value = json.loads(path.read_text("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise CompareError(f"manifest is not valid UTF-8 JSON: {exc}") from exc
    if value.get("schema") != MANIFEST_SCHEMA:
        raise CompareError(f"manifest schema must be {MANIFEST_SCHEMA!r}")
    return value


def compare_merge_stamp(args: argparse.Namespace) -> dict:
    manifest_path = args.manifest.expanduser().resolve()
    output_path = args.output.expanduser().resolve()
    report_path = args.report.expanduser().resolve()
    render_root = args.render_dir.expanduser().resolve()
    manifest = load_manifest(manifest_path, args.max_manifest_bytes)
    if not output_path.is_file():
        raise CompareError("output must be an existing PDF")
    if report_path in {manifest_path, output_path}:
        raise CompareError("report path must differ from manifest and PDF output")
    if not 1 <= args.dpi <= 600 or args.max_pages < 1 or args.timeout < 1 or not 0 <= args.max_dark_ratio_delta <= 1:
        raise CompareError("dpi must be 1..600, max-pages/timeout must be positive, and max-dark-ratio-delta must be 0..1")

    source_records = manifest.get("sources")
    if not isinstance(source_records, list) or not source_records:
        raise CompareError("manifest sources must be a non-empty array")
    sources: dict[str, dict] = {}
    before_hashes: dict[str, str] = {}
    for record in source_records:
        if not isinstance(record, dict):
            raise CompareError("every source must be an object")
        source_id = str(record.get("id", "")).strip()
        if not source_id or source_id in sources:
            raise CompareError("source IDs must be non-empty and unique")
        source_path = Path(str(record.get("path", ""))).expanduser()
        if not source_path.is_absolute():
            source_path = manifest_path.parent / source_path
        source_path = source_path.resolve()
        if not source_path.is_file():
            raise CompareError(f"source {source_id!r} is unavailable")
        sources[source_id] = {"path": source_path}
        before_hashes[source_id] = sha256(source_path)

    output_pages = render_pdf(output_path, render_root / "output", "page", args.dpi, args.timeout)
    if len(output_pages) > args.max_pages:
        raise CompareError(f"output has {len(output_pages)} pages; max-pages is {args.max_pages}")
    for source_id, source in sources.items():
        source["pages"] = render_pdf(source["path"], render_root / "sources" / source_id, "page", args.dpi, args.timeout)

    sequence = manifest.get("sequence")
    if not isinstance(sequence, list) or not sequence:
        raise CompareError("manifest sequence must be a non-empty array")
    page_map: list[tuple[str, int]] = []
    for segment in sequence:
        if not isinstance(segment, dict) or segment.get("source") not in sources:
            raise CompareError("every sequence segment must name a declared source")
        source_id = segment["source"]
        page_count = len(sources[source_id]["pages"])
        pages = list(range(1, page_count + 1)) if segment.get("pages") == "all" else segment.get("pages")
        if not isinstance(pages, list) or not pages or any(not isinstance(page, int) or page < 1 or page > page_count for page in pages):
            raise CompareError(f"sequence pages for {source_id!r} are invalid")
        page_map.extend((source_id, page) for page in pages)
    if len(page_map) != len(output_pages):
        raise CompareError(f"manifest selects {len(page_map)} pages but output has {len(output_pages)}")

    watermark_sources = {
        str(rule.get("source"))
        for rule in manifest.get("watermarks", [])
        if isinstance(rule, dict) and rule.get("source") in sources
    }
    if not watermark_sources:
        raise CompareError("manifest must contain at least one watermark rule for a declared source")

    page_evidence = []
    failures = []
    for output_number, (source_id, source_number) in enumerate(page_map, 1):
        evidence = image_evidence(sources[source_id]["pages"][source_number - 1], output_pages[output_number - 1])
        watermark_expected = source_id in watermark_sources
        evidence.update({"page": output_number, "source": source_id, "sourcePage": source_number, "watermarkExpected": watermark_expected})
        page_failures = []
        if not evidence["sameDimensions"]:
            page_failures.append("pixel dimensions changed")
        if evidence["sourceNonBlank"] and not evidence["outputNonBlank"]:
            page_failures.append("source content became blank")
        if not watermark_expected and evidence["sourceNonBlank"] != evidence["outputNonBlank"]:
            page_failures.append("non-watermarked blank/non-blank state changed")
        if evidence["darkRatioDelta"] > args.max_dark_ratio_delta:
            page_failures.append("dark-pixel ratio increased beyond the declared threshold")
        if watermark_expected and evidence["pixelStable"]:
            page_failures.append("watermarked page did not change")
        if not watermark_expected and not evidence["pixelStable"]:
            page_failures.append("non-watermarked page changed")
        evidence["passed"] = not page_failures
        evidence["failures"] = page_failures
        failures.extend(f"page {output_number}: {failure}" for failure in page_failures)
        page_evidence.append(evidence)

    after_hashes = {source_id: sha256(source["path"]) for source_id, source in sources.items()}
    if before_hashes != after_hashes:
        failures.append("one or more source PDFs changed during rendering")
    payload = {
        "schema": REPORT_SCHEMA,
        "status": "passed" if not failures else "failed",
        "renderer": {"name": "pdftoppm", "dpi": args.dpi},
        "manifest": {"path": str(manifest_path), "sha256": sha256(manifest_path)},
        "output": {"path": str(output_path), "sha256": sha256(output_path), "pages": len(output_pages)},
        "sources": [
            {"id": source_id, "path": str(source["path"]), "sha256": after_hashes[source_id], "pages": len(source["pages"])}
            for source_id, source in sources.items()
        ],
        "policy": {
            "nonWatermarkedPagesPixelStable": True,
            "watermarkedPagesChanged": True,
            "maxDarkRatioDelta": args.max_dark_ratio_delta,
        },
        "pages": page_evidence,
        "failures": failures,
    }
    report_path.parent.mkdir(parents=True, exist_ok=True)
    temporary = report_path.with_name(f".{report_path.name}.tmp")
    temporary.write_text(json.dumps(payload, ensure_ascii=False, indent=2) + "\n", "utf-8")
    temporary.replace(report_path)
    return payload


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)
    merge = subparsers.add_parser("merge-stamp", help="compare a merge-stamp output to its declared source pages")
    merge.add_argument("manifest", type=Path)
    merge.add_argument("output", type=Path)
    merge.add_argument("--report", type=Path, required=True)
    merge.add_argument("--render-dir", type=Path, required=True)
    merge.add_argument("--dpi", type=int, default=144)
    merge.add_argument("--max-pages", type=int, default=500)
    merge.add_argument("--max-manifest-bytes", type=int, default=1024 * 1024)
    merge.add_argument("--max-dark-ratio-delta", type=float, default=0.25)
    merge.add_argument("--timeout", type=int, default=120)
    args = parser.parse_args()
    try:
        payload = compare_merge_stamp(args)
        print(json.dumps(payload, ensure_ascii=False, indent=2))
        return 0 if payload["status"] == "passed" else 2
    except (CompareError, OSError, subprocess.SubprocessError) as exc:
        print(json.dumps({"schema": REPORT_SCHEMA, "status": "error", "error": str(exc)}, ensure_ascii=False), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
