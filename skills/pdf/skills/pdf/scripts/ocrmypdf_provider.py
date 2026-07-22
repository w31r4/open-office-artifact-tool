#!/usr/bin/env python3
"""Add a searchable text layer through a bounded, source-bound OCRmyPDF route."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
from pathlib import Path
import re
import shutil
import signal
import stat
import subprocess
import sys
import tempfile
import threading
import time
from typing import Any

import qpdf_provider as qpdf


SCHEMA = "open-office-artifact-tool.ocrmypdf-ocr.v1"
MINIMUM_VERSION = (17, 8, 0)
MAXIMUM_VERSION_EXCLUSIVE = (17, 9, 0)
MINIMUM_TESSERACT_VERSION = (5, 0, 0)
MAXIMUM_TESSERACT_VERSION_EXCLUSIVE = (6, 0, 0)
HARD_MAX_INPUT_BYTES = 512 * 1024 * 1024
DEFAULT_MAX_INPUT_BYTES = HARD_MAX_INPUT_BYTES
HARD_MAX_OUTPUT_BYTES = 1024 * 1024 * 1024
DEFAULT_MAX_OUTPUT_BYTES = HARD_MAX_OUTPUT_BYTES
HARD_MAX_SIDECAR_BYTES = 64 * 1024 * 1024
DEFAULT_MAX_SIDECAR_BYTES = 16 * 1024 * 1024
HARD_MAX_EXTRACTED_TEXT_BYTES = 128 * 1024 * 1024
DEFAULT_MAX_EXTRACTED_TEXT_BYTES = 32 * 1024 * 1024
HARD_MAX_TIMEOUT_SECONDS = 1800
DEFAULT_TIMEOUT_SECONDS = 600
HARD_MAX_TESSERACT_TIMEOUT_SECONDS = 300
DEFAULT_TESSERACT_TIMEOUT_SECONDS = 60
HARD_MAX_STDOUT_BYTES = 2 * 1024 * 1024
DEFAULT_MAX_STDOUT_BYTES = 256 * 1024
HARD_MAX_STDERR_BYTES = 2 * 1024 * 1024
DEFAULT_MAX_STDERR_BYTES = 512 * 1024
HARD_MAX_IMAGE_MPIXELS = 256
DEFAULT_MAX_IMAGE_MPIXELS = 128
MAX_REQUIRE_TEXT = 32
MAX_REQUIRE_TEXT_CHARS = 256
MAX_LANGUAGE_COUNT = 16
MAX_TESSDATA_DIRECTORIES = 16
MAX_TESSDATA_FILE_BYTES = 128 * 1024 * 1024
MAX_DIAGNOSTIC_LINES = 200
MAX_DIAGNOSTIC_LINE_CHARS = 1_000
MAX_TEXT_PREVIEW_CHARS = 4_096


class ProviderError(RuntimeError):
    pass


def bounded_text(value: Any, limit: int = MAX_DIAGNOSTIC_LINE_CHARS) -> str:
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
        return path.stat().st_size, sha256(path)
    except OSError as exc:
        raise ProviderError(f"{label} became unavailable during OCR: {bounded_text(exc)}") from exc


def positive_int(value: str) -> int:
    parsed = int(value)
    if parsed <= 0:
        raise argparse.ArgumentTypeError("must be a positive integer")
    return parsed


def checked_budget(name: str, value: int, hard_maximum: int) -> int:
    if value > hard_maximum:
        raise ProviderError(f"{name} cannot exceed the hard maximum {hard_maximum}")
    return value


def expected_hash(value: str) -> str:
    normalized = value.strip().lower()
    if not re.fullmatch(r"[0-9a-f]{64}", normalized):
        raise ProviderError("--expected-sha256 must be exactly 64 hexadecimal characters")
    return normalized


def executable_path(environment: str, command: str, label: str) -> Path:
    configured = os.environ.get(environment, "").strip()
    candidate = Path(configured).expanduser() if configured else Path(shutil.which(command) or "")
    if not str(candidate) or not candidate.is_file() or not os.access(candidate, os.X_OK):
        detail = f" configured by {environment}={configured!r}" if configured else " on PATH"
        raise ProviderError(f"{label} executable is unavailable{detail}")
    if not candidate.is_absolute():
        candidate = Path.cwd() / candidate
    return Path(os.path.abspath(candidate))


def managed_tessdata_stream(path: Path):
    """Open one managed language data file without ever following a link."""

    nofollow = getattr(os, "O_NOFOLLOW", None)
    if nofollow is None:  # Managed packs are intentionally unsupported on Windows for now.
        raise ProviderError("managed tessdata requires a platform with O_NOFOLLOW support")
    try:
        descriptor = os.open(path, os.O_RDONLY | nofollow)
    except OSError as exc:
        raise ProviderError(f"managed tessdata entry could not be opened safely: {path}") from exc
    try:
        metadata = os.fstat(descriptor)
        if not stat.S_ISREG(metadata.st_mode) or metadata.st_nlink != 1 or metadata.st_size <= 0 or metadata.st_size > MAX_TESSDATA_FILE_BYTES:
            raise ProviderError(f"managed tessdata entry is unsafe or outside the size budget: {path}")
        return os.fdopen(descriptor, "rb"), metadata.st_size
    except Exception:
        os.close(descriptor)
        raise


def managed_tessdata_digest(path: Path) -> str:
    stream, _ = managed_tessdata_stream(path)
    try:
        digest = hashlib.sha256()
        for chunk in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(chunk)
        return digest.hexdigest()
    finally:
        stream.close()


def copy_managed_tessdata(source: Path, target: Path) -> None:
    stream, expected_bytes = managed_tessdata_stream(source)
    copied_bytes = 0
    try:
        with target.open("xb") as destination:
            for chunk in iter(lambda: stream.read(1024 * 1024), b""):
                copied_bytes += len(chunk)
                if copied_bytes > MAX_TESSDATA_FILE_BYTES:
                    raise ProviderError(f"managed tessdata entry is outside the size budget: {source}")
                destination.write(chunk)
            destination.flush()
            os.fsync(destination.fileno())
    except FileExistsError as exc:
        raise ProviderError(f"managed tessdata has an unexpected duplicate destination: {target.name}") from exc
    finally:
        stream.close()
    if copied_bytes != expected_bytes:
        raise ProviderError(f"managed tessdata entry changed while being copied: {source}")
    os.chmod(target, 0o400)


def managed_tessdata_directory(private_root: Path) -> Path | None:
    """Materialize a bounded, private language union for managed OCR packs.

    A managed language pack has its own immutable cache root, while Tesseract
    accepts only one tessdata directory. The caller may provide a path-list of
    verified managed data directories through OPEN_OFFICE_PDF_TESSDATA_DIRS;
    this adapter never follows their links or mutates their receipts. It copies
    regular traineddata files into the per-operation private root and exposes
    that one union to Tesseract. A missing variable preserves the selected
    system runtime's normal Tesseract data discovery.
    """

    raw = os.environ.get("OPEN_OFFICE_PDF_TESSDATA_DIRS", "").strip()
    if not raw:
        return None
    source_texts = [value.strip() for value in raw.split(os.pathsep) if value.strip()]
    if not source_texts or len(source_texts) > MAX_TESSDATA_DIRECTORIES:
        raise ProviderError(f"OPEN_OFFICE_PDF_TESSDATA_DIRS must contain 1..{MAX_TESSDATA_DIRECTORIES} directories")
    destination = private_root / "tessdata"
    destination.mkdir(mode=0o700, exist_ok=False)
    copied: set[str] = set()
    for source_text in source_texts:
        source = Path(source_text).expanduser()
        try:
            source_stat = source.lstat()
        except OSError as exc:
            raise ProviderError(f"managed tessdata directory is unavailable: {source_text}") from exc
        if stat.S_ISLNK(source_stat.st_mode) or not stat.S_ISDIR(source_stat.st_mode):
            raise ProviderError(f"managed tessdata directory must be a real directory: {source_text}")
        for child in sorted(source.iterdir(), key=lambda entry: entry.name):
            if not re.fullmatch(r"[A-Za-z0-9_-]+\.traineddata", child.name):
                continue
            try:
                child_stat = child.lstat()
            except OSError as exc:
                raise ProviderError(f"managed tessdata entry became unavailable: {child}") from exc
            if stat.S_ISLNK(child_stat.st_mode) or not stat.S_ISREG(child_stat.st_mode) or child_stat.st_nlink != 1 or child_stat.st_size <= 0 or child_stat.st_size > MAX_TESSDATA_FILE_BYTES:
                raise ProviderError(f"managed tessdata entry is unsafe or outside the size budget: {child}")
            target = destination / child.name
            if child.name in copied:
                if managed_tessdata_digest(child) != managed_tessdata_digest(target):
                    raise ProviderError(f"managed tessdata has conflicting duplicate language data: {child.name}")
                continue
            copy_managed_tessdata(child, target)
            copied.add(child.name)
    if not copied:
        raise ProviderError("managed tessdata directories contain no regular traineddata files")
    return destination


def provider_environment(private_root: Path, executables: list[Path]) -> dict[str, str]:
    home = private_root / "home"
    temporary = private_root / "tmp"
    cache = private_root / "cache"
    config = private_root / "config"
    for directory in (home, temporary, cache, config):
        directory.mkdir(mode=0o700, exist_ok=True)
    path_entries: list[str] = []
    for entry in [*(str(executable.parent) for executable in executables), "/usr/bin", "/bin", "/usr/sbin", "/sbin"]:
        if entry not in path_entries and Path(entry).is_dir():
            path_entries.append(entry)
    environment = {
        "HOME": str(home),
        "TMPDIR": str(temporary),
        "TMP": str(temporary),
        "TEMP": str(temporary),
        "XDG_CACHE_HOME": str(cache),
        "XDG_CONFIG_HOME": str(config),
        "PATH": os.pathsep.join(path_entries),
        "LANG": "C.UTF-8",
        "LC_ALL": "C.UTF-8",
        "NO_COLOR": "1",
        "OMP_THREAD_LIMIT": "1",
        "OMP_NUM_THREADS": "1",
        "PYTHONDONTWRITEBYTECODE": "1",
    }
    for name in ("SYSTEMROOT", "WINDIR", "PATHEXT", "COMSPEC"):
        if os.environ.get(name):
            environment[name] = os.environ[name]
    tessdata = managed_tessdata_directory(private_root)
    if tessdata is not None:
        environment["TESSDATA_PREFIX"] = str(tessdata)
    return environment


def terminate_process_tree(process: subprocess.Popen) -> None:
    try:
        if os.name != "nt":
            os.killpg(process.pid, signal.SIGKILL)
        else:  # pragma: no cover - Windows CI is a future validation target
            process.kill()
    except (OSError, ProcessLookupError):
        pass


def run_bounded(
    executable: Path,
    arguments: list[str],
    *,
    label: str,
    timeout_seconds: int,
    max_stdout_bytes: int,
    max_stderr_bytes: int,
    environment: dict[str, str],
    watched_files: list[tuple[Path, int, str]] | None = None,
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
            env=environment,
            start_new_session=os.name != "nt",
        )
    except OSError as exc:
        raise ProviderError(f"{label} could not start: {bounded_text(exc)}") from exc

    def record_violation(message: str) -> None:
        with lock:
            if not violations:
                violations.append(message)
        terminate_process_tree(process)

    def pump(stream: Any, buffer: bytearray, limit: int, stream_label: str) -> None:
        try:
            while True:
                chunk = stream.read(64 * 1024)
                if not chunk:
                    break
                if len(buffer) + len(chunk) > limit:
                    record_violation(f"{label} {stream_label} exceeded the {limit} byte budget")
                    break
                buffer.extend(chunk)
        except OSError as exc:
            record_violation(f"{label} {stream_label} capture failed: {bounded_text(exc)}")
        finally:
            stream.close()

    stdout_thread = threading.Thread(target=pump, args=(process.stdout, stdout, max_stdout_bytes, "stdout"), daemon=True)
    stderr_thread = threading.Thread(target=pump, args=(process.stderr, stderr, max_stderr_bytes, "stderr"), daemon=True)
    stdout_thread.start()
    stderr_thread.start()
    deadline = time.monotonic() + timeout_seconds
    while process.poll() is None:
        if time.monotonic() >= deadline:
            record_violation(f"{label} timed out after {timeout_seconds} seconds")
            break
        for path, maximum, watched_label in watched_files or []:
            try:
                size = path.stat().st_size
            except FileNotFoundError:
                continue
            except OSError as exc:
                record_violation(f"could not inspect {watched_label} during OCR: {bounded_text(exc)}")
                break
            if size > maximum:
                record_violation(f"{watched_label} exceeded the {maximum} byte budget")
                break
        time.sleep(0.05)
    try:
        process.wait(timeout=5)
    except subprocess.TimeoutExpired:
        terminate_process_tree(process)
        process.wait()
    stdout_thread.join()
    stderr_thread.join()
    if violations:
        raise ProviderError(violations[0])
    return subprocess.CompletedProcess(command, process.returncode, bytes(stdout), bytes(stderr))


def decoded(value: bytes) -> str:
    return value.decode("utf-8", "replace")


def semantic_version(value: str, label: str) -> tuple[int, int, int]:
    match = re.search(r"(?<!\d)(\d+)\.(\d+)\.(\d+)", value)
    if not match:
        raise ProviderError(f"{label} returned an unsupported version response: {bounded_text(value)}")
    return tuple(int(part) for part in match.groups())


def version_text(value: tuple[int, int, int]) -> str:
    return ".".join(str(part) for part in value)


def bounded_command(
    executable: Path,
    arguments: list[str],
    *,
    label: str,
    environment: dict[str, str],
    timeout_seconds: int,
    stdout_limit: int = 64 * 1024,
    stderr_limit: int = 64 * 1024,
) -> str:
    result = run_bounded(
        executable,
        arguments,
        label=label,
        timeout_seconds=timeout_seconds,
        max_stdout_bytes=stdout_limit,
        max_stderr_bytes=stderr_limit,
        environment=environment,
    )
    output = "\n".join(part for part in (decoded(result.stdout).strip(), decoded(result.stderr).strip()) if part)
    if result.returncode != 0:
        raise ProviderError(f"{label} failed with exit {result.returncode}: {bounded_text(output)}")
    return output


def component_probe(timeout_seconds: int, private_root: Path) -> dict[str, Any]:
    ocrmypdf = executable_path("OPEN_OFFICE_PDF_OCRMYPDF", "ocrmypdf", "OCRmyPDF")
    tesseract = executable_path("OPEN_OFFICE_PDF_TESSERACT", "tesseract", "Tesseract")
    pdftotext = executable_path("OPEN_OFFICE_PDF_PDFTOTEXT", "pdftotext", "Poppler pdftotext")
    ghostscript = executable_path("OPEN_OFFICE_PDF_GS", "gs", "Ghostscript")
    qpdf_executable = qpdf.qpdf_path()
    environment = provider_environment(private_root, [ocrmypdf, tesseract, pdftotext, ghostscript, qpdf_executable])

    ocr_output = bounded_command(ocrmypdf, ["--version"], label="OCRmyPDF --version", environment=environment, timeout_seconds=timeout_seconds)
    ocr_version = semantic_version(ocr_output, "OCRmyPDF")
    if not MINIMUM_VERSION <= ocr_version < MAXIMUM_VERSION_EXCLUSIVE:
        raise ProviderError(
            f"OCRmyPDF >= {version_text(MINIMUM_VERSION)} and < {version_text(MAXIMUM_VERSION_EXCLUSIVE)} "
            f"is required; received {version_text(ocr_version)}"
        )

    tesseract_output = bounded_command(tesseract, ["--version"], label="Tesseract --version", environment=environment, timeout_seconds=timeout_seconds)
    tesseract_version = semantic_version(tesseract_output, "Tesseract")
    if not MINIMUM_TESSERACT_VERSION <= tesseract_version < MAXIMUM_TESSERACT_VERSION_EXCLUSIVE:
        raise ProviderError(
            f"Tesseract >= {version_text(MINIMUM_TESSERACT_VERSION)} and "
            f"< {version_text(MAXIMUM_TESSERACT_VERSION_EXCLUSIVE)} is required; "
            f"received {version_text(tesseract_version)}"
        )
    language_output = bounded_command(
        tesseract,
        ["--list-langs"],
        label="Tesseract --list-langs",
        environment=environment,
        timeout_seconds=timeout_seconds,
        stdout_limit=1024 * 1024,
    )
    language_lines = [line.strip() for line in language_output.splitlines() if line.strip()]
    languages = [line for line in language_lines if not line.lower().startswith("list of available languages")]
    if not languages or any(len(language) > 128 for language in languages):
        raise ProviderError("Tesseract did not return a bounded non-empty language catalog")

    poppler_output = bounded_command(pdftotext, ["-v"], label="Poppler pdftotext -v", environment=environment, timeout_seconds=timeout_seconds)
    poppler_version = semantic_version(poppler_output, "Poppler pdftotext")
    ghostscript_output = bounded_command(ghostscript, ["--version"], label="Ghostscript --version", environment=environment, timeout_seconds=timeout_seconds)
    ghostscript_version = semantic_version(ghostscript_output, "Ghostscript")
    qpdf_version = qpdf.qpdf_version(qpdf_executable, timeout_seconds)
    return {
        "ocrmypdf": ocrmypdf,
        "tesseract": tesseract,
        "pdftotext": pdftotext,
        "ghostscript": ghostscript,
        "qpdf": qpdf_executable,
        "environment": environment,
        "versions": {
            "ocrmypdf": version_text(ocr_version),
            "tesseract": version_text(tesseract_version),
            "popplerPdftotext": version_text(poppler_version),
            "ghostscript": version_text(ghostscript_version),
            "qpdf": qpdf_version,
        },
        "languages": sorted(set(languages)),
    }


def normalized_languages(requested: list[str], available: list[str]) -> list[str]:
    if not requested:
        requested = ["eng"]
    if len(requested) > MAX_LANGUAGE_COUNT:
        raise ProviderError(f"at most {MAX_LANGUAGE_COUNT} OCR languages may be selected")
    selected: list[str] = []
    for value in requested:
        language = value.strip()
        if not re.fullmatch(r"(?:[A-Za-z0-9_-]+/)?[A-Za-z0-9_-]+", language):
            raise ProviderError(f"invalid Tesseract language identifier: {value!r}")
        if language not in available:
            raise ProviderError(f"Tesseract language {language!r} is unavailable; installed languages: {', '.join(available)}")
        if language not in selected:
            selected.append(language)
    return selected


def normalized_required_text(values: list[str]) -> list[str]:
    if len(values) > MAX_REQUIRE_TEXT:
        raise ProviderError(f"at most {MAX_REQUIRE_TEXT} --require-text gates may be selected")
    result: list[str] = []
    for value in values:
        text = " ".join(value.split())
        if not text or len(text) > MAX_REQUIRE_TEXT_CHARS:
            raise ProviderError(f"each --require-text value must contain 1..{MAX_REQUIRE_TEXT_CHARS} characters")
        result.append(text)
    return result


def normalized_search_text(value: str) -> str:
    return " ".join(value.split()).casefold()


def source_prefix_retained(source: Path, output: Path) -> bool:
    if output.stat().st_size < source.stat().st_size:
        return False
    with source.open("rb") as left, output.open("rb") as right:
        while True:
            expected = left.read(1024 * 1024)
            if not expected:
                return True
            if right.read(len(expected)) != expected:
                return False


def diagnostics(result: subprocess.CompletedProcess, replacements: dict[str, str]) -> list[str]:
    combined = "\n".join(part for part in (decoded(result.stdout).strip(), decoded(result.stderr).strip()) if part)
    for original, replacement in replacements.items():
        combined = combined.replace(original, replacement)
    lines = [bounded_text(line.rstrip()) for line in combined.splitlines() if line.strip()]
    return lines[-MAX_DIAGNOSTIC_LINES:]


def text_evidence(value: bytes) -> dict[str, Any]:
    text = value.decode("utf-8", "replace")
    pages = text.split("\f")
    if pages and not pages[-1].strip():
        pages.pop()
    return {
        "bytes": len(value),
        "sha256": hashlib.sha256(value).hexdigest(),
        "characters": len(text),
        "nonWhitespaceCharacters": sum(1 for character in text if not character.isspace()),
        "pageCharacterCounts": [len(page.strip()) for page in pages],
        "preview": text[:MAX_TEXT_PREVIEW_CHARS],
        "previewTruncated": len(text) > MAX_TEXT_PREVIEW_CHARS,
    }


def inspect_snapshot(
    target: Path,
    *,
    components: dict[str, Any],
    expected_sha256: str | None,
    timeout_seconds: int,
    display_path: Path,
) -> dict[str, Any]:
    report = qpdf.inspect_pdf(
        target,
        executable=components["qpdf"],
        version=components["versions"]["qpdf"],
        expected_sha256=expected_sha256,
        max_json_bytes=qpdf.DEFAULT_MAX_JSON_BYTES,
        timeout_seconds=min(timeout_seconds, qpdf.DEFAULT_TIMEOUT_SECONDS),
        display_path=display_path,
    )
    report["source"]["path"] = str(display_path)
    return report


def validate_mode(args: argparse.Namespace, before: dict[str, Any]) -> list[str]:
    warnings: list[str] = []
    tagged = bool(before["structure"].get("tagged"))
    form_fields = int(before["structure"].get("formFieldCount") or 0)
    annotations = int(before["structure"].get("annotationCount") or 0)
    if args.mode in {"redo", "force"} and not args.allow_structure_loss:
        raise ProviderError(f"--mode {args.mode} requires --allow-structure-loss because OCRmyPDF discards incompatible structure markup")
    if tagged and not args.allow_structure_loss:
        raise ProviderError("Tagged PDF input requires --allow-structure-loss before OCR because newly recognized text cannot be mapped to the existing structure tree")
    if args.mode == "force" and not args.allow_rasterize_all:
        raise ProviderError("--mode force requires --allow-rasterize-all because every page is rasterized")
    if args.mode == "force" and (form_fields or annotations) and not args.allow_interactive_flattening:
        raise ProviderError("--mode force on a PDF with form fields or annotations requires --allow-interactive-flattening")
    if args.allow_interactive_flattening and args.mode != "force":
        raise ProviderError("--allow-interactive-flattening is valid only with --mode force")
    if args.allow_rasterize_all and args.mode != "force":
        raise ProviderError("--allow-rasterize-all is valid only with --mode force")
    if args.allow_structure_loss and not tagged and args.mode == "skip":
        warnings.append("structure-loss acknowledgement was supplied although the source is not tagged")
    if args.mode == "force":
        warnings.append("force mode rasterizes all visible page content and may flatten interactive objects")
    elif args.mode == "redo":
        warnings.append("redo mode replaces hidden OCR text and discards incompatible tagged structure")
    return warnings


def fsync_directory(directory: Path) -> None:
    try:
        descriptor = os.open(directory, os.O_RDONLY)
    except OSError:
        return
    try:
        os.fsync(descriptor)
    except OSError:
        pass
    finally:
        os.close(descriptor)


def ocr_pdf(args: argparse.Namespace) -> dict[str, Any]:
    max_input_bytes = checked_budget("--max-input-bytes", args.max_input_bytes, HARD_MAX_INPUT_BYTES)
    max_output_bytes = checked_budget("--max-output-bytes", args.max_output_bytes, HARD_MAX_OUTPUT_BYTES)
    max_sidecar_bytes = checked_budget("--max-sidecar-bytes", args.max_sidecar_bytes, HARD_MAX_SIDECAR_BYTES)
    max_text_bytes = checked_budget("--max-extracted-text-bytes", args.max_extracted_text_bytes, HARD_MAX_EXTRACTED_TEXT_BYTES)
    timeout_seconds = checked_budget("--timeout-seconds", args.timeout_seconds, HARD_MAX_TIMEOUT_SECONDS)
    tesseract_timeout = checked_budget(
        "--tesseract-timeout-seconds",
        args.tesseract_timeout_seconds,
        HARD_MAX_TESSERACT_TIMEOUT_SECONDS,
    )
    max_stdout_bytes = checked_budget("--max-stdout-bytes", args.max_stdout_bytes, HARD_MAX_STDOUT_BYTES)
    max_stderr_bytes = checked_budget("--max-stderr-bytes", args.max_stderr_bytes, HARD_MAX_STDERR_BYTES)
    max_image_mpixels = checked_budget("--max-image-mpixels", args.max_image_mpixels, HARD_MAX_IMAGE_MPIXELS)
    expected = expected_hash(args.expected_sha256)
    required_text = normalized_required_text(args.require_text)
    source = qpdf.input_path(args.input, max_input_bytes)
    destination = qpdf.output_path(args.output)
    if source == destination:
        raise ProviderError("input and output must be distinct; OCR never overwrites the source")
    source_bytes, source_hash = file_identity(source, "source PDF")
    if source_hash != expected:
        raise ProviderError(f"source SHA-256 mismatch: expected {expected}, received {source_hash}")

    with tempfile.TemporaryDirectory(prefix=f".{destination.name}.ocrmypdf-", dir=destination.parent) as temporary:
        private_root = Path(temporary)
        components = component_probe(min(timeout_seconds, 30), private_root)
        languages = normalized_languages(args.language, components["languages"])
        snapshot = private_root / "source.pdf"
        candidate = private_root / "candidate.pdf"
        sidecar = private_root / "ocr-sidecar.txt"
        shutil.copyfile(source, snapshot)
        with snapshot.open("rb+") as stream:
            os.fsync(stream.fileno())
        os.chmod(snapshot, 0o400)
        if sha256(snapshot) != expected:
            raise ProviderError("private OCR snapshot does not match the expected source SHA-256")

        before = inspect_snapshot(
            snapshot,
            components=components,
            expected_sha256=expected,
            timeout_seconds=timeout_seconds,
            display_path=source,
        )
        if before["structure"]["encrypted"]:
            raise ProviderError("encrypted input is unsupported; select an explicit decryption workflow before OCR")
        signature = before["signaturePolicy"]
        if signature["hasSignatureEvidence"] and not args.invalidate_signatures:
            raise ProviderError("signed or signature-constrained input requires --invalidate-signatures after pyHanko and DocMDP/FieldMDP review")
        warnings = validate_mode(args, before)

        ocr_arguments = [
            "--mode", args.mode,
            "--output-type", "pdf",
            "--optimize", "0",
            "--jobs", "1",
            "--ocr-engine", "tesseract",
            "--rasterizer", "pypdfium",
            "--pdf-renderer", "fpdf2",
            "--language", "+".join(languages),
            "--tesseract-timeout", str(tesseract_timeout),
            "--max-image-mpixels", str(max_image_mpixels),
            "--sidecar", str(sidecar),
            "--tagged-pdf-mode", "ignore" if before["structure"].get("tagged") else "default",
            "--no-overwrite",
            "--quiet",
        ]
        if signature["hasSignatureEvidence"] and args.invalidate_signatures:
            ocr_arguments.append("--invalidate-digital-signatures")
        ocr_arguments.extend([str(snapshot), str(candidate)])
        result = run_bounded(
            components["ocrmypdf"],
            ocr_arguments,
            label="OCRmyPDF",
            timeout_seconds=timeout_seconds,
            max_stdout_bytes=max_stdout_bytes,
            max_stderr_bytes=max_stderr_bytes,
            environment=components["environment"],
            watched_files=[
                (candidate, max_output_bytes, "OCR output PDF"),
                (sidecar, max_sidecar_bytes, "OCR sidecar"),
            ],
        )
        diagnostic_lines = diagnostics(result, {str(snapshot): str(source), str(candidate): str(destination), str(sidecar): "<private-sidecar>"})
        if result.returncode != 0:
            raise ProviderError(f"OCRmyPDF failed with exit {result.returncode}: {bounded_text(' | '.join(diagnostic_lines[-20:]), 4_096)}")
        if not candidate.is_file() or not sidecar.is_file():
            raise ProviderError("OCRmyPDF completed without both the expected PDF and private sidecar outputs")
        candidate_bytes = candidate.stat().st_size
        sidecar_bytes = sidecar.stat().st_size
        if candidate_bytes < 5 or candidate_bytes > max_output_bytes:
            raise ProviderError(f"OCR output size {candidate_bytes} is outside the 5..{max_output_bytes} byte budget")
        if sidecar_bytes > max_sidecar_bytes:
            raise ProviderError(f"OCR sidecar exceeded the {max_sidecar_bytes} byte budget")
        with candidate.open("rb") as stream:
            if stream.read(5) != b"%PDF-":
                raise ProviderError("OCRmyPDF output does not begin with a PDF header")
        if sha256(snapshot) != expected:
            raise ProviderError("OCRmyPDF changed the private read-only source snapshot")
        if source_prefix_retained(snapshot, candidate):
            raise ProviderError("OCR output retained the complete source byte prefix; a full rewrite was required")

        candidate_hash = sha256(candidate)
        after = inspect_snapshot(
            candidate,
            components=components,
            expected_sha256=candidate_hash,
            timeout_seconds=timeout_seconds,
            display_path=destination,
        )
        if after["check"]["status"] != "clean":
            raise ProviderError("OCR output contains qpdf structural warnings")
        if before["structure"]["pageCount"] != after["structure"]["pageCount"]:
            raise ProviderError(
                f"OCR changed pageCount: {before['structure']['pageCount']} -> {after['structure']['pageCount']}"
            )
        for key in ("attachmentCount", "outlineCount"):
            if before["structure"][key] != after["structure"][key]:
                raise ProviderError(f"OCR changed {key}: {before['structure'][key]} -> {after['structure'][key]}")
        if args.mode != "force":
            for key in ("formFieldCount", "annotationCount"):
                if before["structure"][key] != after["structure"][key]:
                    raise ProviderError(
                        f"OCR changed {key} without force-mode acknowledgement: "
                        f"{before['structure'][key]} -> {after['structure'][key]}"
                    )
        if args.mode in {"redo", "force"} and after["structure"].get("tagged"):
            raise ProviderError(f"OCRmyPDF {args.mode} output retained a stale Tagged PDF structure tree")

        extracted = run_bounded(
            components["pdftotext"],
            ["-enc", "UTF-8", str(candidate), "-"],
            label="Poppler pdftotext",
            timeout_seconds=min(timeout_seconds, 180),
            max_stdout_bytes=max_text_bytes,
            max_stderr_bytes=128 * 1024,
            environment=components["environment"],
        )
        if extracted.returncode != 0:
            raise ProviderError(f"Poppler could not extract OCR output text: {bounded_text(decoded(extracted.stderr))}")
        extracted_text = decoded(extracted.stdout)
        if not args.allow_empty_text and not any(not character.isspace() for character in extracted_text):
            raise ProviderError("OCR output contains no extractable non-whitespace text; use --allow-empty-text only for an expected blank scan")
        normalized_output = normalized_search_text(extracted_text)
        unmatched = [value for value in required_text if normalized_search_text(value) not in normalized_output]
        if unmatched:
            raise ProviderError(f"OCR output failed required text gates: {unmatched!r}")
        sidecar_value = sidecar.read_bytes()
        if sha256(candidate) != candidate_hash:
            raise ProviderError("OCR output changed during independent structure/text validation")
        final_source_bytes, final_source_hash = file_identity(source, "source PDF")
        if (final_source_bytes, final_source_hash) != (source_bytes, expected):
            raise ProviderError("source PDF changed during OCR; output was not published")

        with candidate.open("rb") as stream:
            os.fsync(stream.fileno())
        os.chmod(candidate, 0o600)
        qpdf.publish_new_file(candidate, destination)
        fsync_directory(destination.parent)

    output_bytes, output_hash = file_identity(destination, "published OCR output")
    if output_hash != candidate_hash or output_bytes != candidate_bytes:
        raise ProviderError("published OCR output identity differs from the validated candidate")
    if file_identity(source, "source PDF") != (source_bytes, expected):
        try:
            if file_identity(destination, "published OCR output") == (candidate_bytes, candidate_hash):
                destination.unlink()
                fsync_directory(destination.parent)
        except OSError:
            pass
        raise ProviderError("source PDF changed before OCR delivery completed; the new output was withdrawn")
    return {
        "schema": SCHEMA,
        "ok": True,
        "provider": {
            "name": "ocrmypdf",
            "version": components["versions"]["ocrmypdf"],
            "executable": str(components["ocrmypdf"]),
            "components": {
                "tesseract": {"version": components["versions"]["tesseract"], "executable": str(components["tesseract"])},
                "qpdf": {"version": components["versions"]["qpdf"], "executable": str(components["qpdf"])},
                "popplerPdftotext": {"version": components["versions"]["popplerPdftotext"], "executable": str(components["pdftotext"])},
                "ghostscript": {"version": components["versions"]["ghostscript"], "executable": str(components["ghostscript"])},
                "rasterizer": "pypdfium",
                "pdfRenderer": "fpdf2",
            },
        },
        "silentFallback": False,
        "operation": {
            "type": "ocr-searchable-layer",
            "scope": "complete-document",
            "mode": args.mode,
            "languages": languages,
        },
        "savePolicy": "rewrite",
        "source": {"path": str(source), "bytes": source_bytes, "sha256": expected},
        "output": {"path": str(destination), "bytes": output_bytes, "sha256": output_hash},
        "sourceProtected": True,
        "sourcePrefixRetained": False,
        "transaction": {
            "privateSourceSnapshot": True,
            "atomicDistinctOutput": True,
            "outputReplaced": False,
            "privateSidecarRetained": False,
        },
        "inputSecurity": {
            "trustAssertion": args.input_trust,
            "adapterSandboxEnforced": False,
            "providerIsSanitizer": False,
        },
        "fidelityPolicy": {
            "outputType": "pdf",
            "optimize": 0,
            "jobs": 1,
            "visualPreprocessing": False,
            "wholeDocumentRewrite": True,
            "structureLossAcknowledged": bool(args.allow_structure_loss),
            "rasterizeAllAcknowledged": bool(args.allow_rasterize_all),
            "interactiveFlatteningAcknowledged": bool(args.allow_interactive_flattening),
        },
        "signaturePolicyBefore": signature,
        "signaturePolicyAfter": after["signaturePolicy"],
        "signatureInvalidated": bool(signature["hasSignatureEvidence"] and args.invalidate_signatures),
        "structureBefore": before["structure"],
        "structureAfter": after["structure"],
        "textValidation": {
            "finalExtraction": text_evidence(extracted.stdout),
            "ocrSidecar": text_evidence(sidecar_value),
            "requiredText": required_text,
            "requiredTextMatchPolicy": "unicode-casefold-whitespace-collapsed-substring",
            "requiredTextMatched": True,
            "allowEmptyText": bool(args.allow_empty_text),
        },
        "diagnostics": diagnostic_lines,
        "warnings": warnings,
        "requiredNextGates": [
            "fresh MuPDF/qpdf inspection",
            "Poppler render every page and compare visible geometry",
            "manual OCR text and reading-order review",
            *( ["pyHanko validation must record invalidated signatures"] if signature["hasSignatureEvidence"] else [] ),
        ],
    }


def probe(args: argparse.Namespace) -> dict[str, Any]:
    timeout_seconds = checked_budget("--timeout-seconds", args.timeout_seconds, 60)
    with tempfile.TemporaryDirectory(prefix="open-office-ocrmypdf-probe-") as temporary:
        components = component_probe(timeout_seconds, Path(temporary))
    return {
        "ok": True,
        "provider": "ocrmypdf",
        "providerVersion": components["versions"]["ocrmypdf"],
        "integration": "shipped-thin-script-external-cli",
        "silentFallback": False,
        "components": {
            "tesseract": components["versions"]["tesseract"],
            "qpdf": components["versions"]["qpdf"],
            "popplerPdftotext": components["versions"]["popplerPdftotext"],
            "ghostscript": components["versions"]["ghostscript"],
            "pypdfium": "required through OCRmyPDF runtime",
            "fpdf2": "required through OCRmyPDF runtime",
        },
        "languages": components["languages"],
        "modes": ["skip", "redo", "force"],
        "savePolicies": ["rewrite"],
        "outputType": "pdf",
        "optimize": 0,
        "arbitraryProviderFlagsAccepted": False,
        "passwordsAccepted": False,
        "providerIsSanitizer": False,
        "inputTrustAssertionRequired": True,
        "adapterSandboxEnforced": False,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description=__doc__)
    subparsers = parser.add_subparsers(dest="command", required=True)

    probe_parser = subparsers.add_parser("probe", help="require the pinned OCRmyPDF/Tesseract/qpdf/Poppler capability set")
    probe_parser.add_argument("--timeout-seconds", type=positive_int, default=30)

    ocr = subparsers.add_parser("ocr", help="add a searchable OCR layer and publish a distinct full-rewrite PDF")
    ocr.add_argument("input")
    ocr.add_argument("output")
    ocr.add_argument("--expected-sha256", required=True)
    ocr.add_argument("--mode", choices=["skip", "redo", "force"], required=True)
    ocr.add_argument("--language", action="append", default=[])
    ocr.add_argument("--input-trust", choices=["trusted", "caller-isolated"], required=True)
    ocr.add_argument("--require-text", action="append", default=[])
    ocr.add_argument("--allow-empty-text", action="store_true")
    ocr.add_argument("--allow-structure-loss", action="store_true")
    ocr.add_argument("--allow-rasterize-all", action="store_true")
    ocr.add_argument("--allow-interactive-flattening", action="store_true")
    ocr.add_argument("--invalidate-signatures", action="store_true")
    ocr.add_argument("--max-input-bytes", type=positive_int, default=DEFAULT_MAX_INPUT_BYTES)
    ocr.add_argument("--max-output-bytes", type=positive_int, default=DEFAULT_MAX_OUTPUT_BYTES)
    ocr.add_argument("--max-sidecar-bytes", type=positive_int, default=DEFAULT_MAX_SIDECAR_BYTES)
    ocr.add_argument("--max-extracted-text-bytes", type=positive_int, default=DEFAULT_MAX_EXTRACTED_TEXT_BYTES)
    ocr.add_argument("--timeout-seconds", type=positive_int, default=DEFAULT_TIMEOUT_SECONDS)
    ocr.add_argument("--tesseract-timeout-seconds", type=positive_int, default=DEFAULT_TESSERACT_TIMEOUT_SECONDS)
    ocr.add_argument("--max-stdout-bytes", type=positive_int, default=DEFAULT_MAX_STDOUT_BYTES)
    ocr.add_argument("--max-stderr-bytes", type=positive_int, default=DEFAULT_MAX_STDERR_BYTES)
    ocr.add_argument("--max-image-mpixels", type=positive_int, default=DEFAULT_MAX_IMAGE_MPIXELS)
    return parser


def main() -> int:
    from python_runtime import reexec_configured_provider_python
    reexec_configured_provider_python()
    args = build_parser().parse_args()
    try:
        result = probe(args) if args.command == "probe" else ocr_pdf(args)
        print(json.dumps(result, indent=2, sort_keys=True))
        return 0
    except (ProviderError, qpdf.ProviderError, OSError, UnicodeError) as exc:
        print(json.dumps({
            "ok": False,
            "provider": "ocrmypdf",
            "error": str(exc),
            "silentFallback": False,
        }, sort_keys=True), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
