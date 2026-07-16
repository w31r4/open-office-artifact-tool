#!/usr/bin/env python3
"""Fail-closed sensitive-residue scan for a fully rewritten PDF."""

from __future__ import annotations

import argparse
import hashlib
import json
from pathlib import Path
import re
import sys
from typing import Any


class ScanError(RuntimeError):
    pass


RISKY_STRUCTURAL_NAMES = (
    "/AA",
    "/EmbeddedFiles",
    "/ImportData",
    "/JavaScript",
    "/JS",
    "/Launch",
    "/OpenAction",
    "/RichMedia",
    "/SubmitForm",
    "/XFA",
)


def require_pymupdf():
    try:
        import pymupdf
    except ImportError as exc:
        raise ScanError("PyMuPDF is required for decoded object, page, attachment, and OCR residue scanning") from exc
    return pymupdf


def sha256(payload: bytes) -> str:
    return hashlib.sha256(payload).hexdigest()


def serialize(value: Any) -> str:
    try:
        return json.dumps(value, ensure_ascii=False, sort_keys=True, default=str)
    except Exception:
        return str(value)


def structural_name_counts(doc) -> dict[str, int]:
    counts = {name: 0 for name in RISKY_STRUCTURAL_NAMES}
    action_subtypes = {"/ImportData", "/JavaScript", "/Launch", "/SubmitForm"}
    for xref in range(1, doc.xref_length()):
        try:
            keys = doc.xref_get_keys(xref) or []
            for key in keys:
                token = f"/{key}"
                if token not in counts:
                    continue
                value_type, _ = doc.xref_get_key(xref, key)
                if value_type != "null":
                    counts[token] += 1
            obj = doc.xref_object(xref, compressed=False, ascii=False) or ""
            for subtype in action_subtypes:
                if re.search(rf"/S\s+{re.escape(subtype)}(?![A-Za-z0-9])", obj):
                    counts[subtype] += 1
        except Exception:
            continue
    return counts


class Findings:
    def __init__(self, terms: list[str]):
        self.terms = terms
        self.counts = {term: 0 for term in terms}
        self.evidence = {term: [] for term in terms}

    def text(self, category: str, location: str, value: Any) -> None:
        haystack = str(value or "").casefold()
        for term in self.terms:
            count = haystack.count(term.casefold())
            if not count:
                continue
            self.counts[term] += count
            if len(self.evidence[term]) < 25:
                self.evidence[term].append({"category": category, "location": location, "matches": count})

    def bytes(self, category: str, location: str, payload: bytes) -> None:
        for term in self.terms:
            variants = []
            for encoding in ("utf-8", "utf-16-le", "utf-16-be", "latin-1"):
                try:
                    encoded = term.encode(encoding)
                except UnicodeEncodeError:
                    continue
                if encoded and encoded not in variants:
                    variants.append(encoded)
            count = sum(payload.lower().count(encoded.lower()) for encoded in variants)
            if not count:
                continue
            self.counts[term] += count
            if len(self.evidence[term]) < 25:
                self.evidence[term].append({"category": category, "location": location, "matches": count})

    def report(self) -> dict[str, Any]:
        return {
            term: {"matches": self.counts[term], "evidence": self.evidence[term]}
            for term in self.terms
        }

    @property
    def total(self) -> int:
        return sum(self.counts.values())


def scan_pdf(
    input_path: Path,
    terms: list[str],
    *,
    require_ocr: bool = False,
    require_single_revision: bool = False,
    require_inert: bool = False,
    ocr_language: str = "eng",
    ocr_dpi: int = 150,
    max_bytes: int = 512 * 1024 * 1024,
    max_pages: int = 2_000,
    max_xrefs: int = 1_000_000,
    max_decoded_bytes: int = 1024 * 1024 * 1024,
) -> dict[str, Any]:
    pymupdf = require_pymupdf()
    source = input_path.expanduser().resolve()
    if not source.is_file():
        raise ScanError("input must be an existing PDF")
    normalized_terms = []
    for term in terms:
        value = str(term)
        if not value:
            raise ScanError("sensitive terms must not be empty")
        if value not in normalized_terms:
            normalized_terms.append(value)
    if not normalized_terms and not require_inert:
        raise ScanError("at least one --term or --require-inert is required")
    payload = source.read_bytes()
    if len(payload) > max_bytes:
        raise ScanError(f"PDF is {len(payload)} bytes; strict max-bytes is {max_bytes}")

    findings = Findings(normalized_terms)
    incomplete: list[dict[str, Any]] = []
    findings.bytes("raw-bytes", "entire-file", payload)
    findings.text("raw-pdf-text", "latin-1-projection", payload.decode("latin-1", "ignore"))
    eof_markers = len(re.findall(rb"%%EOF\s*", payload))
    prev_markers = len(re.findall(rb"/Prev\s+\d+", payload))

    doc = pymupdf.open(str(source))
    try:
        if doc.needs_pass:
            raise ScanError("encrypted PDF requires an authorized decrypted input for strict residue scanning")
        if doc.page_count > max_pages:
            raise ScanError(f"PDF has {doc.page_count} pages; strict max-pages is {max_pages}")
        xref_length = doc.xref_length()
        if xref_length > max_xrefs:
            raise ScanError(f"PDF has {xref_length} xrefs; strict max-xrefs is {max_xrefs}")

        metadata = dict(doc.metadata or {})
        findings.text("metadata", "document-metadata", serialize(metadata))
        try:
            xml_metadata = doc.get_xml_metadata() or ""
            findings.text("xmp", "document-xmp", xml_metadata)
        except Exception as exc:
            incomplete.append({"category": "xmp", "error": str(exc)})

        decoded_bytes = 0
        for xref in range(1, xref_length):
            try:
                obj = doc.xref_object(xref, compressed=False, ascii=False) or ""
                decoded_bytes += len(obj.encode("utf-8", "replace"))
                if decoded_bytes > max_decoded_bytes:
                    raise ScanError(f"decoded object data exceeds strict max-decoded-bytes {max_decoded_bytes}")
                findings.text("decoded-object", f"xref:{xref}", obj)
                if doc.is_stream(xref):
                    stream = doc.xref_stream(xref) or b""
                    decoded_bytes += len(stream)
                    if decoded_bytes > max_decoded_bytes:
                        raise ScanError(f"decoded stream data exceeds strict max-decoded-bytes {max_decoded_bytes}")
                    findings.bytes("decoded-stream", f"xref:{xref}", stream)
                    findings.text("decoded-stream-text", f"xref:{xref}", stream.decode("utf-8", "ignore"))
                    findings.text("decoded-stream-text", f"xref:{xref}:latin-1", stream.decode("latin-1", "ignore"))
            except ScanError:
                raise
            except Exception as exc:
                incomplete.append({"category": "xref", "location": f"xref:{xref}", "error": str(exc)})

        attachment_names = []
        try:
            attachment_names = list(doc.embfile_names() or [])
            for name in attachment_names:
                findings.text("attachment-name", str(name), name)
                try:
                    attachment = doc.embfile_get(name)
                    findings.bytes("attachment-bytes", str(name), attachment)
                    findings.text("attachment-text", str(name), attachment.decode("utf-8", "ignore"))
                    findings.text("attachment-text", f"{name}:latin-1", attachment.decode("latin-1", "ignore"))
                except Exception as exc:
                    incomplete.append({"category": "attachment", "location": str(name), "error": str(exc)})
        except Exception as exc:
            incomplete.append({"category": "attachments", "error": str(exc)})

        pages = []
        hidden_text_spans = 0
        comment_annotations = 0
        links = 0
        populated_form_values: list[dict[str, Any]] = []
        for page_index in range(doc.page_count):
            page = doc.load_page(page_index)
            page_number = page_index + 1
            page_record: dict[str, Any] = {"page": page_number}
            try:
                text = page.get_text("text") or ""
                findings.text("extracted-text", f"page:{page_number}", text)
                page_record["textChars"] = len(text)
            except Exception as exc:
                incomplete.append({"category": "page-text", "location": f"page:{page_number}", "error": str(exc)})
            try:
                images = page.get_images(full=True) or []
                page_record["images"] = len(images)
            except Exception as exc:
                images = []
                incomplete.append({"category": "page-images", "location": f"page:{page_number}", "error": str(exc)})
            annotations = []
            try:
                for annotation in page.annots() or []:
                    info = dict(annotation.info or {})
                    annotations.append({"type": annotation.type, "info": info, "rect": list(annotation.rect)})
                    findings.text("annotation", f"page:{page_number}:xref:{annotation.xref}", serialize(annotations[-1]))
            except Exception as exc:
                incomplete.append({"category": "annotations", "location": f"page:{page_number}", "error": str(exc)})
            page_record["annotations"] = len(annotations)
            comment_annotations += len(annotations)
            widgets = []
            try:
                for widget in page.widgets() or []:
                    record = {
                        "name": widget.field_name,
                        "value": widget.field_value,
                        "type": widget.field_type_string,
                        "rect": list(widget.rect),
                    }
                    widgets.append(record)
                    findings.text("widget", f"page:{page_number}:{widget.field_name}", serialize(record))
                    if str(widget.field_value or "").strip().casefold() not in {"", "off", "false", "none"}:
                        populated_form_values.append({"page": page_number, "name": widget.field_name, "value": str(widget.field_value)})
            except Exception as exc:
                incomplete.append({"category": "widgets", "location": f"page:{page_number}", "error": str(exc)})
            page_record["widgets"] = len(widgets)
            try:
                page_links = list(page.get_links() or [])
                page_record["links"] = len(page_links)
                links += len(page_links)
            except Exception as exc:
                incomplete.append({"category": "links", "location": f"page:{page_number}", "error": str(exc)})
            try:
                page_hidden = [trace for trace in page.get_texttrace() if int(trace.get("type", 0)) == 3 or float(trace.get("opacity", 1)) == 0]
                page_record["hiddenTextSpans"] = len(page_hidden)
                hidden_text_spans += len(page_hidden)
            except Exception as exc:
                incomplete.append({"category": "hidden-text", "location": f"page:{page_number}", "error": str(exc)})

            if images and require_ocr:
                try:
                    textpage = page.get_textpage_ocr(
                        language=ocr_language,
                        dpi=ocr_dpi,
                        full=True,
                    )
                    ocr_text = page.get_text("text", textpage=textpage) or ""
                    findings.text("image-ocr", f"page:{page_number}", ocr_text)
                    page_record["ocr"] = {"complete": True, "chars": len(ocr_text), "language": ocr_language, "dpi": ocr_dpi}
                except Exception as exc:
                    page_record["ocr"] = {"complete": False, "error": str(exc)}
                    incomplete.append({"category": "image-ocr", "location": f"page:{page_number}", "error": str(exc)})
            elif images:
                page_record["ocr"] = {"complete": False, "required": False}
            pages.append(page_record)

        if require_single_revision:
            if eof_markers != 1:
                incomplete.append({"category": "revisions", "error": f"expected one %%EOF marker, found {eof_markers}"})
            if prev_markers:
                incomplete.append({"category": "revisions", "error": f"found {prev_markers} /Prev revision pointer(s)"})

        names = structural_name_counts(doc)
        metadata_fields = {
            key: str(metadata.get(key) or "")
            for key in ("title", "author", "subject", "keywords", "creator")
            if str(metadata.get(key) or "").strip()
        }
        inert_violations = []
        if require_inert:
            if attachment_names:
                inert_violations.append({"category": "attachments", "count": len(attachment_names)})
            if comment_annotations:
                inert_violations.append({"category": "annotations", "count": comment_annotations})
            if populated_form_values:
                inert_violations.append({"category": "form-values", "values": populated_form_values})
            if metadata_fields:
                inert_violations.append({"category": "personal-metadata", "fields": metadata_fields})
            if hidden_text_spans:
                inert_violations.append({"category": "hidden-text", "count": hidden_text_spans})
            if links:
                inert_violations.append({"category": "links", "count": links})
            remaining_names = {name: count for name, count in names.items() if count}
            if remaining_names:
                inert_violations.append({"category": "active-structure", "names": remaining_names})

        report = {
            "ok": findings.total == 0 and not incomplete and not inert_violations,
            "provider": "pymupdf",
            "providerVersion": pymupdf.__version__,
            "strategy": "read-only",
            "silentFallback": False,
            "source": {"path": str(source), "bytes": len(payload), "sha256": sha256(payload)},
            "terms": findings.report(),
            "summary": {
                "matches": findings.total,
                "pages": doc.page_count,
                "xrefs": xref_length,
                "decodedBytes": decoded_bytes,
                "attachments": len(attachment_names),
                "commentAnnotations": comment_annotations,
                "populatedFormValues": len(populated_form_values),
                "hiddenTextSpans": hidden_text_spans,
                "links": links,
                "incompleteChecks": len(incomplete),
            },
            "activeContent": {
                "required": require_inert,
                "structuralNames": names,
                "personalMetadata": metadata_fields,
                "populatedFormValues": populated_form_values,
                "violations": inert_violations,
            },
            "revisions": {"eofMarkers": eof_markers, "prevPointers": prev_markers, "singleRevisionRequired": require_single_revision},
            "pages": pages,
            "incomplete": incomplete,
        }
        return report
    finally:
        doc.close()


def parser() -> argparse.ArgumentParser:
    root = argparse.ArgumentParser(description=__doc__)
    root.add_argument("input", type=Path)
    root.add_argument("--term", action="append", default=[], help="sensitive token; repeat for multiple tokens")
    root.add_argument("--output", type=Path, help="optional JSON report path")
    root.add_argument("--require-ocr", action="store_true", help="fail if any image-bearing page cannot be OCR scanned")
    root.add_argument("--require-single-revision", action="store_true")
    root.add_argument("--require-inert", action="store_true", help="fail on active actions, attachments, comments, populated forms, personal metadata, links, or hidden text")
    root.add_argument("--ocr-language", default="eng")
    root.add_argument("--ocr-dpi", type=int, default=150)
    root.add_argument("--max-bytes", type=int, default=512 * 1024 * 1024)
    root.add_argument("--max-pages", type=int, default=2_000)
    root.add_argument("--max-xrefs", type=int, default=1_000_000)
    root.add_argument("--max-decoded-bytes", type=int, default=1024 * 1024 * 1024)
    return root


def main() -> int:
    args = parser().parse_args()
    try:
        report = scan_pdf(
            args.input,
            args.term,
            require_ocr=args.require_ocr,
            require_single_revision=args.require_single_revision,
            require_inert=args.require_inert,
            ocr_language=args.ocr_language,
            ocr_dpi=args.ocr_dpi,
            max_bytes=args.max_bytes,
            max_pages=args.max_pages,
            max_xrefs=args.max_xrefs,
            max_decoded_bytes=args.max_decoded_bytes,
        )
        rendered = json.dumps(report, ensure_ascii=False, indent=2, sort_keys=True)
        if args.output:
            destination = args.output.expanduser().resolve()
            if destination == args.input.expanduser().resolve():
                raise ScanError("JSON report path must differ from the source PDF")
            destination.parent.mkdir(parents=True, exist_ok=True)
            destination.write_text(rendered + "\n", "utf-8")
        else:
            print(rendered)
        return 0 if report["ok"] else 2
    except Exception as exc:
        print(json.dumps({"ok": False, "provider": "pymupdf", "error": str(exc), "silentFallback": False}), file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
