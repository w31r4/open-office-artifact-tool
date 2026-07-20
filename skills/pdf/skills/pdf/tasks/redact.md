# Redact and sanitize

High-trust redaction is a destructive security workflow, never a visual overlay.

MuPDF.js can apply `redact_text` or `redact_rect` during a full rewrite and refuses incremental output. That is useful for bounded ordinary editing, but it is not a complete sanitize claim. The high-trust workflow below additionally scrubs active content and metadata, removes old revisions, scans decoded/raw/OCR residue, and renders independently.

## Requirements

- Use PyMuPDF directly on the original PDF.
- Add and apply real redaction annotations.
- Select `sanitize`; incremental save is forbidden.
- Scrub attachments, embedded files, JavaScript/actions, metadata/XMP, hidden text, annotation responses, thumbnails, and other selected residues.
- Perform a full rewrite with garbage collection; the output must not retain the original byte prefix or old incremental revisions.
- Scan raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations/widgets, and OCR of image-bearing pages.
- Render every page from final bytes with Poppler and inspect it.
- For text that exists only in raster pixels, use the typed `redact_ocr_text`
  operation. It requires one explicit page, its expected 0/90/180/270-degree
  rotation, and an expected image-backed match count; it is not a general OCR
  guess-and-delete mode.

## Operations

Before mutation, prove the exact adapter surface and bind the destructive route:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe \
  --accept-license agpl --ocr-language eng --require-ocr
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task redact --provider pymupdf --strategy sanitize \
  --input input.pdf --output tmp/pdfs/sanitized.pdf \
  --accept-license agpl --invalidate-signatures --require-provider
```

```json
[
  { "type": "redact_text", "term": "Customer Secret", "fill": [0, 0, 0] },
  { "type": "redact_ocr_text", "page": 2, "expected_rotation": 90, "term": "IMAGE SECRET 8842", "expected_matches": 1, "fill": [0, 0, 0] },
  { "type": "redact_rect", "page": 2, "rect": [72, 120, 340, 180], "fill": [0, 0, 0] }
]
```

```bash
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/sanitized.pdf \
  --strategy sanitize \
  --operations tmp/pdfs/redactions.json \
  --sensitive-term 'Customer Secret' \
  --sensitive-term 'IMAGE SECRET 8842' \
  --ocr-language eng --ocr-dpi 200 \
  --accept-license agpl \
  --invalidate-signatures
```

`redact_ocr_text` OCRs the selected page through Tesseract, discards matches
that do not overlap a native raster placement by at least 90%, and applies a
real image redaction only when the remaining count equals `expected_matches`.
The adapter caps OCR at 72–300 dpi, 100 million raster pixels per page, 1,000
matches, and 4,096 term characters. It requires an exact `expected_rotation`.
For rotated pages it temporarily clears `/Rotate` only while Tesseract sees the
page upright, keeps all returned/redaction rectangles in canonical unrotated
PyMuPDF page space, restores `/Rotate`, and reports both unrotated and
display-space rectangles for independent QA. OCR uncertainty is never converted
into a broad rectangle or provider fallback.

If an image-bearing page requires OCR and Tesseract/PyMuPDF OCR cannot run, the strict residue gate fails. Install/configure OCR, repeat the scan, and do not deliver an incompletely scanned file. After redaction, the same OCR term must be absent from the fully rewritten output; source bytes remain immutable.

## Active-content public copy

For a public-release sanitize with no requested text redaction, use the typed scrub operation:

```json
[{ "type": "scrub" }]
```

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task sanitize --provider pymupdf --strategy sanitize \
  --input input.pdf --output tmp/pdfs/public-safe.pdf \
  --accept-license agpl --invalidate-signatures --require-provider
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/public-safe.pdf \
  --strategy sanitize --operations tmp/pdfs/scrub.json \
  --accept-license agpl --invalidate-signatures
"$PYTHON_BIN" scripts/residue_scan.py tmp/pdfs/public-safe.pdf \
  --require-inert --require-single-revision
```

`--require-inert` fails on active action names, including inert-looking `null` dictionary entries, attachments, comment annotations, populated form values, personal metadata, links, or invisible text. The typed adapter physically removes those null active-content names after provider scrub and reports the affected xrefs; it refuses an unfamiliar object serialization rather than requiring caller-side object edits. It removes isolated invisible text only when its rectangle does not overlap visible text; otherwise it deletes the transactional output and fails closed. Scrub-only sanitization needs no placeholder `--sensitive-term`. Any `redact_text`, `redact_ocr_text`, `redact_rect`, or `replace_text` operation still requires explicit sensitive terms and the residue/OCR gate.

Opaque black rectangles, annotation-only redactions, incremental writes, or text-extraction-only checks are not acceptable. Keep the original and QA evidence under restricted access according to the user's data-handling requirements.

Emit the canonical [`open-office-artifact-tool.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) record, including residue and render evidence under `validation`, and run `scripts/pdf_audit.py validate` against the final sanitized bytes.
