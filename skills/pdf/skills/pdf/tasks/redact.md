# Redact and sanitize

High-trust redaction is a destructive security workflow, never a visual overlay.

## Requirements

- Use PyMuPDF directly on the original PDF.
- Add and apply real redaction annotations.
- Select `sanitize`; incremental save is forbidden.
- Scrub attachments, embedded files, JavaScript/actions, metadata/XMP, hidden text, annotation responses, thumbnails, and other selected residues.
- Perform a full rewrite with garbage collection; the output must not retain the original byte prefix or old incremental revisions.
- Scan raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations/widgets, and OCR of image-bearing pages.
- Render every page from final bytes with Poppler and inspect it.

## Operations

Before mutation, prove the exact adapter surface and bind the destructive route:

```bash
python3 scripts/pymupdf_edit.py probe --accept-license agpl
python3 scripts/pdf_provider.py plan \
  --task redact --provider pymupdf --strategy sanitize \
  --input input.pdf --output tmp/pdfs/sanitized.pdf \
  --accept-license agpl --invalidate-signatures --require-provider
```

```json
[
  { "type": "redact_text", "term": "Customer Secret", "fill": [0, 0, 0] },
  { "type": "redact_rect", "page": 2, "rect": [72, 120, 340, 180], "fill": [0, 0, 0] }
]
```

```bash
python3 scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/sanitized.pdf \
  --strategy sanitize \
  --operations tmp/pdfs/redactions.json \
  --sensitive-term 'Customer Secret' \
  --accept-license agpl \
  --invalidate-signatures
```

If an image-bearing page requires OCR and Tesseract/PyMuPDF OCR cannot run, the strict residue gate fails. Install/configure OCR, repeat the scan, and do not deliver an incompletely scanned file.

Opaque black rectangles, annotation-only redactions, incremental writes, or text-extraction-only checks are not acceptable. Keep the original and QA evidence under restricted access according to the user's data-handling requirements.

Emit the canonical [`open-office-artifact-tool.pdf-audit.v1`](../references/AUDIT_SCHEMA.md) record, including residue and render evidence under `validation`, and run `scripts/pdf_audit.py validate` against the final sanitized bytes.
