# Provider workflow examples

Run these from the PDF Skill directory. Always substitute task-local source/output paths and inspect the operation JSON before execution.

## Default MuPDF.js inspect, render, and bounded edit

The normal npm installation includes required `mupdf@1.28.0`; the root facade loads its WASM runtime only when the first MuPDF-backed PDF operation runs. Explicit provider probing imports the MuPDF subpath intentionally. Use the thin Skill CLI for the default arbitrary-file path:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
node scripts/mupdf.mjs render input.pdf tmp/pdfs/page-1.png --page 1 --dpi 144
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

The CLI rejects direct or symlink-alias source overwrite and writes output atomically. `incremental` is limited to unsigned, non-destructive operations and must preserve the exact source byte prefix. A MuPDF.js rewrite redaction is real page-content redaction, but it is not the strict metadata/attachment/hidden-layer/OCR sanitization workflow below.

## pypdf attachment quarantine

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pypdf --require
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task extract-attachments --provider pypdf --strategy read-only \
  --input input.pdf --require-provider
"$PYTHON_BIN" scripts/pypdf_edit.py extract-attachments input.pdf outputs/quarantine \
  --manifest outputs/attachments.json
```

Use only the `savedPath` values in the manifest. Raw display names and internal keys are evidence, never filesystem paths. The command preserves duplicate names as separate files, confines traversal names to the quarantine directory, verifies decoded bytes and hashes, and does not open payloads.

## pypdf merge, reorder, and selective watermark

Copy [`merge-stamp-manifest.json`](merge-stamp-manifest.json) to a task-local path and replace its source paths. The sequence must select each source page exactly once.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pypdf --require
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task merge-stamp --provider pypdf --strategy rewrite \
  --input tmp/pdfs/merge-stamp.json --output outputs/merged.pdf --require-provider
"$PYTHON_BIN" scripts/pypdf_edit.py merge-stamp \
  tmp/pdfs/merge-stamp.json outputs/merged.pdf --strategy rewrite
"$PYTHON_BIN" scripts/poppler_compare.py merge-stamp \
  tmp/pdfs/merge-stamp.json outputs/merged.pdf \
  --report tmp/pdfs/merge-visual-qa.json --render-dir tmp/pdfs/merge-rendered
```

The result JSON proves source hashes, final page mapping/geometry, navigation targets, and watermark text/opacity. The comparison report proves the declared source-to-output pixel invariants through Poppler. Bind the manifest plus every PDF source in the canonical audit.

## ReportLab greenfield PDF

```bash
python3 scripts/reportlab_create.py \
  --spec examples/reportlab-report-spec.json \
  --output tmp/pdfs/release-evidence.pdf
python3 scripts/pdfplumber_extract.py tmp/pdfs/release-evidence.pdf \
  --output tmp/pdfs/release-evidence-extraction.json
pdftoppm -png -r 144 tmp/pdfs/release-evidence.pdf tmp/pdfs/release-evidence-page
```

## PyMuPDF specialist positioned-text and image edit

Select this optional specialist route only when the requested operation, such as `insert_textbox`, `insert_image`, or `replace_image`, is outside the default MuPDF.js primitive set.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task edit-content --provider pymupdf --strategy rewrite \
  --input input.pdf --output tmp/pdfs/edited.pdf \
  --accept-license agpl --require-provider
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations examples/pymupdf-edit-operations.json \
  --accept-license agpl
```

## PyMuPDF high-trust replacement

The sample assumes the exact text `Customer Secret` exists in one horizontal source span and the short replacement fits its original box. The primitive preserves the source baseline and defaults, records source/output style plus measured width and fixed numerical tolerance in `operations[].fitChecks`, and fails for cross-span/rotated text or overflow beyond that sub-millipoint bound.

Before editing, run `pymupdf_edit.py probe` and `pdf_provider.py plan --task redact --provider pymupdf --strategy sanitize ... --invalidate-signatures --require-provider`; both must succeed before mutation.

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/sanitized.pdf \
  --strategy sanitize \
  --operations examples/pymupdf-redaction-operations.json \
  --sensitive-term 'Customer Secret' \
  --accept-license agpl \
  --invalidate-signatures
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/residue_scan.py tmp/pdfs/sanitized.pdf \
  --term 'Customer Secret' --require-ocr --require-single-revision
pdftoppm -png -r 144 tmp/pdfs/sanitized.pdf tmp/pdfs/sanitized-page
```

The command itself performs the strict residue scan before promoting its transactional output. The second invocation preserves a standalone JSON-able audit step. Image-bearing files require a working Tesseract installation.

For an active-content public copy without term redaction, use `[ { "type": "scrub" } ]`, plan `--task sanitize --strategy sanitize`, omit placeholder sensitive terms, and run the structural gate after the transactional edit:

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/residue_scan.py tmp/pdfs/public-safe.pdf \
  --require-inert --require-single-revision
```

The gate covers active action names, including null-valued dictionary entries, attachments, comments, populated widgets, personal metadata, links, and invisible text. The typed provider physically removes null active-content names after scrub; callers must not patch xrefs or content streams themselves. If invisible text overlaps visible content, the provider refuses the operation instead of redacting the visible page.

After semantic and Poppler checks, write the canonical audit envelope from [`AUDIT_SCHEMA.md`](../references/AUDIT_SCHEMA.md), then bind it to the delivered bytes:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source input.pdf --artifact tmp/pdfs/sanitized.pdf \
  --require-operation replace_text
```
