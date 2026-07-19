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

## qpdf structure inspection and repair

qpdf is separately installed. The shipped wrapper exposes only bounded inspect,
recovery rewrite, and linearize operations; it does not pass arbitrary flags
through to the provider:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider qpdf --require
"$PYTHON_BIN" scripts/qpdf_provider.py inspect input.pdf \
  > tmp/pdfs/qpdf-inspect.json
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task repair --provider qpdf --strategy rewrite \
  --input input.pdf --output outputs/repaired.pdf --require-provider
"$PYTHON_BIN" scripts/qpdf_provider.py rewrite \
  input.pdf outputs/repaired.pdf \
  --mode repair --expected-sha256 '<sha256-from-inspect>' \
  > tmp/pdfs/qpdf-repair.json
```

The source hash is a required precondition. A recoverable damaged xref may enter
with warning status, but the promoted result must re-inspect cleanly and preserve
page/form/attachment/outline counts. Use `--mode linearize` for an explicit linearized
rewrite. Signature evidence requires `--invalidate-signatures` after pyHanko and
DocMDP review. This route is never sanitize; run Poppler over every final page.

## pyHanko read-only signature validation

Install `pyHanko>=0.35,<0.36` into the selected PDF provider environment. The
core library is enough for validation; the separate `pyhanko-cli` package is
only needed by an explicit signing workflow.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 signed.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider pyhanko --require
"$PYTHON_BIN" scripts/pyhanko_provider.py probe
"$PYTHON_BIN" scripts/pyhanko_provider.py verify signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy explicit-roots \
  --trust-root /trusted/root-ca.pem \
  --revocation-policy none \
  --require-signature --require-all-integrity-valid \
  --require-all-trusted --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json
```

`none` deliberately makes no revocation check; choose `soft-fail`, `hard-fail`,
or `require` when the task and available evidence demand it. The adapter never
fetches network evidence or guesses trust roots. Review every signature's
signed revision, coverage, modification level, DocMDP result, timestamp, and
trust status rather than treating an intact old ByteRange as approval of later
edits. The report is not a complete PAdES conformance certificate.

## veraPDF source-bound machine validation

Install veraPDF 1.30.x separately and select the built-in profile that matches
the delivery requirement. The shipped adapter validates an immutable private
snapshot rather than accepting arbitrary veraPDF flags:

```bash
export OPEN_OFFICE_PDF_VERAPDF="/absolute/path/to/verapdf"
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 output.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task validate-conformance --provider verapdf --strategy read-only \
  --input output.pdf --require-provider
"$PYTHON_BIN" scripts/verapdf_provider.py validate output.pdf \
  --expected-sha256 "$SOURCE_SHA256" --flavour 2u --require-compliant \
  > tmp/pdfs/verapdf-pdfa2u.json
```

Omit `--require-compliant` only when collecting diagnostics rather than gating
delivery: a noncompliant result then completes with `machineRuleCompliant:
false`. PDF/UA profiles (`ua1` and `ua2`) always require a separate human
review of reading order, semantics, alternatives, contrast, and actual
assistive-technology usability.

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
