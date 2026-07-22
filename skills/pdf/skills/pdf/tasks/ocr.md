# Add a searchable OCR layer

Use the shipped `scripts/ocrmypdf_provider.py` adapter for a complete scanned
PDF that needs searchable text. It is a source-bound full rewrite, not a parser,
layout editor, accessibility repair, redaction, or sanitize operation.

## Resolve and probe

OCRmyPDF is an explicit capability route. First resolve `task: "ocr"` with
`provider: "ocrmypdf"`, `savePolicy: "rewrite"`, `mutationAuthorized: true`,
and the requested language(s). The route needs OCRmyPDF, Tesseract language
data, Ghostscript, qpdf, and Poppler `pdftotext`; those facts, their pinned
versions, and their pack closure come from the public catalog rather than this
task page.

Use [provider setup](provider_setup.md) to select a ready managed pack or an
explicit deployment-owned `system-only` runtime. qpdf is already available as a
managed pack, but OCR core and language packs remain unpublished; an OCR route
therefore blocks rather than downloading a replacement. Do not repair that
state with brew, apt, global pip, or a guessed download URL.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider ocrmypdf --require
"$PYTHON_BIN" scripts/ocrmypdf_provider.py probe
```

For an explicitly selected system runtime, set its exact executable with
`OPEN_OFFICE_PDF_OCRMYPDF`, `OPEN_OFFICE_PDF_TESSERACT`,
`OPEN_OFFICE_PDF_QPDF`, `OPEN_OFFICE_PDF_PDFTOTEXT`, or
`OPEN_OFFICE_PDF_GS`. For a ready managed route, pass the resolver-returned
`runtime.managed.environment` unchanged; its `OPEN_OFFICE_PDF_TESSDATA_DIRS`
value is copied into a private temporary Tesseract directory rather than used
in place. The probe reports the resolved versions and installed Tesseract
languages. It never installs a provider or silently selects a substitute.

OCRmyPDF explicitly warns that it is not a malware boundary for attacker-chosen
PDFs. `--input-trust trusted` asserts that the source is trusted.
`--input-trust caller-isolated` asserts that the caller has already placed this
whole adapter process inside an OS/container/VM isolation boundary. The adapter
records that assertion but does not claim to enforce a sandbox itself.

## OCR one complete PDF

Inspect first, retain the source SHA-256, then create a distinct absent output:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/qpdf_provider.py inspect scanned.pdf \
  > tmp/pdfs/scanned-structure.json
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task ocr --provider ocrmypdf --strategy rewrite \
  --input scanned.pdf --output outputs/scanned-searchable.pdf --require-provider
"$PYTHON_BIN" scripts/ocrmypdf_provider.py ocr \
  scanned.pdf outputs/scanned-searchable.pdf \
  --expected-sha256 '<sha256-from-fresh-inspect>' \
  --mode skip --language eng --input-trust trusted \
  --require-text 'expected phrase' \
  > tmp/pdfs/ocr-report.json
```

The adapter always passes `--output-type pdf --optimize 0 --jobs 1`, selects
Tesseract/fpdf2/pypdfium explicitly, disables arbitrary provider flags, and
works on a private read-only source snapshot. It rejects encrypted input,
stale hashes, source overwrite, symlink/collision output, output/source budget
breaches, missing language packs, unexpected page/attachment/outline changes,
qpdf warnings, an incremental source prefix, missing text, and failed
`--require-text` gates. Poppler independently extracts the final text before a
no-replace atomic output promotion. The temporary Tesseract sidecar is hashed
and summarized in the typed report, then removed rather than published.

`--require-text` is a case-insensitive, whitespace-collapsed substring gate.
Repeat it for a small set of phrases that must be recognized. A genuinely blank
scan requires explicit `--allow-empty-text`; do not use that flag merely to hide
an OCR failure.

## Select the mode deliberately

- `skip` is the normal mixed-document route. Pages with existing text remain
  untouched; image-only pages receive an OCR layer.
- `redo` removes an existing hidden OCR layer and requires
  `--allow-structure-loss` because compatible Tagged PDF structure cannot be
  retained.
- `force` rasterizes every page and requires both `--allow-structure-loss` and
  `--allow-rasterize-all`. If forms or annotations are present it also requires
  `--allow-interactive-flattening`.

Any Tagged PDF requires explicit `--allow-structure-loss`: `skip` can preserve
markup on untouched text pages, but OCRmyPDF cannot map newly recognized text to
the existing structure tree. Signed or signature-constrained input requires a
pyHanko/DocMDP/FieldMDP review and explicit `--invalidate-signatures`. The full
rewrite cannot preserve the prior signature's approval.

No preprocessing flags such as deskew, rotate, clean, background removal,
arbitrary page selection, provider plugins, PDF/A conversion, or lossy
optimization are exposed in this first fidelity slice. Add a separate typed
operation if one of those transformations becomes necessary.

## Delivery gates

The OCR report proves provider identity, source/output hashes, fixed flags,
mode acknowledgements, qpdf topology, extracted-text evidence, and atomic
rewrite. It does not prove visual fidelity or OCR correctness. Before delivery:

```bash
pdfinfo outputs/scanned-searchable.pdf
pdftotext outputs/scanned-searchable.pdf -
pdftoppm -png -r 144 scanned.pdf tmp/pdfs/source/page
pdftoppm -png -r 144 outputs/scanned-searchable.pdf tmp/pdfs/output/page
```

Compare every Poppler-rendered page, review text and reading order against the
scan, reopen with MuPDF, and run conformance/signature checks when the delivery
contract requires them. OCRmyPDF can make a searchable PDF; it does not create
reliable author-intent tags or turn the file into a Word-style reflow model.
