# Provider setup and probes

## Required MuPDF.js runtime

`mupdf@1.28.0` is a direct dependency of `open-office-artifact-tool`. A normal npm install is the complete setup:

```bash
npm install open-office-artifact-tool
node scripts/mupdf.mjs probe
```

The WASM runtime is loaded only when a PDF operation first needs it. The project has no PDF `postinstall`, downloader, virtual-environment bootstrapper, or global mutation step.

Use the thin CLI for the default native path:

```bash
node scripts/mupdf.mjs inspect input.pdf
node scripts/mupdf.mjs render input.pdf tmp/pdfs/page-1.png --page 1 --dpi 144
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/operations.json tmp/pdfs/edited.pdf --save-policy rewrite
```

The CLI refuses to overwrite the input, including through a symlink alias. It enforces input, page/object, and render budgets and writes outputs atomically; `PdfFile.importPdf` additionally enforces image budgets. MuPDF.js supports `rewrite` and byte-prefix-preserving `incremental`; redaction and delete operations require `rewrite` and remain narrower than strict sanitize.

## Optional specialist providers

Install a specialist provider yourself only when the task requires a capability outside the MuPDF.js contract:

- ReportLab: greenfield layout-oriented creation.
- pdfplumber: table-oriented text and geometry extraction.
- pypdf: typed attachment quarantine, complex AcroForm appearance handling, and complete-source merge/reorder/stamp workflows.
- PyMuPDF 1.27.2.x: retained specialist adapter for strict scrub, residue/OCR gates, image-backed OCR redaction, and legacy high-level operations not yet migrated to JavaScript.
- Poppler: independent `pdfinfo` and `pdftoppm` file/render QA.
- qpdf 11+: separately installed structural diagnosis, recovery rewrite, and
  linearization through the shipped `scripts/qpdf_provider.py`; configure an
  exact executable with `OPEN_OFFICE_PDF_QPDF` when it is not on `PATH`.
- pyHanko 0.35.x core: source-bound local-PKCS#12 approval/certification signing
  through `scripts/pyhanko_sign_provider.py`, plus exact-source integrity,
  trust, difference, DocMDP, and FieldMDP validation through
  `scripts/pyhanko_provider.py`. TSA/LTV, PKCS#11, and remote signing remain
  external workflows.
- veraPDF 1.30.x: exact-source PDF/A and PDF/UA machine validation through the shipped `scripts/verapdf_provider.py` adapter.
- OCRmyPDF 17.8.x + Tesseract 5.x + qpdf 11+ + Poppler `pdftotext`:
  source-bound complete-document searchable-layer OCR through the shipped
  `scripts/ocrmypdf_provider.py` adapter. Tesseract also supplies image OCR
  evidence for high-trust sanitization.
- pikepdf 10.10.x: source-bound, fixed-profile active/auxiliary structure cleanup through the shipped `scripts/pikepdf_provider.py` adapter.

Install PyMuPDF and Tesseract only for the explicit specialist route. The
adapter accepts PyMuPDF `>=1.27.2,<1.28`; hosted CI pins `1.27.2.3`. Install the
requested Tesseract 5.x language data separately, then bind the virtual
environment executable as provider identity:

```bash
# macOS: brew install tesseract
# Debian/Ubuntu: sudo apt-get install tesseract-ocr tesseract-ocr-eng
uv venv .venv-pymupdf
uv pip install --python .venv-pymupdf/bin/python 'PyMuPDF>=1.27.2,<1.28'
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="$PWD/.venv-pymupdf/bin/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check \
  --provider pymupdf --require
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pymupdf_edit.py probe \
  --accept-license agpl --ocr-language eng --require-ocr
```

Choose `--accept-license commercial` instead only when the deployment has the
corresponding Artifex license. The OCR probe verifies the requested traineddata
before destructive work. `redact_ocr_text` additionally binds one page, its
expected 0/90/180/270-degree rotation, an exact term, and an expected
image-backed match count. Missing OCR, unsafe language names, rotation or match
drift, off-image-only results, and excessive raster work fail before output
publication. See [redact and sanitize](redact.md).

Probe qpdf through both the registry and its executable adapter:

```bash
# macOS/Homebrew: brew install qpdf
# Debian/Ubuntu: sudo apt-get install qpdf
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider qpdf --require
"$PYTHON_BIN" scripts/qpdf_provider.py probe
```

The wrapper requires qpdf JSON v2 (qpdf 11 or newer). It does not install qpdf,
accept passwords, expose arbitrary qpdf flags, or fall back to pikepdf. See
[inspect, repair, and linearize](repair_linearize.md).

Install and probe the bounded pikepdf structure-clean route separately:

```bash
uv venv .venv-pdf
uv pip install --python .venv-pdf/bin/python 'pikepdf>=10.10.0,<10.11.0'
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="$PWD/.venv-pdf/bin/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check \
  --provider pikepdf --require
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pikepdf_provider.py probe
```

The adapter exposes only read-only bounded inspection and two full-rewrite
profiles: `active-content` and `active-and-auxiliary`. It rejects encryption,
parser warnings, stale source hashes, missing trust/isolation declarations,
incremental output, and missing signature-invalidation acknowledgement. It
does not erase metadata, form values, XFA, comments, hidden/OCR text, or visible
page content and is not strict sanitize. See
[active and auxiliary structure cleanup](structure_clean.md).

Install and probe both pyHanko adapters in the same explicit provider environment:

```bash
uv venv .venv-pdf
uv pip install --python .venv-pdf/bin/python \
  'pyHanko>=0.35.0,<0.36.0' 'pyhanko-certvalidator>=0.31.0,<0.32.0'
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="$PWD/.venv-pdf/bin/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check \
  --provider pyhanko --require
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pyhanko_sign_provider.py probe
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pyhanko_provider.py probe
```

This installs no command-line signer; the bounded local-PKCS#12 adapter calls
pyHanko core directly. It requires exact source/credential hashes, explicit
trust or isolation, stdin/no-passphrase, one field/DocMDP choice, exact-prefix
incremental output, and post-sign validation. The read-only adapter accepts only
caller-supplied roots and disables network fetching. Neither claims complete
PAdES conformance. See [sign and verify](sign_verify.md).

Install veraPDF 1.30.x and a compatible Java runtime separately from its
[official distribution](https://docs.verapdf.org/install/). Configure the exact
launcher when it is not on `PATH`, then probe both the registry contract and
the stricter adapter contract:

```bash
export OPEN_OFFICE_PDF_VERAPDF="/absolute/path/to/verapdf"
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider verapdf --require
"$PYTHON_BIN" scripts/verapdf_provider.py probe
```

The adapter accepts veraPDF `>=1.30.0,<1.31.0`, requires all supported built-in
PDF/A/PDF/UA profiles, and exposes only one-file read-only validation. It does
not install Java or veraPDF, pass arbitrary flags, infer a profile, accept a
password/custom profile, repair a file, or fall back to another engine. See
[accessibility and archival conformance](accessibility.md).

Install and probe the searchable-layer OCR route separately:

```bash
# macOS: brew install ocrmypdf poppler
# Debian/Ubuntu: install tesseract-ocr, qpdf, and poppler-utils, then create an
# exact OCRmyPDF 17.8.x CLI environment when the distribution version differs.
export OPEN_OFFICE_PDF_OCRMYPDF="/absolute/path/to/ocrmypdf"
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider ocrmypdf --require
"$PYTHON_BIN" scripts/ocrmypdf_provider.py probe
```

Configure non-PATH components with `OPEN_OFFICE_PDF_TESSERACT`,
`OPEN_OFFICE_PDF_QPDF`, and `OPEN_OFFICE_PDF_PDFTOTEXT`. The adapter exposes
only full-rewrite `skip`/`redo`/`force` modes, fixed standard-PDF/O0/one-job
settings, explicit source/trust/loss preconditions, and no plugins or arbitrary
provider flags. It is not a sanitizer or host-level malware sandbox. See
[scanned-PDF OCR](ocr.md).

When Python tools are installed, set one interpreter as provider identity:

```bash
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="/absolute/path/to/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check --provider all
```

Every shipped Python entry point re-executes through that interpreter when configured. A missing capability must fail explicitly; never route to another provider or model reconstruction automatically. The `integration` field distinguishes shipped adapters, external documented tools, and planned routes.

If the provider lives in a virtual environment, configure that environment's
`bin/python` (or Windows `Scripts/python.exe`) directly. The runtime retains
that executable link instead of resolving it to a base interpreter, so
`pyvenv.cfg` and the environment's installed PyMuPDF/pypdf/etc. modules remain
active across probe, plan, mutation, residue scan, and audit.

The project and required MuPDF.js runtime are GNU AGPL-3.0-or-later. Optional tools retain their own upstream licenses and installation obligations; see the repository's third-party notices.
