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
- PyMuPDF: retained specialist adapter for strict scrub, residue/OCR gates, and legacy high-level operations not yet migrated to JavaScript.
- Poppler: independent `pdfinfo` and `pdftoppm` file/render QA.
- qpdf 11+: separately installed structural diagnosis, recovery rewrite, and
  linearization through the shipped `scripts/qpdf_provider.py`; configure an
  exact executable with `OPEN_OFFICE_PDF_QPDF` when it is not on `PATH`.
- pyHanko 0.35.x core: exact-source read-only signature integrity, trust,
  difference, DocMDP, and FieldMDP validation through the shipped
  `scripts/pyhanko_provider.py`; the separate `pyhanko-cli` package is needed
  only for external signing/timestamp/LTV command workflows.
- veraPDF: PDF/A and PDF/UA machine validation.
- Tesseract: OCR evidence for image-bearing high-trust sanitization.
- pikepdf and OCRmyPDF: planned routes without a shipped mutation adapter in this release.

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

Install and probe the pyHanko read-only validator in the same explicit provider
environment:

```bash
uv venv .venv-pdf
uv pip install --python .venv-pdf/bin/python \
  'pyHanko>=0.35.0,<0.36.0' 'pyhanko-certvalidator>=0.31.0,<0.32.0'
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="$PWD/.venv-pdf/bin/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check \
  --provider pyhanko --require
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pyhanko_provider.py probe
```

This installs no command-line signer. The adapter accepts only caller-supplied
trust roots, disables network fetching, never mutates the source, and does not
claim complete PAdES profile conformance. Install `pyhanko-cli` separately only
for a deliberately selected signing workflow. See [sign and verify](sign_verify.md).

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
