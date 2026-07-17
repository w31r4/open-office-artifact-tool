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
- qpdf: structural diagnosis, recovery, encryption, and rewrite.
- pyHanko: signing, trust validation, timestamps, LTV/PAdES, DocMDP, and FieldMDP.
- veraPDF: PDF/A and PDF/UA machine validation.
- Tesseract: OCR evidence for image-bearing high-trust sanitization.
- pikepdf and OCRmyPDF: planned routes without a shipped mutation adapter in this release.

When Python tools are installed, set one interpreter as provider identity:

```bash
export OPEN_OFFICE_PDF_PROVIDER_PYTHON="/absolute/path/to/python"
"$OPEN_OFFICE_PDF_PROVIDER_PYTHON" scripts/pdf_provider.py check --provider all
```

Every shipped Python entry point re-executes through that interpreter when configured. A missing capability must fail explicitly; never route to another provider or model reconstruction automatically. The `integration` field distinguishes shipped adapters, external documented tools, and planned routes.

The project and required MuPDF.js runtime are GNU AGPL-3.0-or-later. Optional tools retain their own upstream licenses and installation obligations; see the repository's third-party notices.
