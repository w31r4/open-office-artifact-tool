# PDF

PDF is the file-type wrapper plugin for PDF artifact workflows.

This installable Codex plugin is distributed with `open-office-artifact-tool`.

## Included Skills

- `PDF`: read, create, render, verify, and extract content from PDF files.

## Discoverability

Use this plugin for PDF-oriented terms from the file-type naming model: PDF, PDFs, create PDF, edit PDF, render PDF, review PDF, extract PDF, tagged PDF, forms, annotations, signatures, redaction, accessibility, PDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, Poppler, qpdf, pyHanko, veraPDF, and `.pdf`.

## Source

The plugin tree is versioned directly under `skills/pdf` in the public repository.

## Compatibility status

The plugin is a reference-compatible provider router rather than a single backend. `PdfArtifact`/`PdfFile` handle greenfield semantic/tagged authoring and QA. ReportLab, pdfplumber, pypdf, and the project-approved optional PyMuPDF provider have shipped thin scripts for their bounded roles; Poppler remains final native render QA. qpdf, pyHanko, and veraPDF are documented/probed external tools. pikepdf and OCRmyPDF are planned providers without shipped adapters in this release.

Existing PDFs stay as original bytes and go directly to an explicitly selected provider. No provider failure silently falls back to model reconstruction. Rewrite, incremental, and sanitize are distinct save contracts; high-trust redaction requires a full scrub/rewrite, residue/OCR checks, and final page rendering.
