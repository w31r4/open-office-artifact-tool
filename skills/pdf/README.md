# PDF

PDF is the file-type wrapper plugin for PDF artifact workflows.

This installable Skill bundle is distributed with `open-office-artifact-tool`.

## Included Skills

- `PDF`: read, create, render, verify, and extract content from PDF files.

## Discoverability

Use this plugin for PDF-oriented terms from the file-type naming model: PDF, PDFs, create PDF, edit PDF, render PDF, review PDF, extract PDF, tagged PDF, forms, annotations, signatures, redaction, accessibility, MuPDF.js, PDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, Poppler, qpdf, pyHanko, veraPDF, and `.pdf`.

## Source

The plugin tree is versioned directly under `skills/pdf` in the public repository.

## Compatibility status

The plugin is a reference-compatible capability router. `PdfArtifact` handles greenfield semantic/tagged authoring. `PdfFile` and the bundled `mupdf` npm dependency handle arbitrary-file reading, native inspection including source-bound widget/form-field snapshots, PNG/JPEG rendering, and bounded direct-original edits; `scripts/mupdf.mjs` is the thin Skill CLI over those APIs. The runtime stays lazy until the first PDF operation. A direct form update is intentionally narrow: one source-bound non-password text/compatible combo/checkbox widget only; shared or complex fields route explicitly to pypdf.

Existing PDFs stay as original bytes. No failure silently falls back to model reconstruction. Rewrite and byte-prefix-preserving incremental saves are distinct; redaction and delete operations reject incremental output because prior revisions retain the original content. MuPDF rewrite redaction is real removal from the new revision, but it is not the complete high-trust sanitize contract.

## Specialist providers

ReportLab, pdfplumber, pypdf, PyMuPDF, Poppler, qpdf, pyHanko, veraPDF, Tesseract, and related tools remain explicit external routes for layout creation, complex tables/forms/merge, strict scrub and residue/OCR evidence, signatures, conformance, and independent visual QA. The bundle ships a bounded `qpdf_provider.py` adapter for source-bound structure inspection, recovery rewrite, and linearization, plus a read-only `pyhanko_provider.py` adapter for exact-source signature integrity, trust, difference, and DocMDP evidence. Their qpdf 11+ executable and pyHanko 0.35.x Python runtime remain separately installed; signature creation remains an external pyHanko workflow. These tools are never installed by a lifecycle hook or chosen as a silent fallback. The project and required MuPDF.js dependency are licensed under GNU AGPL-3.0-or-later.
