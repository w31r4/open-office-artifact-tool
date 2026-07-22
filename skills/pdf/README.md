# PDF

PDF is the file-type wrapper plugin for PDF artifact workflows.

This installable Skill bundle is distributed with `open-office-artifact-tool`.

## Included Skills

- `PDF`: read, create, render, verify, and extract content from PDF files.

## Discoverability

Use this plugin for PDF-oriented terms from the file-type naming model: PDF, PDFs, create PDF, edit PDF, render PDF, review PDF, extract PDF, tagged PDF, forms, annotations, signatures, redaction, accessibility, MuPDF.js, PDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, Poppler, qpdf, pikepdf, pyHanko, veraPDF, and `.pdf`.

## Source

The plugin tree is versioned directly under `skills/pdf` in the public repository.

## Compatibility status

The plugin is a reference-compatible capability router. `PdfArtifact` handles greenfield semantic/tagged authoring. `PdfFile` and the bundled `mupdf` npm dependency handle arbitrary-file reading, native inspection including source-bound widget/form-field snapshots, PNG/JPEG rendering, and bounded direct-original edits; `scripts/mupdf.mjs` is the thin Skill CLI over those APIs. The runtime stays lazy until the first PDF operation. Source-bound Text-note, unique-text Highlight, and safe-link placement use the inspected `mupdfPage.bbox` in explicit `mupdf-page-space` after the current 0/90/180/270-degree rotation, while raw `MediaBox`/`CropBox` facts remain unrotated PDF coordinates. Native `appearanceBbox` evidence prevents clipped review marks across renderers. A direct form update is intentionally narrow: one source-bound non-password text/compatible combo/checkbox widget only; shared or complex fields route explicitly to pypdf.

Existing PDFs stay as original bytes. No failure silently falls back to model reconstruction. Rewrite and byte-prefix-preserving incremental saves are distinct; redaction and delete operations reject incremental output because prior revisions retain the original content. MuPDF rewrite redaction is real removal from the new revision, but it is not the complete high-trust sanitize contract.

## Specialist providers

ReportLab, pdfplumber, pypdf, PyMuPDF, Poppler, qpdf, pikepdf, pyHanko,
veraPDF, OCRmyPDF, Tesseract, and related tools are explicit provider routes for
layout creation, complex tables/forms/merge, strict scrub and residue/OCR
evidence, signatures, conformance, searchable-layer OCR, and independent
visual QA. The public `open-office-artifact-tool/pdf/providers` resolver owns
their catalogued version, platform, licence, size, and runtime facts. It may
use only an explicitly selected deployment-owned `system-only` runtime or a
policy-authorized, hash-pinned managed pack; it never invokes a package manager,
global Python install, lifecycle hook, unpinned URL, or silent fallback.

The bundle ships bounded adapters for strict PyMuPDF sanitize and image-backed
`redact_ocr_text`, qpdf source-bound structure inspection/recovery/linearization,
pikepdf source-bound active/auxiliary structure cleanup, pyHanko exact-source
local-PKCS#12 signing and signature validation, veraPDF exact-source PDF/A/PDF/UA
machine-rule validation, and OCRmyPDF exact-source complete-document searchable-
layer generation. Non-MuPDF managed release assets are not published yet, so
such a managed route currently reports `blocked` rather than promising an
installation. The OCR redaction primitive requires one explicit page, an exact
expected 0/90/180/270-degree page rotation, an exact term, expected image-backed
match count, full sanitize rewrite, rotation-aware residue scan, and render QA;
it never broadens an uncertain OCR result. The shipped signer supports one
bounded local-PKCS#12 approval/certification revision; TSA/LTV, PKCS#11, remote
signing, and complete PAdES conformance remain external. The pikepdf adapter is
a fixed full-rewrite `structure-clean` route, not redaction, metadata/form/XFA
cleanup, strict sanitize, or a malware sandbox. The project and required
MuPDF.js dependency are licensed under GNU AGPL-3.0-or-later.
