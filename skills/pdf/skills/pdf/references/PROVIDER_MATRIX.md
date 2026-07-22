# PDF provider matrix

This is the human capability boundary. It intentionally does **not** duplicate
versions, package sizes, hashes, URLs, platform availability, or installation
facts: `open-office-artifact-tool/pdf/providers` owns those in its versioned
catalog and policy resolver.

Choose one provider before touching a file. A provider error is not permission
to retry through another route.

| Provider | Primary role | Save policy | Important boundary |
| --- | --- | --- | --- |
| `PdfArtifact` | New tagged semantic authoring, reading order, inspect/verify | `rewrite` | New/trusted model only; never an imported-PDF fidelity editor. |
| MuPDF.js / `PdfFile` | Default arbitrary-PDF parse, inspect, render, bounded direct-original edit | `read-only`, `rewrite`, explicit `incremental` | No Word-style reflow, complete sanitize, or signature authority. Source-bound edits and page evidence remain mandatory. |
| ReportLab | New visual/layout PDF | `rewrite` | Does not inherit the `PdfArtifact` tagged/reading-order contract. |
| pdfplumber | Read-only text, word geometry, tables, lines, rectangles | `read-only` | Extraction is evidence, not an edit representation. |
| pypdf | Attachment quarantine, complete-source merge/reorder/stamp, complex forms/annotations | `read-only`, `rewrite`, explicit `incremental` | Inspect signatures/DocMDP first; an incremental layout is not authorization. |
| PyMuPDF | Strict scrub/residue/OCR redaction and selected advanced bounded edits | explicit `rewrite`, `incremental`, `sanitize` | Explicit specialist only. Sanitization is full-rewrite and residue-scanned; it is not the signing authority. |
| Poppler | Independent file evidence and native raster QA | `read-only` | Renderer/inspector only, not an editor or conformance validator. |
| qpdf | Structural diagnosis, recovery rewrite, linearization, one source-to-AES-256 delivery copy | `read-only`, `rewrite` | AES-256 copy only from an unencrypted source with caller-owned restricted password files; not an encrypted-input opener/decrypter, permission editor, renderer, sanitizer, or signature validator. |
| pikepdf | Fixed-profile active/auxiliary structure cleanup | `read-only`, `rewrite` | Not redaction, metadata/form/XFA cleanup, strict sanitize, rendering, or malware isolation. |
| pyHanko | Local PKCS#12 signing and exact-source signature validation | `read-only`, bounded `incremental` | Certificates, keys, HSMs, TSA/LTV, remote signing, trust roots, and online retrieval remain explicit caller concerns. |
| veraPDF | PDF/A and PDF/UA machine-rule validation | `read-only` | One explicit profile; a green report is not repair or universal accessibility certification. |
| OCRmyPDF / Tesseract | Complete-document searchable layer and OCR residue evidence | `rewrite` / `read-only` | Source-bound rewrite, explicit language data, isolation; not proof of OCR accuracy, sanitization, or PDF/UA repair. |

## Mandatory routing rules

1. Preserve original bytes and open them directly through the selected provider.
   Do not reconstruct an arbitrary PDF and describe it as a faithful edit.
2. Declare `read-only`, `rewrite`, `incremental`, or `sanitize` before work.
3. Resolve/probe the exact provider and capability first. Missing provider,
   unsafe strategy, unsupported operation, encryption, or signature restriction
   blocks the task.
4. Keep output distinct from input. Reopen, verify the intended delta, render
   every final page, and retain audit evidence.
5. `incremental` preserves old bytes; it is not signature authorization.
   Redaction and delete operations must not be incremental.
6. High-trust redaction requires the explicit sanitize route: real redactions,
   scrub, full rewrite, residue/single-revision checks, then render review.

## Delivery and licensing boundary

MuPDF.js is required through normal npm resolution and remains runtime-lazy.
All other routes are selected by the capability resolver: they are either a
ready explicitly provisioned `system-only` runtime, an authorized immutable
managed pack, or a clear `blocked` result. The resolver reports licence
acknowledgements before installation; it never silently downloads a provider or
obtains secrets. See [provider setup](../tasks/provider_setup.md).
