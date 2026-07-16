---
name: "pdf"
description: "Create, read, review, edit, sign, redact, render, and verify PDF files through explicit capability-routed providers. Use for greenfield tagged authoring, imported-PDF native edits, forms, annotations, signatures, accessibility, sanitization, extraction, and visual QA."
---

# PDF Skill

## Architecture

PDF is independent from the OpenChestnut DOCX/XLSX/PPTX codec. Do not create an OpenChestnut PDF codec, put PDF into the Office protobuf/WASM wire, or build another universal PDF parser/writer here.

This Skill routes explicit operations to mature providers. The project-owned layer is deliberately thin: capability probes, raw-source provenance, transactional output, save policy, signature constraints, residue scanning, audit evidence, and semantic/structural/visual/security QA.

Never catch a provider failure and silently retry through another provider. There is no silent fallback. Never reconstruct an arbitrary PDF through PDF.js or `PdfArtifact`, export it, and describe the result as a fidelity-preserving edit.

## Choose The Route First

Read the [provider matrix](references/PROVIDER_MATRIX.md), [save policies](references/SAVE_POLICIES.md), [security checklist](references/SECURITY_CHECKLIST.md), and [product boundaries](references/PRODUCT_BOUNDARIES.md) before mutating an imported file.

| Need | Route |
| --- | --- |
| Greenfield tagged document, explicit reading order, inspect/verify | `PdfArtifact` / `PdfFile` |
| Greenfield visual/layout PDF | ReportLab |
| Text, words, geometry, table candidates | pdfplumber |
| Basic structure, forms, annotations, merge/split/stamp | pypdf |
| Existing-PDF page/content/image edits, forms, annotations, redaction/scrub | PyMuPDF |
| Native page/file evidence and final raster QA | Poppler |
| Structural diagnosis/recovery/rewrite | qpdf; pikepdf is planned but has no shipped adapter |
| Signing, timestamps/LTV, DocMDP/FieldMDP, signature validation | pyHanko |
| PDF/A or PDF/UA machine rules | veraPDF |
| Scanned-PDF OCR | OCRmyPDF is planned; strict image residue OCR currently uses separately installed Tesseract through PyMuPDF |

Probe before work:

```bash
python3 scripts/pdf_provider.py check --provider all
python3 scripts/pdf_provider.py plan \
  --task edit-content --provider pymupdf --strategy rewrite \
  --input input.pdf --output tmp/pdfs/edited.pdf \
  --accept-license agpl --require-provider
```

The probe reports whether a dependency exists and whether the integration is shipped, external/documented, or planned. Availability does not turn a planned provider into a shipped adapter.

## Preserve Source And Provenance

For every imported PDF:

1. Keep the source immutable and record absolute path, bytes, and SHA-256.
2. Inspect encryption, signatures, ByteRange, `/Perms`, DocMDP/FieldMDP, forms, annotations, attachments, metadata/XMP, images, OCR layers, active content, page boxes, and page count.
3. Select one provider and one save strategy: `rewrite`, `incremental`, or `sanitize`.
4. Write to a distinct transactional output path. Never overwrite the source during work.
5. Reopen the output independently, verify the intended delta, render every page with Poppler, and retain an audit record.

An incremental update preserves the exact old byte prefix by design. It does not prove that DocMDP permits the change or that the earlier signer endorses the new revision.

## Greenfield Creation

Use `PdfArtifact` for a new semantic/tagged document with explicit headings, Table/TR/TH/TD structure, Figure alternative text, reading order, inspect/resolve, and modeled verification. Start with the [API quick start](artifact_tool/API_QUICK_START.md) and [public end-to-end example](examples/public-api-end-to-end.mjs).

```js
import { PdfArtifact, PdfFile, verifyArtifact } from "open-office-artifact-tool";

const pdf = PdfArtifact.create({
  metadata: { title: "Readiness report", language: "en-US" },
  pages: [{ text: "Readiness report", width: 612, height: 792 }],
});
// Add headings, flow text, semantic tables, figures/charts, and reading order.
if (!verifyArtifact(pdf).ok) throw new Error("PDF semantic verification failed");
await (await PdfFile.exportPdf(pdf)).save("output/pdf/readiness-report.pdf");
```

Use ReportLab for a greenfield layout-oriented PDF when the `PdfArtifact` tagged contract is not required:

```bash
python3 scripts/reportlab_create.py \
  --spec tmp/pdfs/report-spec.json \
  --output output/pdf/report.pdf
```

ReportLab output still needs extraction, conformance checks when requested, and Poppler review. See [create](tasks/create.md).

## Read And Review

Read-only review does not require model re-export:

```bash
pdfinfo input.pdf
python3 scripts/pdfplumber_extract.py input.pdf \
  --output tmp/pdfs/extraction.json --max-pages 200
pdftoppm -png -r 144 input.pdf tmp/pdfs/pages/page
```

PDF.js through `createPdfjsParser()` is useful for agent-facing extraction, inspect, layout, and QA. Its result is a reconstructed view, not an editable native object graph. Extracted tables are heuristic candidates and must be checked against rendered geometry. See [read and review](tasks/read_review.md).

## Edit An Existing PDF

PyMuPDF is the explicit advanced provider selected for this project. It operates directly on the original bytes/file or a byte-identical transactional copy and supports bounded page operations, positioned text, image insertion/replacement, annotations, AcroForm values, real redactions, scrub, rewrite, and incremental save.

```bash
python3 scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

The shipped operation contract includes `insert_textbox`, `insert_image`, `replace_image`, `add_text_annotation`, `fill_form`, `rotate_page`, `delete_page`, `redact_text`, `redact_rect`, `replace_text`, and `scrub`. Unsupported operations fail; there is no fallback.

General Word-style reflow is not available for an ordinary imported PDF. `replace_text` is a bounded redaction plus same-box overlay and rejects a replacement that does not fit. Use a trusted source model or explicitly create a reconstructed new document when broad reflow is required. See [edit existing](tasks/edit_existing.md).

## Forms And Annotations

Use pypdf for basic AcroForm/annotation operations and PyMuPDF for advanced widget/page integration:

```bash
python3 scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
python3 scripts/pypdf_edit.py fill-form input.pdf tmp/pdfs/filled.pdf \
  --strategy incremental --field 'sender.city=Shanghai'
```

Inspect signatures, DocMDP, FieldMDP, and field locks first. Use `--allow-signed` only after policy review. Use `--flatten` only with `rewrite`, after confirming that interactivity should be removed. See [forms and annotations](tasks/forms_annotations.md).

## Sign And Verify

Use pyHanko for digital signatures, timestamps, trust validation, LTV/PAdES, DocMDP, and FieldMDP. PyMuPDF and pypdf are not complete signing backends.

Validate before and after every later revision. Report separately: cryptographic integrity, signer trust, time/revocation evidence, and whether the exact modification class is permitted. An older signature cannot be claimed to approve arbitrary new edits. See [sign and verify](tasks/sign_verify.md).

## Redact And Sanitize

High-trust redaction always uses `sanitize`:

```bash
python3 scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/sanitized.pdf \
  --strategy sanitize \
  --operations tmp/pdfs/redactions.json \
  --sensitive-term 'Customer Secret' \
  --accept-license agpl \
  --invalidate-signatures
```

The script adds/applies real redactions, runs strict PyMuPDF scrub, performs a full garbage-collected rewrite, proves the old byte prefix is absent, and scans raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations/widgets, image OCR, and revision pointers. Every `redact_text`/`replace_text` token must also be an explicit residue term.

Image-bearing pages require Tesseract-backed OCR. If OCR is unavailable or any scan is incomplete, the operation deletes the transactional output and fails closed. Incremental redaction, opaque overlays, and text-extraction-only checks are forbidden. See [redact and sanitize](tasks/redact.md).

## Accessibility And Conformance

Use `PdfArtifact` for greenfield tagged semantics and `pdf.verify()` plus `PdfFile.inspectPdf(...)` for its modeled contract. Use veraPDF for requested PDF/A/PDF/UA machine rules.

Automation cannot infer arbitrary author intent with certainty. Heading hierarchy, reading order, table meaning, alternative-text quality, link purpose, color/contrast, and other PDF/UA human checkpoints require review when confidence is low. See [accessibility](tasks/accessibility.md).

## Render Every Final Page

Final visual QA uses actual exported bytes through Poppler, never only a model/SVG preview:

```bash
pdfinfo output.pdf > tmp/pdfs/pdfinfo.txt
mkdir -p tmp/pdfs/pages
pdftoppm -png -r 144 output.pdf tmp/pdfs/pages/page
```

Inspect every page for clipping, overlap, blank pages, font substitution, missing glyphs, table overflow, image quality, form/widget appearances, annotations, signatures, redaction geometry, and unexplained page-box changes. Semantic verification can be green while content is visibly clipped. See [render and review](tasks/render_review.md).

## Dynamic And Opaque Features

Dynamic XFA, complex Acrobat JavaScript, 3D annotations, and RichMedia need application-specific runtimes. Default behavior is detect plus opaque preserve when the chosen non-destructive operation allows it. Flattening, specialist processing, or failure must be explicit. No shipped adapter executes these programs.

## Dependencies And License Boundary

The npm package remains MIT licensed and does not bundle Python providers. The selected PyMuPDF provider is separately installed under GNU AGPL or a commercial license. The user has approved the AGPL path for this project, but deployment and redistribution must still satisfy its terms.

```bash
uv pip install --python "$PYTHON" PyMuPDF==1.27.2.3
export OPEN_OFFICE_PDF_PYMUPDF_LICENSE=AGPL
"$PYTHON" scripts/pdf_provider.py check --provider pymupdf --require
```

See [provider setup](tasks/provider_setup.md). Do not add MuPDF.NET alongside PyMuPDF; both wrap the same MuPDF engine. A future .NET-only provider such as iText would be a separate evaluated integration, not the default architecture.

## Temp And Output Conventions

- Use `tmp/pdfs/` for source snapshots, transactional outputs, operation JSON, reports, PNGs, and diffs.
- Write final artifacts under `output/pdf/` in this repository unless the user specifies another location.
- Preserve source hashes and provider/version/save-policy evidence.
- Remove or restrict sensitive intermediate files after the user's data-retention requirements are satisfied.

## Final Gate

- Provider and exact capability were probed; no fallback occurred.
- Source bytes/hash/provenance and save strategy were recorded.
- Signature/DocMDP/FieldMDP policy was checked before mutation.
- Output reopened and the intended semantic/structural delta was verified.
- Sanitize jobs passed strict residue and single-revision gates, including image OCR.
- pyHanko and veraPDF evidence exists when signatures or conformance were requested.
- `pdfinfo` and Poppler rendering succeeded for every final page, followed by visual review.
- Final file and audit evidence are in the requested locations; no temporary file is presented as the deliverable.
