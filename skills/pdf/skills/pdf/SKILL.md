---
name: "pdf"
description: "Create, read, review, edit, sign, redact, render, and verify PDF files through explicit capability-routed providers. Use for greenfield tagged authoring, imported-PDF native edits, forms, annotations, signatures, accessibility, sanitization, extraction, and visual QA."
---

# PDF Skill

## Architecture

PDF is independent from the OpenChestnut DOCX/XLSX/PPTX codec. Do not create an OpenChestnut PDF codec, put PDF into the Office protobuf/WASM wire, or build another universal PDF parser/writer here.

This Skill exposes agent-usable PDF primitives. `PdfArtifact` owns greenfield semantic/tagged authoring. The required, runtime-lazy MuPDF.js dependency owns arbitrary-file parsing, native inspection, raster rendering, and bounded direct-original edits. Specialist providers remain explicit routes for strict sanitize/OCR, complex forms/merge, signatures, conformance, and independent QA.

Never catch a provider failure and silently retry through another provider. There is no silent fallback. Never reconstruct an arbitrary PDF through PDF.js or `PdfArtifact`, export it, and describe the result as a fidelity-preserving edit.

The default native CLI is `scripts/mupdf.mjs`. It calls `PdfFile` rather than copying PDF logic into the Skill:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
node scripts/mupdf.mjs render input.pdf tmp/pdfs/page-1.png --page 1 --dpi 144
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/operations.json tmp/pdfs/edited.pdf --save-policy rewrite
```

The CLI budgets input/page/object/render work, refuses source overwrite including symlink aliases, writes output atomically, and never falls back. `PdfFile.importPdf` additionally enforces per-image and cumulative decoded/retained-image budgets.

## Specialist Python Runtime Contract

Treat `OPEN_OFFICE_PDF_PROVIDER_PYTHON` as part of provider identity. When it is non-empty, every shipped Python entry point automatically re-executes through that exact interpreter before probing or importing a provider. Do not unset it, replace it after a failed system-Python probe, or infer availability from `which python3`; use one interpreter for probe, plan, mutation, residue scan, and audit. An invalid configured path fails closed. Only when the variable is absent does the current interpreter remain authoritative.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" -c 'import sys; print(sys.executable)'
```

## Choose The Route First

Read the [provider matrix](references/PROVIDER_MATRIX.md), [save policies](references/SAVE_POLICIES.md), [audit schema](references/AUDIT_SCHEMA.md), [security checklist](references/SECURITY_CHECKLIST.md), and [product boundaries](references/PRODUCT_BOUNDARIES.md) before mutating an imported file.

| Need | Route |
| --- | --- |
| Greenfield tagged document, explicit reading order, inspect/verify | `PdfArtifact` / `PdfFile` |
| Greenfield visual/layout PDF | ReportLab |
| Text, words, geometry, table candidates | pdfplumber |
| Basic structure, path-safe attachment quarantine, forms, annotations, merge/split/stamp | pypdf |
| Default arbitrary-PDF read/inspect/render and bounded native edit | `PdfFile` + MuPDF.js through `scripts/mupdf.mjs` |
| Strict scrub/residue/OCR and retained high-level edits not yet migrated | optional PyMuPDF specialist path |
| Native page/file evidence and final raster QA | Poppler |
| Structural diagnosis/recovery/rewrite | qpdf; pikepdf is planned but has no shipped adapter |
| Signing, timestamps/LTV, DocMDP/FieldMDP, signature validation | pyHanko |
| PDF/A or PDF/UA machine rules | veraPDF |
| Scanned-PDF OCR | OCRmyPDF is planned; strict image residue OCR currently uses separately installed Tesseract through PyMuPDF |

Probe and validate the route before work. The default MuPDF.js path uses this mandatory preflight:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
```

For an optional Python specialist mutation, the provider-specific probe and `pdf_provider.py plan` remain mandatory and must finish before mutation:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task edit-content --provider pymupdf --strategy rewrite \
  --input input.pdf --output tmp/pdfs/edited.pdf \
  --accept-license agpl --require-provider
```

The probe reports whether a dependency exists and whether the integration is shipped, external/documented, or planned. Availability does not turn a planned provider into a shipped adapter.

## Preserve Source And Provenance

For every imported PDF:

1. Keep the source immutable and record absolute path, bytes, and SHA-256.
2. Inspect encryption, signatures, ByteRange, `/Perms`, DocMDP/FieldMDP, forms, annotations, attachments, metadata/XMP, images, OCR layers, active content, page boxes, and page count.
3. Select one provider and one save strategy: `read-only`, `rewrite`, `incremental`, or `sanitize`.
4. Write to a distinct transactional output path. Never overwrite the source during work.
5. Reopen the output independently, verify the intended delta, render every page with Poppler, and retain an audit record.

An incremental update preserves the exact old byte prefix by design. It does not prove that DocMDP permits the change or that the earlier signer endorses the new revision.

Every mutation audit and security-sensitive read-only extraction audit uses the canonical `open-office-artifact-tool.pdf-audit.v1` envelope. Keep the exact camelCase fields `source`, `output`, `provider.actual`, `provider.version`, `provider.silentFallback`, `savePolicy.strategy`, `preflight`, `operation.type`, and `validation`; do not invent naming aliases. Before delivery, run `scripts/pdf_audit.py validate` against the source and output bytes. See the [audit schema](references/AUDIT_SCHEMA.md).

For a multi-source merge, bind `source` to the exact operation manifest and add canonical `inputs` records for every source PDF. Validate those records with repeated `pdf_audit.py validate --input ...` arguments.

## Greenfield Creation

Use `PdfArtifact` for a new semantic/tagged document with explicit headings, Table/TR/TH/TD structure, Figure alternative text, meaningful tagged links, running artifacts, reading order, inspect/resolve, and modeled verification. Start with the [API quick start](artifact_tool/API_QUICK_START.md) and [public end-to-end example](examples/public-api-end-to-end.mjs). For a six-page CJK report with a constrained cross-page logical table and separate modeled/veraPDF/human evidence, run [the accessible board report example](examples/accessible-board-report.mjs).

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

Use `artifact: true` for repeating headers and footers, `addLink(...)` for meaningful visible URI links, and a shared table `semanticId` only for consecutive-page segments whose continuation is first and non-final segment is last in page reading order. These constraints make the logical structure explicit instead of guessing from layout.

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

MuPDF.js is the default arbitrary-file parser and native inspector. It returns structured text geometry, raster placements, links, annotations, and widgets without turning that reconstruction into an edit graph. PDF.js through `createPdfjsParser()` remains an optional independent read adapter. Extracted tables are heuristic candidates and must be checked against rendered geometry. See [read and review](tasks/read_review.md).

For untrusted embedded files, use the typed read-only quarantine primitive. It inventories document-level and page-level attachments separately, keeps duplicate display names as distinct files, neutralizes path traversal and portable reserved names, enforces count/decoded-byte budgets, verifies every extracted SHA-256, and rechecks that the source PDF did not change. It never opens or executes a payload:

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pdf_provider.py check \
  --provider pypdf --require
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pdf_provider.py plan \
  --task extract-attachments --provider pypdf --strategy read-only \
  --input input.pdf --require-provider
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pypdf_edit.py extract-attachments input.pdf outputs/quarantine \
  --manifest outputs/attachments.json \
  --max-attachments 1000 --max-total-bytes 1073741824
```

Treat `attachments.json` as the authoritative mapping from raw display name/internal key/scope to the sanitized saved path. Do not derive output paths yourself, and do not inspect archive or executable contents unless a later explicitly sandboxed workflow requests it.
Bind the canonical audit `output` to `attachments.json`, set `savePolicy.strategy` to `read-only` and `operation.type` to `extract-attachments`, then run `pdf_audit.py validate --source input.pdf --artifact outputs/attachments.json --require-operation extract-attachments`.

## Edit An Existing PDF

MuPDF.js is the default native provider. It operates directly on original bytes and currently supports `add_text_annotation`, typed text/choice/checkbox `fill_form`, `delete_page`, complete `rearrange_pages`, `set_metadata`, `delete_embedded_file`, `delete_link`, `redact_text`, and `redact_rect`. Unsupported operations, untrusted radio export-value mapping, invalid limits, and unsafe save policies fail closed.

Use JSON operations with the thin CLI:

```bash
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

`incremental` preserves the exact source-byte prefix but is forbidden for redaction, delete operations, and any signed input. A signed rewrite requires deliberate invalidation, and the API reports signature validity as unknown; use pyHanko for actual trust and policy validation.

The optional PyMuPDF specialist script remains available for operations not yet migrated, including bounded positioned text/image workflows and the strict sanitize path. Run its provider probe and route plan above before use. For `replace_text` under `sanitize`, use `--task redact --strategy sanitize --invalidate-signatures` in the plan.

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

The specialist Python operation contract includes `insert_textbox`, `insert_image`, `replace_image`, `add_text_annotation`, `fill_form`, `rotate_page`, `delete_page`, `redact_text`, `redact_rect`, `replace_text`, and `scrub`. It is not a fallback from MuPDF.js; select it only when the task requires that explicit capability.

General Word-style reflow is not available for an ordinary imported PDF. `replace_text` is a bounded redaction plus same-box overlay for one horizontal source span: it preserves the source baseline and default style, reports its measured fit evidence, and allows only a fixed sub-millipoint numerical tolerance. Cross-span, rotated, or genuinely overflowing replacements fail closed. Use a trusted source model or explicitly create a reconstructed new document when broad reflow is required. See [edit existing](tasks/edit_existing.md).

## Merge, Reorder, And Stamp

Use the shipped pypdf manifest-driven primitive for a complete-source merge whose page order may split one source into multiple sequence segments while preserving internal links, outlines, named destinations, page boxes, and orientation. Transparent watermark rules select pages by source ID, not fragile output-page guesses.

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

The sequence must include every source page exactly once. The typed Poppler comparison maps every output page back to its declared source page, requires non-watermarked pages to remain pixel-identical, requires watermarked pages to change, and detects geometry, blank-page, or dark-background drift. Treat its structured pass/fail evidence as the visual delivery gate; a subjective thumbnail impression alone must not override a passing comparison. Encrypted input, ambiguous named-destination collisions, unresolved navigation, duplicate/omitted pages, unsupported rotated-page stamp placement, and unacknowledged signatures fail closed before output promotion. See [merge, reorder, and stamp](tasks/transform.md).

## Forms And Annotations

Use MuPDF.js for bounded text, choice, and checkbox values plus text annotations. Radio buttons currently fail closed because the JavaScript API does not expose a trustworthy widget-to-export-value mapping. Use the typed pypdf workflow for radio/checkbox appearance-state validation and more complex AcroForms:

```bash
python3 scripts/pypdf_edit.py inspect input.pdf \
  --output tmp/pdfs/pypdf-inspect.json
python3 scripts/pypdf_edit.py fill-form input.pdf tmp/pdfs/filled.pdf \
  --strategy incremental --field 'sender.city=Shanghai'
```

Inspect signatures, DocMDP, FieldMDP, field types, appearance states, and field locks first. Use `--allow-signed` only after policy review. Use `--flatten` only with `rewrite`, after confirming that interactivity should be removed. The shipped pypdf primitive resolves radio/checkbox values against actual appearance-state names and fails closed when post-write `/V` and widget `/AS` do not agree. See [forms and annotations](tasks/forms_annotations.md).

## Sign And Verify

Use pyHanko for digital signatures, timestamps, trust validation, LTV/PAdES, DocMDP, and FieldMDP. PyMuPDF and pypdf are not complete signing backends.

Validate before and after every later revision. Report separately: cryptographic integrity, signer trust, time/revocation evidence, and whether the exact modification class is permitted. An older signature cannot be claimed to approve arbitrary new edits. See [sign and verify](tasks/sign_verify.md).

## Redact And Sanitize

MuPDF.js can apply a real text/rectangle redaction during a full rewrite, but that primitive does not claim complete metadata, attachment, hidden-layer, OCR, or revision sanitization. High-trust redaction always uses the explicit specialist `sanitize` workflow:

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/sanitized.pdf \
  --strategy sanitize \
  --operations tmp/pdfs/redactions.json \
  --sensitive-term 'Customer Secret' \
  --accept-license agpl \
  --invalidate-signatures
```

The script adds/applies real redactions, runs strict PyMuPDF scrub, performs a full garbage-collected rewrite, proves the old byte prefix is absent, and scans raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations/widgets, image OCR, and revision pointers. Every `redact_text`/`replace_text` token must also be an explicit residue term.

For an inert public-release copy with no term redaction, use a scrub-only operation and the structural gate:

```bash
printf '[{"type":"scrub"}]' > tmp/pdfs/scrub.json
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/pymupdf_edit.py probe --accept-license agpl
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task sanitize --provider pymupdf --strategy sanitize \
  --input input.pdf --output tmp/pdfs/public-safe.pdf \
  --accept-license agpl --invalidate-signatures --require-provider
"$PYTHON_BIN" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/public-safe.pdf \
  --strategy sanitize --operations tmp/pdfs/scrub.json \
  --accept-license agpl --invalidate-signatures
"$PYTHON_BIN" scripts/residue_scan.py tmp/pdfs/public-safe.pdf \
  --require-inert --require-single-revision
```

This bounded primitive removes root/additional JavaScript, launch/submit actions, attachments, comments, populated widget defaults, personal metadata, and isolated invisible text before the provider scrub. After scrub it also physically removes active-content dictionary names that PyMuPDF represents as `null`; an unfamiliar object serialization fails closed instead of inviting a caller-side object rewrite. Invisible text that overlaps visible text fails closed because rectangle removal would damage the visible page. A scrub-only job does not invent a fake sensitive term; any redaction/replacement operation still requires explicit `--sensitive-term` values.

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

The project and required `mupdf@1.28.0` dependency are GNU AGPL-3.0-or-later. A normal npm installation resolves MuPDF.js; the root facade initializes its WASM runtime only on the first MuPDF-backed PDF operation, while an explicit `open-office-artifact-tool/pdf/mupdf` import intentionally initializes that provider. There is no PDF lifecycle hook, standalone downloader, or implicit Python installation.

Optional Python and system providers are installed separately only for a selected specialist workflow and keep their own upstream terms. See [provider setup](tasks/provider_setup.md). Do not install the unrelated Python `fitz` package. A future .NET-only provider would be a separately evaluated integration, not the default architecture.

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
- `pdf_audit.py validate` accepts the canonical audit and recomputed source/output hashes.
