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

Treat `OPEN_OFFICE_PDF_PROVIDER_PYTHON` as part of provider identity. When it is non-empty, every shipped Python entry point automatically re-executes through that exact interpreter before probing or importing a provider. Point it at the virtual environment executable itself (`bin/python` or `Scripts/python.exe`): the runtime deliberately preserves that link instead of dereferencing its base interpreter, so the environment's `pyvenv.cfg` and installed provider modules stay active. Do not unset it, replace it after a failed system-Python probe, or infer availability from `which python3`; use one interpreter for probe, plan, mutation, residue scan, and audit. An invalid configured path fails closed. Only when the variable is absent does the current interpreter remain authoritative.

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
| Strict scrub/residue/OCR, image-backed OCR redaction, and retained high-level edits not yet migrated | optional PyMuPDF 1.27.2.x specialist path |
| Native page/file evidence and final raster QA | Poppler |
| Structural diagnosis, recovery rewrite, and linearization | separately installed qpdf through `scripts/qpdf_provider.py` |
| Bounded active/auxiliary structure cleanup without page reconstruction | separately installed pikepdf 10.10.x through `scripts/pikepdf_provider.py` |
| Signature inventory plus integrity/trust/difference/DocMDP validation | pyHanko core through `scripts/pyhanko_sign_provider.py` and `scripts/pyhanko_provider.py` |
| Source-bound local PKCS#12 approval/certification signature | `scripts/pyhanko_sign_provider.py`; incremental output, explicit field/DocMDP policy, passphrase on stdin |
| Timestamp-authority, LTV, PKCS#11, or remote signing | explicit external pyHanko workflow |
| PDF/A or PDF/UA machine rules | separately installed veraPDF 1.30.x through `scripts/verapdf_provider.py` |
| Scanned-PDF searchable text layer | separately installed OCRmyPDF 17.8.x through source-bound `scripts/ocrmypdf_provider.py`; strict redaction residue OCR remains the PyMuPDF/Tesseract sanitize route |

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

## Structural Repair And Linearization

The shipped qpdf wrapper is a thin external-provider boundary, not another PDF
model. `inspect` binds the source SHA-256 and reports qpdf warnings, page/object/
form/annotation/attachment counts, tagging, encryption/linearization state, and signature/ByteRange/
DocMDP evidence. `rewrite` accepts only that fresh source SHA-256, works on a
private snapshot, and atomically publishes a distinct clean output:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/qpdf_provider.py probe
"$PYTHON_BIN" scripts/qpdf_provider.py inspect input.pdf > tmp/pdfs/qpdf-inspect.json
"$PYTHON_BIN" scripts/qpdf_provider.py rewrite input.pdf outputs/repaired.pdf \
  --mode repair --expected-sha256 '<sha256-from-inspect>'
```

Use `--mode linearize` only when linearization is the requested rewrite
postcondition. Signature evidence requires explicit `--invalidate-signatures`
after pyHanko/DocMDP review; encrypted rewrites fail closed. qpdf repair does
not remove active content, metadata, attachments, hidden/OCR text, or sensitive
content and must never be described as sanitize or redaction. See
[inspect, repair, and linearize](tasks/repair_linearize.md).

## Active And Auxiliary Structure Cleanup

Use the pikepdf adapter only for its bounded `structure-clean` role. It opens a
private snapshot of the exact imported PDF and exposes two fixed profiles:
`active-content` removes JavaScript, external actions, and multimedia;
`active-and-auxiliary` additionally removes attachments, thumbnails, search
indexes, Web Capture data, private page-piece data, and portfolio presentation.
Neither profile reconstructs visible pages through `PdfArtifact`.

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA="$($PYTHON_BIN -c 'import hashlib,sys; print(hashlib.sha256(open(sys.argv[1], "rb").read()).hexdigest())' input.pdf)"
"$PYTHON_BIN" scripts/pikepdf_provider.py probe
"$PYTHON_BIN" scripts/pikepdf_provider.py inspect input.pdf \
  --expected-sha256 "$SOURCE_SHA" --trusted-input
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task structure-clean --provider pikepdf --strategy rewrite \
  --input input.pdf --output outputs/structure-clean.pdf \
  --invalidate-signatures --require-provider
"$PYTHON_BIN" scripts/pikepdf_provider.py clean \
  input.pdf outputs/structure-clean.pdf \
  --profile active-and-auxiliary --expected-sha256 "$SOURCE_SHA" \
  --trusted-input --invalidate-signatures
```

This is always a source-bound full rewrite with no provider fallback. It is not
redaction, metadata scrub, form/XFA flattening, strict sanitize, signature
validation, or a malware sandbox. DocumentInfo/XMP metadata, form values, XFA,
annotations, hidden/OCR text, and signature appearances remain outside the
primitive. Use caller-managed isolation for attacker-chosen input, then run
qpdf inspection, the intended residue policy, pyHanko when applicable, and a
full Poppler render comparison. See [active and auxiliary structure cleanup](tasks/structure_clean.md).

## Scanned-PDF OCR

The shipped OCRmyPDF adapter owns the bounded searchable-layer route for one
complete imported PDF. It requires the exact source SHA-256 and a distinct
absent output, works only on a private source snapshot, fixes OCRmyPDF to
`--output-type pdf --optimize 0 --jobs 1` with Tesseract/fpdf2/pypdfium, and
uses qpdf plus Poppler `pdftotext` as independent delivery gates:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/ocrmypdf_provider.py probe
"$PYTHON_BIN" scripts/ocrmypdf_provider.py ocr \
  scanned.pdf outputs/scanned-searchable.pdf \
  --expected-sha256 '<sha256-from-fresh-inspect>' \
  --mode skip --language eng --input-trust trusted \
  --require-text 'expected phrase'
```

The only modes are explicit `skip`, `redo`, and `force`. Tagged input and
`redo`/`force` require a structure-loss acknowledgement; `force` additionally
requires rasterize-all and, when forms or annotations exist, interactive-flattening
acknowledgements. Signature evidence requires prior pyHanko/DocMDP review plus
explicit invalidation. Encrypted files, stale hashes, missing language packs,
unexpected topology changes, empty/unmatched text, incremental prefixes, and
provider/budget failures stop without output. OCRmyPDF is not a sanitizer and
is not designed as a malware boundary: the caller must declare trusted input or
an already isolated execution environment. See [scanned-PDF OCR](tasks/ocr.md),
then Poppler-render and manually review every final page.

## Edit An Existing PDF

MuPDF.js is the default native provider. It operates directly on original bytes and currently supports source-bound `add_text_annotation` and `add_text_highlight`, legacy text/choice/checkbox `fill_form`, source-bound `update_form_field`, `delete_page`, source-bound `delete_annotation` and `update_annotation`, complete `rearrange_pages`, raw-coordinate visible-only `set_page_crop`, absolute-quarter-turn `rotate_page`, `set_metadata`, `delete_embedded_file`, source-bound `add_link`, `delete_link`, and URL-only `update_link`, `redact_text`, and `redact_rect`. Native inspection distinguishes two coordinate systems: raw `mediaBox`/`cropBox` are unrotated PDF-space facts, while `mupdfPage.bbox` carries `coordinateSpace: "mupdf-page-space"`, uses an upper-left origin with y increasing downward, and already reflects the current 0/90/180/270-degree page rotation. `add_text_annotation`, `add_text_highlight`, and `add_link` bind the exact source SHA-256 plus that bbox/rotation snapshot and use only that page space. `add_text_annotation` accepts a visible `[x,y]` pin and non-empty `contents` (optional non-empty author/subject). MuPDF owns the normalized Text-note icon rectangle, and the audit also returns a conservative `appearanceBbox` because native `NoZoom`/`NoRotate` note flags can make renderers paint a larger transformed footprint. Requested `bbox`/`rect`, `text` aliases, icon selection, clipped appearance, stale page evidence, and incremental saves fail closed. `add_text_highlight` binds one **unique native text selection**, accepts at most 4,096 characters plus optional RGB/review metadata, and never accepts caller quadrilaterals or rectangles. Zero/multiple native hits, an out-of-window native `appearanceBbox`, stale evidence, and incremental output fail closed. MuPDF verifies one native Highlight plus provider quadrilateral/color/appearance facts before rewrite delivery. `add_link` accepts only an in-page `[x,y,width,height]` and internal `#...` or absolute `http`, `https`, or `mailto` destination. All three placement primitives support right-angle rotated pages, report `coordinateSpace` and `pageRotation`, require rewrite, and must be re-inspected/rendered. `update_form_field` uses the inspect-returned `summary.sourceSha256`, `mupdf-form-field-<xref>` locator, and full field `snapshot`; it permits exactly one non-password text widget, one non-multiselect combo whose display/export options are identical, or one checkbox. The field update may use unsigned byte-prefix-verified `incremental` save, whereas Text annotation/highlight creation, annotation update/deletion, and all three link operations require rewrite only. `delete_annotation` uses `mupdf-annotation-<page>-<xref>` with type/content/name/author/subject/rectangle facts; `update_annotation` may patch only non-empty `contents`, `author`, and `subject`. A Text annotation rectangle is a snapshot precondition, not mutable geometry: MuPDF normalizes that native geometry, so use an explicit delete-plus-add route or a specialist provider to reposition it. `delete_link` and URL-only `update_link` use `mupdf-link-<page>-<fingerprint>` with URL/rectangle/external facts. Link bounds are snapshot evidence rather than a mutable patch: MuPDF's native bounds setter has save/reload coordinate semantics that are not a stable public contract, so move a link by one source-bound `delete_link` + `add_link` rewrite transaction or use a specialist provider. No locator is a persistent document identity, so re-inspect after every output. Shared-widget fields, radio/list/multi-select fields, password fields, mismatched choice export values, stale snapshots, unsupported operations, duplicate visible link additions, rotated-page crop requests, invalid limits, and unsafe save policies fail closed; route complex forms explicitly to pypdf.

Use JSON operations with the thin CLI:

```bash
node scripts/mupdf.mjs edit input.pdf tmp/pdfs/edit-operations.json tmp/pdfs/edited.pdf \
  --save-policy rewrite
```

`set_page_crop` accepts a `[x, y, width, height]` box in the raw unrotated PDF page coordinate system reported by `inspect`; it must lie fully inside `MediaBox`. It changes only `CropBox`, keeps content outside the visible region in the file, and is therefore not redaction or sanitize. `rotate_page` accepts an absolute `0`, `90`, `180`, or `270` degree clockwise `/Rotate` value, reports the prior normalized value, and changes viewer orientation without transforming or deleting page content. Both are unsigned non-destructive operations that may use byte-prefix-verified `incremental` save; it remains forbidden for redaction, delete operations, and any signed input. A signed rewrite requires deliberate invalidation, and the API reports signature validity as unknown; use pyHanko for actual trust and policy validation.

The optional PyMuPDF specialist script remains available for operations not yet migrated, including bounded positioned text/image workflows and the strict sanitize path. Run its provider probe and route plan above before use. For `replace_text` under `sanitize`, use `--task redact --strategy sanitize --invalidate-signatures` in the plan.

```bash
"${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}" scripts/pymupdf_edit.py edit input.pdf tmp/pdfs/edited.pdf \
  --strategy rewrite \
  --operations tmp/pdfs/edit-operations.json \
  --accept-license agpl
```

The specialist Python operation contract includes `insert_textbox`, `insert_image`, `replace_image`, `add_text_annotation`, `fill_form`, `delete_page`, `redact_text`, `redact_ocr_text`, `redact_rect`, `replace_text`, and `scrub`. It is not a fallback from MuPDF.js; select it only when the task requires that explicit capability.

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

Use the shipped source-bound pyHanko adapters for local PKCS#12 signing and
read-only validation. Inspect the exact source first, then sign to a distinct
path. The credential passphrase is accepted only on stdin (or the caller must
explicitly declare an unencrypted PKCS#12); never put it in argv, environment
variables, reports, logs, or repository files:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 input.pdf | awk '{print $1}')"
CREDENTIAL_SHA256="$(shasum -a 256 /secure/signer.p12 | awk '{print $1}')"
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py probe
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py inspect input.pdf \
  --expected-sha256 "$SOURCE_SHA256" --trusted-input \
  > tmp/pdfs/signature-inventory.json
"$PYTHON_BIN" scripts/pyhanko_sign_provider.py sign \
  input.pdf tmp/pdfs/signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" --trusted-input \
  --credential /secure/signer.p12 \
  --credential-sha256 "$CREDENTIAL_SHA256" --passphrase-stdin \
  --field-name Approval --field-mode create-invisible \
  --signature-kind approval --expected-signature-count 0 \
  > tmp/pdfs/signing-report.json
# A terminal gets a hidden prompt. Automation pipes stdin directly from its
# secret manager; never stage the secret in argv, env, or a file.
```

For a visible field, add `--field-mode create-visible --page-index 0 --box
x1,y1,x2,y2`. Certification requires `--signature-kind certification` plus an
explicit `--docmdp-permission no-changes|fill-forms|annotate`,
and is rejected unless it is the first signature. A later approval signature
requires both the exact expected signature count and
`--allow-existing-signatures`; that acknowledgement does not claim an earlier
signer approved the new revision. The signing adapter preserves the exact old
byte prefix, validates every resulting signature before atomic no-replace
promotion, and emits source/output/credential identity without secret material.

Then validate the exact output under an explicit trust policy:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 tmp/pdfs/signed.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/pyhanko_provider.py probe
"$PYTHON_BIN" scripts/pyhanko_provider.py verify tmp/pdfs/signed.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --trust-policy explicit-roots \
  --trust-root /trusted/root-ca.pem \
  --require-signature --require-all-integrity-valid \
  --require-all-trusted --require-docmdp-compliant \
  --require-all-bottom-line \
  > tmp/pdfs/signature-validation.json
```

The validator accepts only explicit trust roots, never fetches network evidence,
works on a private immutable snapshot, and reports ByteRange coverage,
cryptographic integrity, certificate trust, timestamps, modification level,
DocMDP/FieldMDP, and policy gates separately. The output explicitly does not
claim complete PAdES profile conformance. The shipped signer supports only a
local PKCS#12 credential and SHA-256 detached/PAdES subfilters. Timestamping,
LTV updates, PKCS#11, remote signing, and online revocation retrieval remain
explicit external pyHanko workflows;
PyMuPDF, pypdf, MuPDF.js, and qpdf are not complete signing backends.

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

The script adds/applies real redactions, runs strict PyMuPDF scrub, performs a full garbage-collected rewrite, proves the old byte prefix is absent, and scans raw bytes, decoded objects/streams, extracted text, metadata/XMP, attachments, annotations/widgets, image OCR, and revision pointers. Every `redact_text`/`redact_ocr_text`/`replace_text` token must also be an explicit residue term. For raster-only text, `redact_ocr_text` requires one page, an explicit `expected_rotation` of 0, 90, 180, or 270 degrees, an exact term, expected match count, Tesseract language data, bounded 72–300 dpi work, and at least 90% overlap with a native image placement. OCR and residue scanning temporarily normalize the page orientation while all match/redaction coordinates remain in canonical unrotated PyMuPDF page space; the source `/Rotate` value is restored and reported before mutation continues. Any precondition drift or OCR uncertainty fails before publication.

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

Image-bearing pages require Tesseract-backed OCR. Probe it explicitly with `pymupdf_edit.py probe --accept-license agpl --ocr-language eng --require-ocr`. If OCR is unavailable or any scan is incomplete, the operation deletes the transactional output and fails closed. Incremental redaction, opaque overlays, OCR match-count drift, and text-extraction-only checks are forbidden. See [redact and sanitize](tasks/redact.md).

## Accessibility And Conformance

Use `PdfArtifact` for greenfield tagged semantics and `pdf.verify()` plus `PdfFile.inspectPdf(...)` for its modeled contract. Use the shipped source-bound veraPDF adapter for requested PDF/A/PDF/UA machine rules:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 output.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/verapdf_provider.py probe
"$PYTHON_BIN" scripts/verapdf_provider.py validate output.pdf \
  --expected-sha256 "$SOURCE_SHA256" --flavour ua1 --require-compliant \
  > tmp/pdfs/verapdf-ua1.json
```

Choose exactly one built-in profile; the adapter never accepts veraPDF's automatic profile choice, custom profiles, passwords, directories, or arbitrary provider flags. A noncompliant result is still a completed validation unless `--require-compliant` makes it a delivery gate. Preserve the typed report, selected profile, provider component versions, source hash, and failed-rule evidence.

Automation cannot infer arbitrary author intent with certainty. Heading hierarchy, reading order, table meaning, alternative-text quality, link purpose, color/contrast, and other PDF/UA checkpoints require human review when confidence is low. See [accessibility](tasks/accessibility.md).

## Render Every Final Page

Final visual QA uses actual exported bytes through Poppler, never only a model/SVG preview:

```bash
pdfinfo output.pdf > tmp/pdfs/pdfinfo.txt
mkdir -p tmp/pdfs/pages
pdftoppm -png -r 144 output.pdf tmp/pdfs/pages/page
```

Inspect every page for clipping, overlap, blank pages, font substitution, missing glyphs, table overflow, image quality, form/widget appearances, annotations, signatures, redaction geometry, and unexplained page-box changes. Semantic verification can be green while content is visibly clipped. See [render and review](tasks/render_review.md).

## Dynamic And Opaque Features

Dynamic XFA, complex Acrobat JavaScript, 3D annotations, and RichMedia need application-specific runtimes. Default behavior is detect plus opaque preserve when the chosen non-destructive operation allows it. The explicit pikepdf `structure-clean` route can defang JavaScript, external actions, multimedia, and selected auxiliary structures without executing them; it neither flattens XFA/forms nor erases metadata or hidden text. Flattening, another specialist route, or failure must remain explicit. No shipped adapter executes these programs.

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
- The typed pyHanko report passes every requested integrity/trust/DocMDP gate when signatures were requested; the typed source-bound veraPDF report passes the explicitly requested machine-rule gate when conformance was requested, with separate human PDF/UA review.
- `pdfinfo` and Poppler rendering succeeded for every final page, followed by visual review.
- Final file and audit evidence are in the requested locations; no temporary file is presented as the deliverable.
- `pdf_audit.py validate` accepts the canonical audit and recomputed source/output hashes.
