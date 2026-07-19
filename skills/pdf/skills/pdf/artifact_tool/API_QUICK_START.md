# PDF API quick start

Use `open-office-artifact-tool` for greenfield modeled PDF creation, semantic editing of a trusted internal model, tagged export, inspect, resolve, render, and verification. The PDF pipeline is independent from OpenChestnut and does not require Microsoft Office, LibreOffice, or a local .NET SDK. Do not use this model as a fidelity-preserving mutation path for an arbitrary imported PDF.

For an executable six-page accessible report with H1-H3, a Figure alt description, meaningful Link annotation/OBJR, running Artifact text, a constrained cross-page logical Table, CJK font embedding, Poppler rendering, and separate modeled/veraPDF/human evidence, use [`../examples/accessible-board-report.mjs`](../examples/accessible-board-report.mjs).

## Startup

Use a supported Node.js runtime and resolve the installed package through standard Node.js module resolution. Work in a writable task directory and use ES modules.

```js
import {
  FileBlob,
  PdfArtifact,
  PdfFile,
  verifyArtifact,
} from "open-office-artifact-tool";
```

## Create an accessible PDF

```js
const pdf = PdfArtifact.create({
  metadata: { title: "Readiness report", language: "en-US" },
  pages: [{
    text: "Readiness report\nDecision evidence for the launch gate",
    width: 612,
    height: 792,
  }],
});

const decision = pdf.addText("Decision", {
  bbox: [72, 150, 180, 24],
  fontSize: 18,
  bold: true,
  headingLevel: 2,
});
const recommendation = pdf.addText("Approve the controlled rollout.", {
  bbox: [72, 180, 420, 18],
  fontSize: 12,
});
const evidence = pdf.addTable({
  id: "readiness-evidence",
  name: "readiness-evidence",
  values: [
    ["Gate", "Owner", "Status"],
    ["Model", "Artifact Platform", "Pass"],
    ["Native render", "Release QA", "Pending"],
  ],
  bbox: [72, 230, 468, 96],
});
const trend = pdf.addChart({
  name: "readiness-trend",
  title: "Readiness by gate",
  alt: "Bar chart showing readiness increasing from 76 to 96 percent.",
  chartType: "bar",
  categories: ["Model", "Package", "Render"],
  series: [{ name: "Readiness", values: [76, 90, 96], color: "#0F766E" }],
  bbox: [72, 370, 468, 180],
});

const page = pdf.pages[0];
page.setReadingOrder([
  `${page.id}/text`,
  decision,
  recommendation,
  evidence,
  trend,
]);

const report = verifyArtifact(pdf);
if (!report.ok) throw new Error(report.ndjson || JSON.stringify(report.issues));

const output = await PdfFile.exportPdf(pdf, {
  title: "Readiness report",
  language: "en-US",
});
await output.save("readiness-report.pdf");
```

Use explicit H1-H6 `headingLevel` values, meaningful figure `alt` text or `decorative: true`, semantic table cells/spans, and a complete page reading order. `verify()` detects missing or duplicate reading-order targets, invalid heading nesting, inaccessible figures, malformed tables, geometry errors, and other modeled defects.

Current model details that matter in authoring:

- Non-empty `page.text` is painted and contributes an implicit H1. Do not add another H1 unless two top-level headings are intended.
- `${page.id}/text` exists as a reading-order target only when `page.text` is non-empty.
- `addText(...)` is positioned and does not wrap. Use `addFlowText(...)` for wrapped paragraphs and automatic pagination.

## Import, inspect, and edit

PDFs exported by this package carry a clean-room model envelope and round-trip directly:

```js
const input = await FileBlob.load("readiness-report.pdf");
const pdf = await PdfFile.importPdf(input);

const inspection = pdf.inspect({
  kind: "page,text,textItem,readingOrder,table,tableCell,image,chart",
  maxChars: 20_000,
});
console.log(inspection.ndjson);

const table = pdf.pages.flatMap((page) => page.tables)
  .find((candidate) => candidate.name === "readiness-evidence");
if (!table) throw new Error("Readiness evidence table was not found.");
table.getCell(2, 2).value = "Pass";

const edited = await PdfFile.exportPdf(pdf);
await edited.save("readiness-report-final.pdf");
```

Model IDs are locators for the current object graph, not durable identities across unrelated imports. Locate targets again by bounded text, kind, name, page, or table position after importing a different file.

## Arbitrary PDF extraction and native operations

For PDFs not created by this package, MuPDF.js is the default runtime-lazy parser:

```js
const input = await FileBlob.load("third-party.pdf");
const parsed = await PdfFile.importPdf(input);

console.log(parsed.extractText());
console.log(parsed.extractTables());

const inspection = await PdfFile.inspectPdf(input);
const page = await PdfFile.renderPdf(input, { page: 1, dpi: 144 });
await page.save("third-party-page-1.png");
```

Parser-backed import reconstructs a modeled view for extraction, inspect, and QA. It is not the edit representation and must not be exported as a faithful edit. Table reconstruction is heuristic. Direct-original mutations use `PdfFile.editPdf(input, { operations, savePolicy })`; signatures still route to pyHanko, and strict sanitize/OCR or complex forms/merge route to the documented specialist tools. Inject `createPdfjsParser()` only when an independent PDF.js read adapter is specifically required.

For a bounded imported AcroForm update, inspect first and bind the exact source
plus the returned grouped `mupdfFormField` snapshot. This direct route supports
only one non-password text widget, one non-multiselect combo with identical
display/export options, or one checkbox; radio/shared/list/complex fields route
explicitly to pypdf:

```js
const inspection = await PdfFile.inspectPdf(input);
const city = inspection.records.find((record) => record.kind === "mupdfFormField"
  && record.name === "sender.city");
if (!city?.snapshot) throw new Error("Expected one inspectable city field.");

const filled = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{
    type: "update_form_field",
    sourceSha256: inspection.summary.sourceSha256,
    formFieldId: city.id,
    expected: city.snapshot,
    value: "Shanghai",
  }],
});
await filled.save("third-party-form-filled.pdf");
```

The output gets new bytes, so source-bound locators/snapshots cannot be reused:
re-inspect before a second edit. Incremental is available only for unsigned
input and proves the original byte prefix; it never claims signed-document
permission or content sanitization.

For a bounded visible crop, take the raw page box from native inspection and edit the original bytes directly. The operation is intentionally not redaction: it changes `CropBox`, retains off-window content, and supports only unrotated pages.

```js
const pageRecord = inspection.records.find((record) => record.kind === "mupdfPage" && record.page === 1);
if (!pageRecord?.mediaBox) throw new Error("Missing native MediaBox evidence.");

const cropped = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{ type: "set_page_crop", page: 1, bbox: [72, 72, 468, 648] }],
});
await cropped.save("third-party-page-1-cropped.pdf");
```

The requested `[x, y, width, height]` must fit fully inside the inspected raw `MediaBox`. Reopen and render the result; use a rewrite-plus-sanitize route for any task that requires actual removal of sensitive content.

For an orientation-only edit, use the same direct-original route. `rotation` is
an absolute clockwise `/Rotate` value, not a relative turn; it must be `0`,
`90`, `180`, or `270`. It changes viewer orientation without transforming or
removing content, so unsigned byte-prefix-verified incremental save is allowed:

```js
const rotated = await PdfFile.editPdf(input, {
  savePolicy: "incremental",
  operations: [{ type: "rotate_page", page: 1, rotation: 90 }],
});
await rotated.save("third-party-page-1-rotated.pdf");
```

Inspect and render the result before delivery. Rotated-coordinate text/image
editing remains an explicit specialist-provider task.

For a new Text review note on an imported PDF, bind the exact source hash and
target page snapshot first. This is a pin, not a rectangle API: MuPDF owns the
native icon's normalized size and returns its actual rectangle in the audit:

```js
const notePage = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!notePage) throw new Error("Expected one inspectable target page.");

const withReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_annotation",
    page: notePage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: notePage.bbox, rotation: notePage.rotation },
    point: [72, 128],
    contents: "Review this assumption.",
    author: "Reviewer",
  }],
});
await withReviewNote.save("third-party-with-review-note.pdf");
```

The pin must fit within an unrotated inspected `mupdfPage.bbox`; `text`,
`bbox`, `rect`, icon selection, stale evidence, and incremental save are
rejected. Re-inspect the rewrite before relying on the fresh annotation locator.

For a review highlight, give the provider one requested text string instead of
trying to calculate a rectangle or character quadrilaterals. It is accepted
only if native search finds exactly one selection on the same inspected,
unrotated visible page:

```js
const highlightPage = inspection.records.find((record) => record.kind === "mupdfPage"
  && record.page === 1);
if (!highlightPage) throw new Error("Expected one inspectable target page.");

const withHighlight = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_text_highlight",
    page: highlightPage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: highlightPage.bbox, rotation: highlightPage.rotation },
    text: "Revenue assumptions remain provisional",
    color: [1, 0.92, 0.2],
    contents: "Validate before approval.",
    author: "Reviewer",
  }],
});
await withHighlight.save("third-party-with-review-highlight.pdf");
```

`text` is non-empty and at most 4,096 characters. The optional RGB color uses
three `[0,1]` components (yellow by default), while optional `contents`,
`author`, and `subject` carry non-empty review metadata. Caller quads or
rectangles, zero/multiple native hits, a rotated/stale page, a selection beyond
the visible CropBox, and incremental save are rejected. Re-inspect and render
the rewrite before handoff; its `mupdfAnnotation` record returns the native
Highlight quadrilaterals/color and a current-source-only locator.

For an imported annotation, do not use its array index as identity. Inspect the
exact input bytes, retain the returned `summary.sourceSha256`, and delete only
one source-bound annotation locator with a semantic precondition. This is a
rewrite-only operation because a deletion must not leave the original object in
an incremental revision:

```js
const annotation = inspection.records.find((record) =>
  record.kind === "mupdfAnnotation"
  && record.page === 2
  && record.type === "Text"
  && record.contents === "Resolved in board review"
);
if (!annotation?.id || !inspection.summary.sourceSha256) {
  throw new Error("The target annotation was not uniquely inspectable.");
}

const withoutReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "delete_annotation",
    page: annotation.page,
    annotationId: annotation.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      type: annotation.type,
      contents: annotation.contents,
      rect: annotation.rect,
    },
  }],
});
await withoutReviewNote.save("third-party-without-review-note.pdf");
```

`mupdf-annotation-<page>-<xref>` is a locator for these exact source bytes,
not a durable annotation identity. Re-inspect after every rewrite: MuPDF may
renumber or reuse xrefs. A mismatched source hash, page, locator, or expected
snapshot fails closed before output is written.

To update one imported Text annotation instead of deleting it, use the same
inspection evidence. The bounded patch changes only non-empty `contents`,
`author`, and `subject`; it does not claim arbitrary annotation editing or
layout reflow:

```js
const revisedReviewNote = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_annotation",
    page: annotation.page,
    annotationId: annotation.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      type: annotation.type,
      contents: annotation.contents,
      rect: annotation.rect,
    },
    patch: {
      contents: "Resolved in board review",
      author: "Reviewer",
      subject: "Resolved",
    },
  }],
});
await revisedReviewNote.save("third-party-review-note-updated.pdf");
```

The rectangle can guard the source snapshot but cannot appear in `patch`.
MuPDF normalizes native Text annotation geometry, so moving or resizing a note
must be an explicit delete-plus-add transaction from a fresh inspection, or a
specialist-provider task. Re-inspect this rewrite before any later annotation
mutation; the old xref locator is not a persistent identity.

For an imported link, use the same inspect → source-bound locator → snapshot
pattern. Do not select a link by mutable page-array index or URL alone: several
links may share a target URL. A duplicate semantic fingerprint fails closed
rather than choosing one arbitrarily.

```js
const link = inspection.records.find((record) =>
  record.kind === "mupdfLink"
  && record.page === 2
  && record.url === "https://example.com/obsolete-policy"
);
const linkPage = inspection.records.find((record) =>
  record.kind === "mupdfPage" && record.page === link?.page
);
if (!link?.id || !linkPage || !inspection.summary.sourceSha256) {
  throw new Error("The target link was not uniquely inspectable.");
}

const withoutObsoleteLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "delete_link",
    page: link.page,
    linkId: link.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      url: link.url,
      bbox: link.bbox,
      external: link.external,
    },
  }],
});
await withoutObsoleteLink.save("third-party-without-obsolete-link.pdf");
```

To add a new link, bind its rectangle to the same inspected page geometry.
`add_link` accepts only an unrotated visible CropBox and internal `#...` or
absolute `http`, `https`, and `mailto` destinations:

```js
const withCurrentPolicyLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "add_link",
    page: linkPage.page,
    sourceSha256: inspection.summary.sourceSha256,
    expectedPage: { bbox: linkPage.bbox, rotation: linkPage.rotation },
    bbox: [72, 128, 160, 18],
    url: "https://example.com/current-policy",
  }],
});
await withCurrentPolicyLink.save("third-party-current-policy-link.pdf");
```

To replace that same link's target without moving its native rectangle, use the
same locator and snapshot with a URL-only patch:

```js
const updatedPolicyLink = await PdfFile.editPdf(input, {
  savePolicy: "rewrite",
  operations: [{
    type: "update_link",
    page: link.page,
    linkId: link.id,
    sourceSha256: inspection.summary.sourceSha256,
    expected: {
      url: link.url,
      bbox: link.bbox,
      external: link.external,
    },
    patch: { url: "https://example.com/current-policy" },
  }],
});
await updatedPolicyLink.save("third-party-policy-link-updated.pdf");
```

The patch must contain exactly one non-empty safe internal `#...` or absolute
`http`, `https`, or `mailto` `url`; a link rectangle is snapshot evidence, not
mutable geometry. MuPDF's bounds setter does not provide a stable
saved/reloaded coordinate contract for this public API, so move a link through
an explicit `delete_link` + `add_link` transaction in one rewrite using the
same original link/page evidence, or use a specialist provider. Re-inspect the
rewrite before any later link operation.

`mupdf-link-<page>-<fingerprint>` is source-byte-bound, not a persistent link
identity. It has no native xref because the MuPDF link API abstracts that
object away; its fingerprint covers page, URL, rectangle, and externality. A
new output must be re-inspected before a later link operation.

## Render and visual QA

Use the model SVG preview while authoring, then render the exported PDF with Poppler and inspect every page before delivery:

```js
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

const preview = await pdf.render({ pageIndex: 0 });
await preview.save("readiness-report-preview.svg");

const renderer = createPopplerRenderer({ dpi: 144, timeoutMs: 60_000 });
for (let pageIndex = 0; pageIndex < pdf.pages.length; pageIndex += 1) {
  const png = await pdf.render({
    source: "pdf",
    format: "png",
    pageIndex,
    renderer,
  });
  await png.save(`readiness-report-page-${pageIndex + 1}.png`);
}
```

`PdfFile.inspectPdf(...)` verifies byte-level evidence such as PDF version, page/object counts, EOF, tagged structure, reading-order IDs, headings, Figure alternative text, Table/TR/TH/TD roles, spans, and font evidence. It complements semantic `pdf.verify()` and visual page review; none of the three replaces the others.

Use `pdftoppm`/`pdfinfo` as independent native render and file QA tools. The surrounding PDF Skill defines the MuPDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, pyHanko, veraPDF, and OCR routing contract; its shipped `scripts/qpdf_provider.py` provides bounded source-hash-bound qpdf 11+ structure inspection, recovery rewrite, and linearization without exposing qpdf as a generic flag passthrough.
