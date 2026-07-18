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

Use `pdftoppm`/`pdfinfo` as independent native render and file QA tools. The surrounding PDF Skill defines the MuPDF.js, ReportLab, pdfplumber, pypdf, PyMuPDF, qpdf, pyHanko, veraPDF, and OCR routing contract.
