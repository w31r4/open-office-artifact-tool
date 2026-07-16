# PDF API quick start

Use `open-office-artifact-tool` for greenfield modeled PDF creation, semantic editing of a trusted internal model, tagged export, inspect, resolve, render, and verification. The PDF pipeline is independent from OpenChestnut and does not require Microsoft Office, LibreOffice, or a local .NET SDK. Do not use this model as a fidelity-preserving mutation path for an arbitrary imported PDF.

## Startup

Resolve Node.js and the package directory through the Codex workspace dependency loader. Work in a writable task directory, link that loader-provided package directory into the task workspace when necessary, and use ES modules.

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

## Arbitrary PDF extraction with PDF.js

For PDFs not created by this package, inject the optional public PDF.js parser explicitly:

```js
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";

const input = await FileBlob.load("third-party.pdf");
const parsed = await PdfFile.importPdf(input, {
  parser: createPdfjsParser(),
  preferParser: true,
  parserName: "pdfjs",
});

console.log(parsed.extractText());
console.log(parsed.extractTables());
```

Parser-backed import reconstructs a modeled view for extraction, inspect, and QA. It is not a native PDF object-stream editor and must not be exported as an edit to the arbitrary source file. Table reconstruction is heuristic and requires text/geometry/render review. Route mutations directly from original bytes to the explicitly selected pypdf or PyMuPDF provider; route signatures to pyHanko.

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

Use `pdftoppm`/`pdfinfo` as explicit native render and file QA tools. The surrounding PDF Skill defines the ReportLab, pdfplumber, pypdf, PyMuPDF, qpdf, pyHanko, veraPDF, and OCR routing contract for capabilities outside this greenfield model API.
