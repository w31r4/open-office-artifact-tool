# open-office-artifact-tool

Clean-room Office/PDF artifact toolkit inspired by the public behavior of agent's Office/PDF skills.

The goal is not to vendor or copy agent's reference bundle. This package rebuilds a similar agent-facing surface using open implementation code:

- `Workbook` / `SpreadsheetFile` for XLSX-style artifacts
- `Presentation` / `PresentationFile` for PPTX-style artifacts
- `DocumentModel` / `DocumentFile` for DOCX-style artifacts
- `PdfArtifact` / `PdfFile` for PDF artifacts
- shared `FileBlob`
- `inspect(...)`, `resolve(...)`, `help(...)`, render/export-style APIs where practical

## Current status

This is an early MVP. It already creates and imports minimal XLSX/PPTX/DOCX/PDF artifacts, supports stable inspect IDs, and includes tests for all four skill families. The spreadsheet facade includes formula traces plus comments/data-validation/conditional-formatting/table/chart/image/sparkline metadata roundtrips, native XLSX table/chart/image/sparkline/threaded-comment XML parts, and SVG visual previews. The presentation facade includes compose/JSX layout, inspectable shape/table/chart/image objects, native PPTX table/chart/image XML parts, and a geometry-based layout QA detector for overlap/off-canvas/overflow checks. The document facade includes styled paragraphs, real list items, headers/footers, hyperlinks, fields, citations, images, sections, tracked insertions/deletions, tables, comments, DOCX styles/numbering/header/footer/hyperlink/comment/image/section/tracked-change export, and SVG page previews. The PDF facade includes modeled multi-page text/table artifacts, `extractText`, `extractTables`, SVG page render, and metadata roundtrips. Cross-format `verifyArtifact(...)` and `renderArtifact(...)` helpers provide agent-style QA and preview entry points. Fidelity, advanced OOXML, high-quality raster rendering, and template QA are roadmap work.

## Usage

```js
import { Workbook, SpreadsheetFile } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C2").values = [["A", "B", "Sum"], [2, 3, null]];
sheet.getRange("C2").formulas = [["=A2+B2"]];
workbook.recalculate();

console.log((await workbook.inspect({ kind: "table,formula" })).ndjson);
const file = await SpreadsheetFile.exportXlsx(workbook);
await file.save("output.xlsx");
```

Presentation compose-first authoring uses helper nodes that mirror the agent-oriented JSX vocabulary while staying transpiler-free:

```js
import { column, paragraph, Presentation, PresentationFile, row, box } from "open-office-artifact-tool";

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = presentation.slides.add();
slide.compose(
  column({ name: "content-frame", width: "fill", height: "fill", gap: 16, padding: { x: 24, y: 20 } }, [
    paragraph({ id: "sh/stable-headline", name: "primary-heading", className: "text-slate-950 text-4xl font-bold" }, ["Quarterly readiness"]),
    row({ name: "kpi-row", width: "fill", height: 120, gap: 12 }, [
      box({ name: "kpi-card", width: "fill", height: "fill", fill: "slate-50", padding: { x: 12, y: 10 } }, [
        paragraph({ name: "kpi-label" }, ["Pipeline"]),
      ]),
    ]),
  ]),
  { frame: { left: 80, top: 120, width: 760, height: 360 } },
);

console.log(presentation.inspect({ kind: "textbox,shape" }).ndjson);
await (await PresentationFile.exportPptx(presentation)).save("deck.pptx");
```

If you use a JSX transform, the package also exposes presentation-jsx-compatible subpaths:

```js
import { Fragment } from "open-office-artifact-tool/presentation-jsx";
import { jsx, jsxs } from "open-office-artifact-tool/presentation-jsx/jsx-runtime";
import { jsxDEV } from "open-office-artifact-tool/presentation-jsx/jsx-dev-runtime";

const tree = jsxs(Fragment, {
  children: [
    jsx("paragraph", {
      name: "headline",
      className: "text-slate-950 text-4xl font-bold",
      children: "Agent-ready JSX runtime",
    }),
  ],
});
slide.compose(tree, { frame: { left: 80, top: 120, width: 720, height: 180 } });
```

## Design notes

The package deliberately prioritizes agent workflows:

1. inspect compact semantic snapshots instead of dumping raw XML;
2. resolve stable IDs back to editable objects;
3. export both durable files and lightweight layout/preview artifacts;
4. expose bounded help records for API discovery via `helpArtifact(...)` and generated [`docs/api.md`](docs/api.md);
5. verify artifacts with `verifyArtifact(artifact)` or per-artifact `verify()` methods before delivery.

## Development

```sh
npm install
npm test
npm run test:pack
```
