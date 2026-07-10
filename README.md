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

This is an early MVP. It already creates and imports minimal XLSX/PPTX/DOCX/PDF artifacts, supports stable inspect IDs, and includes tests for all four skill families. Fidelity, charts, comments, advanced OOXML, high-quality raster rendering, and template QA are roadmap work.

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

## Design notes

The package deliberately prioritizes agent workflows:

1. inspect compact semantic snapshots instead of dumping raw XML;
2. resolve stable IDs back to editable objects;
3. export both durable files and lightweight layout/preview artifacts;
4. expose bounded help records for API discovery.

## Development

```sh
npm install
npm test
npm run test:pack
```
