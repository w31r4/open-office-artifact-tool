# open-office-artifact-tool

Clean-room Office and PDF artifact toolkit for agent workflows.

Version 0.2 has one Office codec: **OpenChestnut**, the repository's C# Open XML SDK implementation compiled to bundled .NET WebAssembly. `SpreadsheetFile`, `DocumentFile`, and `PresentationFile` always use it for XLSX, DOCX, and PPTX import/export. Installed consumers do not need `dotnet` on `PATH`.

PDF is the fourth, independent format pipeline. It does not use the Office wire codec and keeps its JavaScript creation/import, inspect, render, and verification workflow.

## Format boundary

| Format | File pipeline | Supported authoring/import boundary |
| --- | --- | --- |
| XLSX | OpenChestnut C# WASM | Cells, formulas, static styles, merged cells, row/column sizes, frozen panes, tables, PNG/JPEG images, bar/line/pie charts, dates as Excel serials, basic data validation, basic conditional formatting, and one-level threaded comments. |
| DOCX | OpenChestnut C# WASM | Styles, paragraphs and runs, page/section settings, headers and footers, PAGE/simple fields, PNG/JPEG images, lists, fixed-geometry tables, links, classic comments, and fixed-topology edits of modeled objects. Bookmarks, bibliography, unsupported settings, and opaque blocks are imported read-only. |
| PPTX | OpenChestnut C# WASM | Text boxes and round rectangles, basic fill/line/shadow, line/polyline connectors and arrows, source-free bar/line/pie charts, images, tables, rich text, lists, links, and read-only source-bound Master/Layout preservation. |
| PDF | Independent PDF pipeline | Creation/import, text/table/image extraction and modeling, page geometry, reading order, accessibility metadata, inspect, render, and verify. Optional PDF.js, Poppler, and Playwright adapters extend import/render QA. |

Imported Office objects outside the modeled boundary remain hash-bound to their source package. Leaving them unchanged preserves them; trying to create or semantically edit an unsupported object fails explicitly. If a source-bound opaque object no longer has a trustworthy source snapshot, export fails. There is no lossy fallback.

JavaScript still owns the public object models, calculations, Compose/JSX, normalization, inspect/resolve, explicit low-level OOXML package patching, render adapters, and QA. Those facilities are not a second Office serializer.

## Installation

```sh
npm install open-office-artifact-tool
```

## Office examples

No codec selector is accepted. Facades lazily load the bundled runtime on first Office import/export.

```js
import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:B2").values = [["Metric", "Value"], ["Revenue", 42.5]];
sheet.tables.add({ range: "A1:B2", name: "MetricsTable", style: "TableStyleMedium4" });
const xlsx = await SpreadsheetFile.exportXlsx(workbook, { recalculate: true });
const importedWorkbook = await SpreadsheetFile.importXlsx(xlsx);

const document = DocumentModel.create({ paragraphs: ["OpenChestnut document"] });
const docx = await DocumentFile.exportDocx(document);
const importedDocument = await DocumentFile.importDocx(docx);

const deck = Presentation.create();
deck.slides.add({ name: "Overview" }).shapes.add({
  name: "Title",
  type: "roundRect",
  text: "OpenChestnut presentation",
  position: { left: 60, top: 40, width: 640, height: 80 },
});
const pptx = await PresentationFile.exportPptx(deck);
const importedDeck = await PresentationFile.importPptx(pptx);
```

Office import methods accept only `limits`. DOCX/PPTX export also accepts only `limits`; XLSX export additionally accepts `recalculate`. Legacy `codec`, `allowLossy`, `preferNative`, and `relativeDateAsOf` options throw. The removed `codecs/openxml-wasm` package path is not exported.

Advanced byte-boundary users may call the same implementation through `open-office-artifact-tool/codecs/open-chestnut`. Generated wire bindings are exported at `open-office-artifact-tool/codecs/open-chestnut/wire`. Both use wire protocol version 2 and identify output with `metadata.codec === "open-chestnut"`.

## PDF example

```js
import { PdfArtifact, PdfFile } from "open-office-artifact-tool";

const pdf = PdfArtifact.create({
  title: "Readiness report",
  pages: [{
    text: "Quarterly readiness",
    tables: [{ values: [["Metric", "Value"], ["Coverage", "82%"]], bbox: [72, 160, 320, 80] }],
  }],
});

const file = await PdfFile.exportPdf(pdf);
const inspection = await PdfFile.inspectPdf(file);
const imported = await PdfFile.importPdf(file);
console.log(imported.extractText(), inspection.summary);
```

`open-office-artifact-tool/pdf/pdfjs` supplies the optional PDF.js parser. Poppler and Playwright renderer adapters are exposed separately so runtime behavior remains explicit.

## Inspect, patch, render, and QA

Artifact models expose stable `inspect(...)`, `resolve(...)`, layout, render, and `verify()` surfaces. `verifyArtifact(...)`, `renderArtifact(...)`, and `visualQaArtifact(...)` provide shared gates.

`SpreadsheetFile.inspectXlsx/patchXlsx`, `DocumentFile.inspectDocx/patchDocx`, and `PresentationFile.inspectPptx/patchPptx` are explicit package-level tools. They validate paths, relationships, content types, and source references; the normal facades never call them as a fallback.

Renderer adapters are exported for Playwright, sharp, canvas, Poppler, LibreOffice, and the optional native Office bridge. The bridge is a QA/render sidecar, not an Office codec.

## Reference Skills

The repository ships four runnable agent Skills:

- `skills/documents`
- `skills/spreadsheets`
- `skills/presentations`
- `skills/pdf`

The three Office Skills use the canonical facade without codec options. The PDF Skill remains independent.

## Development

```sh
npm install
npm run proto:generate
npm run build:open-chestnut
npm run test:open-chestnut-dotnet
npm test
npm run test:pack
npm run docs:api
```

`npm run verify:open-chestnut-build` performs the deterministic WASM build check. `npm run release:check` validates release metadata and packaged artifacts. C# source and build tooling live in the repository; the npm package contains the bundled runtime, public proto/generated JavaScript, integrity manifest, SBOM, and licenses.

See [coverage](docs/coverage.md), [runtime architecture](docs/reference-runtime-architecture.md), and [release gates](docs/release.md).

## License

MIT. Runtime third-party notices are in `THIRD_PARTY_NOTICES.md` and the packaged OpenChestnut runtime notice files.
