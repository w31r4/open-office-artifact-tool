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

Spreadsheet formulas support relationship-driven native table import plus a shared structured-reference parser for evaluation, dependency graphs, and trace output. This includes `#This Row`/`@`, unqualified calculated-column references, and apostrophe escaping for special-character headers such as `Sales['#Items]` and `Sales[Bracket'[Value']]`; metadata-free XLSX roundtrips retain table column names and exact row precedents.

The six conditional aggregation functions `COUNTIF(S)`, `SUMIF(S)`, and `AVERAGEIF(S)` share case-insensitive Excel-style comparison and `?`/`*`/`~` wildcard matching. Multi-range `IFS` variants validate equal dimensions and propagate matched aggregation errors instead of silently coercing them.

This is an early MVP. It already creates and imports XLSX/CSV/TSV/PPTX/DOCX/PDF artifacts, supports stable inspect/resolve IDs plus target-sliced and include/exclude-shaped inspect output, and includes tests for all four skill families. The spreadsheet facade includes bounded RFC-style CSV/TSV import/export/inspection, formula traces, workbook defined names, a dependency graph with cycle/missing-sheet reporting, Excel-style table structured references (`TableName[Column]`, `#Headers`/`#Data`/`#All`/`#Totals`, contiguous multi-column ranges, and comma-separated column unions), a broader clean-room formula catalog spanning logical/text/lookup/conditional aggregation, statistical ranking (`MEDIAN`, `MODE.SNGL`, `LARGE`, `SMALL`, `RANK.EQ`), signed decimal rounding, deterministic Excel 1900/1904 date/workday arithmetic (`DATE`, `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `DAYS`, `WEEKDAY`, `NETWORKDAYS`, `WORKDAY`, `NETWORKDAYS.INTL`, `WORKDAY.INTL`) with native `workbookPr@date1904` roundtrip and custom weekend masks, and dynamic-array helpers (`SEQUENCE`, `TRANSPOSE`, `FILTER`, `UNIQUE`, `SORT`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `HSTACK`, `VSTACK`, `EXPAND`), live range style/dimension/autofit formatting, Excel-style `fillDown`/`fillRight` formula translation, native merged-cell creation/import/export, native shared strings/styles, built-in/custom number-format and cell-style inheritance, theme/tint/indexed-color resolution, and per-edge borders through a dedicated SpreadsheetML codec, formatted layout/SVG display values, frozen-pane/gridline worksheet views, column widths/best-fit/hidden state, row heights/hidden state, comments/data-validation/conditional-formatting/table/chart/image/sparkline metadata roundtrips, native XLSX table/chart/image/sparkline/threaded-comment XML parts, and relationship-driven metadata-free import of embedded images plus basic bar/line/pie charts from arbitrary valid drawing targets and one-/two-cell anchors. Dimension-aware SVG previews and workbook/worksheet layout JSON feed semantic and visual QA. The presentation facade includes compose/JSX layout, inspectable shape/table/chart/image/connector objects, speaker notes, threaded comments, bounded PPTX package inspection, native text/fill/line styling, explicit native table cell styling, native PPTX table/chart/image/connector XML export plus notes/comment parts and clean-room native import restoration, and a geometry-based layout QA detector for overlap/off-canvas/overflow/connector checks. The document facade includes styled paragraphs, real list items, headers/footers, hyperlinks, fields, citations, images, sections, tracked insertions/deletions, tables, comments, design presets, page-aware layout JSON, visual layout QA, DOCX styles/numbering/header/footer/hyperlink/comment/image/section/tracked-change export, and SVG page previews. The PDF facade includes modeled multi-page text/table/image/chart artifacts, tagged PDF export with language/title metadata, H1/P/Figure elements, semantic `Table` → `TR` → `TH`/`TD` hierarchy, column-scoped headers and image/chart alt text, positioned text, vector tables/charts and embedded PNG/JPEG images, bounded structural inspection, `extractText`, `extractTables`, SVG/layout previews, metadata roundtrips, and an optional PDF.js parser for page geometry/positioned text/table/image-mask extraction. This is a tagged-PDF baseline, not a claim of full PDF/UA conformance. Cross-format `verifyArtifact(...)`, `visualQaArtifact(...)`, and `renderArtifact(...)` helpers provide agent-facing QA and preview entry points, including pluggable PNG/WebP/JPEG/PDF renderers, decoded PNG/JPEG/WebP/PPM baselines, cross-encoding comparisons, and configurable aligned PNG diff heatmaps for visual regressions. Fidelity, advanced OOXML, Windows Office bridge validation, robust arbitrary-PDF parsing, and template QA remain roadmap work.

Worksheet-backed native PivotTables now participate in the same metadata-free path as tables and drawings. Export writes worksheet→pivotTable, pivotTable→cache-definition, workbook→cache-definition, and cache-definition→records relationships plus workbook `pivotCaches`; import follows arbitrary valid targets and reconstructs source/target ranges, row/column/value fields, aggregation, computed values, inspect/resolve/layout, and SVG rendering.

PPTX legacy comments export with a native `ppt/commentAuthors.xml` registry, presentation relationship, content type, per-author IDs, one-based per-author comment indexes, initials, display colors, and `lastIdx`. Import restores each comment/reply author and follows relationship targets for author, notes, and comment parts even when safe package patches move them away from conventional filenames. Master/layout handling follows the full presentation→master→layout and slide→layout graph, preserves unused layouts, and exports the standard master/layout ID lists plus bidirectional master-layout relationships.

DOCX headers and footers support `default`, `first`, and `even` reference types plus an optional zero-based `sectionIndex`. Each section/type combination gets its own relationship-driven part and is attached to that section's `w:sectPr`; omitting `sectionIndex` preserves the original final-section behavior. Export activates first-page/even-page semantics through section-local `w:titlePg` and `settings.xml`, while native import restores arbitrary relationship targets, reference types, and section indexes. `DocumentFile.patchDocx(...)` header/footer recipes accept the same `sourceReference: { type, sectionIndex }` targeting. Use `DocumentFile.importDocx(..., { preferNative: true })` after package-level patches so native OOXML takes precedence over embedded clean-room metadata.

Classic DOCX comments preserve `author`, deterministic or caller-supplied `initials`, and optional `date` metadata. Paragraph and table targets export with matching `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference` anchors. Native import locates the Comments part through `document.xml.rels`, so a standards-compliant package may place it at an arbitrary internal target instead of `word/comments.xml`; `{ preferNative: true }` restores author metadata and table targets from the package itself. `DocumentFile.patchDocx(...)` also accepts a comments recipe with `sourceReference: { anchors: [{ commentId, target }] }`: targets may select a top-level block, a document-order paragraph, or a table cell, every ID must exist uniquely in the patched Comments part, and deleting the part removes matching document anchors.

DOCX numbering recipes accept `sourceReference: { assignments: [{ numId, level, target }] }`. Each assignment adds or replaces the target paragraph's `w:numPr`, verifies that the Numbering part uniquely declares the requested instance and level, supports the same block/paragraph/table-cell selectors as comment anchors, and removes all `w:numPr` references when that Numbering part is deleted. Native-preferred import restores the resulting list kind, number format, start, level text, numbering ID, and nesting level.

`DocumentFile.importDocx(..., { preferNative: true })` also reconstructs external hyperlinks from their `r:id` relationships, including escaped URL query strings, and records the relationship ID on the block. Both `w:fldSimple` fields and complex begin/instruction/separate/end field sequences recover their instruction and displayed result. Clean-room citation bookmarks with the `OpenOfficeCitation_` prefix recover as citation blocks instead of flattening to paragraphs. Comments attached to hyperlinks, fields, and citations use the same native range/reference anchors.

DOCX lists use real multi-level numbering definitions rather than assuming `numId` 1/2. `document.addListItem(...)` accepts levels 0–8, `numberFormat`, positive `start`, `levelText`, and an optional `numberingId` that groups levels into one list instance. Export emits every referenced level in `abstractNum` plus a `num` instance and points each paragraph at the generated ID. Native import follows the styles and numbering relationships to arbitrary internal part paths, resolves `numId` → `abstractNumId` → `lvl`, applies `lvlOverride`/`startOverride`, and preserves formats such as `upperLetter` and `lowerRoman` for inspect/re-export.

PDF content using Chinese, Cyrillic, Greek, accented Latin, or other non-ASCII characters must provide a legally usable standalone glyf-based TrueType `.ttf` through `PdfFile.exportPdf(..., { font })`. The clean-room writer embeds a Type0/CIDFontType2 font, glyph widths, a `ToUnicode` CMap, and by default a deterministic font subset containing used glyphs plus recursive composite dependencies. It rebuilds `glyf`, long `loca`, Unicode `cmap`, the sfnt directory, table checksums, and `checkSumAdjustment`; `{ subsetFont: false }` preserves the full font only for diagnostics. Missing fonts/glyphs, `.ttc` collections, malformed tables, and oversized inputs fail explicitly rather than corrupting text. Direct Unicode mapping is implemented; complex-script shaping remains roadmap work.

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

// Delimited files flatten one selected sheet/range; calculated values are the default.
await (await SpreadsheetFile.exportCsv(workbook, { sheetName: "Sheet1" })).save("output.csv");
const importedCsv = await SpreadsheetFile.importCsv("Name,Value\r\nRevenue,120", { coerceTypes: true });
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

## Renderer adapters

`renderArtifact(artifact, { format })` returns the artifact's native SVG preview by default. Raster/PDF formats must be supplied by an explicit adapter so the package never pretends to support PNG/WebP/PDF output without a renderer. `visualQaArtifact(...)` records deterministic render metadata and hashes, checks empty/malformed renders, and compares PNG/JPEG/WebP/PPM decoded pixels with `pixelDiff: true`. A `pixelDiff` object can customize `diffPalette` colors/alpha, set `diffAlignment` to `strict`, `top-left`, or `center` for dimension-mismatch heatmaps, and opt into `pixelRegistration: { maxOffset: 2 }` for bounded same-size baseline translation. Registration is sampled under a fixed work budget and reports the chosen offset, before/after mismatches, improvement, and ignored edge pixels.

```js
import { DocumentModel, PdfArtifact, renderArtifact } from "open-office-artifact-tool";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";

const document = DocumentModel.create({ paragraphs: ["Raster-ready report"] });
const pdfArtifact = PdfArtifact.create({ text: "Raster-ready PDF" });
const renderer = createPlaywrightRenderer({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 1 });

const png = await renderArtifact(document, { format: "png", renderer });
const webp = await renderArtifact(document, { format: "webp", renderer });
const pdf = await renderArtifact(document, { format: "pdf", renderer });

// For DOCX-fidelity render gates, feed the real WordprocessingML package into
// a DOCX-capable adapter such as LibreOffice or native Office instead of SVG:
const docxPdf = await renderArtifact(document, { format: "pdf", source: "docx", renderer: libreOfficeOrNativeRenderer });

// For PDF raster gates, feed the real exported PDF into a PDF-capable adapter
// such as Poppler instead of rasterizing the SVG preview:
const pdfPng = await renderArtifact(pdfArtifact, { format: "png", source: "pdf", renderer: popplerRenderer });
```

The Playwright adapter is an optional peer dependency to keep the core npm package light:

```sh
npm install -D playwright
npx playwright install chromium
npm run test:playwright-renderer
```

It accepts SVG or HTML `FileBlob` input, fixes viewport/device scale/timezone/locale, waits for font readiness, disables animations, and blocks network requests by default. Set `allowNetwork: true` only for explicitly trusted local HTML previews.

Additional optional render adapters are available for narrower environments:

```js
import { createSharpRenderer } from "open-office-artifact-tool/renderers/sharp";
import { createCanvasRenderer } from "open-office-artifact-tool/renderers/canvas";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";

const sharpRenderer = createSharpRenderer(); // SVG/PNG/JPEG/WebP -> PNG/WebP/JPEG
const canvasRenderer = createCanvasRenderer(); // SVG/PNG/JPEG/WebP -> PNG/JPEG via node-canvas
const popplerRenderer = createPopplerRenderer(); // PDF FileBlob -> page PNG/PPM/TIFF via pdftoppm
const libreOfficeRenderer = createLibreOfficeRenderer(); // DOCX/XLSX/PPTX/CSV/TSV/HTML -> PDF via soffice
```

Install `sharp` for the sharp adapter. Install the optional `canvas` npm package for the node-canvas adapter (SVG rasterization requires a canvas build with SVG/rsvg support). Install Poppler (`pdftoppm`) on the host for the Poppler adapter, or pass a custom `command` for a compatible CLI. Install LibreOffice (`soffice`) on the host for the LibreOffice adapter, or pass `command`/`args` for a compatible headless conversion CLI.

## Native Office bridge adapter

The core package does not depend on Windows or Microsoft Office. The Node-side wrapper at `open-office-artifact-tool/native/office-bridge` calls the optional C# sidecar in `native/OfficeBridge` (or any compatible JSON stdin/stdout command) with timeout handling, isolated temp files, cleanup, and structured errors.

```js
import { renderFileWithNativeOffice } from "open-office-artifact-tool/native/office-bridge";

const pdf = await renderFileWithNativeOffice(docxBlob, {
  command: "dotnet",
  args: ["run", "--project", "native/OfficeBridge"],
  artifactKind: "document",
  inputType: docxBlob.type,
  outputType: "application/pdf",
  format: "pdf",
  timeoutMs: 60_000,
});
```

Set `OFFICE_BRIDGE_COMMAND` and optional JSON `OFFICE_BRIDGE_ARGS` to configure the bridge without passing command options. The C# Office sidecar supports status checks on every platform and uses Microsoft Office COM automation on Windows when Office is installed; Office-specific integration tests are gated with `OFFICE_NATIVE_TESTS=1`.

```sh
# Non-Office protocol tests, when dotnet is installed
dotnet test native/OfficeBridge

# Optional Windows + Office integration tests
OFFICE_NATIVE_TESTS=1 dotnet test native/OfficeBridge
```

## PDF parsing adapters

`PdfFile.importPdf(blob)` preserves clean-room metadata roundtrips first, then falls back to lightweight visible-text heuristics. For arbitrary PDFs, pass a parser adapter explicitly:

```js
import { PdfFile } from "open-office-artifact-tool";
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";

const parser = createPdfjsParser();
const pdf = await PdfFile.importPdf(fileBlob, { parser, preferParser: true });
console.log(pdf.inspect({ kind: "page,textItem,table,image" }).ndjson);
```

The PDF.js adapter is optional and keeps `pdfjs-dist` out of the core dependency tree. It extracts bounded raster XObjects and converts 1-bit stencil masks to fill-colored RGBA PNGs while preserving their placement and mask metadata:

```sh
npm install -D pdfjs-dist
```

It extracts page size, positioned text items, text-line regions, heuristic tables from pipe-delimited or column-like text, and image operator placeholders. The built-in heuristic parser remains available when no adapter is supplied.

## Safe OOXML package inspection and patching

XLSX, PPTX, and DOCX expose the same bounded package workflow. Inspect records validate safe paths, content types, internal relationships, duplicate IDs, and namespace-aware source XML references under part/byte budgets. Patch methods accept XML, JSON, text, binary, and remove operations; content types and relationships synchronize automatically, with IDs preserved long enough to clean source XML. Set `recipe.sourceReference` for DOCX header/footer references, classic-comment anchors, and numbering assignments; XLSX worksheet/table lists, complete worksheet→drawing→image/chart chains, and workbook/cache/worksheet PivotTable chains; and PPTX slide/master/layout ID lists. Drawing recipes require explicit geometry, DOCX comment/numbering builders require declared IDs and explicit structural targets, pivot cache definitions require a unique `cacheId`, and PPTX master/layout recipes validate schema-bounded unsigned IDs, so the package layer never silently invents semantic identity. It owns namespace prefixes, non-visual IDs, source lists, add/remove cleanup, and final atomic validation. Use `validateResult: false` only for deliberate invalid-package fixtures.

```js
const report = await SpreadsheetFile.inspectXlsx(xlsx, {
  includeText: true,
  maxParts: 5000,
  maxPartBytes: 64 * 1024 * 1024,
  maxTotalBytes: 256 * 1024 * 1024,
});

const patched = await PresentationFile.patchPptx(pptx, [
  {
    path: "customXml/review.json",
    json: { status: "approved" },
    contentType: "application/json",
    relationship: {
      source: "ppt/presentation.xml",
      id: "rIdReview",
      type: "urn:open-office:relationships/review",
    },
  },
  {
    path: "ppt/charts/review.xml",
    xml: "<c:chartSpace xmlns:c=\"http://schemas.openxmlformats.org/drawingml/2006/chart\"><c:chart/></c:chartSpace>",
    recipe: { kind: "chart", source: "ppt/slides/slide1.xml", id: "rIdReviewChart" },
  },
]);
```

An XLSX drawing chain can be created in one validated patch call. The chart/image part contents remain caller-supplied public OOXML or bytes; the package layer builds the source nodes and relationships:

```js
const patchedWorkbook = await SpreadsheetFile.patchXlsx(xlsx, [
  {
    path: "xl/drawings/agent.xml",
    xml: '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>',
    recipe: { kind: "drawing", source: "xl/worksheets/sheet1.xml", sourceReference: true },
  },
  {
    path: "xl/media/status.png",
    bytes: pngBytes,
    recipe: {
      kind: "image",
      source: "xl/drawings/agent.xml",
      sourceReference: {
        name: "Status",
        alt: "Green status indicator",
        anchor: { type: "oneCell", from: { row: 2, col: 5 }, extent: { widthPx: 48, heightPx: 48 } },
      },
    },
  },
]);
```

Equivalent APIs are `PresentationFile.inspectPptx`, `DocumentFile.inspectDocx`, `SpreadsheetFile.patchXlsx`, `PresentationFile.patchPptx`, and `DocumentFile.patchDocx`.

## Examples

Runnable examples live in [`examples/`](examples/):

- `create-docx-report.mjs`
- `create-xlsx-dashboard.mjs`
- `create-pptx-compose.mjs`
- `parse-render-pdf.mjs`
- `render-via-playwright.mjs`
- `render-via-native-office.mjs`

Run all examples with:

```sh
npm run test:examples
```

Outputs are written to `OUTPUT_DIR` or a temp example directory.

## Runnable agent skills

Project-local clean-room agent workflows live under [`skills/`](skills/). [`skills/spreadsheets/SKILL.md`](skills/spreadsheets/SKILL.md) uses this package's public APIs for XLSX/CSV/TSV authoring and import, bounded package/row evidence, semantic verification, layout export, and SVG or Playwright-backed visual QA. [`skills/documents/SKILL.md`](skills/documents/SKILL.md) adds durable DOCX roundtrip, package-part inspection, modeled QA, and optional real LibreOffice PDF + Poppler page-PNG verification. [`skills/presentations/SKILL.md`](skills/presentations/SKILL.md) adds narrative/layout rules, native PPTX package inspection, per-slide modeled/native rendering, montage review, and optional PNG baseline pixel diffs. [`skills/pdf/SKILL.md`](skills/pdf/SKILL.md) adds real PDF byte/object inspection, PDF.js semantic extraction, per-page Playwright and Poppler rendering, and model/native PNG baseline diffs.

```sh
node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/formula-summary.json \
  --output-dir tmp/spreadsheet-skill-fixture

node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input tmp/spreadsheet-skill-fixture/formula-summary.xlsx \
  --output-dir tmp/spreadsheet-skill-fixture/qa-png \
  --sheet Summary \
  --range A1:D4 \
  --render-format png

node skills/documents/scripts/run-fixture.mjs \
  --fixture skills/documents/fixtures/business-brief.json \
  --output-dir tmp/document-skill-fixture \
  --native-render required

node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/agent-readiness.json \
  --output-dir tmp/presentation-skill-fixture \
  --native-render required

node skills/pdf/scripts/run-fixture.mjs \
  --fixture skills/pdf/fixtures/qa-report.json \
  --output-dir tmp/pdf-skill-fixture \
  --native-render required \
  --pdfjs required
```

All four skill directories are included in the npm package.

## Release readiness

Use [`docs/release.md`](docs/release.md) and `npm run release:check` before publishing. The checker runs the required local gates, reports npm authentication/package-version status, and prints blockers. Publishing is intentionally blocked in this environment until `npm whoami` succeeds.

Third-party libraries and separately installed render/native runtimes are documented in [`THIRD_PARTY_NOTICES.md`](THIRD_PARTY_NOTICES.md). The release gate audits every locked npm license expression and every declared dependency against that notice.

## Design notes

The package deliberately prioritizes agent workflows:

1. inspect compact semantic snapshots instead of dumping raw XML;
2. resolve stable IDs back to editable objects;
3. export both durable files and lightweight layout/preview artifacts;
4. expose bounded help records for API discovery via `helpArtifact(...)` and generated [`docs/api.md`](docs/api.md);
5. verify artifacts with `verifyArtifact(artifact)` or per-artifact `verify()` methods before delivery, and use `visualQaArtifact(...)` when a render hash/baseline gate or PNG pixel-diff gate is useful.

## Development

```sh
npm install
npm test
npm run test:examples
npm run test:office-bridge
npm run test:playwright-renderer # skips unless Playwright/Chromium are installed
npm run test:pack
# optional when dotnet is installed:
dotnet test native/OfficeBridge
```
