# Clean-room coverage matrix

This document tracks the target behavior observed from four Office/PDF artifact workflows and the current implementation status of `open-office-artifact-tool`.

Status legend:

- `done`: implemented and covered by smoke tests.
- `partial`: API or minimal behavior exists, but fidelity/edge cases remain.
- `todo`: not implemented yet.

## Cross-cutting agent runtime

| Requirement | Status | Notes |
| --- | --- | --- |
| Shared `FileBlob` with `load`, `save`, `arrayBuffer`, `text` | done | Used by all four artifact families. |
| `inspect(...)` emits bounded NDJSON snapshots | partial | Implemented for workbook, presentation, document, PDF. Needs full token coverage and search/target slicing. |
| Stable anchor IDs + `resolve(...)` | partial | Implemented for workbook/sheets and presentation/shapes. Needs table/chart/image/comment/text-range IDs. |
| `help(query, opts)` bounded NDJSON API discovery | partial | Seed catalog implemented. Needs full API/function catalog. |
| Render/preview loop | partial | SVG previews implemented for sheets, slides, and PDF pages. PNG/WebP rasterization is roadmap. |
| Layout JSON export | partial | Slide layout JSON implemented. Workbook/document layout exports are roadmap. |
| Durable file export/import smoke tests | done | Minimal XLSX/PPTX/DOCX/PDF round trips pass. |

## Spreadsheets skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Workbook.create`, worksheets collection, `getRange`, values/formulas | done | Matrix writes and simple formula recalc implemented. |
| `SpreadsheetFile.exportXlsx` / `importXlsx` | partial | Minimal OOXML XLSX round trip implemented with inline strings/formulas. Needs styles, shared strings, charts, comments, conditional formatting, validations, pivots, images. |
| Formula calculation | partial | Supports arithmetic cell refs and `SUM(range)`. Needs broad Excel formula catalog and dependency graph. |
| `workbook.inspect` table/formula/match/sheet/workbook | partial | Core records implemented. Needs computedStyle, drawing, thread, definedName, richer include/exclude. |
| `workbook.render` visual verification | partial | SVG grid preview only. Needs PNG and chart/drawing rendering. |
| `workbook.trace` | partial | Basic formula precedent tree implemented for same-sheet/cross-sheet A1 references and ranges. Needs richer formula parser, cycle reporting UX, and large-trace summarization. |
| Charts/tables/sparklines/images/comments/data validations/conditional formats | todo | Public collection placeholders exist only as arrays. |

## Presentations skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Presentation.create`, slides collection, `slide.shapes.add` | done | Basic slide/shape/text facade works. |
| `PresentationFile.exportPptx` / `importPptx` | partial | Minimal PPTX round trip for text shapes works. Needs themes, masters, layouts, images, tables, charts, notes, comments, connectors. |
| `presentation.inspect` slide/textbox/shape/layout | partial | Basic stable records implemented. Needs all kinds and target context windows. |
| `presentation.resolve` | partial | Resolves slides/shapes. Needs chart/table/image/thread/text-range. |
| `slide.export({format:'layout'})` | done | Minimal layout JSON implemented. |
| `presentation.export` image preview/montage | partial | SVG preview implemented. Needs PNG/JPEG/WebP and montage. |
| Compose/JSX layout + token parser | partial | Helper-node compose engine implemented for row, column, layers, box, paragraph, run, shape, and rule with fill/hug/fixed sizing, gap, padding, stable names/ids, text class tokens, inspect, resolve, layout JSON, and PPTX roundtrip. `slide.autoLayout` now places existing shapes with horizontal/vertical flow, frame, gap, padding, and alignment. Package exports now include `./presentation-jsx`, `./presentation-jsx/jsx-runtime`, and `./presentation-jsx/jsx-dev-runtime` with `jsx`, `jsxs`, `jsxDEV`, `Fragment`, helper nodes, and function component support. Needs grid/table/chart/image nodes, fuller token parser, and collision detection. |
| Overlap/overflow/template fidelity QA | todo | Required before claiming skill parity. |

## Documents skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `DocumentModel.create`, paragraphs, `inspect` | done | Minimal paragraph document model. |
| `DocumentFile.exportDocx` / `importDocx` | partial | Minimal WordprocessingML round trip implemented. Needs styles, headers/footers, tables, lists, comments, tracked changes, fields, hyperlinks, citations, render gate. |
| Design presets / table geometry / OOXML patch helpers | todo | Required for agent document skill parity. |
| DOCX render-to-page images | todo | Could use LibreOffice/Poppler or a JS/PDF pipeline later. |

## PDF skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `PdfArtifact.create`, `PdfFile.exportPdf`, `PdfFile.importPdf` | partial | Minimal single-page PDF generation and extraction for generated files. Needs robust PDF parsing/creation. |
| PDF visual render verification | partial | SVG preview of modeled page only. Needs Poppler/PDF.js rasterization. |
| Text/table extraction | todo | Required for read/review/extract workflows. |
| Create polished reports with typography/layout/charts | todo | Needs layout engine and image/table/chart primitives. |

## Implementation priorities

1. Replace the hand-written minimal ZIP/XML emitters with a safer OOXML package layer.
2. Add a real formula dependency graph and a broader Excel formula catalog.
3. Extend the presentation compose/layout engine with grid, table, chart, image nodes, stable named layout trees, and collision detection.
4. Add raster rendering via a pluggable renderer (`sharp`, browser, canvas, or Poppler/LibreOffice adapters).
5. Add DOCX style/table/list/comment/tracked-change helpers.
6. Add robust PDF creation/extraction/rendering adapters.
7. Expand `help` catalogs from the observed skill docs into generated, testable API records.
