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
| `SpreadsheetFile.exportXlsx` / `importXlsx` | partial | Minimal OOXML XLSX round trip implemented with inline strings/formulas. Clean-room metadata part preserves comments, data validations, conditional formatting, worksheet tables, worksheet charts, worksheet images, and sparkline groups across roundtrip; worksheet XML also emits simple dataValidation/conditionalFormatting nodes. Needs styles, shared strings, native table/chart/image OOXML, pivots, and native threaded-comment/sparkline OOXML. |
| Formula calculation | partial | Supports arithmetic cell refs and `SUM(range)`. Needs broad Excel formula catalog and dependency graph. |
| `workbook.inspect` table/formula/match/sheet/workbook | partial | Core records implemented for workbook, sheet, table/structured table, formula, match, drawing/chart/image, sparkline, thread, dataValidation, and conditionalFormat with simple search filtering. Needs computedStyle, definedName, and richer include/exclude. |
| `workbook.render` visual verification | partial | SVG grid preview includes cell values, worksheet table outlines, simple chart previews, image placeholders/embedded data URLs, and sparkline previews. Needs PNG output and full drawing rendering. |
| `workbook.trace` | partial | Basic formula precedent tree implemented for same-sheet/cross-sheet A1 references and ranges. Needs richer formula parser, cycle reporting UX, and large-trace summarization. |
| Charts/tables/sparklines/images/comments/data validations/conditional formats | partial | Workbook threaded comment facade (`comments.setSelf/addThread/addReply/resolve/reopen`), `range.dataValidation`, `sheet.dataValidations.add`, `range.conditionalFormats.add/addCustom/deleteAll/clear`, worksheet tables (`sheet.tables.add`, `rows.add`, `getDataRows`, `getHeaderRowRange`, style/toggles), worksheet charts (`sheet.charts.add`, `setData`, `setPosition`, inferred series/category formulas), worksheet images (`sheet.images.add` with dataUrl/URI/prompt + anchors), sparkline groups (`sheet.sparklineGroups.add`, `range.sparklines.add`, delete/edit fields), inspect/resolve/render hints, and roundtrip metadata preservation are implemented. Needs native Excel threaded-comment/table/chart/image/sparkline OOXML, richer conditional-format rendering/evaluation, and pivot/data table support. |

## Presentations skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Presentation.create`, slides collection, `slide.shapes.add` | done | Basic slide/shape/text facade works. |
| `PresentationFile.exportPptx` / `importPptx` | partial | Minimal PPTX round trip for text shapes works; table/chart/image facades export as visible PPTX text placeholders so durable files retain user-visible content. Needs true PPTX table/chart/image OOXML, themes, masters, layouts, notes, comments, connectors. |
| `presentation.inspect` slide/textbox/shape/layout | partial | Stable records implemented for slide, textbox, shape, table, chart, image, and layout, with simple JSON search filtering. Needs notes/thread/text-range kinds, include/exclude shaping, and target context windows. |
| `presentation.resolve` | partial | Resolves slides, shapes, tables, charts, and images. Needs thread/text-range anchors and richer imported-object identity preservation. |
| `slide.export({format:'layout'})` | done | Minimal layout JSON implemented for shapes, textboxes, tables, charts, and images. |
| `presentation.export` image preview/montage | partial | SVG preview implemented for shapes, tables, charts, and image placeholders. Needs PNG/JPEG/WebP and montage. |
| Compose/JSX layout + token parser | partial | Helper-node compose engine implemented for row, column, grid, layers, box, paragraph, run, shape, table, chart, image, and rule with fill/hug/fixed sizing, fr/fixed grid tracks, spans, gaps, padding, stable names/ids, text class tokens, inspect, resolve, layout JSON, and PPTX roundtrip. `slide.autoLayout` now places existing shapes with horizontal/vertical flow, frame, gap, padding, and alignment. Package exports now include `./presentation-jsx`, `./presentation-jsx/jsx-runtime`, and `./presentation-jsx/jsx-dev-runtime` with `jsx`, `jsxs`, `jsxDEV`, `Fragment`, helper nodes, and function component support. Needs fuller token parser and real native PPTX table/chart/image XML. |
| Overlap/overflow/template fidelity QA | partial | `slide.validateLayout()` and `presentation.validateLayout()` detect off-canvas elements, geometry overlaps, basic text overflow, and table cell overflow with bounded NDJSON issue records. Needs visual/render-based QA, connector checks, placeholder/template fidelity gates, and chart/data consistency checks. |

## Documents skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `DocumentModel.create`, paragraphs, lists, headers/footers, hyperlinks, fields, citations, tables, comments, styles, `inspect` | partial | Block model now supports styled paragraph blocks, real list-item blocks, header/footer blocks, hyperlink blocks, field blocks, citation blocks with metadata, table blocks/cells, named styles, comment anchors, inspect/resolve/search, `document.render()` SVG page preview, and help catalog entries. Needs sections, images, tracked changes, and richer layout/page model. |
| `DocumentFile.exportDocx` / `importDocx` | partial | WordprocessingML round trip now writes document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlink relationships, fldSimple fields, citation bookmarks, clean-room metadata, relationships, styled paragraphs, list items, and tables; import restores paragraphs, list items, tables, comments, headers, footers, hyperlinks, fields, and citations. Needs sections, images, tracked changes, and stronger style fidelity. |
| Design presets / table geometry / OOXML patch helpers | todo | Required for agent document skill parity. |
| DOCX render-to-page images | partial | `document.render()` provides a clean-room SVG page preview for headers, footers, paragraphs, hyperlinks, fields, citations, list items, and tables. Needs DOCX-to-PNG/PDF render gate via LibreOffice/Poppler or equivalent, page pagination fidelity, and visual diff workflow. |

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
3. Replace presentation table/chart/image placeholders with true native PPTX OOXML parts while preserving inspect/resolve/layout behavior.
4. Add raster rendering via a pluggable renderer (`sharp`, browser, canvas, or Poppler/LibreOffice adapters).
5. Add DOCX sections, fields, hyperlinks, citations, images, tracked changes, stronger style/table geometry audits, and a real render gate.
6. Add robust PDF creation/extraction/rendering adapters.
7. Expand `help` catalogs from the observed skill docs into generated, testable API records.
