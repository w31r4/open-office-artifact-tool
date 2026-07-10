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
| `help(query, opts)` bounded NDJSON API discovery | partial | Shared `HELP_CATALOG` and `helpArtifact(...)` power Workbook/Presentation/Document/PDF help methods; `docs/api.md` is generated from the catalog. Needs broader examples, schemas, and full formula/API coverage. |
| Render/preview loop | partial | `renderArtifact(...)` plus per-artifact render/export adapters return SVG `FileBlob` previews with metadata for workbooks, presentations, documents, and PDFs. PNG/WebP rasterization and external Poppler/LibreOffice adapters are roadmap. |
| Shared verification API | partial | `verifyArtifact(...)` plus per-artifact `verify()` methods emit bounded NDJSON issues for workbook formula/structure errors, presentation layout issues, document fake-list/broken-link/comment issues, and PDF text/table/page issues. Needs render-backed verification and full skill-specific gates. |
| Layout JSON export | partial | Slide layout JSON implemented. Workbook/document layout exports are roadmap. |
| Durable file export/import smoke tests | done | Minimal XLSX/PPTX/DOCX/PDF round trips pass. |

## Spreadsheets skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Workbook.create`, worksheets collection, `getRange`, values/formulas | done | Matrix writes and simple formula recalc implemented. |
| `SpreadsheetFile.exportXlsx` / `importXlsx` | partial | Minimal OOXML XLSX round trip implemented with inline strings/formulas. Export now writes native Excel table parts (`xl/tables/table*.xml`), worksheet `tableParts`, native image drawing/media parts (`xl/drawings/drawing*.xml`, `xl/media/image*.png`, drawing rels), and content-type overrides/defaults; clean-room metadata still preserves comments, data validations, conditional formatting, worksheet charts, and sparkline groups across roundtrip. Needs styles, shared strings, native chart OOXML, pivots, and native threaded-comment/sparkline OOXML. |
| Formula calculation | partial | Supports arithmetic cell refs and `SUM(range)`. Needs broad Excel formula catalog and dependency graph. |
| `workbook.inspect` table/formula/match/sheet/workbook | partial | Core records implemented for workbook, sheet, table/structured table, formula, match, drawing/chart/image, sparkline, thread, dataValidation, and conditionalFormat with simple search filtering. Needs computedStyle, definedName, and richer include/exclude. |
| `workbook.render` visual verification | partial | SVG grid preview includes cell values, worksheet table outlines, simple chart previews, image placeholders/embedded data URLs, and sparkline previews. Needs PNG output and full drawing rendering. |
| `workbook.trace` | partial | Basic formula precedent tree implemented for same-sheet/cross-sheet A1 references and ranges. Needs richer formula parser, cycle reporting UX, and large-trace summarization. |
| Charts/tables/sparklines/images/comments/data validations/conditional formats | partial | Workbook threaded comment facade (`comments.setSelf/addThread/addReply/resolve/reopen`), `range.dataValidation`, `sheet.dataValidations.add`, `range.conditionalFormats.add/addCustom/deleteAll/clear`, worksheet tables (`sheet.tables.add`, `rows.add`, `getDataRows`, `getHeaderRowRange`, style/toggles, native table XML parts), worksheet charts (`sheet.charts.add`, `setData`, `setPosition`, inferred series/category formulas), worksheet images (`sheet.images.add` with dataUrl/URI/prompt + anchors and native image drawing/media parts for base64 data URLs), sparkline groups (`sheet.sparklineGroups.add`, `range.sparklines.add`, delete/edit fields), inspect/resolve/render hints, and roundtrip metadata preservation are implemented. Needs native Excel threaded-comment/chart/sparkline OOXML, richer conditional-format rendering/evaluation, and pivot/data table support. |

## Presentations skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Presentation.create`, slides collection, `slide.shapes.add` | done | Basic slide/shape/text facade works. |
| `PresentationFile.exportPptx` / `importPptx` | partial | Minimal PPTX round trip for text shapes works; table facades export as native DrawingML table `p:graphicFrame`/`a:tbl` XML, chart facades export as native `ppt/charts/chart*.xml` parts with `c:chart` relationships, and `dataUrl` image facades export as native `ppt/media/image*.ext` parts with slide `p:pic` XML and relationships. Needs themes, masters, layouts, notes, comments, connectors, and chart/image/table import restoration. |
| `presentation.inspect` slide/textbox/shape/layout | partial | Stable records implemented for slide, textbox, shape, table, chart, image, and layout, with simple JSON search filtering. Needs notes/thread/text-range kinds, include/exclude shaping, and target context windows. |
| `presentation.resolve` | partial | Resolves slides, shapes, tables, charts, and images. Needs thread/text-range anchors and richer imported-object identity preservation. |
| `slide.export({format:'layout'})` | done | Minimal layout JSON implemented for shapes, textboxes, tables, charts, and images. |
| `presentation.export` image preview/montage | partial | SVG preview implemented for shapes, tables, charts, and image placeholders. Needs PNG/JPEG/WebP and montage. |
| Compose/JSX layout + token parser | partial | Helper-node compose engine implemented for row, column, grid, layers, box, paragraph, run, shape, table, chart, image, and rule with fill/hug/fixed sizing, fr/fixed grid tracks, spans, gaps, padding, stable names/ids, text class tokens, inspect, resolve, layout JSON, and PPTX roundtrip. `slide.autoLayout` now places existing shapes with horizontal/vertical flow, frame, gap, padding, and alignment. Package exports now include `./presentation-jsx`, `./presentation-jsx/jsx-runtime`, and `./presentation-jsx/jsx-dev-runtime` with `jsx`, `jsxs`, `jsxDEV`, `Fragment`, helper nodes, and function component support. Table nodes emit native PPTX DrawingML tables, chart nodes emit native PPTX chart parts, and `dataUrl` images emit native PPTX image parts. Needs fuller token parser and richer chart schema coverage. |
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
| `PdfArtifact.create`, `PdfFile.exportPdf`, `PdfFile.importPdf` | partial | Multi-page modeled PDF artifacts now support page text, table regions, PDF export, import from embedded clean-room metadata, and heuristic import from visible text/table rows. Needs robust arbitrary-PDF parsing via PDF.js/PDFium/Poppler or similar. |
| PDF visual render verification | partial | SVG page render now draws modeled page text and tables for clean-room PDF artifacts. Needs Poppler/PDF.js rasterization to PNG for arbitrary PDFs and visual diff workflow. |
| Text/table extraction | partial | `extractText()` and `extractTables()` work for modeled/generated PDFs, and import heuristically detects visible pipe-delimited table rows when clean-room metadata is absent. Needs robust text positioning/table extraction for arbitrary PDFs. |
| Create polished reports with typography/layout/charts | partial | Modeled PDF creation supports text pages and table rendering/export. Needs richer report layout, images/charts, pagination, typography controls, and raster QA. |

## Implementation priorities

1. Replace the hand-written minimal ZIP/XML emitters with a safer OOXML package layer.
2. Add a real formula dependency graph and a broader Excel formula catalog.
3. Replace presentation table/chart/image placeholders with true native PPTX OOXML parts while preserving inspect/resolve/layout behavior.
4. Add raster rendering via a pluggable renderer (`sharp`, browser, canvas, or Poppler/LibreOffice adapters).
5. Add DOCX sections, fields, hyperlinks, citations, images, tracked changes, stronger style/table geometry audits, and a real render gate.
6. Add robust PDF creation/extraction/rendering adapters.
7. Expand `help` catalogs from the observed skill docs into generated, testable API records.
