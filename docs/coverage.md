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
| Render/preview loop | partial | `renderArtifact(...)` plus per-artifact render/export adapters return SVG `FileBlob` previews with metadata for workbooks, presentations, documents, and PDFs. `renderArtifact(..., { format, renderer })` supports pluggable PNG/WebP/JPEG/PDF conversion adapters over the SVG/FileBlob source; optional adapters include Playwright SVG/HTML → PNG/WebP/JPEG/PDF, sharp SVG/PNG/JPEG/WebP → PNG/WebP/JPEG, node-canvas SVG/PNG/JPEG/WebP → PNG/JPEG, Poppler PDF → PNG/PPM/TIFF page rasterization, LibreOffice DOCX/XLSX/PPTX/HTML → PDF/office-supported conversions, and native Office via the Node/C# JSON bridge. Raster adapters hard-error when their optional peer dependency or CLI is missing rather than silently faking output. |
| Shared verification API | partial | `verifyArtifact(...)` plus per-artifact `verify()` methods emit bounded NDJSON issues for workbook formula/structure errors, presentation layout issues, document fake-list/broken-link/comment/layout issues, and PDF text/table/page issues. `visualQaArtifact(...)` adds render-backed metadata/hash/baseline checks for SVG or raster/PDF outputs and optional PNG pixel-diff metrics/issues via `pixelDiff: true`. Needs broader image-format diff support and full skill-specific gates. |
| Layout JSON export | partial | Slide layout JSON, document page-aware layout JSON, and workbook/worksheet layout JSON are implemented. Workbook layout JSON includes cells, tables, charts, images, sparklines, and rules with pixel bounding boxes. Needs richer pagination/windowing and target context slicing across all families. |
| Durable file export/import smoke tests | done | Minimal XLSX/PPTX/DOCX/PDF round trips pass. |

## Spreadsheets skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Workbook.create`, worksheets collection, `getRange`, values/formulas | done | Matrix writes and simple formula recalc implemented. |
| `SpreadsheetFile.exportXlsx` / `importXlsx` | partial | Minimal OOXML XLSX round trip implemented with formulas, native shared strings (`xl/sharedStrings.xml`), and basic native style records (`xl/styles.xml` for fills/fonts/number formats). Export now writes native Excel table parts (`xl/tables/table*.xml`), native chart parts (`xl/charts/chart*.xml` via worksheet drawings), native image drawing/media parts (`xl/drawings/drawing*.xml`, `xl/media/image*.png`, drawing rels), native x14 sparkline extension XML (`x14:sparklineGroups` in worksheet `extLst`), native threaded-comment/person parts (`xl/threadedComments/threadedComment*.xml`, `xl/persons/person.xml`), and content-type overrides/defaults; clean-room metadata still preserves comments, data validations, and conditional formatting across roundtrip. Needs pivots, richer style fidelity, shared formula/array formulas, and richer threaded-comment interoperability. |
| Formula calculation | partial | Supports arithmetic cell refs plus a broader clean-room formula catalog across math/statistical/logical/text/lookup helpers (`SUM`, `AVERAGE`, `MIN`, `MAX`, `COUNT`, `IF`, `AND`, `OR`, `NOT`, `ROUND`, `ABS`, `INT`, `CEILING`, `FLOOR`, `COUNTIF`, `SUMIF`, `VLOOKUP`, `XLOOKUP`, `CONCAT`, `TEXTJOIN`, `LEFT`, `RIGHT`, `MID`, `LEN`, `UPPER`, `LOWER`, `TRIM`) over ranges/arguments. Workbook-level formula dependency graph, dependent edges, missing-sheet reporting, and cycle detection now feed `recalculate`, `trace`, `inspect`, and `verify`. Needs array/dynamic formulas, structured references, and closer Excel coercion semantics. |
| `workbook.inspect` table/formula/match/sheet/workbook | partial | Core records implemented for workbook, sheet, table/structured table, formula, formula graph, match, style/computedStyle, drawing/chart/image, sparkline, thread, dataValidation, and conditionalFormat with simple search filtering. Needs definedName and richer include/exclude. |
| `workbook.render` visual verification | partial | SVG grid preview includes cell values, worksheet table outlines, simple chart previews, image placeholders/embedded data URLs, and sparkline previews. `workbook.render({ format: 'layout' })` returns layout JSON for cells, tables, charts, images, sparklines, and validation/format rules. Can be rasterized through the optional Playwright renderer adapter when Playwright/Chromium are installed. Needs full drawing rendering and render-backed QA. |
| `workbook.trace` | partial | Formula tracing now uses the same dependency model as recalculation, exposes precedent trees, graph nodes/edges via `workbook.formulaGraph()` and inspect kinds (`formulaNode`, `formulaEdge`, `formulaCycle`), and reports cycles/missing-sheet refs to verify. Needs richer parser coverage, structured references, and large-trace summarization. |
| Charts/tables/sparklines/images/comments/data validations/conditional formats | partial | Workbook threaded comment facade (`comments.setSelf/addThread/addReply/resolve/reopen` with native threaded-comment/person XML export/import), `range.dataValidation`, `sheet.dataValidations.add`, `range.conditionalFormats.add/addCustom/deleteAll/clear`, worksheet tables (`sheet.tables.add`, `rows.add`, `getDataRows`, `getHeaderRowRange`, style/toggles, native table XML parts), worksheet charts (`sheet.charts.add`, `setData`, `setPosition`, inferred series/category formulas, native chart XML parts), worksheet images (`sheet.images.add` with dataUrl/URI/prompt + anchors and native image drawing/media parts for base64 data URLs), sparkline groups (`sheet.sparklineGroups.add`, `range.sparklines.add`, delete/edit fields, native x14 sparklineGroups XML export/import), inspect/resolve/render hints, roundtrip metadata preservation, and `workbook.verify()` consistency checks for table ranges, chart series/category data, image sources/bounds, sparkline target/source ranges, data-validation ranges/lists, conditional-format ranges/formulas, and comment targets are implemented. Needs richer conditional-format rendering/evaluation, pivot/data table support, and broader Excel compatibility testing. |

## Presentations skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `Presentation.create`, slides collection, `slide.shapes.add` | done | Basic slide/shape/text facade works. |
| `PresentationFile.exportPptx` / `importPptx` | partial | Minimal PPTX round trip for text shapes works; table facades export as native DrawingML table `p:graphicFrame`/`a:tbl` XML, chart facades export as native `ppt/charts/chart*.xml` parts with `c:chart` relationships, `dataUrl` image facades export as native `ppt/media/image*.ext` parts with slide `p:pic` XML and relationships, connector facades export as native `p:cxnSp`, speaker notes export as notesSlide parts, comment threads export as PPTX comments parts, and theme/layout facades export `ppt/theme/theme1.xml`, `ppt/slideMasters/slideMaster1.xml`, and `ppt/slideLayouts/slideLayout*.xml`. Import now restores native text shapes, tables, chart frames/basic chart data, pictures, connectors, notes, comments, themes, layouts, and placeholders from clean-room generated PPTX. Needs richer master/theme fidelity, richer comments interoperability, and broader third-party PPTX fidelity. |
| `presentation.inspect` slide/textbox/shape/layout | partial | Stable records implemented for deck, theme, layout templates, slide, textbox, shape, table, chart, image, connector, notes, comments/threads, and layout, with simple JSON search filtering. Needs text-range kinds, include/exclude shaping, and target context windows. |
| `presentation.resolve` | partial | Resolves presentations, themes, slide layouts, slides, shapes, tables, charts, images, connectors, and comment threads. Needs text-range anchors and richer imported-object identity preservation. |
| `slide.export({format:'layout'})` | done | Minimal layout JSON implemented for shapes, textboxes, tables, charts, and images. |
| `presentation.export` image preview/montage | partial | SVG preview implemented for shapes, tables, charts, connectors, and image placeholders, and can be rasterized to PNG/WebP/PDF through the optional Playwright renderer adapter. Needs montage and richer DrawingML fidelity. |
| Compose/JSX layout + token parser | partial | Helper-node compose engine implemented for row, column, grid, layers, box, paragraph, run, shape, table, chart, image, and rule with fill/hug/fixed sizing, fr/fixed grid tracks, spans, gaps, padding, stable names/ids, text class tokens, inspect, resolve, layout JSON, and PPTX roundtrip. `slide.autoLayout` now places existing shapes with horizontal/vertical flow, frame, gap, padding, and alignment. Package exports now include `./presentation-jsx`, `./presentation-jsx/jsx-runtime`, and `./presentation-jsx/jsx-dev-runtime` with `jsx`, `jsxs`, `jsxDEV`, `Fragment`, helper nodes, and function component support. Table nodes emit native PPTX DrawingML tables, chart nodes emit native PPTX chart parts, and `dataUrl` images emit native PPTX image parts. Needs fuller token parser and richer chart schema coverage. |
| Overlap/overflow/template fidelity QA | partial | `slide.validateLayout()` and `presentation.validateLayout()` detect off-canvas elements, connector endpoints outside the slide, geometry overlaps, basic text overflow, and table cell overflow; `presentation.verify()` also detects dangling comment targets. Needs visual/render-based QA, placeholder/template fidelity gates, and chart/data consistency checks. |

## Documents skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `DocumentModel.create`, paragraphs, lists, headers/footers, hyperlinks, fields, citations, images, sections, tables, comments, styles, `inspect` | partial | Block model now supports styled paragraph blocks, real list-item blocks, header/footer blocks, hyperlink blocks, field blocks, citation blocks with metadata, image blocks with dataUrl/URI/prompt metadata, section-break blocks with page size/orientation/margins, tracked insertion/deletion blocks, table blocks/cells, named styles, design presets, comment anchors, inspect/resolve/search, page-aware layout records, `document.render()` SVG/page-layout preview, and help catalog entries. Needs richer pagination and text measurement. |
| `DocumentFile.exportDocx` / `importDocx` | partial | WordprocessingML round trip now writes document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlink relationships, fldSimple fields, citation bookmarks, native inline image DrawingML/media parts, native section `w:sectPr` page setup, native tracked-change `w:ins`/`w:del` markup, clean-room metadata, relationships, styled paragraphs, list items, and tables; import restores paragraphs, list items, tables, comments, headers, footers, hyperlinks, fields, citations, images, sections, and tracked changes. Needs stronger style fidelity. |
| Design presets / table geometry / OOXML patch helpers | partial | `document.applyDesignPreset()` adds report/memo style sets, and `document.verify()` now covers unknown styles, invalid links/citations, malformed image metadata, invalid section setup, empty/ragged/wide tables, long table cells, and layout/visual overflow when `visualQa` is enabled. Needs richer style fidelity and safe OOXML patch helpers. |
| DOCX render-to-page images | partial | `document.render()` provides a clean-room SVG page preview for headers, footers, paragraphs, hyperlinks, fields, citations, images, section breaks, tracked changes, list items, and tables; `document.render({ format: 'layout' })` returns page-aware layout JSON; SVG previews can be rasterized to PNG/WebP/PDF through renderer adapters. Needs DOCX-to-PNG/PDF render gate via LibreOffice/Poppler or native Office, page pagination fidelity, and visual diff workflow. |

## PDF skill target

| Requirement | Status | Notes |
| --- | --- | --- |
| `PdfArtifact.create`, `PdfFile.exportPdf`, `PdfFile.importPdf` | partial | Multi-page modeled PDF artifacts now support page text, positioned text items, layout regions, table regions, image regions, PDF export, import from embedded clean-room metadata, injected parser adapters for arbitrary PDFs, optional PDF.js parsing (`open-office-artifact-tool/pdf/pdfjs`) for page geometry/text/table/image-placeholder extraction, and heuristic import from visible text/table/image-placeholder rows. Needs stronger arbitrary-PDF fidelity, real image extraction bytes, and Poppler/PDFium alternatives. |
| PDF visual render verification | partial | SVG page render now draws modeled page text, tables, and image regions for clean-room PDF artifacts, and SVG previews can be rasterized through the optional Playwright renderer adapter. Needs Poppler/PDF.js rasterization to PNG for arbitrary PDFs and visual diff workflow. |
| Text/table extraction | partial | `extractText()` and `extractTables()` work for modeled/generated PDFs; import heuristically detects visible pipe-delimited table rows when clean-room metadata is absent; injected parser adapters can populate positioned text items, layout regions, tables, and image placeholders; optional PDF.js adapter extracts page text geometry and heuristic tables. Needs robust table reconstruction for arbitrary PDFs and byte-level image extraction. |
| Create polished reports with typography/layout/charts | partial | Modeled PDF creation supports text pages, image regions, and table rendering/export. Needs richer report layout, charts, pagination, typography controls, and raster QA. |

## Native Office bridge target

| Requirement | Status | Notes |
| --- | --- | --- |
| Optional Node wrapper without core Windows/Office dependency | done | `open-office-artifact-tool/native/office-bridge` exposes `callOfficeBridge`, `renderFileWithNativeOffice`, `createNativeOfficeRenderer`, `nativeOfficeStatus`, temp-file isolation, cleanup, timeout handling, and structured `OfficeBridgeError` responses. Covered by mock-backed npm smoke tests. |
| C# JSON stdin/stdout sidecar | partial | `native/OfficeBridge` contains a .NET 8 sidecar that reads one JSON request from stdin and writes one structured JSON response to stdout. It supports `status`, `render`/`convert`/`export`, and Office-specific operation names. Local verification is pending in this environment because `dotnet` is unavailable. |
| Windows Microsoft Office automation | partial | C# sidecar uses late-bound COM automation for Word DOCX PDF export/update fields/accept/reject revisions, Excel XLSX recalculation/autofit/PDF export, and PowerPoint PPTX PDF/PNG export. Office-specific integration tests remain gated by `OFFICE_NATIVE_TESTS=1` and require Windows + Office. |
| Dotnet tests for bridge protocol | partial | `native/OfficeBridge/tests` covers status, structured errors, JSON serialization, CLI stdin/stdout status, and graceful no-Office behavior. Tests cannot be run locally here because `dotnet` is not installed; CI runs them when `dotnet` is available. |

## Docs, examples, and CI

| Requirement | Status | Notes |
| --- | --- | --- |
| README clean-room/family/renderer/native bridge docs | partial | README covers clean-room goal, four artifact families, renderer adapters, optional Playwright renderer, PDF parser adapters, Node-side native Office wrapper, C# sidecar usage/tests, examples, release-check flow, and development commands. `docs/release.md` records current npm auth/publish blockers. Needs update once publishing/tagging succeeds. |
| Examples for DOCX/XLSX/PPTX/PDF/rendering | partial | `examples/` includes create DOCX report, create XLSX dashboard with charts/sparklines/comments, create PPTX compose deck with notes/comments/connectors, parse/render PDF, render via Playwright, and render via native Office bridge. Native bridge and Playwright examples skip gracefully when optional runtimes are unavailable. |
| Basic CI without Microsoft Office | partial | `.github/workflows/ci.yml` runs npm install/test/docs/release-check/pack and conditionally skips dotnet bridge tests when `native/OfficeBridge` or `dotnet` is unavailable. Needs release/tag automation after npm auth is configured. |

## Implementation priorities

1. Replace the hand-written minimal ZIP/XML emitters with a safer OOXML package layer.
2. Add a real formula dependency graph and a broader Excel formula catalog.
3. Replace presentation table/chart/image placeholders with true native PPTX OOXML parts while preserving inspect/resolve/layout behavior.
4. Add additional raster/render adapters (`sharp`, canvas, Poppler/LibreOffice, native Office) and render-backed QA on top of the shipped Playwright adapter. (sharp, canvas, Poppler, LibreOffice, native Office adapters shipped; render-backed QA via `visualQaArtifact` is partial — PNG pixel diff is available, broader image/PDF visual diff remains.)
5. Add DOCX real render gate, design presets, and richer style fidelity.
6. Add robust PDF creation/extraction/rendering adapters.
7. Expand `help` catalogs from the observed skill docs into generated, testable API records.
