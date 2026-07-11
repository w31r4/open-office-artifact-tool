# API catalog

Generated from `HELP_CATALOG` in `src/index.mjs`.

## document

| Name | Kind | Summary |
| --- | --- | --- |
| `document.addChange` | api | Append a tracked insertion or deletion block backed by native DOCX w:ins/w:del revision markup. |
| `document.addCitation` | api | Append a citation block with visible text and structured metadata preserved through clean-room DOCX metadata. |
| `document.addComment` | api | Attach a comment to a paragraph or table block using a stable target ID. |
| `document.addDeletion` | api | Append a tracked deletion with author/date metadata and native DOCX w:del/w:delText export. |
| `document.addField` | api | Append a Word field block exported as w:fldSimple with instruction text such as PAGE, REF, PAGEREF, or TOC. |
| `document.addFooter` | api | Add footer text exported as a DOCX footer part and referenced from section properties. |
| `document.addHeader` | api | Add header text exported as a DOCX header part and referenced from section properties. |
| `document.addHyperlink` | api | Append an external hyperlink backed by a DOCX relationship and w:hyperlink element. |
| `document.addImage` | api | Append an inspectable image block; dataUrl images export as native DOCX media parts with DrawingML inline pictures. |
| `document.addInsertion` | api | Append a tracked insertion with author/date metadata and native DOCX w:ins export. |
| `document.addListItem` | api | Append a real numbered or bulleted list item backed by DOCX numbering definitions. |
| `document.addParagraph` | api | Append a styled paragraph block and return an inspectable/resolveable paragraph object. |
| `document.addSection` | api | Append a DOCX section break with page size, orientation, margin, and break-type metadata backed by w:sectPr. |
| `document.addTable` | api | Append a Word-style table block with rows, columns, cell values, and style metadata. |
| `document.verify` | api | Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, and prose-like table cells. |
| `DocumentFile.exportDocx` | api | Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata. |
| `DocumentModel.create` | api | Create a document with paragraph, list, table, header/footer, style, and comment blocks. |

## pdf

| Name | Kind | Summary |
| --- | --- | --- |
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders. |
| `pdf.addImage` | api | Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box. |
| `pdf.extractTables` | api | Extract modeled table values and bounding boxes across all pages or a selected page. |
| `pdf.extractText` | api | Extract modeled text across all pages or a selected page. |
| `pdf.inspect` | api | Emit bounded NDJSON for pages, text, and table records. |
| `pdf.render` | api | Render a modeled PDF page to SVG in the current clean-room MVP. |
| `pdf.verify` | api | Return QA issues for empty pages, Unicode dashes, malformed tables, and out-of-bounds table boxes. |
| `PdfArtifact.create` | api | Create a modeled PDF artifact with pages, text, table regions, and image regions. |
| `PdfFile.exportPdf` | api | Export a modeled PDF artifact to a minimal PDF with visible text/table rows and embedded clean-room metadata. |
| `PdfFile.importPdf` | api | Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, or fall back to heuristic visible-text/table extraction. |

## presentation

| Name | Kind | Summary |
| --- | --- | --- |
| `compose.column` | api | Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels. |
| `compose.paragraph` | api | Create an editable text block with name, className/style text tokens, and stable inspect output. |
| `Presentation.create` | api | Create a deck with a default or explicit slide size. |
| `presentation.export` | api | Export a slide preview, deck montage, or layout JSON. |
| `presentation.inspect` | api | Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, and layout. |
| `presentation.resolve` | api | Map stable inspect anchor IDs back to editable facade objects. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. |
| `presentation.verify` | api | Return presentation QA issues, currently including layout validation issues. |
| `slide.addNotes` | api | Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' }). |
| `slide.autoLayout` | api | Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options. |
| `slide.charts.add` | api | Add an inspectable chart facade with chartType, title, categories, series, layout JSON, SVG preview, and PPTX placeholder output. |
| `slide.comments.addThread` | api | Attach threaded comments to slide elements; exported as PPTX comments parts and verified for dangling targets. |
| `slide.compose` | api | Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects. |
| `slide.connectors.add` | api | Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA. |
| `slide.images.add` | api | Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output. |
| `slide.shapes.add` | api | Add a shape/textbox with geometry, position, fill, line, and text. |
| `slide.tables.add` | api | Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output. |

## shared

| Name | Kind | Summary |
| --- | --- | --- |
| `createPlaywrightRenderer` | api | Create an optional Playwright renderer adapter from open-office-artifact-tool/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default. |
| `renderArtifact` | api | Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output. |
| `verifyArtifact` | api | Run an artifact's verify() method and return a bounded NDJSON QA report. |

## workbook

| Name | Kind | Summary |
| --- | --- | --- |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.PMT` | formula | Calculate a loan payment for constant payments and constant interest rate. |
| `fx.SUM` | formula | Sum numeric values across arguments and ranges. |
| `range.conditionalFormats.add` | api | Add a conditional formatting rule to a range; addCustom(expression, format) creates expression rules. |
| `range.dataValidation` | api | Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule }). |
| `sheet.charts.add` | api | Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas. |
| `sheet.images.add` | api | Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents. |
| `sheet.sparklineGroups.add` | api | Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand. |
| `sheet.tables.add` | api | Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles. |
| `SpreadsheetFile.exportXlsx` | api | Serialize a Workbook facade to an XLSX FileBlob. |
| `SpreadsheetFile.importXlsx` | api | Load an XLSX file into a Workbook facade. |
| `workbook.comments.addThread` | api | Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...'). |
| `Workbook.create` | api | Create an empty workbook; add worksheets before editing. |
| `workbook.formulaGraph` | api | Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet or range in the current clean-room MVP. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged. |
| `workbook.verify` | api | Return bounded QA issues for sheets, formulas, tables, charts, and comments. |
| `worksheet.getRange` | api | Select an A1 range for values, formulas, formatting, merge, fill, and copy operations. |

