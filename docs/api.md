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
| `document.applyDesignPreset` | api | Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews. |
| `document.inspect` | api | Emit bounded NDJSON for document blocks, comments, styles, headers/footers, and layout; narrow with search/target anchors and shape fields with include/exclude. |
| `document.layoutJson` | api | Return page-aware layout JSON with block bounding boxes, page records, style IDs, and design preset metadata. |
| `document.verify` | api | Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, visual layout overflow, and prose-like table cells. |
| `DocumentFile.exportDocx` | api | Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata. |
| `DocumentModel.create` | api | Create a document with paragraph, list, table, header/footer, style, and comment blocks. |

## pdf

| Name | Kind | Summary |
| --- | --- | --- |
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders. |
| `pdf.addImage` | api | Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box. |
| `pdf.extractTables` | api | Extract modeled table values and bounding boxes across all pages or a selected page. |
| `pdf.extractText` | api | Extract modeled text across all pages or a selected page. |
| `pdf.inspect` | api | Emit bounded NDJSON for pages, text, positioned text items, layout regions, tables, and images; narrow with search/target anchors and shape fields with include/exclude. |
| `pdf.layoutJson` | api | Return modeled PDF page layout JSON with page text, positioned text items, layout regions, tables, and images. |
| `pdf.render` | api | Render a modeled PDF page to SVG or return page layout JSON when called with { format: 'layout' }. |
| `pdf.resolve` | api | Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, layout regions, tables, and images. |
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
| `presentation.inspect` | api | Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, comments, and layout; narrow with search/target anchors and shape fields with include/exclude. |
| `presentation.layouts.add` | api | Create a reusable slide layout with placeholders; export writes slideLayout and slideMaster parts for clean-room PPTX roundtrip. |
| `presentation.resolve` | api | Map stable inspect anchor IDs back to editable facade objects. |
| `presentation.theme` | api | Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. |
| `presentation.verify` | api | Return presentation QA issues, currently including layout validation issues. |
| `slide.addNotes` | api | Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' }). |
| `slide.applyLayout` | api | Apply a slide layout to materialize editable placeholder shapes and preserve layout identity for inspect, verify, and PPTX export. |
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
| `createCanvasRenderer` | api | Create an optional node-canvas renderer adapter from open-office-artifact-tool/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG. |
| `createLibreOfficeRenderer` | api | Create a LibreOffice CLI renderer adapter from open-office-artifact-tool/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF. |
| `createNativeOfficeRenderer` | api | Create a native Office renderer adapter from open-office-artifact-tool/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors. |
| `createPlaywrightRenderer` | api | Create an optional Playwright renderer adapter from open-office-artifact-tool/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default. |
| `createPopplerRenderer` | api | Create a Poppler CLI renderer adapter from open-office-artifact-tool/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF. |
| `createSharpRenderer` | api | Create an optional sharp renderer adapter from open-office-artifact-tool/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG. |
| `renderArtifact` | api | Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output. |
| `renderFileWithNativeOffice` | api | Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output. |
| `verifyArtifact` | api | Run an artifact's verify() method and return a bounded NDJSON QA report. |
| `visualQaArtifact` | api | Render an artifact, record deterministic render metadata/hash, validate empty or malformed render output, optionally compare against a baseline render, and compute PNG pixel-diff metrics when requested. |

## workbook

| Name | Kind | Summary |
| --- | --- | --- |
| `fx.ABS` | formula | Return the absolute value of a number. |
| `fx.AND` | formula | Return TRUE when all conditions are true. |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.CEILING` | formula | Round a number up to the nearest significance. |
| `fx.CONCAT` | formula | Concatenate text values and ranges. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.COUNTIF` | formula | Count values in a range that match a criterion. |
| `fx.FLOOR` | formula | Round a number down to the nearest significance. |
| `fx.IF` | formula | Return one value when a condition is true and another when false. |
| `fx.INT` | formula | Round a number down to the nearest integer. |
| `fx.LEFT` | formula | Return characters from the start of a text value. |
| `fx.LEN` | formula | Return the length of a text value. |
| `fx.LOWER` | formula | Convert text to lowercase. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MID` | formula | Return characters from the middle of a text value. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.OR` | formula | Return TRUE when any condition is true. |
| `fx.PMT` | formula | Calculate a loan payment for constant payments and constant interest rate. |
| `fx.RIGHT` | formula | Return characters from the end of a text value. |
| `fx.ROUND` | formula | Round a numeric value to a fixed number of decimal places. |
| `fx.SUM` | formula | Sum numeric values across arguments and ranges. |
| `fx.SUMIF` | formula | Sum values whose corresponding criteria range entries match a criterion. |
| `fx.TEXTJOIN` | formula | Join text values with a delimiter and optional empty-value skipping. |
| `fx.TRIM` | formula | Trim leading/trailing whitespace and collapse internal whitespace. |
| `fx.UPPER` | formula | Convert text to uppercase. |
| `fx.VLOOKUP` | formula | Look up a value in the first column of a table range and return a value from another column. |
| `fx.XLOOKUP` | formula | Look up a value in one range and return the corresponding value from another range. |
| `range.conditionalFormats.add` | api | Add a conditional formatting rule to a range; addCustom(expression, format) creates expression rules. |
| `range.dataValidation` | api | Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule }). |
| `range.format` | api | Assign basic cell style metadata such as fill, font, and numberFormat; XLSX export writes native styles.xml and cell style indexes. |
| `sheet.charts.add` | api | Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas. |
| `sheet.images.add` | api | Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents. |
| `sheet.sparklineGroups.add` | api | Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand. |
| `sheet.tables.add` | api | Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles. |
| `SpreadsheetFile.exportXlsx` | api | Serialize a Workbook facade to an XLSX FileBlob. |
| `SpreadsheetFile.importXlsx` | api | Load an XLSX file into a Workbook facade. |
| `workbook.comments.addThread` | api | Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...'). |
| `Workbook.create` | api | Create an empty workbook; add worksheets before editing. |
| `workbook.formulaGraph` | api | Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude. |
| `workbook.layoutJson` | api | Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, and rule bounding boxes in pixels. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }. |
| `workbook.structuredReferences` | formula | Evaluate a clean-room subset of Excel structured references such as TableName[Column] in formulas, expanding them to table data-body cell precedents. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged. |
| `workbook.verify` | api | Return bounded QA issues for sheets, formulas, tables, charts, and comments. |
| `worksheet.getRange` | api | Select an A1 range for values, formulas, formatting, merge, fill, and copy operations. |

