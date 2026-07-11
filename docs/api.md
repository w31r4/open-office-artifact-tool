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
| `document.addParagraph` | api | Append a styled paragraph block with optional run-level styles and return an inspectable/resolveable paragraph object. |
| `document.addSection` | api | Append a DOCX section break with page size, orientation, margin, and break-type metadata backed by w:sectPr. |
| `document.addTable` | api | Append a Word-style table block with rows, columns, cell values, and style metadata. |
| `document.applyDesignPreset` | api | Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews. |
| `document.inspect` | api | Emit bounded NDJSON for document blocks, comments, styles, headers/footers, and layout; narrow with search/target anchors and shape fields with include/exclude. |
| `document.layoutJson` | api | Return page-aware layout JSON with block bounding boxes, page records, style IDs, design preset metadata, and target/search context slicing. |
| `document.render` | api | Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs. |
| `document.resolve` | api | Resolve stable document, block, header/footer, comment, style, and editable text-range IDs. |
| `document.styles.effective` | api | Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata. |
| `document.textRange` | api | Inspect or resolve stable textRange anchors such as blockId/text for editable document block, header/footer, and comment text. |
| `document.verify` | api | Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, visual layout overflow, and prose-like table cells. |
| `DocumentFile.exportDocx` | api | Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata. |
| `DocumentFile.importDocx` | api | Import DOCX bytes into the clean-room document facade, restoring native parts and embedded metadata when available. |
| `DocumentFile.inspectDocx` | api | Inspect a DOCX zip package as bounded NDJSON part records with safe part paths, sizes, content types, and optional XML/JSON previews. |
| `DocumentFile.patchDocx` | api | Apply safe in-package DOCX XML/JSON/binary patches with path traversal validation and return a patched DOCX FileBlob. |
| `DocumentModel.create` | api | Create a document with paragraph, list, table, header/footer, style, and comment blocks. |

### document details

#### `document.addChange`

Append a tracked insertion or deletion block backed by native DOCX w:ins/w:del revision markup.

**Schema parameters:**

- `changeType` (string) required — insert or delete.
- `text` (string) required — Revision text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended tracked-change block.

#### `document.addCitation`

Append a citation block with visible text and structured metadata preserved through clean-room DOCX metadata.

**Schema parameters:**

- `text` (string) required — Visible citation text.
- `metadata` (object) — Structured citation metadata.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `citation` (DocumentCitationBlock) — Appended citation block.

#### `document.addComment`

Attach a comment to a paragraph or table block using a stable target ID.

**Schema parameters:**

- `target` (string|object) required — Stable block ID or block facade.
- `text` (string) required — Comment text.
- `author` (string) — Comment author.
- `resolved` (boolean) — Initial resolution state.

**Schema returns:**

- `comment` (DocumentComment) — Attached comment with stable ID.

#### `document.addDeletion`

Append a tracked deletion with author/date metadata and native DOCX w:del/w:delText export.

**Schema parameters:**

- `text` (string) required — Deleted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended tracked deletion.

#### `document.addField`

Append a Word field block exported as w:fldSimple with instruction text such as PAGE, REF, PAGEREF, or TOC.

**Schema parameters:**

- `instruction` (string) required — Word field instruction such as PAGE, REF, PAGEREF, or TOC.
- `display` (string) — Visible fallback/result text.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `field` (DocumentFieldBlock) — Appended field block.

#### `document.addFooter`

Add footer text exported as a DOCX footer part and referenced from section properties.

**Schema parameters:**

- `text` (string) required — Footer text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.

**Schema returns:**

- `footer` (DocumentHeaderFooterBlock) — Appended footer block.

#### `document.addHeader`

Add header text exported as a DOCX header part and referenced from section properties.

**Schema parameters:**

- `text` (string) required — Header text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.

**Schema returns:**

- `header` (DocumentHeaderFooterBlock) — Appended header block.

#### `document.addHyperlink`

Append an external hyperlink backed by a DOCX relationship and w:hyperlink element.

**Schema parameters:**

- `text` (string) required — Visible link text.
- `url` (string) required — External hyperlink URL.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `hyperlink` (DocumentHyperlinkBlock) — Appended external hyperlink block.

#### `document.addImage`

Append an inspectable image block; dataUrl images export as native DOCX media parts with DrawingML inline pictures.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `widthPx` (number) — Rendered width in pixels.
- `heightPx` (number) — Rendered height in pixels.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `image` (DocumentImageBlock) — Appended image block.

#### `document.addInsertion`

Append a tracked insertion with author/date metadata and native DOCX w:ins export.

**Schema parameters:**

- `text` (string) required — Inserted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended tracked insertion.

#### `document.addListItem`

Append a real numbered or bulleted list item backed by DOCX numbering definitions.

**Schema parameters:**

- `text` (string) required — List item text.
- `listType` (string) — bullet or numbered.
- `level` (number) — Zero-based list nesting level.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `listItem` (DocumentListItemBlock) — Appended native-numbering list item.

#### `document.addParagraph`

Append a styled paragraph block with optional run-level styles and return an inspectable/resolveable paragraph object.

**Schema parameters:**

- `text` (string) required — Paragraph text.
- `styleId` (string) — Named paragraph style ID.
- `name` (string) — Inspectable block name.
- `runs` (object[]) — Optional run-level text/style spans.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Appended paragraph block with stable ID.

#### `document.addSection`

Append a DOCX section break with page size, orientation, margin, and break-type metadata backed by w:sectPr.

**Schema parameters:**

- `breakType` (string) — Section break type such as nextPage or continuous.
- `orientation` (string) — portrait or landscape.
- `pageSize` (object) — Page width/height in twentieths of a point.
- `margins` (object) — Top/right/bottom/left margins in twentieths of a point.

**Schema returns:**

- `section` (DocumentSectionBlock) — Appended section break block.

#### `document.addTable`

Append a Word-style table block with rows, columns, cell values, and style metadata.

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `name` (string) — Inspectable table name.
- `styleId` (string) — Table style ID.
- `widthDxa` (number) — Table width in twentieths of a point.
- `columnWidthsDxa` (number[]) — Column widths in twentieths of a point.
- `cellMarginsDxa` (object) — Cell margins in twentieths of a point.
- `borderColor` (string) — Table border color.
- `headerFill` (string) — Header-row fill color.

**Schema returns:**

- `table` (DocumentTableBlock) — Appended table block.

#### `document.applyDesignPreset`

Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews.

**Schema parameters:**

- `name` (string) required — report, memo, or a custom preset name.
- `styles` (object) — Style overrides merged into the preset.

**Schema returns:**

- `document` (DocumentModel) — The mutated document facade.

#### `document.inspect`

Emit bounded NDJSON for document blocks, comments, styles, headers/footers, and layout; narrow with search/target anchors and shape fields with include/exclude.

**Examples:**

- document.inspect({ kind: 'paragraph,comment', target: comment.id, maxChars: 4000 })

**Options:**

- kind
- search
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- maxChars

**Schema parameters:**

- `kind` (string) — Comma-separated block/comment/style/textRange/layout kinds.
- `search` (string) — Case-insensitive record filter.
- `target` (string) — Stable target ID/anchor.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.
- `include` (string) — Comma-separated fields to keep.
- `exclude` (string) — Comma-separated fields to omit.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Bounded { ndjson, truncated } inspection result.

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `document.layoutJson`

Return page-aware layout JSON with block bounding boxes, page records, style IDs, design preset metadata, and target/search context slicing.

**Schema parameters:**

- `pageWidth` (number) — Modeled page width in pixels.
- `pageHeight` (number) — Modeled page height in pixels.
- `margin` (number) — Modeled page margin in pixels.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive element filter.
- `before` (number) — Context elements before matches.
- `after` (number) — Context elements after matches.

**Schema returns:**

- `layout` (object) — Page-aware document layout tree.

#### `document.render`

Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs.

**Schema parameters:**

- `format` (string) — svg by default, layout, docx, pdf, png, or another renderer output.
- `source` (string) — Set to docx to render exported DOCX bytes.
- `renderer` (function) — Optional LibreOffice/native/raster renderer adapter.
- `pageWidth` (number) — Modeled SVG/layout page width.
- `pageHeight` (number) — Modeled SVG/layout page height.

**Schema returns:**

- `blob` (FileBlob) — SVG, layout JSON, DOCX, or converted renderer output.

#### `document.resolve`

Resolve stable document, block, header/footer, comment, style, and editable text-range IDs.

**Schema parameters:**

- `id` (string) required — Stable document, block, header/footer, comment, style, or text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `document.styles.effective`

Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata.

**Schema parameters:**

- `styleId` (string) required — Named style ID to resolve through basedOn inheritance.

**Schema returns:**

- `style` (object|undefined) — Resolved effective style or undefined.

#### `document.textRange`

Inspect or resolve stable textRange anchors such as blockId/text for editable document block, header/footer, and comment text.

**Schema parameters:**

- `id` (string) required — Stable text range ID ending in /text.

**Schema returns:**

- `textRange` (TextRange|undefined) — Editable text-range facade or undefined.

#### `document.verify`

Return QA issues for fake lists, invalid links/citations, unknown styles, malformed tables, bad image dimensions/data URLs, section setup, dangling comments, visual layout overflow, and prose-like table cells.

**Schema parameters:**

- `visualQa` (boolean) — Include modeled layout overflow checks.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Document semantic/layout QA result.

#### `DocumentFile.exportDocx`

Export DocumentModel to a DOCX package with document.xml, styles.xml, comments.xml, numbering.xml, header/footer parts, hyperlinks, fields, citations, and metadata.

**Schema parameters:**

- `document` (DocumentModel) required — Document facade to serialize.

**Schema returns:**

- `blob` (FileBlob) — DOCX package bytes.

#### `DocumentFile.importDocx`

Import DOCX bytes into the clean-room document facade, restoring native parts and embedded metadata when available.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.

**Schema returns:**

- `document` (DocumentModel) — Imported editable document facade.

#### `DocumentFile.inspectDocx`

Inspect a DOCX zip package as bounded NDJSON part records with safe part paths, sizes, content types, and optional XML/JSON previews.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `includeText` (boolean) — Include bounded XML/JSON/relationship previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual part.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — DOCX package part records and bounded NDJSON.

#### `DocumentFile.patchDocx`

Apply safe in-package DOCX XML/JSON/binary patches with path traversal validation and return a patched DOCX FileBlob.

**Examples:**

- await DocumentFile.patchDocx(docx, [{ path: 'customXml/review-note.xml', text: '<review>ok</review>' }])

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `patches` (array|object) required — Path-validated package part edits with text/xml/json/bytes/remove.
- `maxPatchBytes` (number) — Per-part patch size limit.

**Schema returns:**

- `docx` (FileBlob) — Patched DOCX FileBlob with metadata.patchedParts.

#### `DocumentModel.create`

Create a document with paragraph, list, table, header/footer, style, and comment blocks.

**Schema parameters:**

- `name` (string) — Document name.
- `designPreset` (string) — Initial design preset name.
- `styles` (object) — Named style definitions.
- `paragraphs` (string[]) — Convenience paragraph list; the first paragraph uses Title style.
- `blocks` (object[]) — Ordered paragraph/list/table/link/field/citation/image/section/change block models.
- `headers` (object[]) — Header block models.
- `footers` (object[]) — Footer block models.
- `comments` (object[]) — Comment models targeting stable block IDs.

**Schema returns:**

- `document` (DocumentModel) — Editable document facade.

## pdf

| Name | Kind | Summary |
| --- | --- | --- |
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders. |
| `pdf.addChart` | api | Add a modeled bar/line chart region with categories, series, title, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip. |
| `pdf.addImage` | api | Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box. |
| `pdf.addPage` | api | Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, and charts. |
| `pdf.addTable` | api | Add a modeled table with cell values and a page-space bounding box to the first PDF page. |
| `pdf.addText` | api | Add positioned PDF text with page-space bbox, font metadata, inspect/resolve/layout records, and SVG preview rendering. |
| `pdf.extractTables` | api | Extract modeled table values and bounding boxes across all pages or a selected page. |
| `pdf.extractText` | api | Extract modeled text across all pages or a selected page. |
| `pdf.inspect` | api | Emit bounded NDJSON for pages, text, positioned text items, layout regions, tables, images, and charts; narrow with search/target anchors and shape fields with include/exclude. |
| `pdf.layoutJson` | api | Return modeled PDF page layout JSON with page text, positioned text items, layout regions, tables, images, charts, and target/search context slicing. |
| `pdf.render` | api | Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters. |
| `pdf.resolve` | api | Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, layout regions, tables, images, and charts. |
| `pdf.verify` | api | Return QA issues for empty pages, Unicode dashes, text extraction sanity, page geometry, text/region/table/image/chart bounds, invalid image data URLs, malformed tables, and chart data. |
| `PdfArtifact.create` | api | Create a modeled PDF artifact with pages, text, table regions, and image regions. |
| `PdfFile.exportPdf` | api | Export a modeled artifact as a real multi-page PDF with positioned text, vector tables/charts, embedded PNG images, and clean-room metadata. |
| `PdfFile.importPdf` | api | Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, normalize parser image bytes/base64 into data URLs, reconstruct tables from positioned text geometry when explicit tables are absent, or fall back to heuristic visible-text/table extraction. |
| `PdfFile.inspectPdf` | api | Inspect PDF bytes as bounded file/object records including version, byte size, page/object counts, embedded clean-room model presence, and EOF integrity. |

### pdf details

#### `createPdfjsParser`

Create an optional PDF.js parser adapter from open-office-artifact-tool/pdf/pdfjs to extract page geometry, positioned text, heuristic tables, and image placeholders.

**Examples:**

- const parser = createPdfjsParser({ getDocumentOptions: { useSystemFonts: true } })

**Schema parameters:**

- `pdfjs` (object) — Injected PDF.js module; otherwise pdfjs-dist is loaded.
- `getDocumentOptions` (object) — Options merged into PDF.js getDocument().
- `textContentOptions` (object) — Options merged into getTextContent().

**Schema returns:**

- `parser` (function) — Parser adapter for PdfFile.importPdf().

#### `pdf.addChart`

Add a modeled bar/line chart region with categories, series, title, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip.

**Examples:**

- pdf.addChart({ pageIndex: 0, chartType: 'bar', categories: ['A', 'B'], series: [{ name: 'Score', values: [2, 4] }], bbox: [72, 430, 468, 180] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `chartType` (string) — bar or line.
- `title` (string) — Visible chart title.
- `categories` (string[]) required — Category labels.
- `series` (object[]) required — Series with name, numeric values, and optional color.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.

**Schema returns:**

- `chart` (PdfChart) — Inspectable chart facade with stable ID.

#### `pdf.addImage`

Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box.

**Examples:**

- pdf.addImage({ pageIndex: 0, dataUrl, alt: 'Approval mark', bbox: [430, 60, 64, 64] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `dataUrl` (string) — Embedded image data URL; generated PDF export currently supports PNG.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Image generation/extraction prompt metadata.
- `alt` (string) — Alternative text.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `fit` (string) — contain or cover intent metadata.

**Schema returns:**

- `image` (PdfImage) — Inspectable image facade with stable ID.

#### `pdf.addPage`

Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, and charts.

**Examples:**

- pdf.addPage({ width: 612, height: 792, text: 'Appendix' })

**Schema parameters:**

- `width` (number) — Page width in points; defaults to 612.
- `height` (number) — Page height in points; defaults to 792.
- `text` (string) — Extractable page text.
- `textItems` (object[]) — Positioned text item models.
- `regions` (object[]) — Inspectable page-space regions.
- `tables` (object[]) — Modeled page tables.
- `images` (object[]) — Modeled page images.
- `charts` (object[]) — Modeled page charts.

**Schema returns:**

- `page` (PdfPage) — Appended editable page facade.

#### `pdf.addTable`

Add a modeled table with cell values and a page-space bounding box to the first PDF page.

**Examples:**

- pdf.addTable({ name: 'gates', values: [['Gate', 'Status'], ['PDF.js', 'pass']], bbox: [72, 140, 468, 80] })

**Schema parameters:**

- `name` (string) — Inspectable table name.
- `values` (unknown[][]) required — Rectangular or ragged cell value matrix.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `source` (string) — Optional extraction/source provenance.

**Schema returns:**

- `table` (PdfTable) — Inspectable table facade with stable ID.

#### `pdf.addText`

Add positioned PDF text with page-space bbox, font metadata, inspect/resolve/layout records, and SVG preview rendering.

**Examples:**

- pdf.addText({ pageIndex: 0, text: 'Status', bbox: [72, 72, 200, 24], fontSize: 18, bold: true })

**Schema parameters:**

- `text` (string) required — Text content.
- `pageIndex` (number) — Zero-based target page index.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `fontName` (string) — Font family metadata.
- `fontSize` (number) — Font size in points.
- `color` (string) — Text color.
- `bold` (boolean) — Bold text flag.
- `italic` (boolean) — Italic text flag.

**Schema returns:**

- `textItem` (object) — Positioned text item with stable ID.

#### `pdf.extractTables`

Extract modeled table values and bounding boxes across all pages or a selected page.

**Examples:**

- pdf.extractTables({ page: 1 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `tables` (object[]) — Table records with page, ID, name, values, and bbox.

#### `pdf.extractText`

Extract modeled text across all pages or a selected page.

**Examples:**

- pdf.extractText({ page: 2 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `text` (string) — Selected page text or all page text joined with blank lines.

#### `pdf.inspect`

Emit bounded NDJSON for pages, text, positioned text items, layout regions, tables, images, and charts; narrow with search/target anchors and shape fields with include/exclude.

**Schema parameters:**

- `kind` (string) — Comma-separated page, text, textItem, region, table, image, and chart record kinds.
- `search` (string) — Case-insensitive record filter.
- `target` (string) — Stable ID/anchor target; targetId, id, and anchor are aliases.
- `before` (number) — Records of context before target matches.
- `after` (number) — Records of context after target matches.
- `include` (string) — Comma-separated fields to keep.
- `exclude` (string) — Comma-separated fields to omit.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Bounded { ndjson, truncated } inspection result.

#### `pdf.layoutJson`

Return modeled PDF page layout JSON with page text, positioned text items, layout regions, tables, images, charts, and target/search context slicing.

**Examples:**

- pdf.layoutJson({ page: 1, target: table.id, context: 1 })

**Schema parameters:**

- `page` (number) — Optional one-based page selector.
- `pageIndex` (number) — Optional zero-based page selector.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive layout-record filter.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.

**Schema returns:**

- `layout` (object) — Point-based PDF page layout tree and optional slice metadata.

#### `pdf.render`

Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters.

**Examples:**

- await pdf.render({ pageIndex: 0 })
- await pdf.render({ source: 'pdf', format: 'png', renderer: createPopplerRenderer() })

**Schema parameters:**

- `pageIndex` (number) — Zero-based page index for modeled SVG rendering.
- `page` (number) — One-based page selector used by layout/native renderer workflows.
- `format` (string) — svg by default, layout, pdf, png, ppm, or tiff with a renderer.
- `source` (string) — Set to pdf to render exported PDF bytes.
- `renderer` (function) — Optional PDF-capable renderer adapter.

**Schema returns:**

- `blob` (FileBlob) — SVG, layout JSON, PDF, or renderer output.

#### `pdf.resolve`

Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, layout regions, tables, images, and charts.

**Examples:**

- pdf.resolve('pg-1/txt/1')

**Schema parameters:**

- `id` (string) required — Stable artifact, page, text, text-item, region, table, image, or chart ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `pdf.verify`

Return QA issues for empty pages, Unicode dashes, text extraction sanity, page geometry, text/region/table/image/chart bounds, invalid image data URLs, malformed tables, and chart data.

**Examples:**

- pdf.verify({ maxChars: 12000 })

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — PDF semantic QA result with ok, issues, ndjson, and truncated.

#### `PdfArtifact.create`

Create a modeled PDF artifact with pages, text, table regions, and image regions.

**Examples:**

- const pdf = PdfArtifact.create({ pages: [{ width: 612, height: 792, text: 'Report' }] })

**Schema parameters:**

- `id` (string) — Optional stable artifact ID.
- `metadata` (object) — Clean-room metadata preserved through generated-PDF roundtrip.
- `text` (string) — Convenience text for a single default page.
- `pages` (object[]) — Page models with width, height, text, textItems, regions, tables, images, and charts.

**Schema returns:**

- `pdf` (PdfArtifact) — Editable modeled PDF artifact.

#### `PdfFile.exportPdf`

Export a modeled artifact as a real multi-page PDF with positioned text, vector tables/charts, embedded PNG images, and clean-room metadata.

**Examples:**

- const blob = await PdfFile.exportPdf(pdf)

**Schema parameters:**

- `pdf` (PdfArtifact) required — Modeled PDF artifact to serialize.

**Schema returns:**

- `blob` (FileBlob) — application/pdf bytes with modeled content and clean-room metadata.

#### `PdfFile.importPdf`

Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, normalize parser image bytes/base64 into data URLs, reconstruct tables from positioned text geometry when explicit tables are absent, or fall back to heuristic visible-text/table extraction.

**Examples:**

- await PdfFile.importPdf(blob, { parser: createPdfjsParser() })

**Schema parameters:**

- `blob` (FileBlob|Uint8Array) required — PDF input bytes.
- `parser` (function) — Optional parser adapter returning pages/textItems/tables/images.
- `preferParser` (boolean) — Use parser even if clean-room metadata is embedded.
- `parserName` (string) — Name recorded in artifact metadata.

**Schema returns:**

- `pdf` (PdfArtifact) — Modeled PDF artifact with inspect/resolve/render/verify APIs.

#### `PdfFile.inspectPdf`

Inspect PDF bytes as bounded file/object records including version, byte size, page/object counts, embedded clean-room model presence, and EOF integrity.

**Examples:**

- await PdfFile.inspectPdf(pdf, { maxObjects: 200, maxChars: 12000 })

**Schema parameters:**

- `pdf` (FileBlob|Uint8Array) required — PDF file bytes.
- `maxObjects` (number) — Maximum indirect object records to inspect.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — PDF file summary plus bounded indirect object records.

## presentation

| Name | Kind | Summary |
| --- | --- | --- |
| `compose.column` | api | Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels. |
| `compose.paragraph` | api | Create an editable text block with name, className/style text tokens, and stable inspect output. |
| `Presentation.create` | api | Create a deck with a default or explicit slide size. |
| `presentation.export` | api | Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON. |
| `presentation.inspect` | api | Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, comments, and layout; narrow with search/target anchors and shape fields with include/exclude. |
| `presentation.layouts.add` | api | Create a reusable slide layout with placeholders; export writes slideLayout and slideMaster parts for clean-room PPTX roundtrip. |
| `presentation.resolve` | api | Map stable inspect anchor IDs back to editable facade objects. |
| `presentation.textRange` | api | Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames. |
| `presentation.theme` | api | Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. |
| `presentation.verify` | api | Return presentation QA issues for layout validation, placeholder/template fidelity, chart/data consistency, table shape, image data, and dangling comments. |
| `PresentationFile.inspectPptx` | api | Inspect a PPTX zip package as bounded NDJSON part records with paths, sizes, content types, and optional XML/relationship previews. |
| `slide.addNotes` | api | Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' }). |
| `slide.applyLayout` | api | Apply a slide layout to materialize editable placeholder shapes and preserve layout identity for inspect, verify, and PPTX export. |
| `slide.autoLayout` | api | Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options. |
| `slide.charts.add` | api | Add an inspectable bar/line/pie chart facade with chartType, title, categories, series colors, axes, legend, data labels, layout JSON, SVG preview, and PPTX chart output. |
| `slide.comments.addThread` | api | Attach threaded comments to slide elements; exported as PPTX comments parts and verified for dangling targets. |
| `slide.compose` | api | Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects. |
| `slide.connectors.add` | api | Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA. |
| `slide.images.add` | api | Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output. |
| `slide.shapes.add` | api | Add a shape/textbox with geometry, position, fill, line, and text. |
| `slide.tables.add` | api | Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output. |

### presentation details

#### `presentation.inspect`

Emit NDJSON for deck, slides, textboxes, shapes, tables, charts, images, notes, comments, and layout; narrow with search/target anchors and shape fields with include/exclude.

**Examples:**

- presentation.inspect({ kind: 'image,comment', target: image.id, include: 'alt,bbox' })

**Options:**

- kind
- search
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- maxChars

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `PresentationFile.inspectPptx`

Inspect a PPTX zip package as bounded NDJSON part records with paths, sizes, content types, and optional XML/relationship previews.

**Examples:**

- await PresentationFile.inspectPptx(pptx, { includeText: true, maxChars: 12000 })

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `includeText` (boolean) — Include bounded XML, relationship, and JSON text previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual package part.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — PPTX package and part records with paths, sizes, content types, and optional previews.

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
| `visualQaArtifact` | api | Render an artifact, record deterministic render metadata/hash, validate empty or malformed render output, optionally compare against a baseline render, and compute PNG/PPM pixel-diff metrics when requested. |

### shared details

#### `createCanvasRenderer`

Create an optional node-canvas renderer adapter from open-office-artifact-tool/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG.

**Examples:**

- const renderer = createCanvasRenderer({ width: 1200, height: 800, background: 'white' })

**Schema parameters:**

- `canvas` (object) — Injected node-canvas compatible module.
- `width` (number) — Output width override.
- `height` (number) — Output height override.
- `background` (string) — Canvas background color.
- `outputOptions` (object) — node-canvas encoder options.

**Schema returns:**

- `renderer` (function) — SVG/PNG/JPEG/WebP to PNG/JPEG renderer adapter.

#### `createLibreOfficeRenderer`

Create a LibreOffice CLI renderer adapter from open-office-artifact-tool/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF.

**Examples:**

- const renderer = createLibreOfficeRenderer({ command: 'soffice', timeoutMs: 60000 })

**Schema parameters:**

- `command` (string) — soffice/LibreOffice executable path or command name.
- `format` (string) — Default target format, normally pdf.
- `convertTo` (string) — Explicit LibreOffice --convert-to filter value.
- `timeoutMs` (number) — CLI timeout.
- `tempRoot` (string) — Temporary directory root.
- `argsBuilder` (function) — Custom LibreOffice argument builder.
- `keepTemp` (boolean) — Keep temporary files for diagnostics.

**Schema returns:**

- `renderer` (function) — Office/HTML conversion renderer adapter.

#### `createNativeOfficeRenderer`

Create a native Office renderer adapter from open-office-artifact-tool/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors.

**Examples:**

- const renderer = createNativeOfficeRenderer({ command: 'dotnet', args: ['OfficeBridge.dll'], timeoutMs: 60000 })

**Schema parameters:**

- `command` (string) — Native Office bridge executable.
- `args` (string[]) — Arguments passed before the bridge reads its JSON request from stdin.
- `timeoutMs` (number) — Bridge request timeout.
- `format` (string) — Default requested output format.
- `inputType` (string) — Default input MIME type.
- `outputType` (string) — Default output MIME type.
- `nativeOptions` (object) — Operation-specific native Office options.

**Schema returns:**

- `renderer` (function) — DOCX/XLSX/PPTX/PDF native Office renderer adapter.

#### `createPlaywrightRenderer`

Create an optional Playwright renderer adapter from open-office-artifact-tool/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default.

**Examples:**

- const renderer = createPlaywrightRenderer({ viewport: { width: 900, height: 1200 }, deviceScaleFactor: 1 })

**Options:**

- viewport
- deviceScaleFactor
- allowNetwork
- timeoutMs
- format

**Schema parameters:**

- `viewport` (object) — Chromium viewport width and height; SVG geometry is inferred when omitted.
- `deviceScaleFactor` (number) — Chromium device scale factor.
- `allowNetwork` (boolean) — Permit network requests; disabled by default for deterministic rendering.
- `timeoutMs` (number) — Navigation and rendering timeout.
- `background` (string) — Page background CSS color.
- `chromium` (object) — Injected Playwright Chromium launcher for tests or custom runtimes.

**Schema returns:**

- `renderer` (function) — SVG/HTML to PNG/WebP/JPEG/PDF renderer adapter.

**Returns:**

renderer adapter function for renderArtifact(...)

#### `createPopplerRenderer`

Create a Poppler CLI renderer adapter from open-office-artifact-tool/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF.

**Examples:**

- const renderer = createPopplerRenderer({ command: 'pdftoppm', dpi: 150 })

**Schema parameters:**

- `command` (string) — pdftoppm executable path or command name.
- `dpi` (number) — Raster resolution.
- `page` (number) — One-based PDF page number; pageIndex is the zero-based alias.
- `timeoutMs` (number) — CLI timeout.
- `tempRoot` (string) — Temporary directory root.
- `argsBuilder` (function) — Custom pdftoppm argument builder.
- `keepTemp` (boolean) — Keep temporary input/output files for diagnostics.

**Schema returns:**

- `renderer` (function) — PDF to PNG/PPM/TIFF page renderer adapter.

#### `createSharpRenderer`

Create an optional sharp renderer adapter from open-office-artifact-tool/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG.

**Examples:**

- const renderer = createSharpRenderer({ resize: { width: 1200 }, flatten: true })

**Schema parameters:**

- `sharp` (function) — Injected sharp factory; otherwise the optional peer dependency is loaded.
- `resize` (object) — sharp resize options.
- `flatten` (boolean|object) — Flatten transparency using background options.
- `background` (string|object) — Flatten background color.
- `pngOptions` (object) — sharp PNG encoder options.
- `webpOptions` (object) — sharp WebP encoder options.
- `jpegOptions` (object) — sharp JPEG encoder options.

**Schema returns:**

- `renderer` (function) — SVG/PNG/JPEG/WebP raster renderer adapter.

#### `renderArtifact`

Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output.

**Examples:**

- await renderArtifact(document, { format: 'png', renderer: createPlaywrightRenderer() })

**Options:**

- format
- renderer/rasterRenderer/renderAdapter
- page/pageIndex
- slide
- sheetName
- range

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact facade to render through its native preview/export path.
- `format` (string) — svg, png, webp, jpeg, pdf, layout, or an output MIME type.
- `renderer` (function) — Optional pluggable renderer adapter for raster/PDF conversion.
- `source` (string) — Optional native source such as docx or pdf for renderer gates.

**Schema returns:**

- `blob` (FileBlob) — Rendered output with normalized metadata.

**Returns:**

FileBlob with normalized render metadata

#### `renderFileWithNativeOffice`

Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output.

**Examples:**

- await renderFileWithNativeOffice(docx, { command, format: 'pdf', artifactKind: 'document' })

**Schema parameters:**

- `input` (FileBlob|Uint8Array) required — Office/PDF input bytes.
- `command` (string) required — Native Office bridge executable.
- `args` (string[]) — Arguments passed to the bridge executable.
- `operation` (string) — Bridge operation, defaulting to render.
- `format` (string) — Requested output format.
- `artifactKind` (string) — document, workbook, presentation, or pdf.
- `timeoutMs` (number) — Bridge request timeout.
- `nativeOptions` (object) — Operation-specific native Office options.
- `keepTemp` (boolean) — Keep temporary files for diagnostics.

**Schema returns:**

- `blob` (FileBlob) — Native Office bridge output bytes and renderer metadata.

#### `verifyArtifact`

Run an artifact's verify() method and return a bounded NDJSON QA report.

**Examples:**

- verifyArtifact(workbook, { maxChars: 12000 })

**Options:**

- maxChars

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact exposing a verify() method.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `report` (object) — Semantic QA result with artifactKind, ok, issues, ndjson, and truncated.

**Returns:**

{ artifactKind, ok, issues, ndjson, truncated }

#### `visualQaArtifact`

Render an artifact, record deterministic render metadata/hash, validate empty or malformed render output, optionally compare against a baseline render, and compute PNG/PPM pixel-diff metrics when requested.

**Examples:**

- await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })

**Options:**

- baseline/expected/baselineBlob
- pixelDiff
- PNG/PPM raster pixel comparison
- allowChange
- minBytes
- maxBytes
- maxChars

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact to render and compare.
- `format` (string) — Requested render format such as svg, png, ppm, jpeg, webp, or pdf.
- `renderer` (function) — Optional renderer adapter used for format conversion.
- `baseline` (FileBlob|Uint8Array) — Expected render bytes; expected and baselineBlob are aliases.
- `pixelDiff` (boolean|object) — Enable PNG/PPM pixel comparison and optional thresholds.
- `allowChange` (boolean) — Allow baseline byte/pixel changes without emitting issues.
- `minBytes` (number) — Warn when the render is smaller than this byte count.
- `maxBytes` (number) — Warn when the render exceeds this byte count.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `report` (object) — Visual QA result with ok, blob, summary, issues, ndjson, and truncation metadata.

**Returns:**

{ ok, blob, summary, issues, ndjson }

## workbook

| Name | Kind | Summary |
| --- | --- | --- |
| `fx.ABS` | formula | Return the absolute value of a number. |
| `fx.AND` | formula | Return TRUE when all conditions are true. |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.AVERAGEIF` | formula | Average values whose corresponding criteria range entries match a criterion. |
| `fx.AVERAGEIFS` | formula | Average values where all supplied criteria ranges match their criteria. |
| `fx.CEILING` | formula | Round a number up to the nearest significance. |
| `fx.CONCAT` | formula | Concatenate text values and ranges. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.COUNTIF` | formula | Count values in a range that match a criterion. |
| `fx.COUNTIFS` | formula | Count rows where multiple criteria ranges all match their criteria. |
| `fx.FILTER` | formula | Filter rows from a source range with a boolean or comparison include array and spill the matching rows. |
| `fx.FLOOR` | formula | Round a number down to the nearest significance. |
| `fx.HLOOKUP` | formula | Look up a value in the first row of a table range and return a value from another row. |
| `fx.IF` | formula | Return one value when a condition is true and another when false. |
| `fx.IFERROR` | formula | Return a fallback value when an expression evaluates to a formula error. |
| `fx.INDEX` | formula | Return a value from a range by 1-based row and optional column index. |
| `fx.INT` | formula | Round a number down to the nearest integer. |
| `fx.ISBLANK` | formula | Return TRUE when a referenced value is empty. |
| `fx.ISERROR` | formula | Return TRUE when a value is any recognized formula error. |
| `fx.ISNUMBER` | formula | Return TRUE when a value is numeric. |
| `fx.ISTEXT` | formula | Return TRUE when a value is text and not a formula error. |
| `fx.LEFT` | formula | Return characters from the start of a text value. |
| `fx.LEN` | formula | Return the length of a text value. |
| `fx.LOWER` | formula | Convert text to lowercase. |
| `fx.MATCH` | formula | Return the 1-based position of a lookup value in a range, with exact match and basic ascending/descending approximate modes. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MID` | formula | Return characters from the middle of a text value. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.OR` | formula | Return TRUE when any condition is true. |
| `fx.PMT` | formula | Calculate a loan payment for constant payments and constant interest rate. |
| `fx.RIGHT` | formula | Return characters from the end of a text value. |
| `fx.ROUND` | formula | Round a numeric value to a fixed number of decimal places. |
| `fx.SEQUENCE` | formula | Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine. |
| `fx.SORT` | formula | Sort a range by a 1-based column index and spill the sorted rows. |
| `fx.SUM` | formula | Sum numeric values across arguments and ranges. |
| `fx.SUMIF` | formula | Sum values whose corresponding criteria range entries match a criterion. |
| `fx.SUMIFS` | formula | Sum values where all supplied criteria ranges match their criteria. |
| `fx.SUMPRODUCT` | formula | Multiply corresponding numeric values in equally sized arrays and return the sum of those products. |
| `fx.TEXTJOIN` | formula | Join text values with a delimiter and optional empty-value skipping. |
| `fx.TRANSPOSE` | formula | Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata. |
| `fx.TRIM` | formula | Trim leading/trailing whitespace and collapse internal whitespace. |
| `fx.UNIQUE` | formula | Return unique rows from a range as a spilled dynamic array. |
| `fx.UPPER` | formula | Convert text to uppercase. |
| `fx.VLOOKUP` | formula | Look up a value in the first column of a table range and return a value from another column. |
| `fx.XLOOKUP` | formula | Look up a value in one range and return the corresponding value from another range. |
| `range.conditionalFormats.add` | api | Add a conditional formatting rule; cellIs/expression/containsText/colorScale rules are evaluated into computedStyle inspect records, layout JSON hints, and SVG preview fills. |
| `range.dataValidation` | api | Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule }). |
| `range.format` | api | Assign basic cell style metadata such as fill, font, numberFormat, alignment, and borders; XLSX export writes native styles.xml and cell style indexes. |
| `sheet.charts.add` | api | Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas. |
| `sheet.images.add` | api | Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents. |
| `sheet.pivotTables.add` | api | Create a clean-room pivot table facade over a source range with row/value fields, computed summary values, inspect/resolve/layout records, verification, and metadata roundtrip. |
| `sheet.sparklineGroups.add` | api | Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand. |
| `sheet.tables.add` | api | Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles. |
| `SpreadsheetFile.exportXlsx` | api | Serialize a Workbook facade to an XLSX FileBlob. |
| `SpreadsheetFile.importXlsx` | api | Load an XLSX file into a Workbook facade. |
| `workbook.comments.addThread` | api | Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...'). |
| `Workbook.create` | api | Create an empty workbook; add worksheets before editing. |
| `workbook.definedNames.add` | api | Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData). |
| `workbook.formulaGraph` | api | Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude. |
| `workbook.layoutJson` | api | Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }. |
| `workbook.sharedArrayFormulas` | formula | Import and export native XLSX shared formulas (t=shared) by translating relative A1 references and surface native array formulas (t=array) with formulaType/sharedRef/arrayRef inspect metadata. |
| `workbook.structuredReferences` | formula | Evaluate Excel-style table structured references such as TableName[Column], TableName[#Headers], TableName[[#Data],[Column]], and TableName[[#Data],[First]:[Last]] in formulas, expanding them to stable table cell precedents. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged. |
| `workbook.verify` | api | Return bounded QA issues for sheets, formulas, tables, charts, and comments. |
| `worksheet.getRange` | api | Select an A1 range for values, formulas, formatting, merge, fill, and copy operations. |

### workbook details

#### `fx.ABS`

Return the absolute value of a number.

**Examples:**

- =ABS(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ABS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AND`

Return TRUE when all conditions are true.

**Examples:**

- =AND(A1>0,B1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGE`

Average numeric values across arguments and ranges in the clean-room formula engine.

**Examples:**

- =AVERAGE(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGEIF`

Average values whose corresponding criteria range entries match a criterion.

**Examples:**

- =AVERAGEIF(A1:A10,"East",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGEIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGEIFS`

Average values where all supplied criteria ranges match their criteria.

**Examples:**

- =AVERAGEIFS(C1:C10,A1:A10,"East",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGEIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CEILING`

Round a number up to the nearest significance.

**Examples:**

- =CEILING(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CEILING(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CONCAT`

Concatenate text values and ranges.

**Examples:**

- =CONCAT(A1,"-",B1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CONCAT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNT`

Count numeric values across arguments and ranges.

**Examples:**

- =COUNT(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIF`

Count values in a range that match a criterion.

**Examples:**

- =COUNTIF(A1:A10,">0")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIFS`

Count rows where multiple criteria ranges all match their criteria.

**Examples:**

- =COUNTIFS(A1:A10,"East",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FILTER`

Filter rows from a source range with a boolean or comparison include array and spill the matching rows.

**Examples:**

- =FILTER(A2:C10,B2:B10="East")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FILTER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.FLOOR`

Round a number down to the nearest significance.

**Examples:**

- =FLOOR(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FLOOR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.HLOOKUP`

Look up a value in the first row of a table range and return a value from another row.

**Examples:**

- =HLOOKUP("Revenue",A1:D4,3,FALSE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.IF`

Return one value when a condition is true and another when false.

**Examples:**

- =IF(A1>0,"ok","bad")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.IFERROR`

Return a fallback value when an expression evaluates to a formula error.

**Examples:**

- =IFERROR(XLOOKUP("missing",A1:A10,B1:B10),"not found")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFERROR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.INDEX`

Return a value from a range by 1-based row and optional column index.

**Examples:**

- =INDEX(A2:C4,2,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =INDEX(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.INT`

Round a number down to the nearest integer.

**Examples:**

- =INT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =INT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISBLANK`

Return TRUE when a referenced value is empty.

**Examples:**

- =ISBLANK(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISBLANK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISERROR`

Return TRUE when a value is any recognized formula error.

**Examples:**

- =ISERROR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISERROR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISNUMBER`

Return TRUE when a value is numeric.

**Examples:**

- =ISNUMBER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISNUMBER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.ISTEXT`

Return TRUE when a value is text and not a formula error.

**Examples:**

- =ISTEXT(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISTEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.LEFT`

Return characters from the start of a text value.

**Examples:**

- =LEFT(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LEFT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.LEN`

Return the length of a text value.

**Examples:**

- =LEN(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LEN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.LOWER`

Convert text to lowercase.

**Examples:**

- =LOWER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LOWER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.MATCH`

Return the 1-based position of a lookup value in a range, with exact match and basic ascending/descending approximate modes.

**Examples:**

- =MATCH("Beta",A2:A4,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MATCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MAX`

Return the maximum numeric value across arguments and ranges.

**Examples:**

- =MAX(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MAX(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MID`

Return characters from the middle of a text value.

**Examples:**

- =MID(A1,2,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MID(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.MIN`

Return the minimum numeric value across arguments and ranges.

**Examples:**

- =MIN(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.OR`

Return TRUE when any condition is true.

**Examples:**

- =OR(A1>0,B1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =OR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.PMT`

Calculate a loan payment for constant payments and constant interest rate.

**Examples:**

- =PMT(rate,nper,pv)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Catalog entry only in MVP; full financial formula evaluation is roadmap.

#### `fx.RIGHT`

Return characters from the end of a text value.

**Examples:**

- =RIGHT(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RIGHT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUND`

Round a numeric value to a fixed number of decimal places.

**Examples:**

- =ROUND(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SEQUENCE`

Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine.

**Examples:**

- =SEQUENCE(2,3,10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SEQUENCE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.SORT`

Sort a range by a 1-based column index and spill the sorted rows.

**Examples:**

- =SORT(A2:C10,3,-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SORT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.SUM`

Sum numeric values across arguments and ranges.

**Examples:**

- =SUM(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMIF`

Sum values whose corresponding criteria range entries match a criterion.

**Examples:**

- =SUMIF(A1:A10,"East",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMIFS`

Sum values where all supplied criteria ranges match their criteria.

**Examples:**

- =SUMIFS(C1:C10,A1:A10,"East",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMPRODUCT`

Multiply corresponding numeric values in equally sized arrays and return the sum of those products.

**Examples:**

- =SUMPRODUCT(A1:A10,B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMPRODUCT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTJOIN`

Join text values with a delimiter and optional empty-value skipping.

**Examples:**

- =TEXTJOIN("/",TRUE,A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTJOIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TRANSPOSE`

Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata.

**Examples:**

- =TRANSPOSE(A1:C2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRANSPOSE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.TRIM`

Trim leading/trailing whitespace and collapse internal whitespace.

**Examples:**

- =TRIM(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TRIM(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.UNIQUE`

Return unique rows from a range as a spilled dynamic array.

**Examples:**

- =UNIQUE(A2:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UNIQUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.UPPER`

Convert text to uppercase.

**Examples:**

- =UPPER(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =UPPER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.VLOOKUP`

Look up a value in the first column of a table range and return a value from another column.

**Examples:**

- =VLOOKUP("Beta",A2:B4,2,FALSE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.XLOOKUP`

Look up a value in one range and return the corresponding value from another range.

**Examples:**

- =XLOOKUP("Gamma",A2:A4,B2:B4,"missing")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `range.conditionalFormats.add`

Add a conditional formatting rule; cellIs/expression/containsText/colorScale rules are evaluated into computedStyle inspect records, layout JSON hints, and SVG preview fills.

**Examples:**

- range.conditionalFormats.add('cellIs', { operator: 'greaterThan', formula: 10, format: { fill: 'green' } })
- range.conditionalFormats.addColorScale({ colors: ['#fee2e2', '#fef3c7', '#22c55e'] })

**Schema parameters:**

- `ruleType` (string) required — cellIs, expression, containsText, or colorScale.
- `formula` (string|number) — Rule formula or scalar threshold.
- `operator` (string) — Comparison operator for cellIs rules.
- `format` (object) — Style patch applied when the rule matches.
- `colors` (string[]) — Two or three colors for colorScale rules.

**Schema returns:**

- `conditionalFormat` (object) — Inspectable conditional-format rule with stable id.

#### `range.format`

Assign basic cell style metadata such as fill, font, numberFormat, alignment, and borders; XLSX export writes native styles.xml and cell style indexes.

**Examples:**

- sheet.getRange('A1:D1').format = { fill: '#0f172a', font: { bold: true }, alignment: { horizontal: 'center' }, border: { style: 'thin' } }

**Schema parameters:**

- `fill` (string) — Cell background color token or hex color.
- `font` (object) — Font properties: bold, italic, color, size, name.
- `numberFormat` (string) — Excel number format code.
- `alignment` (object) — horizontal, vertical, and wrapText alignment options.
- `border` (object) — Basic border style and color.

**Schema returns:**

- `range` (Range) — The formatted range facade.

#### `workbook.definedNames.add`

Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData).

**Examples:**

- workbook.definedNames.add('RevenueData', 'Sheet1!G2:G4')
- sheet.getRange('E3').formulas = [['=SUM(RevenueData)']]

**Options:**

- name
- refersTo
- scope/sheetName
- comment

**Returns:**

DefinedName facade with id/name/refersTo/scope

#### `workbook.inspect`

Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude.

**Examples:**

- workbook.inspect({ kind: 'formula', target: 'Sheet1!E2', include: 'formula,value,precedents' })

**Options:**

- kind
- search/searchTerm
- target/targetId/id/anchor
- before/after/context
- include/fields
- exclude/omit
- maxChars

**Schema parameters:**

- `kind` (string) — Comma-separated record kinds such as formula, table, style, computedStyle, chart, image.
- `target` (string) — Stable ID, anchor, or A1 cell/range to slice results around.
- `search` (string) — Case-insensitive text filter over inspect records.
- `include` (string) — Comma-separated top-level fields to keep.
- `exclude` (string) — Comma-separated top-level fields to omit.
- `maxChars` (number) — Maximum NDJSON output size before truncation notice.

**Schema returns:**

- `ndjson` (string) — Bounded newline-delimited JSON records.
- `truncated` (boolean) — True when maxChars truncated the output.

**Returns:**

{ ndjson, truncated } bounded NDJSON records

#### `workbook.structuredReferences`

Evaluate Excel-style table structured references such as TableName[Column], TableName[#Headers], TableName[[#Data],[Column]], and TableName[[#Data],[First]:[Last]] in formulas, expanding them to stable table cell precedents.

**Examples:**

- =SUM(TasksTable[Revenue])
- =TEXTJOIN("|",TRUE,TasksTable[#Headers])
- =SUM(TasksTable[[#Data],[Revenue]])
- =SUM(TasksTable[[#Data],[Revenue]:[Cost]])
- =TEXTJOIN("|",TRUE,TasksTable[[#Data],[Region],[Code]])

**Notes:**

- Current clean-room subset supports #Headers/#Data/#All/#Totals sections, single-column selectors, contiguous column ranges, and comma-separated column unions; special escaping for headers containing brackets remains roadmap.

