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
| `DocumentFile.inspectDocx` | api | Inspect a DOCX package as bounded part records with safe paths, content types, optional previews, decompression budgets, and relationship/content-type consistency issues. |
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

Inspect a DOCX package as bounded part records with safe paths, content types, optional previews, decompression budgets, and relationship/content-type consistency issues.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `includeText` (boolean) — Include bounded XML/JSON/relationship previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual part.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — DOCX package result with ok, issues, parts, records, and bounded NDJSON.

#### `DocumentFile.patchDocx`

Apply safe in-package DOCX XML/JSON/binary patches with path traversal validation and return a patched DOCX FileBlob.

**Examples:**

- await DocumentFile.patchDocx(docx, [{ path: 'customXml/review-note.xml', text: '<review>ok</review>' }])

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `patches` (array|object) required — Path-validated package part edits with text/xml/json/bytes/remove.
- `maxPatchBytes` (number) — Per-part patch size limit.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; relationships accepts an array.

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
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter to extract page geometry, positioned text, heuristic tables, and bounded embedded raster or stencil-mask PNG images with placement boxes. |
| `pdf.addChart` | api | Add a modeled bar/line chart region with categories, series, title, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip. |
| `pdf.addFlowText` | api | Wrap long text into positioned lines and automatically append pages when the configured content box is full. |
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
| `PdfFile.exportPdf` | api | Export a modeled artifact as a real multi-page PDF with positioned text, vector tables/charts, embedded PNG/JPEG images, and clean-room metadata. |
| `PdfFile.importPdf` | api | Import clean-room generated PDFs from metadata, use an injected parser adapter for arbitrary PDFs, normalize parser image bytes/base64 into data URLs, reconstruct tables from positioned text geometry when explicit tables are absent, or fall back to heuristic visible-text/table extraction. |
| `PdfFile.inspectPdf` | api | Inspect PDF bytes as bounded file/object records including version, byte size, page/object counts, embedded clean-room model presence, and EOF integrity. |

### pdf details

#### `createPdfjsParser`

Create an optional PDF.js parser adapter to extract page geometry, positioned text, heuristic tables, and bounded embedded raster or stencil-mask PNG images with placement boxes.

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

#### `pdf.addFlowText`

Wrap long text into positioned lines and automatically append pages when the configured content box is full.

**Examples:**

- pdf.addFlowText(longReport, { fontSize: 11, margins: { top: 72, right: 72, bottom: 72, left: 72 } })

**Schema parameters:**

- `text` (string) required — Paragraph text separated by newlines.
- `pageIndex` (number) — Zero-based starting page index; defaults to the first page.
- `margins` (number|object) — Uniform margin or top/right/bottom/left page margins in points.
- `left` (number) — Explicit content-box left edge overriding margins.left.
- `top` (number) — Explicit first-page top edge overriding margins.top.
- `width` (number) — Explicit content width; defaults to page width minus horizontal margins.
- `fontSize` (number) — Line font size in points.
- `lineHeight` (number) — Line advance in points.
- `paragraphGap` (number) — Extra vertical space after each paragraph.

**Schema returns:**

- `flow` (object) — Flow ID, positioned items, page IDs, page indexes, and line count.

#### `pdf.addImage`

Add a modeled PDF image region with dataUrl/URI/prompt metadata, alt text, and page-space bounding box.

**Examples:**

- pdf.addImage({ pageIndex: 0, dataUrl, alt: 'Approval mark', bbox: [430, 60, 64, 64] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `dataUrl` (string) — Embedded PNG or JPEG image data URL.
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

Export a modeled artifact as a real multi-page PDF with positioned text, vector tables/charts, embedded PNG/JPEG images, and clean-room metadata.

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
| `presentation.slides.add` | api | Append an editable slide with optional name, layout identity, and speaker notes. |
| `presentation.textRange` | api | Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames. |
| `presentation.theme` | api | Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. |
| `presentation.verify` | api | Return presentation QA issues for layout validation, placeholder/template fidelity, chart/data consistency, table shape, image data, and dangling comments. |
| `PresentationFile.exportPptx` | api | Serialize a presentation facade to a native OOXML PPTX FileBlob. |
| `PresentationFile.importPptx` | api | Import PPTX bytes into the clean-room presentation facade, restoring native parts and embedded metadata when available. |
| `PresentationFile.inspectPptx` | api | Inspect a PPTX package as bounded part records with content types, optional previews, decompression budgets, and relationship/content-type consistency issues. |
| `PresentationFile.patchPptx` | api | Apply path-validated XML/JSON/binary PPTX part patches with part-count and byte budgets. |
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

#### `compose.column`

Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels.

**Schema parameters:**

- `children` (object[]) — Ordered child compose nodes.
- `width` (string|number) — fill, hug, or fixed pixel width.
- `height` (string|number) — fill, hug, or fixed pixel height.
- `gap` (number) — Child gap in pixels.
- `padding` (number|object) — Container padding.

**Schema returns:**

- `node` (object) — Vertical compose node.

#### `compose.paragraph`

Create an editable text block with name, className/style text tokens, and stable inspect output.

**Schema parameters:**

- `text` (string) required — Editable paragraph text.
- `name` (string) — Stable element name.
- `className` (string) — Text style token string.
- `style` (object) — Explicit text style metadata.

**Schema returns:**

- `node` (object) — Paragraph compose node.

#### `Presentation.create`

Create a deck with a default or explicit slide size.

**Schema parameters:**

- `slideSize` (object) — Slide width and height in pixels; defaults to 1280x720.
- `theme` (object) — Theme name, colors, and major/minor fonts.
- `layouts` (object[]) — Reusable slide layout definitions.

**Schema returns:**

- `presentation` (Presentation) — Editable presentation facade.

#### `presentation.export`

Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON.

**Schema parameters:**

- `format` (string) — svg by default, montage, or layout.
- `slide` (Slide) — Slide facade to export; defaults to the first slide.
- `columns` (number) — Montage column count.
- `scale` (number) — Montage thumbnail scale.
- `gap` (number) — Montage gap in pixels.

**Schema returns:**

- `blob` (FileBlob) — SVG montage/slide preview or layout JSON.

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

**Schema parameters:**

- `kind` (string) — Comma-separated deck/theme/layout/slide/textbox/textRange/shape/table/chart/image/connector/comment/notes kinds.
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

#### `presentation.layouts.add`

Create a reusable slide layout with placeholders; export writes slideLayout and slideMaster parts for clean-room PPTX roundtrip.

**Schema parameters:**

- `name` (string) required — Layout name.
- `type` (string) — Layout type.
- `masterId` (string) — Master identity.
- `placeholders` (object[]) — Placeholder type/name/frame/text/required/style definitions.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Appended reusable layout facade.

#### `presentation.resolve`

Map stable inspect anchor IDs back to editable facade objects.

**Schema parameters:**

- `id` (string) required — Stable deck, theme, layout, slide, element, comment, or text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `presentation.slides.add`

Append an editable slide with optional name, layout identity, and speaker notes.

**Schema parameters:**

- `name` (string) — Inspectable slide name.
- `layout` (string|object) — Layout ID/name or layout facade.
- `notes` (string) — Initial speaker notes.

**Schema returns:**

- `slide` (Slide) — Appended editable slide.

#### `presentation.textRange`

Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames.

**Schema parameters:**

- `id` (string) required — Stable shape text-range ID ending in /text.

**Schema returns:**

- `textRange` (TextRange|undefined) — Editable slide text-range facade or undefined.

#### `presentation.theme`

Configure inspectable theme colors and major/minor fonts; export writes a real ppt/theme/theme1.xml part.

**Schema parameters:**

- `name` (string) — Theme name.
- `colors` (object) — Theme accent/background/text color map.
- `fonts` (object) — Major and minor font families.

**Schema returns:**

- `theme` (PresentationTheme) — Mutable presentation theme facade.

#### `presentation.validateLayout`

Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow.

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area in square pixels before reporting.
- `boundsPadding` (number) — Allowed padding outside the slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Layout QA result with ok, issues, ndjson, and truncated.

#### `presentation.verify`

Return presentation QA issues for layout validation, placeholder/template fidelity, chart/data consistency, table shape, image data, and dangling comments.

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area for layout QA.
- `boundsPadding` (number) — Allowed padding outside slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Presentation semantic/layout QA result.

#### `PresentationFile.exportPptx`

Serialize a presentation facade to a native OOXML PPTX FileBlob.

**Schema parameters:**

- `presentation` (Presentation) required — Presentation facade to serialize.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML PPTX package bytes.

#### `PresentationFile.importPptx`

Import PPTX bytes into the clean-room presentation facade, restoring native parts and embedded metadata when available.

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.

**Schema returns:**

- `presentation` (Presentation) — Imported editable presentation facade.

#### `PresentationFile.inspectPptx`

Inspect a PPTX package as bounded part records with content types, optional previews, decompression budgets, and relationship/content-type consistency issues.

**Examples:**

- await PresentationFile.inspectPptx(pptx, { includeText: true, maxChars: 12000 })

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `includeText` (boolean) — Include bounded XML, relationship, and JSON text previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual package part.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — PPTX package result with ok, issues, parts, records, and bounded NDJSON.

#### `PresentationFile.patchPptx`

Apply path-validated XML/JSON/binary PPTX part patches with part-count and byte budgets.

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched PPTX FileBlob with patchedParts metadata.

#### `slide.addNotes`

Set speaker notes for a slide; exported as a PPTX notesSlide part and surfaced through inspect({ kind: 'notes' }).

**Schema parameters:**

- `text` (string) required — Speaker notes text.

**Schema returns:**

- `notes` (object) — Mutable speaker-notes record.

#### `slide.applyLayout`

Apply a slide layout to materialize editable placeholder shapes and preserve layout identity for inspect, verify, and PPTX export.

**Schema parameters:**

- `layout` (string|SlideLayoutTemplate) required — Layout name/ID or layout facade.

**Schema returns:**

- `shapes` (Shape[]) — Materialized editable placeholder shapes.

#### `slide.autoLayout`

Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options.

**Schema parameters:**

- `shapes` (object[]) required — Existing editable slide elements.
- `frame` (string|object) — slide, a frame object, or an element facade.
- `direction` (string) — horizontal or vertical.
- `horizontalGap` (number|string) — Horizontal gap or auto.
- `verticalGap` (number|string) — Vertical gap or auto.
- `horizontalPadding` (number) — Left/right inset.
- `verticalPadding` (number) — Top/bottom inset.
- `align` (string) — Cross-axis alignment.

**Schema returns:**

- `shapes` (object[]) — The positioned input elements.

#### `slide.charts.add`

Add an inspectable bar/line/pie chart facade with chartType, title, categories, series colors, axes, legend, data labels, layout JSON, SVG preview, and PPTX chart output.

**Schema parameters:**

- `chartType` (string) — bar, line, or pie.
- `title` (string) — Chart title.
- `categories` (string[]) required — Category labels.
- `series` (object[]) required — Series with names, numeric values, and optional colors.
- `position` (object) — Pixel left/top/width/height frame.
- `axes` (object) — Axis titles/options.
- `legend` (object) — Legend options.
- `dataLabels` (object) — Data-label options.

**Schema returns:**

- `chart` (ChartElement) — Appended editable native-chart facade.

#### `slide.comments.addThread`

Attach threaded comments to slide elements; exported as PPTX comments parts and verified for dangling targets.

**Schema parameters:**

- `target` (string|object) required — Stable element ID or element facade.
- `text` (string) required — Initial comment text.
- `author` (string) — Comment author.
- `resolved` (boolean) — Initial resolution state.

**Schema returns:**

- `thread` (SlideCommentThread) — Attached comment thread.

#### `slide.compose`

Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph, shape, table, chart, image, and rule nodes into editable slide objects.

**Schema parameters:**

- `node` (object) required — Compose tree rooted in row, column, grid, layers, box, paragraph, shape, table, chart, image, or rule.
- `frame` (object) — Pixel materialization frame; defaults to an inset slide frame.

**Schema returns:**

- `elements` (object[]) — Materialized editable slide elements.

#### `slide.connectors.add`

Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA.

**Schema parameters:**

- `from` (string|object) — Start element/ID or point.
- `to` (string|object) — End element/ID or point.
- `start` (object) — Explicit start point {x,y}.
- `end` (object) — Explicit end point {x,y}.
- `connectorType` (string) — Connector geometry, currently straight by default.
- `line` (object) — Line color, width, and arrow metadata.

**Schema returns:**

- `connector` (ConnectorElement) — Appended editable connector.

#### `slide.images.add`

Add an inspectable image facade with alt text, prompt/URI/data URL metadata, fit, frame, layout JSON, SVG preview, and PPTX placeholder output.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `fit` (string) — contain or cover intent.
- `position` (object) — Pixel left/top/width/height frame.

**Schema returns:**

- `image` (ImageElement) — Appended editable image facade.

#### `slide.shapes.add`

Add a shape/textbox with geometry, position, fill, line, and text.

**Schema parameters:**

- `name` (string) — Inspectable shape name.
- `geometry` (string) — Shape geometry such as rect or ellipse.
- `position` (object) — Pixel left/top/width/height frame.
- `text` (string) — Shape text.
- `fill` (string|object) — Shape fill.
- `line` (object) — Line color, width, dash, and arrow metadata.
- `placeholder` (object) — Optional layout placeholder metadata.

**Schema returns:**

- `shape` (Shape) — Appended editable shape/textbox.

#### `slide.tables.add`

Add an inspectable native-style table facade with rows, columns, values, cells, layout JSON, and SVG/PPTX placeholder output.

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `name` (string) — Inspectable table name.
- `position` (object) — Pixel left/top/width/height frame.
- `style` (object) — Table/cell fill, margins, borders, and text style.

**Schema returns:**

- `table` (TableElement) — Appended editable table facade.

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
| `visualQaArtifact` | api | Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, and return a configurable aligned PNG diff heatmap. |

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

Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, and return a configurable aligned PNG diff heatmap.

**Examples:**

- await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })

**Options:**

- baseline/expected/baselineBlob
- pixelDiff
- diffImage
- diffPalette
- diffAlignment
- PNG/JPEG/WebP/PPM raster pixel comparison
- allowChange
- minBytes
- maxBytes
- maxChars

**Schema parameters:**

- `artifact` (Workbook|Presentation|DocumentModel|PdfArtifact) required — Artifact to render and compare.
- `format` (string) — Requested render format such as svg, png, ppm, jpeg, webp, or pdf.
- `renderer` (function) — Optional renderer adapter used for format conversion.
- `baseline` (FileBlob|Uint8Array) — Expected render bytes; expected and baselineBlob are aliases.
- `pixelDiff` (boolean|object) — Enable PNG/JPEG/WebP/PPM pixel comparison, optional channel thresholds, and decoded-pixel limits.
- `diffImage` (boolean) — Set false to disable PNG heatmap generation for changed raster baselines.
- `diffPalette` (object) — Optional changed/unchanged RGB colors and alpha values for the PNG heatmap.
- `diffAlignment` (string) — Dimension-mismatch behavior: strict (no heatmap), top-left, or center alignment on a union canvas.
- `allowChange` (boolean) — Allow baseline byte/pixel changes without emitting issues.
- `minBytes` (number) — Warn when the render is smaller than this byte count.
- `maxBytes` (number) — Warn when the render exceeds this byte count.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `report` (object) — Visual QA result with ok, blob, optional diffBlob PNG heatmap, summary, issues, ndjson, and truncation metadata.

**Returns:**

{ ok, blob, diffBlob, summary, issues, ndjson }

## workbook

| Name | Kind | Summary |
| --- | --- | --- |
| `fx.ABS` | formula | Return the absolute value of a number. |
| `fx.AND` | formula | Return TRUE when all conditions are true. |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.AVERAGEIF` | formula | Average values whose corresponding criteria range entries match a criterion. |
| `fx.AVERAGEIFS` | formula | Average values where all supplied criteria ranges match their criteria. |
| `fx.CEILING` | formula | Round a number up to the nearest significance. |
| `fx.CHOOSECOLS` | formula | Select and reorder one or more 1-based or negative column indexes from an array. |
| `fx.CHOOSEROWS` | formula | Select and reorder one or more 1-based or negative row indexes from an array. |
| `fx.CONCAT` | formula | Concatenate text values and ranges. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.COUNTIF` | formula | Count values in a range that match a criterion. |
| `fx.COUNTIFS` | formula | Count rows where multiple criteria ranges all match their criteria. |
| `fx.DROP` | formula | Drop rows and optional columns from the start or end of an array and spill the remainder. |
| `fx.EXPAND` | formula | Expand an array to requested row and column dimensions with optional padding. |
| `fx.FILTER` | formula | Filter rows from a source range with a boolean or comparison include array and spill the matching rows. |
| `fx.FLOOR` | formula | Round a number down to the nearest significance. |
| `fx.HLOOKUP` | formula | Look up a value in the first row of a table range and return a value from another row. |
| `fx.HSTACK` | formula | Append arrays horizontally, padding shorter arrays with #N/A to the maximum row count. |
| `fx.IF` | formula | Return one value when a condition is true and another when false. |
| `fx.IFERROR` | formula | Return a fallback value when an expression evaluates to a formula error. |
| `fx.IFNA` | formula | Return a fallback only when an expression evaluates to #N/A; preserve every other result or error. |
| `fx.INDEX` | formula | Return a value from a range by 1-based row and optional column index. |
| `fx.INT` | formula | Round a number down to the nearest integer. |
| `fx.ISBLANK` | formula | Return TRUE when a referenced value is empty. |
| `fx.ISERR` | formula | Return TRUE for recognized formula errors other than #N/A. |
| `fx.ISERROR` | formula | Return TRUE when a value is any recognized formula error. |
| `fx.ISNA` | formula | Return TRUE only when a value is the #N/A error. |
| `fx.ISNUMBER` | formula | Return TRUE when a value is numeric. |
| `fx.ISTEXT` | formula | Return TRUE when a value is text and not a formula error. |
| `fx.LEFT` | formula | Return characters from the start of a text value. |
| `fx.LEN` | formula | Return the length of a text value. |
| `fx.LOWER` | formula | Convert text to lowercase. |
| `fx.MATCH` | formula | Return the 1-based position of a lookup value in a range, with exact match and basic ascending/descending approximate modes. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MID` | formula | Return characters from the middle of a text value. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.NA` | formula | Return the #N/A error value to mark unavailable data explicitly. |
| `fx.NOT` | formula | Reverse the truth value of a condition. |
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
| `fx.TAKE` | formula | Take rows and optional columns from the start or end of an array and spill the result. |
| `fx.TEXTJOIN` | formula | Join text values with a delimiter and optional empty-value skipping. |
| `fx.TOCOL` | formula | Flatten an array into one spilled column, optionally ignoring blanks or errors and scanning by column. |
| `fx.TOROW` | formula | Flatten an array into one spilled row, optionally ignoring blanks or errors and scanning by column. |
| `fx.TRANSPOSE` | formula | Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata. |
| `fx.TRIM` | formula | Trim leading/trailing whitespace and collapse internal whitespace. |
| `fx.UNIQUE` | formula | Return unique rows from a range as a spilled dynamic array. |
| `fx.UPPER` | formula | Convert text to uppercase. |
| `fx.VLOOKUP` | formula | Look up a value in the first column of a table range and return a value from another column. |
| `fx.VSTACK` | formula | Append arrays vertically, padding narrower arrays with #N/A to the maximum column count. |
| `fx.WRAPCOLS` | formula | Wrap a one-dimensional vector into columns of a requested height, padding the final column when needed. |
| `fx.WRAPROWS` | formula | Wrap a one-dimensional vector into rows of a requested width, padding the final row when needed. |
| `fx.XLOOKUP` | formula | Look up a value in one range and return the corresponding value from another range. |
| `fx.XMATCH` | formula | Return a 1-based lookup position with exact, next-smaller, next-larger, wildcard, and forward or reverse search modes. |
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
| `SpreadsheetFile.inspectXlsx` | api | Inspect an XLSX package as bounded, content-type-aware part records with decompression budgets and relationship/content-type consistency issues. |
| `SpreadsheetFile.patchXlsx` | api | Apply path-validated XML/JSON/binary XLSX part patches with part-count and byte budgets. |
| `workbook.comments.addThread` | api | Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...'). |
| `Workbook.create` | api | Create an empty workbook; add worksheets before editing. |
| `workbook.definedNames.add` | api | Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData). |
| `workbook.formulaGraph` | api | Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, sheets, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude. |
| `workbook.layoutJson` | api | Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing. |
| `workbook.recalculate` | api | Recalculate workbook formulas, dynamic-array spills, dependency edges, cycles, and errors. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }. |
| `workbook.resolve` | api | Resolve stable workbook, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs. |
| `workbook.sharedArrayFormulas` | formula | Import and export native XLSX shared formulas (t=shared) by translating relative A1 references and surface native array formulas (t=array) with formulaType/sharedRef/arrayRef inspect metadata. |
| `workbook.structuredReferences` | formula | Evaluate Excel-style table structured references such as TableName[Column], TableName[#Headers], TableName[[#Data],[Column]], and TableName[[#Data],[First]:[Last]] in formulas, expanding them to stable table cell precedents. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged. |
| `workbook.verify` | api | Return bounded QA issues for sheets, formulas, tables, charts, and comments. |
| `workbook.worksheets.add` | api | Append an editable worksheet with a stable name and ID. |
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

#### `fx.CHOOSECOLS`

Select and reorder one or more 1-based or negative column indexes from an array.

**Examples:**

- =CHOOSECOLS(A2:C10,3,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CHOOSECOLS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.CHOOSEROWS`

Select and reorder one or more 1-based or negative row indexes from an array.

**Examples:**

- =CHOOSEROWS(A2:C10,3,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CHOOSEROWS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

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

#### `fx.DROP`

Drop rows and optional columns from the start or end of an array and spill the remainder.

**Examples:**

- =DROP(A2:C10,1,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DROP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.EXPAND`

Expand an array to requested row and column dimensions with optional padding.

**Examples:**

- =EXPAND(A2:B3,4,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EXPAND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

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

#### `fx.HSTACK`

Append arrays horizontally, padding shorter arrays with #N/A to the maximum row count.

**Examples:**

- =HSTACK(A2:B4,D2:E3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HSTACK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

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

#### `fx.IFNA`

Return a fallback only when an expression evaluates to #N/A; preserve every other result or error.

**Examples:**

- =IFNA(XLOOKUP("missing",A1:A10,B1:B10),"not found")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFNA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.ISERR`

Return TRUE for recognized formula errors other than #N/A.

**Examples:**

- =ISERR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISERR(...).
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

#### `fx.ISNA`

Return TRUE only when a value is the #N/A error.

**Examples:**

- =ISNA(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ISNA(...).
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

#### `fx.NA`

Return the #N/A error value to mark unavailable data explicitly.

**Examples:**

- =NA()

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.NOT`

Reverse the truth value of a condition.

**Examples:**

- =NOT(A1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NOT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.TAKE`

Take rows and optional columns from the start or end of an array and spill the result.

**Examples:**

- =TAKE(A2:C10,3,-2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TAKE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.TEXTJOIN`

Join text values with a delimiter and optional empty-value skipping.

**Examples:**

- =TEXTJOIN("/",TRUE,A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTJOIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TOCOL`

Flatten an array into one spilled column, optionally ignoring blanks or errors and scanning by column.

**Examples:**

- =TOCOL(A2:C10,1,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TOCOL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.TOROW`

Flatten an array into one spilled row, optionally ignoring blanks or errors and scanning by column.

**Examples:**

- =TOROW(A2:C10,1,TRUE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TOROW(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

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

#### `fx.VSTACK`

Append arrays vertically, padding narrower arrays with #N/A to the maximum column count.

**Examples:**

- =VSTACK(A2:B4,A7:A9)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VSTACK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.WRAPCOLS`

Wrap a one-dimensional vector into columns of a requested height, padding the final column when needed.

**Examples:**

- =WRAPCOLS(A2:A10,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WRAPCOLS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.WRAPROWS`

Wrap a one-dimensional vector into rows of a requested width, padding the final row when needed.

**Examples:**

- =WRAPROWS(A2:A10,3,"n/a")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WRAPROWS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.XLOOKUP`

Look up a value in one range and return the corresponding value from another range.

**Examples:**

- =XLOOKUP("Gamma",A2:A4,B2:B4,"missing")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.XMATCH`

Return a 1-based lookup position with exact, next-smaller, next-larger, wildcard, and forward or reverse search modes.

**Examples:**

- =XMATCH("Beta*",A2:A10,2,-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XMATCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

#### `range.dataValidation`

Assign a validation rule to a range or use sheet.dataValidations.add({ range, rule }).

**Schema parameters:**

- `type` (string) required — Validation type such as list, whole, decimal, date, or custom.
- `values` (unknown[]) — Allowed list values.
- `formula1` (string|number) — Primary validation formula/value.
- `formula2` (string|number) — Secondary formula/value for between rules.
- `operator` (string) — Comparison operator.
- `allowBlank` (boolean) — Allow blank cells.

**Schema returns:**

- `validation` (object) — Inspectable data-validation rule anchored to the range.

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

#### `sheet.charts.add`

Create an inspectable worksheet chart from a range or config; setData(range) infers categories and series formulas.

**Schema parameters:**

- `chartType` (string) required — Chart type such as bar, line, or pie.
- `source` (Range|object) — Source range or explicit chart config.
- `title` (string) — Chart title.
- `categories` (string[]) — Explicit categories.
- `series` (object[]) — Explicit series definitions.
- `position` (object) — Pixel chart frame.

**Schema returns:**

- `chart` (WorksheetChart) — Editable worksheet chart facade.

#### `sheet.images.add`

Create an inspectable worksheet image placeholder from a data URL, URI, or prompt with 0-based cell anchors and pixel extents.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `anchor` (object) — Zero-based cell anchor and pixel extent.
- `fit` (string) — contain or cover intent.

**Schema returns:**

- `image` (WorksheetImage) — Editable worksheet image facade.

#### `sheet.pivotTables.add`

Create a clean-room pivot table facade over a source range with row/value fields, computed summary values, inspect/resolve/layout records, verification, and metadata roundtrip.

**Schema parameters:**

- `name` (string) — Stable pivot name.
- `sourceRange` (string|Range) required — Source data range.
- `targetRange` (string|Range) required — Destination anchor/range.
- `rowFields` (string[]) — Row field names.
- `columnFields` (string[]) — Column field names.
- `valueFields` (object[]) — Value field and aggregation definitions.
- `filters` (object) — Pivot filter metadata.

**Schema returns:**

- `pivot` (WorksheetPivotTable) — Editable clean-room pivot facade.

#### `sheet.sparklineGroups.add`

Create line/column/stacked sparklines from sourceData into a targetRange; range.sparklines.add is a shorthand.

**Schema parameters:**

- `type` (string) — line, column, or stacked.
- `targetRange` (string|Range) required — Destination range.
- `sourceData` (string|Range) required — Source data range.
- `dateAxisRange` (string|Range) — Optional date-axis range.
- `seriesColor` (string) — Series color.
- `markers` (object) — Marker visibility/style metadata.
- `axis` (object) — Axis metadata.

**Schema returns:**

- `sparkline` (SparklineGroup) — Editable sparkline group facade.

#### `sheet.tables.add`

Create an inspectable worksheet table over an A1 range with rows.add, getDataRows, getHeaderRowRange, style, and visibility toggles.

**Schema parameters:**

- `range` (string|Range) required — A1 range or range facade.
- `hasHeaders` (boolean) — Whether the first row contains headers.
- `name` (string) — Stable Excel table name.
- `style` (string) — Table style name.

**Schema returns:**

- `table` (WorksheetTable) — Editable worksheet table facade.

#### `SpreadsheetFile.exportXlsx`

Serialize a Workbook facade to an XLSX FileBlob.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to recalculate and serialize.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML XLSX package bytes.

#### `SpreadsheetFile.importXlsx`

Load an XLSX file into a Workbook facade.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.inspectXlsx`

Inspect an XLSX package as bounded, content-type-aware part records with decompression budgets and relationship/content-type consistency issues.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `includeText` (boolean) — Include bounded XML/JSON/relationship previews.
- `maxPreviewChars` (number) — Maximum preview characters per textual part.
- `maxParts` (number) — Maximum package part count.
- `maxPartBytes` (number) — Maximum uncompressed bytes per part.
- `maxTotalBytes` (number) — Maximum total uncompressed package bytes.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `package` (object) — XLSX package result with ok, issues, parts, records, and bounded NDJSON.

#### `SpreadsheetFile.patchXlsx`

Apply path-validated XML/JSON/binary XLSX part patches with part-count and byte budgets.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched XLSX FileBlob with patchedParts metadata.

#### `workbook.comments.addThread`

Create threaded comments after comments.setSelf({ displayName }); resolve with wb.resolve('th/...').

**Schema parameters:**

- `target` (Range|object) required — Target single-cell range or cell descriptor.
- `text` (string) required — Initial comment text.

**Schema returns:**

- `thread` (CommentThread) — Attached threaded comment using comments.setSelf author identity.

#### `Workbook.create`

Create an empty workbook; add worksheets before editing.

**Schema returns:**

- `workbook` (Workbook) — Empty editable workbook facade.

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

**Schema parameters:**

- `name` (string) required — Defined name.
- `refersTo` (string) required — Sheet-qualified A1 reference.
- `scope` (string) — Optional worksheet scope.
- `comment` (string) — Optional description.

**Schema returns:**

- `definedName` (DefinedName) — Created or updated defined-name facade.

**Returns:**

DefinedName facade with id/name/refersTo/scope

#### `workbook.formulaGraph`

Return a dependency graph of formula nodes, edges, dependents, cycles, and formula errors for workbook QA.

**Schema parameters:**

- `recalculate` (boolean) — Recalculate before reading the graph; defaults to true.
- `maxChars` (number) — Maximum bounded NDJSON graph-record size.

**Schema returns:**

- `graph` (object) — Formula nodes, edges, cycles, errors, and bounded NDJSON.

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

#### `workbook.layoutJson`

Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing.

**Schema parameters:**

- `sheetName` (string) — Optional worksheet selector.
- `range` (string) — Optional A1 layout range.
- `target` (string) — Stable target ID/anchor.
- `search` (string) — Case-insensitive layout-record filter.
- `before` (number) — Context records before matches.
- `after` (number) — Context records after matches.

**Schema returns:**

- `layout` (object) — Workbook/worksheet layout tree with cells and drawing/rule bounds.

#### `workbook.recalculate`

Recalculate workbook formulas, dynamic-array spills, dependency edges, cycles, and errors.

**Schema returns:**

- `graph` (object) — Updated formula dependency graph including cycles and errors.

#### `workbook.render`

Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }.

**Schema parameters:**

- `sheetName` (string) — Worksheet name; defaults to the active worksheet.
- `range` (string) — A1 preview range.
- `format` (string) — svg by default or layout.
- `target` (string) — Stable layout target ID/anchor.
- `search` (string) — Case-insensitive layout filter.

**Schema returns:**

- `blob` (FileBlob) — Worksheet SVG preview or workbook layout JSON.

#### `workbook.resolve`

Resolve stable workbook, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs.

**Schema parameters:**

- `id` (string) required — Stable workbook, sheet, table, pivot, chart, image, sparkline, rule, comment, or defined-name ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `workbook.sharedArrayFormulas`

Import and export native XLSX shared formulas (t=shared) by translating relative A1 references and surface native array formulas (t=array) with formulaType/sharedRef/arrayRef inspect metadata.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) — XLSX bytes containing shared or array formula records.
- `formula` (string) — Shared/array formula expression.
- `ref` (string) — Shared or spill A1 range.

**Schema returns:**

- `metadata` (object) — formulaType/sharedRef/arrayRef/spill inspect metadata.

#### `workbook.structuredReferences`

Evaluate Excel-style table structured references such as TableName[Column], TableName[#Headers], TableName[[#Data],[Column]], and TableName[[#Data],[First]:[Last]] in formulas, expanding them to stable table cell precedents.

**Examples:**

- =SUM(TasksTable[Revenue])
- =TEXTJOIN("|",TRUE,TasksTable[#Headers])
- =SUM(TasksTable[[#Data],[Revenue]])
- =SUM(TasksTable[[#Data],[Revenue]:[Cost]])
- =TEXTJOIN("|",TRUE,TasksTable[[#Data],[Region],[Code]])

**Schema parameters:**

- `formula` (string) required — Formula containing an Excel table structured reference.
- `table` (string) required — Worksheet table name.
- `selector` (string) required — Column/section/range/union selector inside brackets.

**Schema returns:**

- `value` (unknown) — Calculated scalar/array value with stable table-cell precedents.

**Notes:**

- Current clean-room subset supports #Headers/#Data/#All/#Totals sections, single-column selectors, contiguous column ranges, and comma-separated column unions; special escaping for headers containing brackets remains roadmap.

#### `workbook.trace`

Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references flagged.

**Schema parameters:**

- `reference` (string|Range) required — Target A1 reference, optionally sheet-qualified, or range facade.
- `maxDepth` (number) — Maximum precedent recursion depth; defaults to 8.
- `maxChars` (number) — Maximum bounded NDJSON trace size.

**Schema returns:**

- `trace` (object) — Precedent tree plus bounded flat NDJSON trace.

#### `workbook.verify`

Return bounded QA issues for sheets, formulas, tables, charts, and comments.

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Workbook formula/structure/drawing/rule QA result.

#### `workbook.worksheets.add`

Append an editable worksheet with a stable name and ID.

**Schema parameters:**

- `name` (string) — Unique worksheet name; defaults to SheetN.

**Schema returns:**

- `worksheet` (Worksheet) — Appended editable worksheet.

#### `worksheet.getRange`

Select an A1 range for values, formulas, formatting, merge, fill, and copy operations.

**Schema parameters:**

- `address` (string) required — A1 cell or range address such as A1:D10.

**Schema returns:**

- `range` (Range) — Editable range facade for values, formulas, formatting, and rules.

