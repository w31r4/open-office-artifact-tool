# open-office-artifact-tool

Clean-room Office/PDF artifact toolkit inspired by public Office/PDF artifact workflows.

This package is implemented independently using open implementation code:

- `Workbook` / `SpreadsheetFile` for XLSX-style artifacts
- `Presentation` / `PresentationFile` for PPTX-style artifacts
- `DocumentModel` / `DocumentFile` for DOCX-style artifacts
- `PdfArtifact` / `PdfFile` for PDF artifacts
- shared `FileBlob`
- `inspect(...)`, `resolve(...)`, `help(...)`, render/export-style APIs where practical

The long-term Office I/O architecture is documented in [docs/reference-runtime-architecture.md](docs/reference-runtime-architecture.md). The agent-facing model remains JavaScript. **OpenChestnut** is this project's independent, source-built C# Open XML SDK codec and bundled .NET WebAssembly runtime for bounded XLSX, DOCX, and PPTX slices; advanced authoring semantics still use the existing JavaScript codecs as the default migration fallback. The Windows Microsoft Office bridge is optional render/compatibility QA, not the core file codec.

## Current status

Spreadsheet formulas support relationship-driven native table import plus a shared structured-reference parser for evaluation, dependency graphs, and trace output. This includes `#This Row`/`@`, unqualified calculated-column references, space intersections over the common cells of two or more table references, and apostrophe escaping for special-character headers such as `Sales['#Items]` and `Sales[Bracket'[Value']]`; metadata-free XLSX roundtrips retain table column names and exact effective precedents.

The six conditional aggregation functions `COUNTIF(S)`, `SUMIF(S)`, and `AVERAGEIF(S)` share case-insensitive Excel-style comparison and `?`/`*`/`~` wildcard matching. Multi-range `IFS` variants validate equal dimensions and propagate matched aggregation errors instead of silently coercing them.

This is an early MVP. It already creates and imports XLSX/CSV/TSV/PPTX/DOCX/PDF artifacts, supports stable inspect/resolve IDs plus target-sliced and include/exclude-shaped inspect output, and includes tests for all four skill families. The spreadsheet facade includes bounded RFC-style CSV/TSV import/export/inspection, formula traces, workbook defined names, a dependency graph with cycle/missing-sheet reporting, Excel-style table structured references (`TableName[Column]`, `#Headers`/`#Data`/`#All`/`#Totals`, contiguous multi-column ranges, comma-separated column unions, and space intersections), a broader clean-room formula catalog spanning logical/text/lookup/conditional aggregation, statistical ranking (`MEDIAN`, `MODE.SNGL`, `LARGE`, `SMALL`, `RANK.EQ`), signed decimal rounding, deterministic Excel 1900/1904 date/workday arithmetic (`DATE`, `DATEVALUE`, `TIME`, `TIMEVALUE`, `HOUR`, `MINUTE`, `SECOND`, `YEAR`, `MONTH`, `DAY`, `EDATE`, `EOMONTH`, `DAYS`, `WEEKDAY`, `NETWORKDAYS`, `WORKDAY`, `NETWORKDAYS.INTL`, `WORKDAY.INTL`) with native `workbookPr@date1904` roundtrip, custom weekend masks, and host-locale-independent ISO/English date-text plus ASCII numeric `VALUE` coercion, and dynamic-array helpers (`SEQUENCE`, `TRANSPOSE`, `FILTER`, `UNIQUE`, `SORT`, `TAKE`, `DROP`, `CHOOSECOLS`, `CHOOSEROWS`, `TOCOL`, `TOROW`, `WRAPROWS`, `WRAPCOLS`, `HSTACK`, `VSTACK`, `EXPAND`), live range style/dimension/autofit formatting, Excel-style `fillDown`/`fillRight` formula translation, native merged-cell creation/import/export, native shared strings/styles, built-in/custom number-format and cell-style inheritance, theme/tint/indexed-color resolution, and per-edge borders through a dedicated SpreadsheetML codec, formatted layout/SVG display values, frozen-pane/gridline worksheet views, column widths/best-fit/hidden state, row heights/hidden state, comments/data-validation/conditional-formatting/table/chart/image/sparkline metadata roundtrips, native XLSX table/chart/image/sparkline/threaded-comment XML parts, and relationship-driven metadata-free import of embedded images plus basic bar/line/pie charts from arbitrary valid drawing targets and one-/two-cell anchors. Dimension-aware SVG previews and workbook/worksheet layout JSON feed semantic and visual QA. The presentation facade includes compose/JSX layout, inspectable shape/table/chart/image/connector objects, speaker notes, threaded comments, bounded PPTX package inspection, structured paragraphs and styled runs, bullets/auto-numbering, nine-level master/layout/slide list inheritance, native text/fill/line styling, explicit native table cell styling, native PPTX table/chart/image/connector XML export plus notes/comment parts and clean-room native import restoration, and a geometry-based layout QA detector for overlap/off-canvas/overflow/connector checks. The document facade includes styled paragraphs, real list items with paragraph/numbering-style inheritance and character or relationship-backed picture bullets, headers/footers, external and internal hyperlinks, inspectable bookmark ranges, fields, citations, images, sections, tracked insertions/deletions, tables, comments, design presets, page-aware layout JSON, visual layout QA, DOCX styles/numbering/header/footer/hyperlink/bookmark/comment/image/section/tracked-change export, and SVG page previews. The PDF facade includes modeled multi-page text/table/image/chart artifacts, explicit stable-ID reading order that drives the tagged structure tree without changing visual paint order, tagged PDF export with language/title metadata, H1/P/Figure elements, semantic `Table` → `TR` → `TH`/`TD` hierarchy, column-scoped headers and image/chart alt text, positioned text, vector tables/charts and embedded PNG/JPEG images, bounded structural inspection, `extractText`, `extractTables`, SVG/layout previews, metadata roundtrips, and an optional PDF.js parser for page geometry/positioned text/table/image-mask extraction. This is a tagged-PDF baseline, not a claim of full PDF/UA conformance. Cross-format `verifyArtifact(...)`, `visualQaArtifact(...)`, and `renderArtifact(...)` helpers provide agent-facing QA and preview entry points, including pluggable PNG/WebP/JPEG/PDF renderers, decoded PNG/JPEG/WebP/PPM baselines, cross-encoding comparisons, and configurable aligned PNG diff heatmaps for visual regressions. Bundled clean-room Open XML SDK WebAssembly slices for XLSX, DOCX, and PPTX are available experimentally; advanced Office fidelity, optional Windows Office validation, robust arbitrary-PDF parsing, and template QA remain roadmap work.

Worksheet-backed native PivotTables now participate in the same metadata-free path as tables and drawings. Export writes worksheet→pivotTable, pivotTable→cache-definition, workbook→cache-definition, and, when cached data is enabled, cache-definition→records relationships plus workbook `pivotCaches`; import follows arbitrary valid targets and reconstructs source/target ranges, row/column/value/calculated/group fields, aggregation, computed values, inspect/resolve/layout, and SVG rendering. Row/column axes produce editable cross-tabs; axis item `include`/`exclude` filters map to native hidden pivot items. `groupFields` derives deterministic Year/Quarter/Month labels from ISO, `Date`, or 1900/1904 serial source dates; multi-level groups export standard `fieldGroup base/par`, `rangePr`, and `groupItems` cache structures, survive metadata-free import and second export, and remain usable as filtered axes. Unsupported imported grouping levels are preserved for re-export while modeled output and verification expose `#NAME?`/`pivotGroupFieldUnsupported`. Absolute date filters support `dateEqual`, `dateNotEqual`, `dateOlderThan`, `dateOlderThanOrEqual`, `dateNewerThan`, `dateNewerThanOrEqual`, `dateBetween`, and `dateNotBetween`; they use whole-day ISO dates by default or accept `useWholeDay: false` with ISO date-time/`Date` thresholds at UTC-second precision. Precise filters roundtrip through standard Pivot `filters`, compatible hidden items, and the public x14 `pivotFilter useWholeDay="0"` extension, including metadata-free import, alternate namespace prefixes, and a second export. Relative UTC filters cover yesterday/today/tomorrow, last/this/next ISO week/month/quarter/year, and year-to-date; they remain whole-day and accept deterministic `asOf` anchors for modeled evaluation. Cache policy roundtrips `refreshOnLoad`, `saveData`, `enableRefresh`, `invalid`, `missingItemsLimit`, `refreshedBy`, and `refreshedDateIso`. `calculatedFields` operates on grouped source-field sums, matching Excel's calculated-field rather than calculated-item semantics. The bounded parser supports arithmetic, percent, concatenation, comparisons, strings/booleans, 12 numeric functions (`ABS`, `SUM`, `MIN`, `MAX`, `AVERAGE`, `ROUND`, `PRODUCT`, `POWER`, `SQRT`, `MOD`, `SIGN`, `INT`), `AND`/`OR`/`NOT`, lazy `IF`/`IFERROR`/`IFNA`, `NA`, and `ISERROR`/`ISNUMBER`/`ISTEXT`. Its AST validates every branch without dynamic code execution while lazy evaluation prevents unused errors from leaking; literal text beginning with `#` is no longer misclassified as an Excel error token. Export emits application-created cache fields with `databaseField="0"` plus native formula attributes. Metadata-free import preserves formulas outside the supported subset for lossless re-export, marks them unsupported, returns `#NAME?` for modeled values, and emits an explicit verification issue. Setting `saveData: false` intentionally omits cache records while retaining a refreshable worksheet-backed definition.

PPTX comments support two native contracts. Legacy export (the compatibility default) writes `ppt/commentAuthors.xml`, per-author indexes, initials, display colors, and `lastIdx`. `Presentation.create({ commentFormat: "modern" })` writes Office 2021 `p188` Comment/Author parts with GUID identities, nested replies, dates, and active/resolved/closed status. Shapes, tables, charts, images, connectors, and grouped shapes use persistent `a16:creationId` identities. `slide.groups.add(...)` creates nested `p:grpSp` trees whose outer frame and local child coordinate rectangle map to DrawingML `off/ext` and `chOff/chExt`; child shapes, connectors, groups, native tables, relationship-backed charts, and pictures participate in inspect, resolve, layout JSON, SVG, native import/export, and layout QA. Imported top-level or grouped `p:contentPart`, OLE, SmartArt/diagram, and otherwise unsupported graphic frames are exposed as read-only `nativeObject` facades. Their raw element, direct relationships, recursively reachable OPC parts, content types, external targets, and binary bytes survive metadata-free second export, including path collisions with modeled images. Modern group targets emit `grpSpMk`; nested shapes use `spMk`, grouped tables/charts use `graphicFrameMk`, grouped pictures use `picMk`, and nested `shapeId/text` targets append `txMk@cp/@len` after the complete ancestor path. Imported subranges and optional context hashes survive a second export without private metadata. Package inspection rejects invalid/duplicate identities, mismatched ancestor paths, missing/out-of-bounds ranges, invalid authors/dates/status, and missing anchors. Unsupported native-object comment targets still use `unknownAnchor`. The multiple-master/theme/layout graph remains relationship-driven and metadata-free across second export.

Presentation themes expose the complete `tx1/bg1/tx2/bg2`, `accent1`–`accent6`, hyperlink color scheme, Latin/East-Asian/complex-script major/minor fonts, master title/body/other text styles, and semantic color mapping. `shape.text.set(...)` accepts plain text or structured paragraphs with styled runs, levels 0–8, Unicode bullets, DrawingML auto-number schemes, marker font/color/fixed-or-percentage size, explicit follow-text marker overrides, alignment, indents, point/percentage spacing, and font styling. Export writes complete DrawingML color/font/format schemes plus Slide Master `clrMap` and nine-level text styles; native import restores both direct text and the master → placeholder → layout → slide paragraph cascade even when valid namespace prefixes differ. Format-specific paragraph/list parsing and serialization lives in `src/presentation/text-paragraphs.mjs`, while theme XML remains in `src/presentation/ooxml-theme.mjs`.

DOCX themes and run formatting are native WordprocessingML rather than metadata-only hints. `DocumentModel.create({ theme, defaultRunStyle, styles })` models all 12 scheme colors, major/minor script fonts, document-wide `w:docDefaults`, and paragraph/character styles with `basedOn` inheritance. Runs preserve `w:rStyle` references, direct `w:rFonts`, theme/script font attributes, direct or theme/tint/shade colors, and paired `w:b`/`w:bCs`, `w:i`/`w:iCs`, and `w:sz`/`w:szCs`. Metadata-free import applies the Word run cascade in precedence order—document defaults, paragraph style chain, character style chain, then direct formatting—while retaining source references and resolved colors/fonts for inspect, layout JSON, SVG, and a second export. Alternate valid prefixes in styles and Theme parts are accepted. The codec lives in `src/ooxml/docx-run-styles.mjs`.

DOCX headers and footers support `default`, `first`, and `even` reference types plus an optional zero-based `sectionIndex`. Each declared section/type combination gets its own relationship-driven part and is attached to that section's `w:sectPr`; omitting `sectionIndex` preserves the final-section authoring behavior. Declaration and activation are separate: `document.setSectionSettings(index, { differentFirstPage })` controls section-local `w:titlePg`, `settings.evenAndOddHeaders` controls even-page selection, and `activateVariant: false` can retain a dormant first/even reference during clean-room construction. A missing reference inherits the previous section's reference of the same type, while a missing reference in the first section produces a blank variant. `document.layoutJson()` reports the effective reference type, source section, inherited state, and blank state for each modeled page. The dedicated `src/ooxml/docx-sections.mjs` planner owns these rules across export, prefix-agnostic native import, layout, and SVG preview, so metadata-free roundtrips preserve dormant references instead of silently enabling them. `DocumentFile.patchDocx(...)` header/footer recipes accept the same `sourceReference: { type, sectionIndex }` targeting. Use `DocumentFile.importDocx(..., { preferNative: true })` after package-level patches so native OOXML takes precedence over embedded clean-room metadata.

Classic DOCX comments preserve `author`, deterministic or caller-supplied `initials`, and optional `date` metadata. Paragraph and table targets export with matching `w:commentRangeStart`, `w:commentRangeEnd`, and `w:commentReference` anchors. Native import locates the Comments part through `document.xml.rels`, so a standards-compliant package may place it at an arbitrary internal target instead of `word/comments.xml`; `{ preferNative: true }` restores author metadata and table targets from the package itself. `DocumentFile.patchDocx(...)` also accepts a comments recipe with `sourceReference: { anchors: [{ commentId, target }] }`: targets may select a top-level block, a document-order paragraph, or a table cell, every ID must exist uniquely in the patched Comments part, and deleting the part removes matching document anchors.

Threaded DOCX review metadata now spans the native Word generations without private metadata. Office 2013 `commentsExtended` maps each classic comment's final `w14:paraId` to reply parents and resolved state. Office 2019 `commentsIds` maps that paragraph identity to a persistent eight-hex-digit `durableId`, while Office 2021 `commentsExtensible` carries the same durable identity plus normalized UTC dates and follow-up placeholder state. The Word `people` part preserves each author's provider/user presence identity. Export plans the four parts as one graph; relationship-driven, prefix-agnostic import restores `paraId`, `durableId`, `dateUtc`, `person`, replies, and resolution, and package inspection rejects missing/duplicate/range-invalid identities, unresolved mappings, invalid UTC dates, reply placeholders, conflicting people, wrong roots/content types/sources, multiplicity, and orphan parts.

DOCX numbering recipes accept `sourceReference: { assignments: [{ numId, level, target }] }`. Each assignment adds or replaces the target paragraph's `w:numPr`, verifies that the Numbering part uniquely declares the requested instance and level, supports the same block/paragraph/table-cell selectors as comment anchors, and removes all `w:numPr` references when that Numbering part is deleted. Native-preferred import restores the resulting list kind, number format, start, level text, numbering ID, and nesting level.

DOCX settings are agent-facing through `DocumentModel.create({ settings })`, `document.setSettings(...)`, `inspect({ kind: "settings" })`, native export/import, and package recipes. A settings recipe accepts `sourceReference: { trackRevisions, updateFields, evenAndOddHeaders, mirrorMargins, documentProtection }`, mutates an arbitrary relationship-backed Settings part in CT_Settings schema order, and preserves unrelated children. Passwordless protection modes are `none`, `readOnly`, `comments`, `trackedChanges`, and `forms`; this is an editing safeguard, not encryption or a security boundary.

DOCX bookmark and hyperlink semantics live in `src/ooxml/docx-links.mjs`. `document.addBookmark(target, name, { endTarget, nativeId })` creates a block-backed or exact `table.getCell(row, column)` range that participates in inspect/resolve/verify and emits paired `w:bookmarkStart`/`w:bookmarkEnd` markers in document order. `document.addHyperlink(...)` accepts an external URL, `#bookmark`, a bookmark facade, or `{ anchor }`; internal links use `w:anchor` without creating a package relationship, while external links use `r:id`. Metadata-free import restores block/table-cell ranges, native numeric IDs, anchors, tooltips, history state, external URLs, and relationship IDs. Package inspection rejects duplicate, unpaired, reversed, out-of-range bookmarks and dangling internal anchors.

The experimental OpenChestnut DOCX path additionally treats bounded whole-paragraph hyperlinks and simple fields as source-bound editable content. Its C# Open XML SDK codec can change hyperlink text, tooltip, history, and external/internal target mode while changing only relationships owned by that hyperlink. It can also edit one-run `w:fldSimple` instructions/results for a bounded non-fetching catalog such as `PAGE`, `NUMPAGES`, dates, document properties, and counts. Residual hashes retain native run/paragraph formatting; mixed-content, multi-run, complex, and unsupported fields remain read-only instead of being flattened or executed.

Native Word bibliography semantics live in `src/ooxml/docx-bibliography.mjs`. `document.addBibliographySource(...)` creates resolvable `b:Source` entries with Word source types, personal or corporate authors, deterministic GUIDs, and structured publication fields; `document.addCitation(...)` writes a real `CITATION` field keyed by the source tag. Export stores cited and uncited sources in a relationship-owned `customXml/item*.xml` `b:Sources` part. Metadata-free import follows arbitrary customXml targets, accepts alternate namespace prefixes, restores source/style metadata and visible field results, and preserves them on second export. Package inspection rejects invalid or duplicate sources, unlinked bibliography parts, and citation tags without a matching source. Generic `w:fldSimple` and complex fields continue to recover their instruction/display values, and comments attached to hyperlinks, fields, and citations use the same native range/reference anchors.

DOCX lists use real multi-level numbering definitions rather than assuming `numId` 1/2. `document.addListItem(...)` accepts levels 0–8, `numberFormat`, positive `start`, `levelText`, and an optional `numberingId` that groups levels into one list instance. Export emits every referenced level in `abstractNum` plus a `num` instance and points each paragraph at the generated ID. Native import follows the styles and numbering relationships to arbitrary internal part paths, resolves `numId` → `abstractNumId` → `lvl`, applies `lvlOverride`/`startOverride`, and preserves formats such as `upperLetter` and `lowerRoman` for inspect/re-export.

PDF content using Chinese, Cyrillic, Greek, accented Latin, or other non-ASCII characters must provide a legally usable standalone glyf-based TrueType `.ttf` through `PdfFile.exportPdf(..., { font })`. The clean-room writer embeds a Type0/CIDFontType2 font, glyph widths, a `ToUnicode` CMap, and by default a deterministic font subset containing used glyphs plus recursive composite dependencies. It rebuilds `glyf`, long `loca`, Unicode `cmap`, the sfnt directory, table checksums, and `checkSumAdjustment`; `{ subsetFont: false }` preserves the full font only for diagnostics. Missing fonts/glyphs, `.ttc` collections, malformed tables, and oversized inputs fail explicitly rather than corrupting text. Direct Unicode mapping is implemented; complex-script shaping remains roadmap work.

Modeled PDF tables accept zero-based `cells` overrides with `rowSpan`, `columnSpan`, `TH`/`TD` roles, `Row`/`Column`/`Both` scope, and explicit header-cell IDs. `table.getCell(row, column)`, `inspect({ kind: "tableCell" })`, `resolve`, extraction, layout JSON, SVG, and vector PDF output share one normalized grid. Tagged PDF 1.7 export assigns every origin cell a structure ID and emits the standard Table-owner `RowSpan`, `ColSpan`, `Scope`, and `Headers` attributes; verification rejects out-of-range/overlapping spans and missing or non-header associations.

Call `page.setReadingOrder([...])` with every semantic page target exactly once. Use `${page.id}/text` for non-positioned body text, followed by positioned-text, table, image, and chart IDs in the desired logical sequence. The declared order is round-tripped in the model, exposed through `inspect({ kind: "readingOrder" })`, `resolve`, and layout JSON, validated fail-closed, and emitted through `StructTreeRoot`; `PdfFile.inspectPdf()` reports the resulting top-level structure IDs. Visual paint order remains unchanged.

Every informative PDF image and chart must provide concise, meaningful `alt` text. Set `decorative: true` only for content that should remain visible but be ignored by assistive technology; decorative figures are excluded from reading order and emitted as `/Artifact BMC` marked content instead of `Figure` structure elements. `pdf.verify()` rejects missing or generic alternatives, and the runnable PDF verifier compares modeled Figure/Artifact expectations with real exported bytes.

Set `headingLevel: 1` through `6` on positioned PDF text that acts as a heading. Visual style remains independent from semantics: the value is preserved through inspect/layout/model roundtrip and exported as an H1-H6 structure role. `pdf.verify()` reports a heading sequence that starts below H1 or skips a level, and the runnable verifier compares modeled heading counts with the real tagged bytes.

The bundled codec also has bounded DOCX and PPTX slices. The complex Documents fixture crosses DOCX with source-preserving advanced WordprocessingML and real page-render QA. PPTX models top-level rectangle/ellipse shapes plus a bounded paragraph subset: ordered text, field, and styled line-break inlines; paragraph tab stops, direct left margins, hanging indents, spacing, and `a:defRPr` default-run styles; shape-local `a:lstStyle` defaults for levels 0 through 8 using the same bounded paragraph-property subset; `a:bodyPr` left/top/right/bottom pixel insets, top/center/bottom anchoring, square/none wrapping, none/shrink-text/resize-shape AutoFit, -360 through 360 degree rotation, horizontal/vertical/vertical270 text, horizontal/vertical overflow, one through sixteen columns with pixel spacing/RTL flow, and upright text; level/alignment; character/auto-number/no-bullet markers; embedded PNG/JPEG/GIF/safe-SVG or non-fetched external picture bullets; direct marker font, solid RGB or one of 16 DrawingML theme colors, fixed/percentage/follow-text marker size; bold/italic, font size/family, solid RGB text color; and bounded click links. Default-run styles retain solid RGB or untransformed theme-color identity; explicit deletion clears only modeled bold/italic/size/Latin-font/color fields and preserves unknown underline, script fonts, and color transforms. Embedded marker identity is a content-addressed protobuf `Asset`; each slide owns only its local Open Packaging Convention relationship, so identical bytes share one media part across slide owners. Source-bound edits retain transformed or unresolved picture bullets, transformed/theme-external marker colors, attributed normal-AutoFit parameters, unknown field paragraph properties, unknown break run properties, and other unmodeled DrawingML when paragraph/inline kinds are unchanged; direct paragraph coordinates, spacing, modeled default-run style, text-body list levels, and bounded body properties can be replaced or explicitly removed while unknown native attributes/children remain source-bound. Attempts to replace an unknown choice or change text/field/break topology fail closed. WordArt vertical modes/warp/3D, body anchor-centering/compatibility flags, master/layout text-style editing, pictures, charts, groups, connectors, content parts, notes, masters, themes, and their relationship graphs remain read-only and preserved until their editable semantics land. Direct authoring outside each modeled subset fails closed.

## Usage

```js
import { Workbook, SpreadsheetFile } from "open-office-artifact-tool";

const workbook = Workbook.create({ theme: { name: "Report Theme", colors: { accent1: "#0F766E", accent2: "#7C3AED" } } });
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C2").values = [["A", "B", "Sum"], [2, 3, null]];
sheet.getRange("C2").formulas = [["=A2+B2"]];
sheet.getRange("A1:C1").format = { fill: { patternType: "lightTrellis", foreground: { theme: 4, tint: 0.4 }, background: "#F8FAFC" }, font: { bold: true, color: { theme: 4, tint: -0.25 } } };
workbook.recalculate();

console.log((await workbook.inspect({ kind: "table,formula" })).ndjson);
const file = await SpreadsheetFile.exportXlsx(workbook);
await file.save("output.xlsx");

// Delimited files flatten one selected sheet/range; calculated values are the default.
await (await SpreadsheetFile.exportCsv(workbook, { sheetName: "Sheet1" })).save("output.csv");
const importedCsv = await SpreadsheetFile.importCsv("Name,Value\r\nRevenue,120", { coerceTypes: true });
```

The experimental OpenChestnut subpath authors a bounded workbook feature set directly. For imported XLSX files, it also carries a budget-checked, hash-bound source-package snapshot and updates modeled workbook/sheet/cell fields in place, so unmodeled styles, tables, pivots, drawings, comments, and arbitrary legal OPC targets survive the second export byte-for-byte at the opaque-part boundary. New advanced workbooks still use the JavaScript codec until their semantics enter the public schema. Consumption uses only bundled runtime assets; a local .NET SDK is not required:

```js
import { Workbook } from "open-office-artifact-tool";
import { exportXlsxWithOpenChestnut, importXlsxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";

const workbook = Workbook.create({ dateSystem: "1904" });
workbook.worksheets.add("Summary").getRange("A1:B2").values = [["Metric", "Value"], ["Revenue", 42.5]];
const xlsx = await exportXlsxWithOpenChestnut(workbook);
const roundtripped = await importXlsxWithOpenChestnut(xlsx);
```

The same subpath authors a bounded DOCX semantic slice. Imported advanced blocks are hash-bound to the original package: unchanged content is preserved, while edits to not-yet-modeled structures fail explicitly.

```js
import { DocumentModel } from "open-office-artifact-tool";
import { exportDocxWithOpenChestnut, importDocxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";

const document = DocumentModel.create({ paragraphs: ["OpenChestnut document"] });
const docx = await exportDocxWithOpenChestnut(document);
const importedDocument = await importDocxWithOpenChestnut(docx);
```

PPTX uses the same loss-aware boundary. A simple deck can be authored directly; importing an advanced deck carries hash-bound master/layout/slide/shape bindings so unsupported native objects survive safe modeled edits. The WebAssembly slice exposes every Slide Master and Slide Layout as a stable package locator, preserves the Slide → Layout → Master chain, and permits bounded title/body/other `p:txStyles`, direct Master/Layout solid or theme-reference backgrounds, and owner-local Master/Layout placeholder text/body/list/link edits. Effective inheritance stays in the JavaScript model; placeholder geometry/fill/shape style/identity, placeholder topology, and slide rebinding remain read-only.

```js
import { Presentation } from "open-office-artifact-tool";
import { exportPptxWithOpenChestnut, importPptxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";

const deck = Presentation.create({
  master: {
    name: "Evidence Master",
    background: { fill: "accent1", mode: "reference", index: 1001 },
    textParagraphStyles: {
      title: { 0: { alignment: "center", style: { bold: true, fontSize: 40, color: "accent1" } } },
    },
  },
});
deck.slides.add({ name: "Overview" }).shapes.add({
  name: "Title",
  text: [
    { alignment: "center", lineSpacing: 1.2, spaceAfter: 8, runs: [
      { text: "OpenChestnut ", style: { bold: true, fontSize: 36, color: "#0F172A" } },
      { text: "presentation", style: { italic: true, fontSize: 36 }, link: { uri: "https://example.com/evidence", tooltip: "Open evidence" } },
    ] },
    { autoNumber: { type: "arabicPeriod", startAt: 1 }, bulletFont: "Aptos", bulletColor: "#2563EB", bulletSizePercent: 1.25, runs: ["Validated list marker"] },
    { bulletImage: { dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABgAAAAYCAYAAADgdz34AAAACXBIWXMAAAPoAAAD6AG1e1JrAAAAMUlEQVR4nGMQW+z1n5aYYdQCsdEgEhtNRWKjGW3xaFGxeLQ0XTxa4XiNVpleg7tVAQAE34S9d7s/SwAAAABJRU5ErkJggg==" }, runs: ["Content-addressed picture marker"] },
  ],
  position: { left: 60, top: 40, width: 860, height: 70 },
});
const pptx = await exportPptxWithOpenChestnut(deck);
const importedDeck = await importPptxWithOpenChestnut(pptx);
```

Presentation compose-first authoring uses helper nodes that mirror the agent-oriented JSX vocabulary while staying transpiler-free:

```js
import { column, paragraph, Presentation, PresentationFile, row, box } from "open-office-artifact-tool";

const presentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  master: {
    name: "Clean Room Master",
    background: { fill: "bg1", mode: "reference", index: 1001 },
    placeholders: [{ type: "title", idx: 1, position: { left: 42, top: 36, width: 1196, height: 80 }, style: { fontSize: 42, bold: true, color: "accent1" } }],
  },
  layouts: [{ id: "layout/content", name: "Content", type: "titleAndContent", background: "#FFFFFF" }],
});
const slide = presentation.slides.add({ layoutId: "layout/content" });
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

### OpenChestnut naming compatibility

`open-office-artifact-tool/codecs/open-chestnut` is the canonical codec subpath and owns all implementation/runtime behavior. The former `open-office-artifact-tool/codecs/openxml-wasm` subpath remains as a deprecated, compatibility-only re-export for existing consumers; its legacy function names are aliases of the OpenChestnut functions and report `metadata.codec === "open-chestnut"`. New code, fixtures, build commands, and documentation should use OpenChestnut. The public wire namespace remains `open_office.artifact.v1` and is independent of either JavaScript module name.

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

XLSX, PPTX, and DOCX expose the same bounded package workflow. Inspect records validate safe paths, content types, internal relationships, duplicate IDs, and namespace-aware source XML references under part/byte budgets. PPTX inspection additionally validates notes plus legacy/Office 2021 comment relationship ownership and multiplicity, roots, singleton author registries, author references, native IDs, dates, status, anchors, per-author index uniqueness, and `lastIdx` bounds. Patch methods accept XML, JSON, text, binary, and remove operations; content types and relationships synchronize automatically, with IDs preserved long enough to clean source XML. Set `recipe.sourceReference` for DOCX settings, header/footer references, classic-comment anchors, and numbering assignments; XLSX worksheet/table lists, complete worksheet→drawing→image/chart chains, and workbook/cache/worksheet PivotTable chains; and PPTX slide/master/layout ID lists plus slide image/chart DrawingML. Drawing recipes require explicit geometry, DOCX settings reject unknown/password fields, comment/numbering builders require declared IDs and explicit structural targets, pivot cache definitions require a unique `cacheId`, and PPTX builders validate schema-bounded IDs and non-visual object IDs, so the package layer never silently invents semantic identity. It owns namespace prefixes, non-visual IDs, source lists/shape trees, add/remove cleanup, and final atomic validation. Use `validateResult: false` only for deliberate invalid-package fixtures.

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
    recipe: {
      kind: "chart",
      source: "ppt/slides/slide1.xml",
      id: "rIdReviewChart",
      sourceReference: {
        objectId: 12,
        name: "Review chart",
        position: { left: 180, top: 120, width: 640, height: 360 },
      },
    },
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

Project-local clean-room agent workflows live under [`skills/`](skills/). [`skills/spreadsheets/SKILL.md`](skills/spreadsheets/SKILL.md) uses this package's public APIs for XLSX/CSV/TSV authoring and import, bounded package/row evidence, semantic verification, layout export, and SVG or Playwright-backed visual QA. [`skills/documents/SKILL.md`](skills/documents/SKILL.md) adds durable DOCX roundtrip, package-part inspection, modeled QA, and optional real LibreOffice PDF + Poppler page-PNG verification. [`skills/presentations/SKILL.md`](skills/presentations/SKILL.md) adds narrative/layout rules, native PPTX package inspection, per-slide modeled/native rendering, montage review, and optional PNG baseline pixel diffs. [`skills/pdf/SKILL.md`](skills/pdf/SKILL.md) adds real PDF byte/object inspection, PDF.js semantic extraction, per-page Playwright and Poppler rendering, and model/native PNG baseline diffs. All four workflows share the shipped `skills/shared/visual-baselines.mjs` lifecycle; when `--baseline-dir` is supplied, missing, empty, or non-contiguous baseline sets fail QA and must be initialized explicitly with `--write-baseline true`.

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
