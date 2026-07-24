# API catalog

Generated from `HELP_CATALOG` in `src/help/index.mjs`.

## document

| Name | Kind | Summary |
| --- | --- | --- |
| `document.addBibliographySource` | api | Add a canonical Word bibliography source for inspect, resolve, and native b:Sources authoring. Recognized imports allow bounded source content edits while source order, IDs, and tags remain source-bound. |
| `document.addBlockTextContentControl` | api | Append one canonical block-level Word plain-text content control around exactly one modeled paragraph and one ordinary run. The handle reports placement=block; OpenChestnut preserves the w:sdt wrapper and binds native identity/topology after import. |
| `document.addBookmark` | api | Wrap exactly one paragraph-like block in a native Word bookmark for inspect, resolve, and internal hyperlinks. Recognized imported whole-block bookmarks are exposed with source identity but remain fixed-topology/read-only; cross-block, nested, crossing, table-cell, and otherwise complex ranges stay opaque-preserved and fail closed on mutation. |
| `document.addChange` | api | Append one bounded whole-paragraph tracked insertion or deletion. OpenChestnut authors native w:ins/w:del markup and permits fixed-topology imported text/author/date edits; mixed or nested revision graphs remain source-bound. |
| `document.addCitation` | api | Add a whole-paragraph bibliography-backed citation exported as a native w:fldSimple CITATION field plus a bounded bookmark. Recognized imports allow display-text edits while source tags and topology remain fixed. |
| `document.addComment` | api | Attach a whole-paragraph Word comment. Classic roots remain minimal; bounded modern roots may carry resolved, durable/UTC, and provider-person metadata through OpenChestnut. |
| `document.addDeletion` | api | Append one bounded whole-paragraph tracked deletion using native w:del/w:delText markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile. |
| `document.addEndnote` | api | Append one native plain-text endnote at the end of one paragraph or list item. Recognized imported canonical endnotes permit body-text edits only; anchor, kind, native ID, and note topology remain source-bound. |
| `document.addField` | api | Append a bounded w:fldSimple block for PAGE, NUMPAGES, SECTION, date/time, and selected document-property commands. External-content and arbitrary reference commands fail closed. |
| `document.addFooter` | api | Add a default, first-page, or even-page DOCX footer, optionally section-scoped; first/even activation is independent from the preserved relationship reference. On an imported package, only a direct unformatted text paragraph in a uniquely used source part advertises editable=true; fields, rich/shared, and irregular parts stay read-only. |
| `document.addFootnote` | api | Append one native plain-text footnote at the end of one paragraph or list item. Recognized imported canonical footnotes permit body-text edits only; anchor, kind, native ID, and note topology remain source-bound. |
| `document.addHeader` | api | Add a default, first-page, or even-page DOCX header, optionally section-scoped; first/even activation is independent from the preserved relationship reference. On an imported package, only a direct unformatted text paragraph in a uniquely used source part advertises editable=true; fields, rich/shared, and irregular parts stay read-only. |
| `document.addHyperlink` | api | Append a native w:hyperlink backed by an external relationship or internal bookmark anchor; native import restores URL/anchor, relationship identity, tooltip, and history state. |
| `document.addImage` | api | Append an inspectable embedded PNG/JPEG image. Images are inline by default; an explicit bounded placement authors a native foreground wp:anchor with square or top-and-bottom wrapping. |
| `document.addInsertion` | api | Append one bounded whole-paragraph tracked insertion using native w:ins markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile. |
| `document.addListItem` | api | Append a numbered, character-bulleted, or bounded picture-bulleted list item using native DOCX numbering definitions. Picture markers are shared numbering-level resources: every item using the same numberingId and level must agree, and recognized imported edits must update the complete group without changing embedded-versus-external source kind. |
| `document.addParagraph` | api | Append a styled paragraph with optional run spans and bounded direct paragraph formatting, including presence-aware line-number suppression. |
| `document.addSection` | api | Append a DOCX section break with page size, orientation, margins, binding gutter, canonical equal-width or explicit-width columns, bounded page-number start/format, and break-type metadata backed by w:sectPr. Imported geometry and page numbering are writable only when their native markup is canonical. |
| `document.addTable` | api | Append a Word-style table with physical cell values, optional logical merge geometry, and fixed-layout width/margin/border/header formatting. |
| `document.addTableOfContents` | api | Append one canonical one-paragraph complex TOC field with bounded heading levels/switches and enable the native updateFields-on-open hint by default. Refreshed cross-paragraph result graphs remain opaque/source-bound and read-only. |
| `document.addWatermark` | api | Add one canonical VML text watermark to a section/header-reference scope. Recognized imported watermarks permit text-only edits or whole-object removal; adding to an imported package, changing scope, shared headers, multiple objects, DrawingML, images, and irregular VML fail closed. |
| `document.applyDesignPreset` | api | Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews. |
| `document.contentControls` | api | List typed mutable handles for recognized inline or table-cell plain-text, checkbox, drop-down, combo-box, and date controls plus block plain-text controls, with explicit placement and model/native identity. |
| `document.fillContentControls` | api | Transactionally fill every recognized block, inline, or table-cell plain-text control matching an object or Map of tag-to-string entries. Checkbox, drop-down, combo-box, and date tags do not silently accept text. |
| `document.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of document theme and explicit run/style font families. |
| `document.inspect` | api | Emit bounded NDJSON for document blocks including typed block/inline plain-text and inline checkbox/list/date content controls with explicit placement, fields, tracked changes, bookmark ranges, footnotes/endnotes, bibliography sources, comments, styles, headers/footers with sourceBound/editable evidence, canonical text watermarks, and layout; narrow with search/target anchors and fields with include/exclude. |
| `document.layoutJson` | api | Return page-aware layout JSON with block bounding boxes, section/page ordinals, effective inherited header/footer selections, styles, and target/search slicing. |
| `document.materializeFields` | api | Transactionally compute canonical inline SEQ counters and REF cached results from native bookmark targets, with dry-run evidence and strict missing-target failure. PAGEREF remains skipped because trustworthy page numbers require a real pagination host. |
| `document.render` | api | Render an SVG preview by default, return layout JSON with { format: 'layout' }, or use { source: 'docx', renderer } to feed native DOCX into LibreOffice/native Office render adapters for PDF/PNG outputs. |
| `document.replyToComment` | api | Add one source-free direct reply to a root comment. OpenChestnut authors the bounded commentsExtended graph; nested replies and imported topology changes fail closed. |
| `document.resolve` | api | Resolve stable document, block, table-cell, content-control, bookmark, footnote/endnote, bibliography source ID/tag, header/footer, watermark, comment, style, and advertised text-range IDs. |
| `document.setCheckboxContentControls` | api | Transactionally set every recognized canonical checkbox control matching an object or Map of tag-to-boolean entries. Other control types do not silently coerce. |
| `document.setComboBoxContentControls` | api | Transactionally set every recognized canonical combo-box control from a tag-to-value string mapping. Values may select a declared choice or provide bounded custom text; unknown tags and invalid values fail before mutation. |
| `document.setDateContentControls` | api | Transactionally set every recognized canonical date control from a tag-to-YYYY-MM-DD mapping. Invalid Gregorian dates, unknown tags, and other control types fail before mutation. |
| `document.setDropdownContentControls` | api | Transactionally set every recognized canonical drop-down control from a tag-to-choice-value string mapping. Unknown tags or values outside the declared choice table fail before mutation. |
| `document.setSectionSettings` | api | Set per-section Word behavior such as different-first-page header/footer activation without changing preserved header/footer references. |
| `document.setSettings` | api | Set model settings. evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields refresh hint, and bounded passwordless documentProtection are inside the OpenChestnut 0.3 DOCX boundary. Irregular page-margin mode markup and password/cryptographic protection variants stay source-owned and fail closed on replacement. |
| `document.styles.effective` | api | Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata. |
| `document.textRange` | api | Inspect or resolve stable textRange anchors such as blockId/text and tableId/cell/row/column/text. Assignment is limited to fully editable text; replace() also supports explicitly advertised source-bound literal patches. |
| `document.verify` | api | Return QA issues for invalid/duplicate content-control IDs and native IDs, malformed tags/aliases, invalid block-control profiles, fake lists, invalid links/citations/bibliography sources, malformed tracked changes, duplicate/dangling/reversed bookmark ranges, invalid footnotes/endnotes, unknown styles, malformed tables, bad images/sections, invalid watermark IDs/scopes/text, dangling comments, visual overflow, and prose-like table cells. |
| `documentComment.reopen` | api | Clear the resolved state of a bounded modern comment without changing its root/reply topology or durable identity. |
| `documentComment.resolve` | api | Set resolved=true for a bounded modern comment. Imported edits re-prove source hashes and commentsExtended topology while keeping thread identity fixed. |
| `DocumentFile.addTrackedReplacement` | api | Add one exact replacement inside a direct body paragraph or bounded table-cell paragraph to hash-bound DOCX source bytes as adjacent native w:del/w:ins runs. A structured paragraph/tableCell selector, full expected text, and one unique literal contained in either one ordinary run or adjacent run fragments with identical w:rPr preserve source formatting; mixed formatting and broader topologies fail closed with exact changed-part audit. |
| `DocumentFile.exportDocx` | api | Export DocumentModel to DOCX through the single bundled OpenChestnut codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly. |
| `DocumentFile.finalizeRevisions` | api | Accept or reject bounded direct whole-paragraph one-run revisions and exact adjacent in-paragraph w:del + w:ins pairs from source bytes, including same-format fragmented deletions in direct body paragraphs or bounded table-cell paragraphs. Mandatory SHA-256 binding, decompression budgets, exact changed-part audit, and fail-closed graph checks prevent silent model reconstruction or broad package mutation. |
| `DocumentFile.importDocx` | api | Import relationship-driven core DOCX semantics through the single bundled OpenChestnut codec. An imported header/footer advertises editable only for one direct unformatted text paragraph in a uniquely used source part; it is hash-bound, allows at most one text edit per part, and leaves PAGE/simple fields, rich, shared, inherited, and irregular page furniture read-only. Recognized inline controls, fields, revisions, notes, citations, simple tables, and exclusive canonical VML text-watermark paragraphs are fixed-topology editable; otherwise read-only paragraphs and complex table cells separately advertise textPatchable when at least one direct ordinary native text node can participate in a bounded literal patch. A unique literal may span adjacent same-format runs without rebuilding the surrounding graph. |
| `DocumentFile.inspectDocx` | api | Inspect bounded DOCX parts, content types, relationships, and namespace-aware source XML r:id/r:embed/r:link references under decompression budgets. |
| `DocumentFile.patchDocx` | api | Apply DOCX part patches with path traversal validation for settings, classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people parts, and numbering assignments; atomically reject dangling packages and invalid comment graphs. |
| `DocumentModel.create` | api | Create a document with paragraph/character styles, formatted paragraphs/runs, canonical inline and one-paragraph table-cell plain-text, checkbox, drop-down, combo-box, and ISO/Gregorian date content controls, one-paragraph block plain-text controls, canonical inline SEQ/REF/PAGEREF fields, sections, headers/footers, canonical VML text watermarks, lists, TableGrid fixed-geometry tables, links, bounded whole-block bookmarks, plain-text footnotes/endnotes, canonical bibliography-backed citations, simple fields, a canonical complex TOC placeholder, bounded whole-paragraph tracked insertions/deletions, classic comments, bounded modern root/direct-reply threads, and PNG/JPEG images. Nested/irregular modern threads, rich comment bodies, multi-paragraph/rich/inline-within-cell/nested/data-bound/locked/placeholder table-cell SDTs, other nested/data-bound/locked/placeholder SDTs, irregular lists, localized dates, custom checkbox symbols, image/DrawingML/irregular VML watermarks, other complex field graphs, arbitrary table-style graphs, complex bookmark/note/revision graphs, and advanced settings remain unsupported or source-bound. |
| `documentTableCell.addCheckboxContentControl` | api | Wrap one source-free rectangular table cell in a canonical Word 2010+ checkbox w:sdt. OpenChestnut owns the visible glyph and symbols; recognized imports permit checked/tag/alias edits while identity, type, placement, symbols, and topology remain fixed. |
| `documentTableCell.addComboBoxContentControl` | api | Wrap one source-free rectangular table cell in a canonical standard combo-box w:sdt with ordered choices and a declared-or-custom typed value. Recognized imports permit value/tag/alias edits while the choice table and topology remain fixed. |
| `documentTableCell.addDateContentControl` | api | Wrap one source-free rectangular table cell in the canonical ISO/Gregorian date w:sdt profile. Recognized imports permit dateValue/tag/alias edits while native date metadata, placement, and topology remain fixed. |
| `documentTableCell.addDropdownContentControl` | api | Wrap one source-free rectangular table cell in a canonical standard drop-down w:sdt with ordered choices and a typed selectedValue. Recognized imports permit selectedValue/tag/alias edits while the choice table and topology remain fixed. |
| `documentTableCell.addTextContentControl` | api | Wrap one source-free rectangular table cell's existing text in a canonical cell-level plain-text w:sdt. The handle reports placement=tableCell plus row/column; recognized imported controls permit fixed-topology text/tag/alias edits, while adding or removing imported control topology fails closed. |
| `documentTableCell.replaceText` | api | Apply a literal source-bound text patch to one table cell that advertises textPatchable. The search must resolve exactly once inside one ordinary native w:t node or adjacent non-empty direct runs with byte-identical w:rPr. Whole-cell replacement, mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and ambiguous matches fail closed. |
| `documentWatermark.remove` | api | Remove one modeled or recognized source-bound canonical watermark as a complete header paragraph. The source-bound operation re-proves exact element and header residual hashes and never heuristically deletes arbitrary header graphics. |
| `exportDocxWithOpenChestnut` | api | Export bounded DocumentModel paragraphs/runs, fields, tables, bookmarks, notes, citations, tracked changes, comments, images, canonical text watermarks, sections, numbering, and settings; recognized imports permit exact-profile semantic edits plus hash-bound literal patches to one unique ordinary paragraph or table-cell span inside one direct w:r/w:t or adjacent same-format runs while preserving all surrounding native markup. |
| `importDocxWithOpenChestnut` | api | Import DOCX bytes through OpenChestnut with source-bound blocks, recognized exclusive canonical VML text-watermark paragraphs, and source-bound header/footer editable evidence. A header/footer edit is limited to one direct unformatted text paragraph in one uniquely used source part; fields, rich/shared/inherited page furniture, scope changes, and multiple edits to one part fail closed. Literal body/table patch capability never implies whole-paragraph/cell editability; only adjacent non-empty direct runs with byte-identical w:rPr may form one patch span, while mixed-format, gapped, cross-paragraph, ambiguous, field/control/revision text remains fail-closed. |
| `paragraph.addCheckboxContentControl` | api | Append one canonical Word 2010+ checkbox content control with typed checked state; OpenChestnut owns its visible glyph and w14 symbol declarations. |
| `paragraph.addComboBoxContentControl` | api | Append one canonical inline Word combo-box content control with ordered displayText/value choices and a typed value that may be a declared choice or bounded custom text. OpenChestnut derives the visible projection. |
| `paragraph.addDateContentControl` | api | Append one canonical inline Word date picker from a real Gregorian YYYY-MM-DD value. OpenChestnut owns the fixed ISO display, UTC-midnight fullDate, language, mapping, and calendar projection. |
| `paragraph.addDropdownContentControl` | api | Append one canonical inline Word drop-down content control with an ordered displayText/value choice table and typed selectedValue. OpenChestnut derives visible text from the selected choice. |
| `paragraph.addField` | api | Append a logical inline SEQ, REF, or PAGEREF field run. A SEQ run may add a bookmark around only its cached result for real caption-number targets. OpenChestnut authors/imports the canonical native graph; imported field position, instruction, and bookmark identity remain source-bound while cached display text is editable. |
| `paragraph.addTextContentControl` | api | Append one inline plain-text Word content-control run with agent ID, tag, alias, text, and optional run formatting. OpenChestnut assigns native w:id identity and authors canonical w:sdt markup. |
| `paragraph.replaceText` | api | Replace literal paragraph text without flattening formatting boundaries. Fully editable one-run paragraphs update their existing run; imported source-bound paragraphs advertise textPatchable when OpenChestnut can replace one unique ordinary w:r/w:t node or adjacent non-empty direct runs with byte-identical w:rPr while preserving all native topology and surrounding markup. Mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and duplicate matches fail closed. |

### document details

#### `document.addBibliographySource`

Add a canonical Word bibliography source for inspect, resolve, and native b:Sources authoring. Recognized imports allow bounded source content edits while source order, IDs, and tags remain source-bound.

**Schema parameters:**

- `tag` (string) required — Unique Word source tag used by CITATION fields: 1 through 255 ASCII letters, digits, periods, underscores, colons, or hyphens.
- `sourceType` (string) required — Word bibliography source type such as Book, Report, JournalArticle, InternetSite, or Misc.
- `title` (string) — Source title.
- `authors` (object[]|string[]) — Personal contributors with first/middle/last names.
- `corporateAuthor` (string) — Corporate author used when personal authors are absent.
- `year` (string|number) — Publication year.
- `publisher` (string) — Publisher.
- `url` (string) — Source URL.
- `fields` (object) — Additional supported Word bibliography fields such as city, journalName, volume, issue, pages, edition, and standardNumber.

**Schema returns:**

- `source` (DocumentBibliographySource) — Canonical b:Source entry. Recognized imports permit bounded field/author edits with fixed source order, ID, and tag.

#### `document.addBlockTextContentControl`

Append one canonical block-level Word plain-text content control around exactly one modeled paragraph and one ordinary run. The handle reports placement=block; OpenChestnut preserves the w:sdt wrapper and binds native identity/topology after import.

**Schema parameters:**

- `text` (string) required — Initial visible paragraph text, including the empty string when the template is intentionally blank.
- `blockId` (string) — Optional agent-facing paragraph block ID; generated when omitted.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Block plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, 1 to 255 characters; defaults to tag.
- `styleId` (string) — Optional modeled paragraph style ID.
- `paragraphFormat` (object) — Optional modeled paragraph formatting for the wrapped paragraph, including presence-aware boolean suppressLineNumbers with the same direct/style inheritance and fail-closed source rules as document.addParagraph.
- `runStyle` (object) — Optional modeled formatting for the single ordinary run.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Appended canonical body-level block w:sdt around one paragraph/run. Multi-run, inline-field/control, non-text, nested, locked, placeholder, repeating-section, and data-bound profiles fail closed; use documentTableCell.addTextContentControl for the separate canonical cell-level profile.

#### `document.addBookmark`

Wrap exactly one paragraph-like block in a native Word bookmark for inspect, resolve, and internal hyperlinks. Recognized imported whole-block bookmarks are exposed with source identity but remain fixed-topology/read-only; cross-block, nested, crossing, table-cell, and otherwise complex ranges stay opaque-preserved and fail closed on mutation.

**Schema parameters:**

- `target` (string|object) required — Paragraph-like block ID/facade to wrap. Canonical authoring does not accept table cells or multi-block ranges.
- `name` (string) required — Unique case-insensitive Word bookmark name: ASCII letter first, then letters, digits, or underscore, at most 40 characters.
- `endTarget` (string|object) — Optional end block. Canonical authoring requires it to be the same block as target.
- `nativeId` (number) — Optional unsigned 32-bit Word bookmark numeric identity for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `bookmark` (DocumentBookmark) — Native whole-block bookmark. Recognized imports are inspectable/resolvable but fixed-topology and read-only.

#### `document.addChange`

Append one bounded whole-paragraph tracked insertion or deletion. OpenChestnut authors native w:ins/w:del markup and permits fixed-topology imported text/author/date edits; mixed or nested revision graphs remain source-bound.

**Schema parameters:**

- `changeType` (string) required — insert or delete.
- `text` (string) required — Revision text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded tracked-change block authored as native whole-paragraph w:ins or w:del markup.

#### `document.addCitation`

Add a whole-paragraph bibliography-backed citation exported as a native w:fldSimple CITATION field plus a bounded bookmark. Recognized imports allow display-text edits while source tags and topology remain fixed.

**Schema parameters:**

- `text` (string) required — Visible citation text.
- `metadata` (object) required — Structured citation metadata containing a bounded ASCII tag that resolves to document.bibliographySources.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `citation` (DocumentCitationBlock) — Native whole-paragraph w:fldSimple CITATION block. Imported display text is editable while its tag and topology remain source-bound.

#### `document.addComment`

Attach a whole-paragraph Word comment. Classic roots remain minimal; bounded modern roots may carry resolved, durable/UTC, and provider-person metadata through OpenChestnut.

**Schema parameters:**

- `target` (string|object) required — Stable block ID or block facade.
- `text` (string) required — Comment text.
- `author` (string) — Comment author.
- `initials` (string) — Author initials written to w:initials; derived deterministically from author when omitted.
- `date` (string) — Optional ISO-style comment timestamp written to w:date.
- `resolved` (boolean) — Optional w15:done state. Its presence selects the bounded modern comment graph.
- `parentId` (string) — Root comment model ID for a direct reply; prefer document.replyToComment().
- `paraId` (string) — Optional w14/w15 paragraph identity from 00000001 through 7FFFFFFF; generated deterministically when omitted for a modern source-free graph.
- `durableId` (string) — Optional Office 2019 durable identity from 00000001 through 7FFFFFFE; generated for the complete thread when required.
- `dateUtc` (string) — Optional ISO 8601 Office 2021 UTC metadata.
- `person` (object) — Optional presence identity for this author: providerId is 1-100 characters and userId is 1-300. Every comment with the same author must use the same identity or omit it consistently.
- `intelligentPlaceholder` (boolean) — Optional Office 2021 intelligent-placeholder flag.

**Schema returns:**

- `comment` (DocumentComment) — Attached classic or bounded modern whole-paragraph root comment. Rich bodies and irregular support-part graphs fail closed.

#### `document.addDeletion`

Append one bounded whole-paragraph tracked deletion using native w:del/w:delText markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile.

**Schema parameters:**

- `text` (string) required — Deleted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded whole-paragraph tracked deletion.

#### `document.addEndnote`

Append one native plain-text endnote at the end of one paragraph or list item. Recognized imported canonical endnotes permit body-text edits only; anchor, kind, native ID, and note topology remain source-bound.

**Schema parameters:**

- `target` (string|DocumentParagraphBlock|DocumentListItemBlock) required — Paragraph or list-item ID/facade whose final run receives the native endnote reference.
- `text` (string) required — Plain-text endnote body, 1 through 1,000,000 XML-safe characters.
- `name` (string) — Optional inspectable note name.
- `nativeId` (number) — Optional positive 32-bit native endnote ID for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `endnote` (DocumentNote) — Native bounded endnote. Canonical imports allow text-only edits with fixed topology.

#### `document.addField`

Append a bounded w:fldSimple block for PAGE, NUMPAGES, SECTION, date/time, and selected document-property commands. External-content and arbitrary reference commands fail closed.

**Schema parameters:**

- `instruction` (string) required — Bounded simple Word field instruction such as PAGE, NUMPAGES, SECTION, DATE, or a supported document-property command.
- `display` (string) — Visible fallback/result text.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `field` (DocumentFieldBlock) — Appended field block.

#### `document.addFooter`

Add a default, first-page, or even-page DOCX footer, optionally section-scoped; first/even activation is independent from the preserved relationship reference. On an imported package, only a direct unformatted text paragraph in a uniquely used source part advertises editable=true; fields, rich/shared, and irregular parts stay read-only.

**Schema parameters:**

- `text` (string) required — Footer text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.
- `referenceType` (string) — default, first, or even section reference type.
- `sectionIndex` (number) — Zero-based target section. Omit to bind to the final section for backward compatibility.
- `activateVariant` (boolean) — Set false to preserve a dormant first/even reference without enabling different-first-page or even/odd behavior.

**Schema returns:**

- `footer` (DocumentHeaderFooterBlock) — Appended footer block.

#### `document.addFootnote`

Append one native plain-text footnote at the end of one paragraph or list item. Recognized imported canonical footnotes permit body-text edits only; anchor, kind, native ID, and note topology remain source-bound.

**Schema parameters:**

- `target` (string|DocumentParagraphBlock|DocumentListItemBlock) required — Paragraph or list-item ID/facade whose final run receives the native footnote reference.
- `text` (string) required — Plain-text footnote body, 1 through 1,000,000 XML-safe characters.
- `name` (string) — Optional inspectable note name.
- `nativeId` (number) — Optional positive 32-bit native footnote ID for source-free authoring; imported identity is source-bound.

**Schema returns:**

- `footnote` (DocumentNote) — Native bounded footnote. Canonical imports allow text-only edits with fixed topology.

#### `document.addHeader`

Add a default, first-page, or even-page DOCX header, optionally section-scoped; first/even activation is independent from the preserved relationship reference. On an imported package, only a direct unformatted text paragraph in a uniquely used source part advertises editable=true; fields, rich/shared, and irregular parts stay read-only.

**Schema parameters:**

- `text` (string) required — Header text.
- `name` (string) — Inspectable block name.
- `styleId` (string) — Named style ID.
- `referenceType` (string) — default, first, or even section reference type.
- `sectionIndex` (number) — Zero-based target section. Omit to bind to the final section for backward compatibility.
- `activateVariant` (boolean) — Set false to preserve a dormant first/even reference without enabling different-first-page or even/odd behavior.

**Schema returns:**

- `header` (DocumentHeaderFooterBlock) — Appended header block.

#### `document.addHyperlink`

Append a native w:hyperlink backed by an external relationship or internal bookmark anchor; native import restores URL/anchor, relationship identity, tooltip, and history state.

**Schema parameters:**

- `text` (string) required — Visible link text.
- `url` (string|DocumentBookmark) — External HTTP(S) URL, #bookmark name, or bookmark facade.
- `anchor` (string|DocumentBookmark) — Internal bookmark name/facade; mutually exclusive with an external URL.
- `tooltip` (string) — Optional Word hyperlink tooltip, at most 260 characters.
- `history` (boolean) — Whether Word records the hyperlink as visited; defaults to true.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `hyperlink` (DocumentHyperlinkBlock) — Appended external or internal hyperlink block.

#### `document.addImage`

Append an inspectable embedded PNG/JPEG image. Images are inline by default; an explicit bounded placement authors a native foreground wp:anchor with square or top-and-bottom wrapping.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `widthPx` (number) — Rendered width in pixels.
- `heightPx` (number) — Rendered height in pixels.
- `styleId` (string) — Named paragraph style ID.
- `placement` (object) — Optional image placement. Omit it (or use { type: 'inline' }) for inline flow. The bounded floating profile is { type: 'floating', horizontal: { relativeTo: 'margin'|'page'|'column', offsetPx }, vertical: { relativeTo: 'margin'|'page'|'paragraph', offsetPx }, wrap: 'square'|'topAndBottom', wrapSide?: 'bothSides'|'left'|'right'|'largest', distanceFromTextPx?: { top, right, bottom, left } }. wrapSide is square-only; offsets are bounded to +/-10000 px and text distances to 0..10000 px.

**Schema returns:**

- `image` (DocumentImageBlock) — Appended embedded image block. Recognized imported floating images permit only fixed-topology placement edits; inline/floating transitions and unsupported anchor graphs fail closed.

#### `document.addInsertion`

Append one bounded whole-paragraph tracked insertion using native w:ins markup. For one exact in-paragraph replacement in existing source bytes, use DocumentFile.addTrackedReplacement; mixed, moved, nested, and property-level revisions remain outside the bounded profile.

**Schema parameters:**

- `text` (string) required — Inserted text.
- `author` (string) — Revision author.
- `date` (string) — Revision timestamp.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `change` (DocumentChangeBlock) — Appended bounded whole-paragraph tracked insertion.

#### `document.addListItem`

Append a numbered, character-bulleted, or bounded picture-bulleted list item using native DOCX numbering definitions. Picture markers are shared numbering-level resources: every item using the same numberingId and level must agree, and recognized imported edits must update the complete group without changing embedded-versus-external source kind.

**Schema parameters:**

- `text` (string) required — List item text.
- `listType` (string) — bullet or numbered.
- `level` (number) — Zero-based list nesting level.
- `numberFormat` (string) — OOXML numbering format such as bullet, decimal, upperLetter, lowerRoman, or ordinal.
- `start` (number) — Positive starting value for this numbering level.
- `levelText` (string) — OOXML level text template using placeholders such as %1 or %2.
- `numberingId` (number|string) — Optional list-instance identity used to group levels during export and preserved by native import.
- `abstractNumberingId` (number|string) — Optional abstract numbering identity used to share one compatible multilevel definition across list instances; preserved by native import.
- `numberingStyleId` (string) — Optional Word numbering-style identity resolved through styleLink/numStyleLink and flattened safely on second export.
- `pictureBullet` (string|object) — Optional embedded PNG/JPEG/GIF base64 data URL, absolute HTTP(S) URI, or { dataUrl|uri, widthPt|sizePt, heightPt, alt } marker. Width and height are 4 through 72 points; external resources are referenced but never fetched. All list items sharing numberingId and level must use the same marker. Recognized imported markers allow only a coherent full-group edit with the original embedded/external source kind; irregular VML and broader inherited numbering graphs fail closed.
- `styleId` (string) — Named paragraph style ID.

**Schema returns:**

- `listItem` (DocumentListItemBlock) — Appended native-numbering list item, including normalized pictureBullet metadata when configured.

#### `document.addParagraph`

Append a styled paragraph with optional run spans and bounded direct paragraph formatting, including presence-aware line-number suppression.

**Schema parameters:**

- `text` (string) required — Paragraph text.
- `styleId` (string) — Named paragraph style ID.
- `name` (string) — Inspectable block name.
- `paragraphFormat` (object) — Optional modeled paragraph formatting. suppressLineNumbers accepts only boolean true or false: true excludes this paragraph's lines from section line-number display and calculation; false is an explicit direct override of inherited style suppression; omission inherits the named style/default. Canonical direct or style w:suppressLineNumbers leaves are editable, while duplicate, child-bearing, extension-bearing, or invalid lexical markup stays source-owned and semantic replacement fails closed.
- `runs` (object[]) — Optional run spans whose style may include runStyleId plus direct/theme formatting. A run may carry a bounded contentControl { id, tag, alias, nativeId?, controlType?, checked?, choices?, selectedValue?, value? } or inlineField { instruction, bookmarkName?, bookmarkNativeId? }.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Appended paragraph block with stable ID.

#### `document.addSection`

Append a DOCX section break with page size, orientation, margins, binding gutter, canonical equal-width or explicit-width columns, bounded page-number start/format, and break-type metadata backed by w:sectPr. Imported geometry and page numbering are writable only when their native markup is canonical.

**Schema parameters:**

- `breakType` (string) — Section break type such as nextPage or continuous.
- `orientation` (string) — portrait or landscape.
- `pageSize` (object) — Page width/height in twentieths of a point.
- `margins` (object) — Top/right/bottom/left margins plus optional non-negative binding gutter in twentieths of a point. document.settings.gutterAtTop chooses top-edge versus binding-side placement.
- `columns` (object) — Optional canonical text columns. Equal-width profile: { count: 1–45, spacing, separator }. Explicit-width profile: { definitions: [{ width, spacing }], separator }, with 1–45 ordered definitions, positive widths, and non-negative spacing-after values. All values are twentieths of a point; margins, binding-side gutter, widths, and gaps must fit the page content width. The two profiles cannot be mixed; ambiguous or extension-bearing w:cols graphs stay source-owned.
- `lineNumbering` (object) — Optional canonical line numbering before each text column: { countBy?: integer 1..32767, start?: integer 0..32767, distance?: integer 0..31680, restart?: 'newPage'|'newSection'|'continuous' }. An empty object defaults countBy to 1; start is the zero-based native value, so the first displayed line is start + 1. distance is in twentieths of a point. Use paragraphFormat.suppressLineNumbers for presence-aware paragraph/style suppression. Duplicate leaves, children, unknown values, or extension-bearing w:lnNumType markup stay source-owned.
- `pageNumbering` (object) — Optional canonical section numbering: { start?: integer 0..2147483647, format?: 'decimal'|'upperRoman'|'lowerRoman'|'upperLetter'|'lowerLetter' }. At least one property is required; omitting start continues the prior sequence. This controls PAGE-field presentation but does not add or refresh a field. Chapter numbering, unsupported formats, duplicate leaves, children, or extension-bearing w:pgNumType markup stay source-owned.

**Schema returns:**

- `section` (DocumentSectionBlock) — Appended section break block. inspect reports editable=false when imported mirrorMargins/gutterAtTop mode markup, section-column topology, line-number markup, or page-number markup is not canonical.

#### `document.addTable`

Append a Word-style table with physical cell values, optional logical merge geometry, and fixed-layout width/margin/border/header formatting.

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `gridColumns` (number) — Logical Word table-grid width. Required for explicit authored geometry; otherwise derived from values.
- `cells` (object[]) — One record per physical value cell with zero-based row/column, gridColumn, columnSpan, rowSpan, verticalMerge none/restart/continue, and editability evidence. OpenChestnut can author complete, contiguous, conforming geometry and keeps imported geometry source-bound.
- `name` (string) — Inspectable table name.
- `styleId` (string) — Table style ID.
- `widthDxa` (number) — Table width in twentieths of a point.
- `indentDxa` (number) — Leading table indent in twentieths of a point.
- `columnWidthsDxa` (number[]) — One width per logical table-grid column in twentieths of a point; values must sum to widthDxa.
- `cellMarginsDxa` (object) — Cell margins in twentieths of a point.
- `borderColor` (string) — Table border color.
- `borderSize` (number) — Uniform border width in eighths of a point; zero disables borders.
- `headerFill` (string) — Header-row fill color.

**Schema returns:**

- `table` (DocumentTableBlock) — Appended table block.

#### `document.addTableOfContents`

Append one canonical one-paragraph complex TOC field with bounded heading levels/switches and enable the native updateFields-on-open hint by default. Refreshed cross-paragraph result graphs remain opaque/source-bound and read-only.

**Schema parameters:**

- `levels` (string) — Ascending heading-level range such as 1-3; defaults to 1-3.
- `minLevel` (number) — Minimum level when levels is omitted.
- `maxLevel` (number) — Maximum level when levels is omitted.
- `hyperlinks` (boolean) — Include the canonical \h switch; defaults to true.
- `hidePageNumbersInWeb` (boolean) — Include the canonical \z switch; defaults to true.
- `useOutlineLevels` (boolean) — Include the canonical \u switch; defaults to true.
- `display` (string) — Cached placeholder result shown until a compatible host refreshes the TOC.
- `updateFields` (boolean) — Enable the updateFields-on-open hint; defaults to true.
- `styleId` (string) — Paragraph style ID.

**Schema returns:**

- `field` (DocumentFieldBlock) — Canonical complex TOC placeholder with complex=true.

#### `document.addWatermark`

Add one canonical VML text watermark to a section/header-reference scope. Recognized imported watermarks permit text-only edits or whole-object removal; adding to an imported package, changing scope, shared headers, multiple objects, DrawingML, images, and irregular VML fail closed.

**Schema parameters:**

- `text` (string) required — Nonblank XML-safe watermark text, 1 through 256 characters.
- `id` (string) — Optional object ID. IDs locate this model object; they are not persistent document identity across unrelated imports.
- `referenceType` (string) — default, first, or even header reference scope; defaults to default.
- `sectionIndex` (number) — Zero-based target section; defaults to 0.

**Schema returns:**

- `watermark` (DocumentWatermark) — One canonical VML text watermark. Only one object is allowed per section/reference scope.

#### `document.applyDesignPreset`

Apply a clean-room report or memo design preset that updates named styles for consistent DOCX export and SVG/layout previews.

**Schema parameters:**

- `name` (string) required — report, memo, or a custom preset name.
- `styles` (object) — Style overrides merged into the preset.

**Schema returns:**

- `document` (DocumentModel) — The mutated document facade.

#### `document.contentControls`

List typed mutable handles for recognized inline or table-cell plain-text, checkbox, drop-down, combo-box, and date controls plus block plain-text controls, with explicit placement and model/native identity.

**Schema returns:**

- `controls` (DocumentContentControlHandle[]) — Fresh typed handles for recognized block/inline/table-cell text, checkbox, drop-down, combo-box, and date controls. placement is block, inline, or tableCell; runIndex is present only for inline controls and row/column only for table-cell controls. Tag/alias plus type-specific text, checked, selectedValue, value, or dateValue are mutable; list choices, controlType, nativeId, native date profile, symbol declarations, placement, and topology are source identity.

#### `document.fillContentControls`

Transactionally fill every recognized block, inline, or table-cell plain-text control matching an object or Map of tag-to-string entries. Checkbox, drop-down, combo-box, and date tags do not silently accept text.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string value mapping. Duplicate tags fill every matching control.
- `strict` (boolean) — Unknown tags fail before mutation; defaults to true. Checkbox, drop-down, combo-box, and date tags are never matched by this text primitive.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of document theme and explicit run/style font families.

**Schema returns:**

- `families` (string[]) — Font-family inventory; mutating the returned array does not mutate the document.

#### `document.inspect`

Emit bounded NDJSON for document blocks including typed block/inline plain-text and inline checkbox/list/date content controls with explicit placement, fields, tracked changes, bookmark ranges, footnotes/endnotes, bibliography sources, comments, styles, headers/footers with sourceBound/editable evidence, canonical text watermarks, and layout; narrow with search/target anchors and fields with include/exclude.

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

- `kind` (string) — Comma-separated block/tableCell/comment/watermark/style/textRange/layout kinds; paragraph and table-cell records expose textEditable/textPatchable capability evidence.
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

Return page-aware layout JSON with block bounding boxes, section/page ordinals, effective inherited header/footer selections, styles, and target/search slicing.

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

#### `document.materializeFields`

Transactionally compute canonical inline SEQ counters and REF cached results from native bookmark targets, with dry-run evidence and strict missing-target failure. PAGEREF remains skipped because trustworthy page numbers require a real pagination host.

**Schema parameters:**

- `types` (string|string[]) — SEQ and/or REF cached-result types; defaults to both. PAGEREF is rejected when requested.
- `dryRun` (boolean) — Plan and report every cache change without mutating the document.
- `strict` (boolean) — Reject unresolved or duplicate bookmark targets before any mutation; defaults to true.

**Schema returns:**

- `result` (object) — Structured { dryRun, updated, wouldUpdate, seqFields, refFields, skippedPageReferences, missingBookmarks, changes } result.

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

#### `document.replyToComment`

Add one source-free direct reply to a root comment. OpenChestnut authors the bounded commentsExtended graph; nested replies and imported topology changes fail closed.

**Schema parameters:**

- `parent` (string|DocumentComment) required — Existing parent comment ID or facade.
- `text` (string) required — Reply text.
- `author` (string) — Reply author.
- `initials` (string) — Reply author initials.
- `date` (string) — Optional reply timestamp.
- `resolved` (boolean) — Reply resolution state.
- `durableId` (string) — Optional preserved Office 2019 durable comment identity.
- `dateUtc` (string) — Optional Office 2021 UTC timestamp.
- `person` (object) — Optional providerId/userId presence identity for the reply author.

**Schema returns:**

- `comment` (DocumentComment) — Source-free direct reply authored through commentsExtended. Replies to replies and additions/removals in an imported fixed-topology thread fail closed.

#### `document.resolve`

Resolve stable document, block, table-cell, content-control, bookmark, footnote/endnote, bibliography source ID/tag, header/footer, watermark, comment, style, and advertised text-range IDs.

**Schema parameters:**

- `id` (string) required — Stable document, block, table-cell, header/footer, watermark, comment, style, or advertised text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `document.setCheckboxContentControls`

Transactionally set every recognized canonical checkbox control matching an object or Map of tag-to-boolean entries. Other control types do not silently coerce.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-boolean checked-state mapping. Duplicate tags update every matching checkbox.
- `strict` (boolean) — Unknown checkbox tags fail before mutation; defaults to true.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setComboBoxContentControls`

Transactionally set every recognized canonical combo-box control from a tag-to-value string mapping. Values may select a declared choice or provide bounded custom text; unknown tags and invalid values fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string value mapping. Each value may match one declared choice or be XML-safe custom text of 1 to 255 characters; duplicate tags update every matching combo-box.
- `strict` (boolean) — Unknown combo-box tags fail before mutation; defaults to true. All values are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setDateContentControls`

Transactionally set every recognized canonical date control from a tag-to-YYYY-MM-DD mapping. Invalid Gregorian dates, unknown tags, and other control types fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string date mapping. Every value must be a real Gregorian date in exact YYYY-MM-DD form; duplicate tags update every matching date control.
- `strict` (boolean) — Unknown date tags fail before mutation; defaults to true. All dates are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setDropdownContentControls`

Transactionally set every recognized canonical drop-down control from a tag-to-choice-value string mapping. Unknown tags or values outside the declared choice table fail before mutation.

**Schema parameters:**

- `values` (object|Map) required — Tag-to-string selected-value mapping. Each value must exactly match one declared choice; duplicate tags update every matching drop-down.
- `strict` (boolean) — Unknown drop-down tags fail before mutation; defaults to true. All selected values are validated before any control changes.

**Schema returns:**

- `result` (object) — Structured { updated, matchedTags, missingTags } result.

#### `document.setSectionSettings`

Set per-section Word behavior such as different-first-page header/footer activation without changing preserved header/footer references.

**Schema parameters:**

- `sectionIndex` (number) required — Zero-based section index from 0 through the number of section-break blocks.
- `differentFirstPage` (boolean) — Whether the section activates first-page header/footer references through w:titlePg.

**Schema returns:**

- `document` (DocumentModel) — Document facade with normalized per-section settings.

#### `document.setSettings`

Set model settings. evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields refresh hint, and bounded passwordless documentProtection are inside the OpenChestnut 0.3 DOCX boundary. Irregular page-margin mode markup and password/cryptographic protection variants stay source-owned and fail closed on replacement.

**Schema parameters:**

- `settings` (object) required — Partial settings object. evenAndOddHeaders, mirrorMargins, gutterAtTop, updateFields, and trackRevisions are booleans. documentProtection accepts false/null/off to remove the element, none/readOnly/comments/trackedChanges/forms, or { edit, enforcement, formatting }; password hashes, cryptographic attributes, IRM, permission exceptions, and irregular mirrorMargins/gutterAtTop markup are unsupported/source-owned. Structurally irregular page-margin mode markup also blocks sibling settings edits and makes imported section geometry read-only when exact reserialization cannot be proved.

**Schema returns:**

- `document` (DocumentModel) — Document facade with normalized facing-page/binding-gutter/header/tracking/refresh settings; updateFields is a refresh request, and passwordless documentProtection is an editing restriction rather than encryption or access control.

#### `document.styles.effective`

Resolve a named document style through basedOn inheritance so inspect/layout/render/DOCX export share the same effective style metadata.

**Schema parameters:**

- `styleId` (string) required — Named style ID to resolve through basedOn inheritance.

**Schema returns:**

- `style` (object|undefined) — Resolved effective style or undefined.

#### `document.textRange`

Inspect or resolve stable textRange anchors such as blockId/text and tableId/cell/row/column/text. Assignment is limited to fully editable text; replace() also supports explicitly advertised source-bound literal patches.

**Schema parameters:**

- `id` (string) required — Stable blockId/text or tableId/cell/row/column/text range ID.

**Schema returns:**

- `textRange` (TextRange|undefined) — Advertised text-range facade. Assignment requires textEditable; replace() may instead use the narrower textPatchable contract.

#### `document.verify`

Return QA issues for invalid/duplicate content-control IDs and native IDs, malformed tags/aliases, invalid block-control profiles, fake lists, invalid links/citations/bibliography sources, malformed tracked changes, duplicate/dangling/reversed bookmark ranges, invalid footnotes/endnotes, unknown styles, malformed tables, bad images/sections, invalid watermark IDs/scopes/text, dangling comments, visual overflow, and prose-like table cells.

**Schema parameters:**

- `visualQa` (boolean) — Include modeled layout overflow checks.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Document semantic/layout QA result.

#### `documentComment.reopen`

Clear the resolved state of a bounded modern comment without changing its root/reply topology or durable identity.

**Schema returns:**

- `comment` (DocumentComment) — The same comment facade with resolved=false; root/reply, paragraph, durable, UTC, and people identity remain fixed.

#### `documentComment.resolve`

Set resolved=true for a bounded modern comment. Imported edits re-prove source hashes and commentsExtended topology while keeping thread identity fixed.

**Schema returns:**

- `comment` (DocumentComment) — The same comment facade with resolved=true; imported edits re-prove source hashes and commentsExtended topology.

#### `DocumentFile.addTrackedReplacement`

Add one exact replacement inside a direct body paragraph or bounded table-cell paragraph to hash-bound DOCX source bytes as adjacent native w:del/w:ins runs. A structured paragraph/tableCell selector, full expected text, and one unique literal contained in either one ordinary run or adjacent run fragments with identical w:rPr preserve source formatting; mixed formatting and broader topologies fail closed with exact changed-part audit.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array|ArrayBuffer) required — Original DOCX bytes. OpenChestnut edits this package directly and never rebuilds it from the imported JavaScript model.
- `target` (object) — Preferred structured selector: { kind: 'paragraph', blockIndex } or { kind: 'tableCell', blockIndex, row, column }. Table row/column are zero-based physical indexes from the exact imported table block.
- `targetBlockIndex` (number) — Compatibility selector for a direct body paragraph. Omit when target is supplied; the two forms are mutually exclusive.
- `expectedText` (string) required — Exact full text of the target paragraph or single-paragraph table cell; stale text fails closed before mutation.
- `search` (string) required — Non-empty literal that must occur exactly once. It may occupy one ordinary w:r/w:t or adjacent non-empty ordinary runs only when their exact w:rPr markup is identical; duplicate, empty-run-gap, and mixed-format spans fail closed.
- `replacement` (string) required — Non-empty replacement text written in a native adjacent w:ins run with the source run formatting.
- `author` (string) required — Revision author, 1 through 255 characters without control characters.
- `date` (string) — Optional ISO 8601 revision timestamp.
- `expectedSourceSha256` (string) required — Lowercase 64-hex SHA-256 of the exact input bytes; JavaScript and OpenChestnut both verify it.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Source-preserving DOCX with metadata.trackedReplacement containing the re-proved structured target, source/output and paragraph-element hashes, UTF-16 text hashes/counts, matchedSourceRunCount, package-local native revision IDs, semantic/body indexes, and the exact changed-part list. Only word/document.xml may change.

#### `DocumentFile.exportDocx`

Export DocumentModel to DOCX through the single bundled OpenChestnut codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly.

**Schema parameters:**

- `document` (DocumentModel) required — Document facade to serialize.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — DOCX package bytes.

#### `DocumentFile.finalizeRevisions`

Accept or reject bounded direct whole-paragraph one-run revisions and exact adjacent in-paragraph w:del + w:ins pairs from source bytes, including same-format fragmented deletions in direct body paragraphs or bounded table-cell paragraphs. Mandatory SHA-256 binding, decompression budgets, exact changed-part audit, and fail-closed graph checks prevent silent model reconstruction or broad package mutation.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array|ArrayBuffer) required — Original DOCX bytes. The native codec operates directly on this package rather than rebuilding it from a JavaScript model.
- `mode` (string) required — accept or reject.
- `expectedSourceSha256` (string) required — Lowercase 64-hex SHA-256 of the exact input bytes; JavaScript and OpenChestnut both verify it.
- `keepTracking` (boolean) — Preserve an existing trackRevisions setting after finalization. Defaults to false and never enables a setting that was absent.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Rewritten DOCX with metadata.revisionFinalization containing source/output hashes, insertion/deletion counts, tracking before/after, and exact changed parts. Direct body whole-paragraph one-run revisions plus exact adjacent deletion/insertion pairs in direct body paragraphs or bounded table cells are accepted; the deletion may retain multiple adjacent fragments only when every fragment and the single insertion have identical w:rPr. Mixed-format, nested, moved, property-level, non-body-story, irregular-table, malformed, or absent revisions fail closed.

#### `DocumentFile.importDocx`

Import relationship-driven core DOCX semantics through the single bundled OpenChestnut codec. An imported header/footer advertises editable only for one direct unformatted text paragraph in a uniquely used source part; it is hash-bound, allows at most one text edit per part, and leaves PAGE/simple fields, rich, shared, inherited, and irregular page furniture read-only. Recognized inline controls, fields, revisions, notes, citations, simple tables, and exclusive canonical VML text-watermark paragraphs are fixed-topology editable; otherwise read-only paragraphs and complex table cells separately advertise textPatchable when at least one direct ordinary native text node can participate in a bounded literal patch. A unique literal may span adjacent same-format runs without rebuilding the surrounding graph.

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `document` (DocumentModel) — Imported document facade with editable core blocks, hash-bound direct-text header/footer capability evidence, recognized canonical text watermarks, and source-bound read-only advanced content.

#### `DocumentFile.inspectDocx`

Inspect bounded DOCX parts, content types, relationships, and namespace-aware source XML r:id/r:embed/r:link references under decompression budgets.

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

Apply DOCX part patches with path traversal validation for settings, classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people parts, and numbering assignments; atomically reject dangling packages and invalid comment graphs.

**Examples:**

- await DocumentFile.patchDocx(docx, [{ path: 'customXml/review-note.xml', text: '<review>ok</review>' }])

**Schema parameters:**

- `docx` (FileBlob|Uint8Array) required — DOCX package bytes.
- `patches` (array|object) required — Path-validated package part edits with text/xml/json/bytes/remove.
- `maxPatchBytes` (number) — Per-part patch size limit.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types and relationships atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; DOCX supports settings mutations, section-scoped header/footer references, batch classic-comment anchors, commentsExtended/commentsIds/commentsExtensible/people relationships, and numbering assignments for block, paragraph, or table-cell targets.
- `sourceReference` (boolean|object) — Opt-in semantic XML mutation. Settings accepts trackRevisions/updateFields/evenAndOddHeaders/mirrorMargins/gutterAtTop booleans and passwordless documentProtection; comments accepts { anchors: [...] }; numbering accepts { assignments: [...] }.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `docx` (FileBlob) — Patched DOCX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `DocumentModel.create`

Create a document with paragraph/character styles, formatted paragraphs/runs, canonical inline and one-paragraph table-cell plain-text, checkbox, drop-down, combo-box, and ISO/Gregorian date content controls, one-paragraph block plain-text controls, canonical inline SEQ/REF/PAGEREF fields, sections, headers/footers, canonical VML text watermarks, lists, TableGrid fixed-geometry tables, links, bounded whole-block bookmarks, plain-text footnotes/endnotes, canonical bibliography-backed citations, simple fields, a canonical complex TOC placeholder, bounded whole-paragraph tracked insertions/deletions, classic comments, bounded modern root/direct-reply threads, and PNG/JPEG images. Nested/irregular modern threads, rich comment bodies, multi-paragraph/rich/inline-within-cell/nested/data-bound/locked/placeholder table-cell SDTs, other nested/data-bound/locked/placeholder SDTs, irregular lists, localized dates, custom checkbox symbols, image/DrawingML/irregular VML watermarks, other complex field graphs, arbitrary table-style graphs, complex bookmark/note/revision graphs, and advanced settings remain unsupported or source-bound.

**Schema parameters:**

- `name` (string) — Document name.
- `designPreset` (string) — Initial design preset name.
- `theme` (object) — Word theme name, 12 scheme colors, and major/minor Latin, East-Asian, and complex-script fonts.
- `defaultRunStyle` (object) — Document-wide run properties serialized as w:docDefaults/w:rPrDefault and applied before named styles.
- `styles` (object) — Named paragraph/character styles plus imported table/numbering style records with optional basedOn inheritance and numberingId/numberingLevel linkage. Source-free table blocks may select TableGrid; arbitrary custom table-style graphs are not materialized.
- `paragraphs` (string[]) — Convenience paragraph list; the first paragraph uses Title style.
- `blocks` (object[]) — Ordered paragraph/list/table/link/field/citation/image/section/change block models. Paragraph runs may carry canonical inline SEQ/REF/PAGEREF fields; bibliography-backed citations and one-paragraph complex TOC placeholders also cross OpenChestnut. Other field graphs remain source-bound.
- `bookmarks` (object[]) — Whole-block bookmark ranges. Source-free authoring requires one unique valid Word name around exactly one paragraph-like block; imported bookmarks are fixed-topology/read-only.
- `notes` (object[]) — Plain-text footnote/endnote records. The bounded profile permits one note at the end of each paragraph or list item; imported note text may change, but kind, anchor, native ID, and topology are source-bound.
- `bibliography` (object) — Canonical Word bibliography SelectedStyle, StyleName, and URI metadata authored in one b:Sources Custom XML part.
- `bibliographySources` (object[]) — Bounded Word bibliography sources with ordinary personal or corporate Author data and supported scalar fields. Imported source order, IDs, and tags remain source-bound.
- `headers` (object[]) — Header block models. Imported items expose sourceBound/editable; only a uniquely used one-run direct text paragraph is text-editable, and at most one edit may target each source Header part.
- `footers` (object[]) — Footer block models. Imported items expose sourceBound/editable; only a uniquely used one-run direct text paragraph is text-editable, and at most one edit may target each source Footer part.
- `sectionSettings` (object[]) — Per-section settings with zero-based sectionIndex and differentFirstPage activation state.
- `comments` (object[]) — Classic whole-paragraph comments. Parent/reply, resolved, durable-ID, UTC/person, and modern extension metadata are outside the OpenChestnut 0.2 boundary.
- `settings` (object) — evenAndOddHeaders, mirrorMargins, gutterAtTop, trackRevisions, the updateFields-on-open refresh hint, and bounded passwordless documentProtection are authorable. mirrorMargins toggles facing-page inside/outside margins; gutterAtTop chooses whether each section's gutter is added at the top edge or binding side. Irregular page-margin mode markup stays source-owned and makes section geometry read-only. Password/cryptographic protection variants cannot be replaced through the semantic model.

**Schema returns:**

- `document` (DocumentModel) — Editable document facade.

#### `documentTableCell.addCheckboxContentControl`

Wrap one source-free rectangular table cell in a canonical Word 2010+ checkbox w:sdt. OpenChestnut owns the visible glyph and symbols; recognized imports permit checked/tag/alias edits while identity, type, placement, symbols, and topology remain fixed.

**Schema parameters:**

- `checked` (boolean) — Initial checked state; defaults to false.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell checkbox SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical Word 2010+ checkbox around the cell's existing single paragraph/run. Source-free rectangular cells may add it once. Visible glyph and symbols are codec-owned; recognized imports keep native ID, type, placement, row/column, symbol profile, and topology fixed.

#### `documentTableCell.addComboBoxContentControl`

Wrap one source-free rectangular table cell in a canonical standard combo-box w:sdt with ordered choices and a declared-or-custom typed value. Recognized imports permit value/tag/alias edits while the choice table and topology remain fixed.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `value` (string) — Initial value, 1 to 255 XML-safe characters; defaults to the first choice. A matching choice uses its displayText, while custom text is shown verbatim.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell combo-box SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical standard combo box around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, ordered choice table, and topology fixed.

#### `documentTableCell.addDateContentControl`

Wrap one source-free rectangular table cell in the canonical ISO/Gregorian date w:sdt profile. Recognized imports permit dateValue/tag/alias edits while native date metadata, placement, and topology remain fixed.

**Schema parameters:**

- `dateValue` (string) required — Real proleptic Gregorian date in exact YYYY-MM-DD form, from 0001-01-01 through 9999-12-31. Date objects and locale-formatted strings are rejected.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell date SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical ISO/Gregorian date picker around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, native date profile, and topology fixed.

#### `documentTableCell.addDropdownContentControl`

Wrap one source-free rectangular table cell in a canonical standard drop-down w:sdt with ordered choices and a typed selectedValue. Recognized imports permit selectedValue/tag/alias edits while the choice table and topology remain fixed.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `selectedValue` (string) — Initial internal choice value; defaults to the first choice.
- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell drop-down SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical standard drop-down around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imports keep native ID, type, placement, row/column, ordered choice table, and topology fixed.

#### `documentTableCell.addTextContentControl`

Wrap one source-free rectangular table cell's existing text in a canonical cell-level plain-text w:sdt. The handle reports placement=tableCell plus row/column; recognized imported controls permit fixed-topology text/tag/alias edits, while adding or removing imported control topology fails closed.

**Schema parameters:**

- `id` (string) — Agent-facing content-control ID; generated when omitted.
- `tag` (string) required — Table-cell plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Non-empty human title/alias, at most 255 characters; defaults to tag.

**Schema returns:**

- `control` (DocumentContentControlHandle) — Canonical plain-text control around the cell's existing single paragraph/run. Source-free rectangular cells may add it once; recognized imported controls keep native ID, type, placement, row/column, and topology fixed.

#### `documentTableCell.replaceText`

Apply a literal source-bound text patch to one table cell that advertises textPatchable. The search must resolve exactly once inside one ordinary native w:t node or adjacent non-empty direct runs with byte-identical w:rPr. Whole-cell replacement, mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and ambiguous matches fail closed.

**Schema parameters:**

- `search` (string) required — Non-empty literal text that must occur exactly once in the visible cell. A source-bound match may occupy one ordinary direct w:r/w:t or adjacent non-empty direct runs only when their exact w:rPr markup is identical and it never crosses a paragraph boundary.
- `replacement` (string) required — XML-safe replacement text, up to 1,000,000 characters.

**Schema returns:**

- `cell` (DocumentTableCell) — Mutated table-cell facade with one pending source-bound text patch.

#### `documentWatermark.remove`

Remove one modeled or recognized source-bound canonical watermark as a complete header paragraph. The source-bound operation re-proves exact element and header residual hashes and never heuristically deletes arbitrary header graphics.

**Schema returns:**

- `watermark` (undefined) — Removes the complete recognized watermark paragraph after source/residual revalidation on export.

#### `exportDocxWithOpenChestnut`

Export bounded DocumentModel paragraphs/runs, fields, tables, bookmarks, notes, citations, tracked changes, comments, images, canonical text watermarks, sections, numbering, and settings; recognized imports permit exact-profile semantic edits plus hash-bound literal patches to one unique ordinary paragraph or table-cell span inside one direct w:r/w:t or adjacent same-format runs while preserving all surrounding native markup.

**Schema parameters:**

- `document` (DocumentModel) required — Document facade within the OpenChestnut paragraph/run/style, inline SEQ/REF/PAGEREF field, section, header/footer, canonical text-watermark, image, list, hyperlink, whole-block bookmark, plain-text footnote/endnote, simple-field, comment, and fixed-table boundary. Advanced imported content remains source-bound; unsupported edits fail closed.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — DOCX bytes produced by the bundled Open XML SDK WebAssembly codec, with codec diagnostics in metadata.

#### `importDocxWithOpenChestnut`

Import DOCX bytes through OpenChestnut with source-bound blocks, recognized exclusive canonical VML text-watermark paragraphs, and source-bound header/footer editable evidence. A header/footer edit is limited to one direct unformatted text paragraph in one uniquely used source part; fields, rich/shared/inherited page furniture, scope changes, and multiple edits to one part fail closed. Literal body/table patch capability never implies whole-paragraph/cell editability; only adjacent non-empty direct runs with byte-identical w:rPr may form one patch span, while mixed-format, gapped, cross-paragraph, ambiguous, field/control/revision text remains fail-closed.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — DOCX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `document` (DocumentModel) — Imported document facade carrying source/opaque evidence. Canonical footnote/endnote bodies, exclusive VML text watermarks, and one direct unformatted header/footer paragraph in a uniquely used source part are text-editable with fixed source-bound anchors; whole-block bookmarks are fixed-topology/read-only, and other complex graphs remain source-bound.

#### `paragraph.addCheckboxContentControl`

Append one canonical Word 2010+ checkbox content control with typed checked state; OpenChestnut owns its visible glyph and w14 symbol declarations.

**Schema parameters:**

- `checked` (boolean) — Initial checked state; defaults to false.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Checkbox SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the canonical visible glyph.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical checkbox content-control metadata.

#### `paragraph.addComboBoxContentControl`

Append one canonical inline Word combo-box content control with ordered displayText/value choices and a typed value that may be a declared choice or bounded custom text. OpenChestnut derives the visible projection.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `value` (string) — Initial value, 1 to 255 XML-safe characters; defaults to the first choice. A matching choice uses its displayText, while custom text is shown verbatim.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Combo-box SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the derived visible value.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical combo-box content-control metadata.

#### `paragraph.addDateContentControl`

Append one canonical inline Word date picker from a real Gregorian YYYY-MM-DD value. OpenChestnut owns the fixed ISO display, UTC-midnight fullDate, language, mapping, and calendar projection.

**Schema parameters:**

- `dateValue` (string) required — Real proleptic Gregorian date in exact YYYY-MM-DD form, from 0001-01-01 through 9999-12-31. Date objects and locale-formatted strings are rejected.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Date SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the codec-owned ISO visible date.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical date content-control metadata.

#### `paragraph.addDropdownContentControl`

Append one canonical inline Word drop-down content control with an ordered displayText/value choice table and typed selectedValue. OpenChestnut derives visible text from the selected choice.

**Schema parameters:**

- `choices` (Array<string|object>) required — Ordered 1 to 256 choice table. A string uses the same displayText and value; objects require unique XML-safe displayText and value strings of 1 to 255 characters.
- `selectedValue` (string) — Initial internal choice value; defaults to the first choice.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Drop-down SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting for the derived visible choice text.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded canonical drop-down content-control metadata.

#### `paragraph.addField`

Append a logical inline SEQ, REF, or PAGEREF field run. A SEQ run may add a bookmark around only its cached result for real caption-number targets. OpenChestnut authors/imports the canonical native graph; imported field position, instruction, and bookmark identity remain source-bound while cached display text is editable.

**Schema parameters:**

- `instruction` (string) required — Canonical SEQ <label> \* ARABIC, REF <bookmark> \h, or PAGEREF <bookmark> \h instruction using a bounded Word-compatible name.
- `display` (string) — Cached visible result before host refresh; defaults to 0.
- `bookmarkName` (string) — Optional unique Word bookmark name for a SEQ field; wraps only the cached-result run so REF/PAGEREF can target the caption number.
- `bookmarkNativeId` (number) — Optional unsigned 32-bit native bookmark ID for source-free authoring; imported identity is source-bound.
- `style` (object) — Optional modeled formatting for the cached result run.

**Schema returns:**

- `run` (object) — Logical inline field run. Imported position, instruction, and optional bookmark identity are source-bound; cached display text remains editable.

#### `paragraph.addTextContentControl`

Append one inline plain-text Word content-control run with agent ID, tag, alias, text, and optional run formatting. OpenChestnut assigns native w:id identity and authors canonical w:sdt markup.

**Schema parameters:**

- `text` (string) required — Initial visible control text.
- `id` (string) — Agent-facing model ID; generated when omitted.
- `tag` (string) required — Plain-text SDT tag, 1 to 64 characters without controls.
- `alias` (string) — Human title/alias, at most 255 characters; defaults to tag.
- `style` (object) — Optional modeled run formatting.

**Schema returns:**

- `run` (object) — Appended paragraph run carrying bounded inline plain-text content-control metadata.

#### `paragraph.replaceText`

Replace literal paragraph text without flattening formatting boundaries. Fully editable one-run paragraphs update their existing run; imported source-bound paragraphs advertise textPatchable when OpenChestnut can replace one unique ordinary w:r/w:t node or adjacent non-empty direct runs with byte-identical w:rPr while preserving all native topology and surrounding markup. Mixed formatting, empty-run gaps, paragraph boundaries, fields, controls, revisions, and duplicate matches fail closed.

**Schema parameters:**

- `search` (string) required — Non-empty literal text that must occur exactly once. A source-bound match may occupy one ordinary direct w:r/w:t or adjacent non-empty direct runs only when their exact w:rPr markup is identical.
- `replacement` (string) required — XML-safe replacement text, up to 1,000,000 characters.

**Schema returns:**

- `paragraph` (DocumentParagraphBlock) — Mutated paragraph facade. Source-bound patches are applied only after native-node and source-hash validation during export.

## pdf

| Name | Kind | Summary |
| --- | --- | --- |
| `createPdfjsParser` | api | Create an optional PDF.js parser adapter to extract page geometry, positioned text, heuristic tables, and bounded embedded raster or stencil-mask PNG images with placement boxes. |
| `pdf.addChart` | api | Add a modeled bar/line chart region with categories, series, title, meaningful alternative text or decorative-artifact semantics, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip. |
| `pdf.addFlowText` | api | Wrap long text into positioned lines and automatically append pages when the configured content box is full. |
| `pdf.addImage` | api | Add a modeled PDF image region with dataUrl/URI/prompt metadata, meaningful alternative text or explicit decorative-artifact semantics, and a page-space bounding box. |
| `pdf.addLink` | api | Add a meaningful visible http, https, or mailto link with stable ID, page-space bounding box, tagged Link structure, URI annotation, OBJR association, and explicit reading-order participation. |
| `pdf.addPage` | api | Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, charts, and links. |
| `pdf.addTable` | api | Add a modeled table with cell values, row/column spans, TH/TD roles, scopes, header associations, stable cell IDs, a page-space bounding box, and optional semanticId joining constrained consecutive-page segments into one logical Table. |
| `pdf.addText` | api | Add positioned PDF text with page-space bbox, font metadata, optional semantic H1-H6 heading level or decorative Artifact semantics, inspect/resolve/layout records, and SVG preview rendering. |
| `pdf.extractTables` | api | Extract modeled table values, normalized spanning-cell/header records, and bounding boxes across all pages or a selected page. |
| `pdf.extractText` | api | Extract modeled text across all pages or a selected page. |
| `pdf.inspect` | api | Emit bounded NDJSON for pages, text, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links; narrow with search/target anchors and shape fields with include/exclude. |
| `pdf.layoutJson` | api | Return modeled PDF page layout JSON with page text, positioned text items, explicit/effective reading order, layout regions, normalized table cells/spans/header IDs, images, charts, links, and target/search context slicing. |
| `pdf.page.setReadingOrder` | api | Declare the complete logical reading sequence of a page's body text, positioned text, tables, images, charts, and links by stable ID without changing visual paint order. |
| `pdf.render` | api | Render a modeled PDF page to SVG by default, return page layout JSON with { format: 'layout' }, or use { source: 'pdf', renderer } to feed the exported PDF into Poppler/PDF-capable raster adapters. |
| `pdf.resolve` | api | Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links. |
| `pdf.verify` | api | Return QA issues for invalid H1-H6 nesting, missing/generic Figure alternative text, meaningless/unsafe links, cross-page logical-table continuity, incomplete/duplicate/unknown reading-order targets, empty pages, text extraction sanity, geometry/bounds, invalid images, table semantics, and chart data. |
| `PdfArtifact.create` | api | Create a modeled PDF artifact with pages, text, span-aware accessible table regions, image regions, charts, links, and explicit reading order. |
| `PdfFile.editPdf` | api | Apply bounded direct-original MuPDF.js operations with explicit rewrite or byte-prefix-verified incremental save, object-level signature detection, atomic caller-controlled output, and fail-closed rejection of incremental page grafting/redaction/deletion/source-bound annotation or link operations, signed incremental edits, ambiguous radio export values, rotated-page crop requests, unsafe link destinations, clipped native appearances, and unsupported operations. add_text_annotation, add_text_highlight, and add_link bind one exact source hash plus the inspected mupdfPage bbox/rotation snapshot and accept coordinates in its explicit mupdf-page-space after the current 0/90/180/270-degree rotation; raw MediaBox/CropBox remain unrotated PDF-space facts. Text-note and Highlight records expose provider appearanceBbox evidence, and placement fails before publication if the full native appearance would leave the visible page. add_text_highlight accepts one unique native text-search selection plus optional RGB/review metadata, never caller quads or rectangles. duplicate_page binds the exact source/page snapshot and copies one ordinary non-interactive untagged right-angle page to a 1-based output position in a single-operation rewrite; it rejects unsupported page-bound graphs or projected page/object budget overflow and does not synthesize navigation. delete_annotation, update_annotation, delete_link, and update_link require an inspect-returned source hash, source-bound locator, and snapshot precondition; update_annotation changes only non-empty contents, author, or subject fields of a native Text annotation, while update_link changes only a safe non-empty URL. Native link geometry is never patched. set_page_crop changes only the visible CropBox on an unrotated page, while rotate_page writes an absolute right-angle /Rotate value; neither removes hidden original content. |
| `PdfFile.exportPdf` | api | Export a modeled artifact as a real multi-page tagged PDF 1.7 whose logical structure follows explicit page reading order without changing paint order, emits semantic H1-H6 headings, meaningful Figure /Alt text, Link annotations with OBJR associations, /Artifact marked content, and constrained logical Tables spanning consecutive pages, and preserves language/title, Table/TR/TH/TD hierarchy, optional Unicode TrueType embedding, positioned text, vector charts, and PNG/JPEG images. |
| `PdfFile.importPdf` | api | Reopen package-generated metadata losslessly or lazily use required MuPDF.js for arbitrary PDFs, producing a bounded reconstructed extraction/QA view with text geometry, raster placements and transforms, links, annotations, widgets, and heuristic table candidates; the view is never an edit representation. |
| `PdfFile.inspectPdf` | api | Inspect a path or PDF bytes after a pre-WASM input budget, combining native MuPDF page/object/annotation/widget/link, source SHA-256, source-bound annotation and link locators, native annotation quadrilateral/color facts, raw MediaBox/CropBox facts, and effective normalized page rotation with bounded tagged-PDF, language, reading-order, heading, Figure, Link, Artifact, font, and table-structure evidence. |
| `PdfFile.renderPdf` | api | Render one page from original PDF bytes through runtime-lazy MuPDF.js as PNG or JPEG, enforcing input, page/object, DPI, and preallocation pixel budgets before returning a FileBlob. |
| `PdfProviders.ensure` | api | Install only a previously installable, policy- and catalog-bound capability resolution into the project-private cache, then return a fresh probe. It uses only catalog-pinned release bytes, validates size/hash/archive/receipt boundaries, and never chooses another provider or obtains credentials. A blocked or ready resolution cannot be forced through this API. |
| `PdfProviders.probe` | api | Probe exactly one selected PDF provider under the requested policy without downloading, mutating the cache, importing MuPDF, or trying fallback providers. The result reports ready or blocked runtime evidence together with the pinned pack plan. |
| `PdfProviders.resolve` | api | Resolve one explicit PDF task and selected/default provider against the immutable capability catalog and project policy. It is read-only: no MuPDF initialization, network access, cache mutation, credential acquisition, or automatic provider fallback occurs. The result is ready, installable, or blocked with exact packs, platform, sizes, licenses, runtime prerequisites, consents, and operation limits. |

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

Add a modeled bar/line chart region with categories, series, title, meaningful alternative text or decorative-artifact semantics, bbox, inspect/resolve/layout records, SVG preview, and PDF metadata roundtrip.

**Examples:**

- pdf.addChart({ pageIndex: 0, chartType: 'bar', categories: ['A', 'B'], series: [{ name: 'Score', values: [2, 4] }], bbox: [72, 430, 468, 180] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `chartType` (string) — bar or line.
- `title` (string) — Visible chart title.
- `alt` (string) — Meaningful alternative text describing the chart; required unless decorative is true.
- `decorative` (boolean) — Mark the chart as decorative PDF Artifact content and exclude it from logical reading order.
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

Add a modeled PDF image region with dataUrl/URI/prompt metadata, meaningful alternative text or explicit decorative-artifact semantics, and a page-space bounding box.

**Examples:**

- pdf.addImage({ pageIndex: 0, dataUrl, alt: 'Approval mark', bbox: [430, 60, 64, 64] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `dataUrl` (string) — Embedded PNG or JPEG image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Image generation/extraction prompt metadata.
- `alt` (string) — Meaningful alternative text; required unless decorative is true.
- `decorative` (boolean) — Mark the image as decorative PDF Artifact content and exclude it from logical reading order.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `fit` (string) — contain or cover intent metadata.

**Schema returns:**

- `image` (PdfImage) — Inspectable image facade with stable ID.

#### `pdf.addLink`

Add a meaningful visible http, https, or mailto link with stable ID, page-space bounding box, tagged Link structure, URI annotation, OBJR association, and explicit reading-order participation.

**Examples:**

- pdf.addLink({ pageIndex: 0, text: 'W3C accessibility guidance', url: 'https://www.w3.org/WAI/', bbox: [72, 700, 240, 18] })

**Schema parameters:**

- `pageIndex` (number) — Zero-based target page index.
- `text` (string) required — Visible text that meaningfully describes the destination; generic text and raw URLs fail verification.
- `url` (string) required — Absolute http, https, or mailto destination.
- `bbox` (number[]) — Page-space [left, top, width, height] in points for visible text and the URI annotation.

**Schema returns:**

- `link` (PdfLink) — Inspectable link facade with stable ID.

#### `pdf.addPage`

Append a modeled PDF page with explicit point dimensions and optional text, positioned items, regions, tables, images, charts, and links.

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
- `links` (object[]) — Modeled visible URI links.
- `readingOrder` (string[]|object[]) — Optional complete logical order of all semantic page items as stable IDs or objects with IDs.

**Schema returns:**

- `page` (PdfPage) — Appended editable page facade.

#### `pdf.addTable`

Add a modeled table with cell values, row/column spans, TH/TD roles, scopes, header associations, stable cell IDs, a page-space bounding box, and optional semanticId joining constrained consecutive-page segments into one logical Table.

**Examples:**

- pdf.addTable({ name: 'gates', values: [['Evidence', '', 'Status'], ['Model', 'Native', ''], ['PDF.js', 'Poppler', 'pass']], cells: [{ row: 0, column: 0, columnSpan: 2 }, { row: 0, column: 2, rowSpan: 2 }], bbox: [72, 140, 468, 96] })

**Schema parameters:**

- `name` (string) — Inspectable table name.
- `values` (unknown[][]) required — Rectangular or ragged cell value matrix.
- `semanticId` (string) — Optional logical table identity shared by constrained segments on consecutive pages. A continuation must be first and a non-final segment last in page reading order.
- `cells` (object[]) — Optional zero-based cell overrides with id, row, column, value, rowSpan, columnSpan, TH/TD role, Row/Column/Both scope, and header ID array.
- `bbox` (number[]) — Page-space [left, top, width, height] in points.
- `source` (string) — Optional extraction/source provenance.

**Schema returns:**

- `table` (PdfTable) — Inspectable table facade with stable cell IDs and getCell(row, column).

#### `pdf.addText`

Add positioned PDF text with page-space bbox, font metadata, optional semantic H1-H6 heading level or decorative Artifact semantics, inspect/resolve/layout records, and SVG preview rendering.

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
- `headingLevel` (number) — Optional semantic PDF heading level from 1 through 6; visual styling remains independent.
- `artifact` (boolean) — Mark repeating/decorative text such as running headers and footers as PDF Artifact content and exclude it from reading order. Cannot be combined with headingLevel.

**Schema returns:**

- `textItem` (object) — Positioned text item with stable ID.

#### `pdf.extractTables`

Extract modeled table values, normalized spanning-cell/header records, and bounding boxes across all pages or a selected page.

**Examples:**

- pdf.extractTables({ page: 1 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `tables` (object[]) — Table records with page, ID, name, values, normalized cells, and bbox.

#### `pdf.extractText`

Extract modeled text across all pages or a selected page.

**Examples:**

- pdf.extractText({ page: 2 })

**Schema parameters:**

- `page` (number) — Optional one-based page number.

**Schema returns:**

- `text` (string) — Selected page text or all page text joined with blank lines.

#### `pdf.inspect`

Emit bounded NDJSON for pages, text, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links; narrow with search/target anchors and shape fields with include/exclude.

**Schema parameters:**

- `kind` (string) — Comma-separated page, text, textItem, readingOrder, region, table, tableCell, image, chart, and link record kinds.
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

Return modeled PDF page layout JSON with page text, positioned text items, explicit/effective reading order, layout regions, normalized table cells/spans/header IDs, images, charts, links, and target/search context slicing.

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

#### `pdf.page.setReadingOrder`

Declare the complete logical reading sequence of a page's body text, positioned text, tables, images, charts, and links by stable ID without changing visual paint order.

**Examples:**

- page.setReadingOrder([`${page.id}/text`, image.id, heading.id, table.id, chart.id])

**Schema parameters:**

- `order` (string[]|object[]) required — Complete page sequence containing each semantic body-text, positioned-text, table, image, chart, and link target exactly once; artifact text and decorative figures are excluded.

**Schema returns:**

- `page` (PdfPage) — The same editable page facade for chaining.

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

Resolve stable PDF artifact IDs for pages, page text blocks, positioned text items, reading-order entries, layout regions, tables/table cells, images, charts, and links.

**Examples:**

- pdf.resolve('pg-1/txt/1')

**Schema parameters:**

- `id` (string) required — Stable artifact, page, text, text-item, reading-order, region, table, table-cell, image, chart, or link ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `pdf.verify`

Return QA issues for invalid H1-H6 nesting, missing/generic Figure alternative text, meaningless/unsafe links, cross-page logical-table continuity, incomplete/duplicate/unknown reading-order targets, empty pages, text extraction sanity, geometry/bounds, invalid images, table semantics, and chart data.

**Examples:**

- pdf.verify({ maxChars: 12000 })

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — PDF semantic QA result with ok, issues, ndjson, and truncated.

#### `PdfArtifact.create`

Create a modeled PDF artifact with pages, text, span-aware accessible table regions, image regions, charts, links, and explicit reading order.

**Examples:**

- const pdf = PdfArtifact.create({ pages: [{ width: 612, height: 792, text: 'Report' }] })

**Schema parameters:**

- `id` (string) — Optional stable artifact ID.
- `metadata` (object) — Clean-room metadata preserved through generated-PDF roundtrip.
- `text` (string) — Convenience text for a single default page.
- `pages` (object[]) — Page models with width, height, text, textItems, regions, tables, images, charts, links, and optional complete readingOrder ID arrays.

**Schema returns:**

- `pdf` (PdfArtifact) — Editable modeled PDF artifact.

#### `PdfFile.editPdf`

Apply bounded direct-original MuPDF.js operations with explicit rewrite or byte-prefix-verified incremental save, object-level signature detection, atomic caller-controlled output, and fail-closed rejection of incremental page grafting/redaction/deletion/source-bound annotation or link operations, signed incremental edits, ambiguous radio export values, rotated-page crop requests, unsafe link destinations, clipped native appearances, and unsupported operations. add_text_annotation, add_text_highlight, and add_link bind one exact source hash plus the inspected mupdfPage bbox/rotation snapshot and accept coordinates in its explicit mupdf-page-space after the current 0/90/180/270-degree rotation; raw MediaBox/CropBox remain unrotated PDF-space facts. Text-note and Highlight records expose provider appearanceBbox evidence, and placement fails before publication if the full native appearance would leave the visible page. add_text_highlight accepts one unique native text-search selection plus optional RGB/review metadata, never caller quads or rectangles. duplicate_page binds the exact source/page snapshot and copies one ordinary non-interactive untagged right-angle page to a 1-based output position in a single-operation rewrite; it rejects unsupported page-bound graphs or projected page/object budget overflow and does not synthesize navigation. delete_annotation, update_annotation, delete_link, and update_link require an inspect-returned source hash, source-bound locator, and snapshot precondition; update_annotation changes only non-empty contents, author, or subject fields of a native Text annotation, while update_link changes only a safe non-empty URL. Native link geometry is never patched. set_page_crop changes only the visible CropBox on an unrotated page, while rotate_page writes an absolute right-angle /Rotate value; neither removes hidden original content.

**Examples:**

- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_text_annotation', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, point: [72, 72], contents: 'Review' }] })
- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 1); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'add_text_highlight', page: 1, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, text: 'Review target', color: [1, 1, 0] }] })
- const inspection = await PdfFile.inspectPdf(pdf); const page = inspection.records.find((record) => record.kind === 'mupdfPage' && record.page === 2); await PdfFile.editPdf(pdf, { savePolicy: 'rewrite', operations: [{ type: 'duplicate_page', page: 2, sourceSha256: inspection.summary.sourceSha256, expectedPage: { bbox: page.bbox, rotation: page.rotation }, insertAt: 4 }] })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — Original PDF path or bytes.
- `operations` (object[]) required — Typed MuPDF operations: source-bound add_text_annotation/add_text_highlight, legacy fill_form, source-bound update_form_field, delete_page, source-bound duplicate_page, delete_annotation, update_annotation, rearrange_pages, set_page_crop, rotate_page, set_metadata, delete_embedded_file, add_link, delete_link, update_link, redact_text, or redact_rect. Native placement operations require the exact inspect-returned sourceSha256 plus expectedPage={bbox,rotation}. The inspected bbox has coordinateSpace=mupdf-page-space: upper-left origin, y downward, with the current 0/90/180/270-degree page rotation already applied; raw MediaBox/CropBox remain unrotated PDF-space facts. add_text_annotation accepts a [x,y] pin and non-empty contents with optional non-empty author/subject, verifies exactly one native Text annotation, and returns its provider-normalized rect plus a conservative appearanceBbox. A requested bbox, text alias, icon selection, stale evidence, an appearance that could clip outside the visible box, or incremental save is rejected. add_text_highlight takes one non-empty <=4096-character text string that native search finds exactly once on the visible page, optional RGB [red,green,blue] components in [0,1], and optional non-empty contents/author/subject. It verifies one native Highlight plus quadrilaterals/color/appearanceBbox and rejects caller quads/rectangles, zero/multiple matches, stale evidence, an out-of-window appearance, and incremental save. add_link accepts an in-page-space [x,y,width,height], a non-duplicate target, and a safe internal #... or absolute http/https/mailto URL. These three operations support right-angle rotated pages, report coordinateSpace/pageRotation, require rewrite, and require fresh output inspection. duplicate_page takes the same exact source/page snapshot plus an optional 1-based insertAt, defaults to immediately after the source page, and copies one right-angle page only when the PDF is untagged and the page has no annotations, links, widgets/forms, page actions, associated files, article beads, transitions, or template steps. It must be the only operation in a rewrite, adds no navigation, and requires fresh inspect/render evidence. update_form_field requires the exact source hash, one mupdf-form-field-<xref> locator, and its full field snapshot; it supports exactly one non-password text widget, one combo field whose display/export options are identical, or one checkbox, then verifies the saved in-memory field state. Shared-widget groups, radio/list/multi-select fields, mismatched choice exports, stale snapshots, and unsafe values fail closed and route to pypdf. delete_annotation and update_annotation each require one inspect-returned sourceSha256, one source-bound mupdf-annotation-<page>-<xref> locator, and a matching expected snapshot. update_annotation supports only native Text annotations and accepts only non-empty contents, author, or subject patch fields; rectangle is an expected snapshot guard, never mutable geometry. delete_link and update_link require sourceSha256, a source-bound mupdf-link-<page>-<fingerprint> locator, and matching expected url/bbox/external facts. update_link accepts only one safe non-empty URL patch field; link bounds are snapshot evidence, never mutable geometry. set_page_crop remains unrotated-only and accepts a raw unrotated PDF-space bbox [x,y,width,height] fully inside the inspected MediaBox; it changes only CropBox and is never redaction. rotate_page accepts an absolute 0, 90, 180, or 270 degree /Rotate value; it does not transform or remove content.
- `savePolicy` (string) — rewrite or incremental. Incremental is forbidden for page duplication, redaction, source-bound annotation/link creation or mutation (including add_text_highlight), delete operations, and signed input; source-bound single-widget form-field update, set_page_crop, and rotate_page are bounded unsigned operations that may be byte-prefix-verified incremental.
- `allowSigned` (boolean) — Acknowledge signed input after external review; never bypasses the incremental prohibition.
- `invalidateSignatures` (boolean) — Required with allowSigned for a deliberate signed-PDF rewrite.
- `password` (string) — Password for an encrypted PDF.
- `limits` (object) — Input/page/object budgets.

**Schema returns:**

- `blob` (FileBlob) — Edited PDF bytes with provider, save policy, signature state, byte counts, and applied-operation evidence.

#### `PdfFile.exportPdf`

Export a modeled artifact as a real multi-page tagged PDF 1.7 whose logical structure follows explicit page reading order without changing paint order, emits semantic H1-H6 headings, meaningful Figure /Alt text, Link annotations with OBJR associations, /Artifact marked content, and constrained logical Tables spanning consecutive pages, and preserves language/title, Table/TR/TH/TD hierarchy, optional Unicode TrueType embedding, positioned text, vector charts, and PNG/JPEG images.

**Examples:**

- const blob = await PdfFile.exportPdf(pdf, { language: 'en-US', title: 'Accessible report' })

**Schema parameters:**

- `pdf` (PdfArtifact) required — Modeled PDF artifact to serialize.
- `tagged` (boolean) — Emit StructTreeRoot/ParentTree/MCID tagging; defaults to true.
- `language` (string) — Catalog language; defaults to artifact metadata language or en-US.
- `title` (string) — Document Info title; defaults to artifact metadata title or first text line.
- `font` (string|FileBlob|Uint8Array|ArrayBuffer|object) — Optional standalone glyf-based TrueType .ttf source for Unicode Type0/CIDFontType2 embedding; accepts a path, bytes, FileBlob, or {path|bytes|base64}.
- `maxFontBytes` (number) — Maximum accepted embedded font input size; defaults to 16 MiB.
- `subsetFont` (boolean) — Subset the embedded TrueType font to used glyphs and composite dependencies; defaults to true. Set false only for diagnostics/interoperability comparison.

**Schema returns:**

- `blob` (FileBlob) — application/pdf bytes with modeled content, clean-room metadata, and tagged-export metadata.

#### `PdfFile.importPdf`

Reopen package-generated metadata losslessly or lazily use required MuPDF.js for arbitrary PDFs, producing a bounded reconstructed extraction/QA view with text geometry, raster placements and transforms, links, annotations, widgets, and heuristic table candidates; the view is never an edit representation.

**Examples:**

- await PdfFile.importPdf('third-party.pdf', { limits: { maxBytes: 64 * 1024 * 1024 }, includeImages: true })
- await PdfFile.importPdf(blob, { parser: createPdfjsParser(), preferParser: true })

**Schema parameters:**

- `blob` (string|FileBlob|Uint8Array|ArrayBuffer) required — PDF path or input bytes. Paths and Blob-like inputs are size-checked before materialization.
- `parser` (function) — Optional parser adapter returning pages/textItems/tables/images.
- `preferParser` (boolean) — Use parser even if clean-room metadata is embedded.
- `parserName` (string) — Name recorded in artifact metadata.
- `password` (string) — Password for an encrypted PDF.
- `includeImages` (boolean) — Extract bounded raster placements; defaults to true.
- `limits` (object) — maxBytes, maxPages, maxObjects, maxImages, maxImagePixels, maxTotalImagePixels, and maxTotalImageBytes budgets.

**Schema returns:**

- `pdf` (PdfArtifact) — Modeled PDF artifact with inspect/resolve/render/verify APIs.

#### `PdfFile.inspectPdf`

Inspect a path or PDF bytes after a pre-WASM input budget, combining native MuPDF page/object/annotation/widget/link, source SHA-256, source-bound annotation and link locators, native annotation quadrilateral/color facts, raw MediaBox/CropBox facts, and effective normalized page rotation with bounded tagged-PDF, language, reading-order, heading, Figure, Link, Artifact, font, and table-structure evidence.

**Examples:**

- await PdfFile.inspectPdf(pdf, { maxObjects: 200, maxChars: 12000 })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — PDF path or bytes.
- `limits` (object) — Input, page, object, annotation/widget, and link budgets applied before or during native inspection.
- `maxObjects` (number) — Maximum indirect object records to inspect.
- `maxLinks` (number) — Maximum native link records to inspect or reconstruct.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — PDF file summary with sourceSha256, tagged/language/structure evidence, bounded indirect object records, source-bound mupdfAnnotation/mupdfLink/mupdfWidget records, and grouped mupdfFormField snapshots. Native Text-note and Highlight annotations expose a provider appearanceBbox; Highlights also expose their quadrilateral selection and RGB color. Native page records include raw unrotated PDF-space MediaBox/CropBox [x,y,width,height] values, normalized right-angle rotation, and an effective visible bbox with coordinateSpace=mupdf-page-space after that rotation. Re-inspect after any rewrite because locators cannot identify a later byte sequence.

#### `PdfFile.renderPdf`

Render one page from original PDF bytes through runtime-lazy MuPDF.js as PNG or JPEG, enforcing input, page/object, DPI, and preallocation pixel budgets before returning a FileBlob.

**Examples:**

- await PdfFile.renderPdf(pdf, { page: 1, dpi: 144, format: 'png' })

**Schema parameters:**

- `pdf` (string|FileBlob|Uint8Array|ArrayBuffer) required — Original PDF path or bytes.
- `page` (number) — One-based page number; defaults to 1.
- `dpi` (number) — Resolution greater than 0 and no more than 1200; defaults to 144.
- `format` (string) — png or jpeg.
- `quality` (number) — JPEG quality from 1 through 100.
- `password` (string) — Password for an encrypted PDF.
- `limits` (object) — Input/page/object and maxRenderPixels budgets.

**Schema returns:**

- `blob` (FileBlob) — Native PNG or JPEG page bytes with provider, page, DPI, and dimensions metadata.

#### `PdfProviders.ensure`

Install only a previously installable, policy- and catalog-bound capability resolution into the project-private cache, then return a fresh probe. It uses only catalog-pinned release bytes, validates size/hash/archive/receipt boundaries, and never chooses another provider or obtains credentials. A blocked or ready resolution cannot be forced through this API.

**Examples:**

- const installed = await PdfProviders.ensure({ resolution, policyPath: '.open-office-artifact-tool/pdf-providers.json' })

**Schema parameters:**

- `resolution` (object) required — The exact current-package resolution returned with status=installable. Its catalog digest, policy fingerprint, provider, and pack plan must still match.
- `policyPath` (string) — The same persistent project policy file used to resolve. It is re-read before any cache mutation.

**Schema returns:**

- `result` (object) — Fresh ready/blocked provider probe plus verified installation receipts. It can only use pinned catalog assets, a bounded project-private cache, hash/size checks, safe extraction, and atomic publication; it never downloads credentials or falls back.

#### `PdfProviders.probe`

Probe exactly one selected PDF provider under the requested policy without downloading, mutating the cache, importing MuPDF, or trying fallback providers. The result reports ready or blocked runtime evidence together with the pinned pack plan.

**Examples:**

- const state = await PdfProviders.probe({ provider: 'qpdf', task: 'repair', policyPath: '.open-office-artifact-tool/pdf-providers.json' })

**Schema parameters:**

- `provider` (string) required — One exact catalog provider ID; probing does not search for alternates.
- `task` (string) — Optional catalog task used to include OCR language-pack requirements in the plan.
- `policyPath` (string) — Explicit project policy file; default-missing remains disabled.
- `languages` (string[]) — Explicit OCR language list, required for an OCR task and checked against policy plus catalogued language packs.

**Schema returns:**

- `state` (object) — Ready or blocked status with local/system/managed runtime evidence and the selected pack plan. It performs no network request, cache write, MuPDF import, or provider fallback.

#### `PdfProviders.resolve`

Resolve one explicit PDF task and selected/default provider against the immutable capability catalog and project policy. It is read-only: no MuPDF initialization, network access, cache mutation, credential acquisition, or automatic provider fallback occurs. The result is ready, installable, or blocked with exact packs, platform, sizes, licenses, runtime prerequisites, consents, and operation limits.

**Examples:**

- const resolution = await PdfProviders.resolve({ task: 'repair', provider: 'qpdf', savePolicy: 'rewrite', mutationAuthorized: true, invalidateSignaturesAuthorized: true, policyPath: '.open-office-artifact-tool/pdf-providers.json' })

**Schema parameters:**

- `task` (string) required — One catalog task such as inspect, repair, ocr, sign, sanitize, or validate-conformance.
- `provider` (string) — Optional provider ID. A task default is a declared preference only; the resolver never substitutes a different provider when it is unavailable.
- `inspection` (object) — Exact-source inspection/preflight evidence. Required for every existing-PDF task except inspect; it must carry a 64-hex sourceSha256 at inspection.summary.sourceSha256 or sourceSha256. A failed MuPDF parse may use a bounded preflight hash record only to route explicit repair.
- `savePolicy` (string) required — One strategy allowed by the selected task, such as read-only, rewrite, incremental, or sanitize.
- `policyPath` (string) — Explicit project policy file. The conventional path is .open-office-artifact-tool/pdf-providers.json; a missing conventional file means disabled, never implicit authorization.
- `languages` (string[]) — Explicit OCR languages. eng and chi_sim are policy defaults; every language must be policy-authorized and catalogued.
- `mutationAuthorized` (boolean) — Required true for a task that mutates source PDF bytes.
- `invalidateSignaturesAuthorized` (boolean) — Required true for a task whose operation can invalidate signatures.
- `credentials` (string[]) — Caller-declared credential kinds such as local-pkcs12. Credentials, private keys, HSMs, remote-signing access, and TSA/LTV access are never installed or acquired.

**Schema returns:**

- `resolution` (object) — Read-only ready, installable, or blocked resolution with one provider, catalog digest, policy fingerprint, no-fallback guarantee, precise pack/platform/download/unpacked/license/runtime plan, and required consents.

## presentation

| Name | Kind | Summary |
| --- | --- | --- |
| `compose.column` | api | Create a vertical compose container. Use width/height fill, hug, or fixed pixels; gap and padding are in pixels. |
| `compose.paragraph` | api | Create an editable text block with name, className/style text tokens, and stable inspect output. |
| `compose.text` | api | Create the same editable paragraph node through the reference-template-compatible children-first text(children, props) helper. |
| `exportPptxWithOpenChestnut` | api | Export bounded direct slide backgrounds, textbox/rectangle/roundRect/ellipse shapes, rich text and lists, basic fills/lines/shadows, straight/elbow connectors and arrows, embedded pictures with native crop/contain/cover semantics, fixed-grid plain-text tables, recursive native p:grpSp trees, relationship-free rich speaker notes, legacy annotations, Office 2021 modern root/direct-reply threads, source-free bar/line/pie charts, the bounded literal clustered bar+line combo profile with either shared primary axes or a canonical secondary line pair, validated payload-only replacement for eligible imported OLE workbooks, and bounded source-bound text updates for canonical SmartArt document nodes. Recognized imported modern threads allow only existing text/status edits; their identity, author/date metadata, anchor/range, position, topology, relationships, and source hashes remain fixed. Inherited or complex graphs remain preserved and fail closed on unsupported mutation. |
| `importPptxWithOpenChestnut` | api | Import PPTX bytes with editable bounded direct slide backgrounds, shapes, rich text, recognized owner-local SlidePart placeholder text, rectangular pictures and native source rectangles, tables, connectors, recursive canonical p:grpSp groups, bar/line/pie charts, the canonical literal clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and a re-proven addable capability for eligible notes-absent slides, unchanged-only legacy comments, fixed-topology modern comment text/status edits, defensive payload access for eligible OLE workbooks, and a source-bound SmartArt plain-node text capability only for the canonical closed four-part one-paragraph/one-run DiagramDataPart profile. Reactions/task fields, nested replies/anchors, connected comment parts, inherited Master/Layout graphs, complex backgrounds/blips/groups, field/link/picture-bullet/layout-bearing notes, irregular combos, ambiguous OLE graphs, rich/multi-run/connected SmartArt, and other unsupported content remain source-bound and read-only. |
| `nativeObject.getEmbeddedWorkbook` | api | Read a defensive FileBlob copy of the XLSX payload from an eligible source-bound top-level OLE object without exposing arbitrary native-part mutation. |
| `nativeObject.replaceEmbeddedWorkbook` | api | Replace only the XLSX payload of an eligible imported top-level OLE object. OpenChestnut validates the new workbook and immutable source binding, preserves the OLE shell, relationships, preview, and all other native parts, and fails closed for malformed or ambiguous graphs. |
| `nativeObject.setDiagramNodeText` | api | Replace text in one existing source-bound SmartArt document node only when its top-level four-part graph and DiagramDataPart prove the canonical one-paragraph/one-run plain-text profile. Node IDs, graph topology, layout/style/colors, geometry, and every non-data part remain fixed; unsupported diagrams reject without fallback. |
| `nativeObject.setName` | api | Native OLE, SmartArt/diagram, contentPart, and media objects imported through OpenChestnut are source-bound and read-only for names; setName rejects instead of mutating the preserved package graph. A separate bounded SmartArt plain-node text capability is exposed only as nativeObject.setDiagramNodeText. |
| `nativeObject.setPosition` | api | Native OLE, SmartArt/diagram, contentPart, and media objects imported through OpenChestnut are source-bound and read-only; setPosition rejects instead of rewriting their geometry or payload graph. |
| `Presentation.create` | api | Create a deck model whose canonical OpenChestnut export supports ordinary slides, bounded direct fade/push transitions, direct solid/style-reference slide backgrounds, shapes, rich text, tables, images, connectors, recursive native p:grpSp groups, plain-text speaker notes, native custom shows with canonical run links, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and a bounded literal clustered bar+line combo profile. Combo bars stay on the primary pair; all lines share either that pair or the canonical secondary top/right pair. Formula/external chart data, custom themes, Master/Layout authoring, comments, custom-show topology mutation, advanced plot geometry, mixed line groups, secondary bars, irregular combo graphs, and other package-level features remain outside the source-free PPTX boundary. |
| `presentation.customShows.add` | api | Define an ordered native p:custShowLst playback route for source-free OpenChestnut export. Text runs may target a show by exact name with optional returnToSlide. Canonical imported shows may change only their name and ordered retained-slide membership; fixed native identity keeps existing run links bound across a rename, while irregular graphs stay opaque. |
| `presentation.customShows.getItem` | api | Resolve a source-free or canonical imported custom show by zero-based index, stable facade ID, or exact name. |
| `presentation.export` | api | Export a slide SVG preview, deck SVG montage via { format: 'montage' }, or target/search-sliced layout JSON. |
| `presentation.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of explicitly used presentation text and bullet font families. |
| `presentation.inspect` | api | Emit NDJSON for deck, custom shows, PowerPoint sections, slides, direct slide transitions, textboxes, shapes, grouped shapes, tables, charts, images, and native contentPart/OLE/diagram/media objects with bounded editability, relationship-reference, root-relationship, preserved-part, and eligible embedded-workbook summaries; narrow with search/target anchors and shape fields with include/exclude. |
| `presentation.layout.clearBackground` | api | Clear a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed. |
| `presentation.layout.placeholders.add` | api | Append a direct-frame title/body/ctrTitle/subTitle text placeholder to a source-free layout. It becomes a native p:ph and must be materialized on each slide through applyLayout/setLayout; object/media/chart/table placeholders remain source-bound. |
| `presentation.layout.placeholders.summary` | api | Return a defensive layout-placeholder discovery snapshot with stable IDs, names, native types/indexes, required flags, and direct-frame presence/geometry; editing the snapshot cannot mutate the model. |
| `presentation.layout.setBackground` | api | Set a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed. |
| `presentation.layouts.add` | api | Create one bounded source-free layout under the canonical master. Use blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only. |
| `presentation.layouts.getById` | api | Resolve a layout by its stable ID without falling back to a same-named or same-typed layout. |
| `presentation.master` | api | Access the one canonical source-free Slide Master. It may author a direct background, bounded text styles, and direct-frame title/body/ctrTitle/subTitle placeholders; imported Master graphs remain source-bound and read-only. |
| `presentation.master.clearBackground` | api | Clear the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed. |
| `presentation.master.setBackground` | api | Set the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed. |
| `presentation.master.setTheme` | api | Set a model-level master theme override for preview only. Canonical PPTX export rejects that source-free override; imported-master mutation remains source-bound and fails closed. |
| `presentation.masters.add` | api | Append a model-level Slide Master. Source-free PPTX authoring requires exactly one master, so use Presentation.create({ master }) or presentation.master for the canonical profile; multiple masters and imported-master edits fail closed. |
| `presentation.masters.getItem` | api | Resolve a model-level or imported Slide Master by stable ID or name. |
| `presentation.resolve` | api | Map stable inspect anchor IDs back to facade objects, including custom shows, PowerPoint sections, and slide transitions; imported advanced package objects may be read-only. |
| `presentation.sections.add` | api | Define a native PowerPoint p14:sectionLst entry for source-free OpenChestnut export. Sections together must form the complete ordered slide partition. Canonical imported sections may change only existing names and contiguous boundaries while count, order, stable facade identity, and native GUID stay fixed; irregular graphs remain opaque. |
| `presentation.sections.getItem` | api | Resolve a source-free or canonical imported PowerPoint section by zero-based index, stable facade ID, or exact name. |
| `presentation.slides.add` | api | Append an editable core slide with an optional bounded source-free layout, direct fade/push transition, solid/style-reference background, and plain-text speaker notes. A supplied layout is resolved and materialized transactionally; effective imported Layout/Master inheritance is never flattened. |
| `presentation.slides.insert` | api | Insert a source-free slide after an existing Slide or 0-based index, or at the beginning with after: null. It uses the same transactional layout materialization, bounded direct fade/push transition, and notes/background profile as slides.add; imported additions fail closed, while slide.duplicate and slide.delete each have their own narrow source-preserving OPC profiles. |
| `presentation.slideSize` | api | Read or set the deck canvas in pixels. On a trusted imported PPTX, a changed size is a deliberately canvas-only source-bound operation: OpenChestnut updates only ppt/presentation.xml p:sldSz, clears an old preset type, and leaves slide, layout, master, chart, and shape coordinates unchanged. It never silently rescales or reflows content; callers must make any layout edits explicitly. |
| `presentation.textRange` | api | Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames. |
| `presentation.theme` | api | Inspect the model theme and theme inheritance. Custom source-free themes are not authored by OpenChestnut 0.2, and imported themes are source-bound and read-only. |
| `presentation.validateLayout` | api | Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow. |
| `presentation.verify` | api | Return QA issues for layout validation, missing master/layout references, placeholder fidelity, chart/data consistency, table shape, image data, and dangling comments. |
| `presentation.view` | api | Control local editor gridline/guide visibility and inspect imported PowerPoint grid spacing, snap settings, and read-only slide guides. Visibility is local model state; imported viewProps.xml metadata remains source/hash-bound and unchanged by canonical export. |
| `PresentationFile.exportPptx` | api | Serialize PPTX through the single bundled OpenChestnut codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly. |
| `PresentationFile.importPptx` | api | Import PPTX through the single bundled OpenChestnut codec with source-bound opaque preservation, speaker-notes edit/add capability evidence, bounded text-only edits for recognized local SlidePart placeholders and canonical SmartArt plain document nodes, eligible OLE workbook payload access/replacement, and fail-closed unsupported edits. |
| `PresentationFile.inspectPptx` | api | Inspect bounded PPTX parts, content types, relationships, namespace-aware source XML references, legacy notes/comments evidence, and Office 2021 modern author/thread/anchor semantics under decompression budgets. |
| `PresentationFile.patchPptx` | api | Apply path-validated PPTX part patches, including safe slide/master/layout ID lists and slide image/chart DrawingML mutations, and atomically reject dangling package references or invalid notes/comments semantics. |
| `shape.text.set` | api | Set plain or structured text with ordered text, field, and line-break inlines; bounded run formatting; character, picture-bullet, or auto-numbered lists; levels, indents, spacing; and external URI, internal-slide, relative-action, or existing custom-show hyperlinks. Missing, opaque, malformed, relationship-bearing, or dangling custom-show targets and unmodeled text graphs fail closed in canonical PPTX export. |
| `shape.useBackgroundFill` | api | Read the presence-aware imported PresentationML p:sp useBgFill flag. It affects preview paint but remains source-bound and read-only; source-free authoring or wire mutation fails closed. |
| `slide.addNotes` | api | Set speaker notes as text or relationship-free paragraph/run data for inspect, preview, and canonical PPTX output. OpenChestnut authors source-free notes, preserves the legacy text-only edit path, and edits a fixed imported rich paragraph/run topology; fields, hyperlinks, picture bullets, notes-body list styles/layout, and unsafe NotesMaster graphs remain source-bound and fail closed. |
| `slide.applyLayout` | api | Bind a slide to a bounded source-free layout and materialize its effective direct-frame placeholder shapes. Applying the same layout is idempotent; switching a materialized layout fails closed. The resulting p:ph identities and direct frames export natively; imported Layout relationships remain preservation-only. |
| `slide.autoLayout` | api | Place existing shapes inside a frame using horizontal or vertical flow, gap, padding, and alignment options. |
| `slide.charts.add` | api | Add a source-free literal bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or clustered bar+line combo chart. Category families use shared literal categories; scatter and bubble use aligned per-series numeric X/Y values, with positive area-based bubble sizes. Supported variants retain title, legend, bounded axes, basic series styling, chart-level data labels, layout JSON, SVG preview, and native ChartPart output across import/edit/re-export. Formula/external data, advanced family geometry, topology changes, and unsupported styling fail closed rather than being flattened. |
| `slide.clearBackground` | api | Remove the direct slide background so preview and PPTX output inherit from the preserved Layout/Master chain. Unsupported imported background graphs fail closed rather than being flattened or discarded. |
| `slide.clearTransition` | api | Remove one canonical direct imported or source-free slide transition. A transition-absent imported slide remains a no-op until an explicit capability-approved add; timing, sound, extension, and opaque-effect graphs remain byte-preserved and reject mutation. |
| `slide.comments.addThread` | api | Create either a bounded legacy PPTX annotation or an Office 2021 modern thread. A comment-free imported presentation may add canonical legacy review comments only when comments.capability.addable is true; existing legacy records remain source-bound and read-only. Modern mode supports a top-level element/text-range/textMatch anchor, one root, direct replies, independent people/timestamps, and active/resolved/closed state; imported modern graphs permit only fixed-topology text/status edits. |
| `slide.comments.capability` | api | Inspect defensive source-bound comment-family evidence before authoring. A comment-free imported presentation may advertise legacy addability; existing legacy records remain read-only and modern graphs retain their separate fixed-topology edit contract. |
| `slide.compose` | api | Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph/text, shape, table, chart, image, and rule nodes into editable slide objects. |
| `slide.connectors.add` | api | Add an inspectable connector line between points or element IDs with SVG preview, layout JSON, PPTX p:cxnSp export, and off-canvas QA. |
| `slide.delete` | api | Remove this slide. Source-free decks may remove any non-final slide. An imported PPTX performs a real OPC deletion only for an isolated slide with exactly its layout relationship and no inbound/package-identity references; media, notes, comments, charts, OLE, hyperlinks, custom shows, sections, extensions, and all clone requests fail closed. |
| `slide.duplicate` | api | Clone one original imported PPTX slide only when its unchanged graph contains canonical shapes, canonical inline fixed-grid tables with bounded rectangular merges, recognized closed literal-data charts, eligible top-level embedded-XLSX OLE frames, canonical top-level four-part SmartArt frames, canonical top-level closed InkML content parts, canonical top-level embedded-MP4 media pictures, embedded rectangular images, bounded canonical straight/elbow connectors, and recursively canonical groups containing only the non-native-graph leaf kinds, exactly one layout relationship, picture-bound image relationships, canonical run-level external/internal/relative-action links plus relationship-free custom-show links bound to an existing stable native show ID, and optional closed NotesSlide-to-NotesMaster/back-to-slide plus bounded legacy-comments leaves. Relationship-backed links keep exact IDs and targets; custom-show actions add no relationship and the clone is never inserted into show membership. Every accepted chart frame uniquely consumes one internal relationship to a numbered ChartPart whose child, external, hyperlink, and data relationship sets are empty. Every accepted OLE frame uniquely consumes one internal package relationship to a closed, uniquely inbound XLSX EmbeddedPackagePart and one internal preview ImagePart relationship. Every accepted SmartArt frame owns exactly one internal dm/lo/qs/cs relationship set to closed relationship-free diagram data, layout, quick-style, and colors parts. Every accepted media picture owns one canonical video/media relationship pair to a uniquely inbound, non-empty, relationship-free video/mp4 part plus one poster ImagePart. Every present connector endpoint must resolve to an element in the same copied SlidePart tree. Accepted tables are inline-only and cannot add a fill, link, or another package edge; accepted groups and connectors add no relationship themselves, and every nested picture must consume one exact verified ImagePart relationship. The pending clone resolves connector targets to fresh clone-local elements, while export privately preserves the source-bound endpoint identities. Export creates a distinct SlidePart and presentation relationship, allocates distinct byte-identical ChartPart, EmbeddedPackagePart, four typed diagram parts, and SDK MediaDataPart payloads for the accepted closed leaves, shares the verified layout, immutable ordinary/OLE-preview/media-poster ImageParts, NotesMaster, and presentation-wide CommentAuthorsPart, copies accepted NotesSlide and SlideComments XML byte-for-byte, and repoints only the notes back-reference at the clone while retaining the origin. The clone must remain untouched until export and reimport; its ChartParts, OLE workbook packages, SmartArt parts, InkML parts, and MP4 parts are then independent. Supported chart or OLE-workbook edits on the clone cannot affect the origin; a separately recognized canonical plain-node SmartArt diagram exposes only source-bound node-text replacement, while other SmartArt, InkML, and media remain source-bound/read-only after reimport. Malformed, shared, external, non-XLSX, nested, relationship-bearing, or replacement-pending OLE graphs, nested/noncanonical/connected SmartArt, InkML, or media graphs, malformed/relationship-bearing/dangling custom-show actions, unsupported connector forms or targets, formula/external-data/embedded-workbook/connected/orphan chart graphs, shape-level/hover/unknown/orphan hyperlinks, external or irregular images, and other complex graphs fail closed. |
| `slide.groups.add` | api | Author recursive native DrawingML p:grpSp trees with outer off/ext and local chOff/chExt coordinates. The bounded profile supports modeled shapes, connectors, images, tables, charts, and nested groups; canonical imported groups allow fixed-topology semantic edits, while group-level fills/effects, locks, transforms, extensions, or unsupported descendants remain opaque and read-only. |
| `slide.images.add` | api | Add an inspectable image facade with alt text, embedded data, contain/cover/stretch fitting, explicit crop, frame, direct rotation/flips, layout JSON, crop-aware SVG preview, and PPTX output. OpenChestnut maps the bounded rectangular profile to native DrawingML a:srcRect. |
| `slide.moveTo` | api | Move this slide to an existing 0-based deck index. On an imported PPTX, OpenChestnut rewrites only the retained source SlidePart order in the presentation slide-ID list; unrelated topology changes and broad graph clones remain fail-closed. |
| `slide.placeholders.getItem` | api | Resolve a slide placeholder shape by stable ID, name, placeholder type, or numeric index. Imported placeholder.textEditable reports a verified local SlidePart text capability; identity, geometry, formatting, layout binding, and inherited Master/Layout graphs remain source-bound. |
| `slide.setBackground` | api | Set a direct slide background to a six-digit RGB/theme color solid fill or a native style reference. Recognized imported direct backgrounds are hash-bound and editable; inherited Layout/Master backgrounds remain inherited. |
| `slide.setLayout` | api | Alias of slide.applyLayout(layout): bind and materialize a bounded source-free layout for native PPTX export. |
| `slide.setTransition` | api | Set a direct p:transition to bounded fade or directional push behavior with slow/medium/fast speed plus click/timer advancement. Source-free slides may author it; imported slides may replace one canonical existing direct transition or add one only when transition.capability.addable is true. Timing, sound, extension, opaque-effect, and every other source graph fail closed. |
| `slide.shapes.add` | api | Add a shape/textbox with preset or bounded literal custom geometry, position, optional center-based rotation/flips, fill, line, text, and DrawingML text-body layout. |
| `slide.speakerNotes.capability` | api | Return defensive sourceBound, partPresent, editable, and addable evidence. addable identifies an imported notes-absent slide whose source NotesMaster/SlideMaster Theme graph can safely receive a canonical NotesSlide. Export independently re-proves the package graph, so mutating model or wire data cannot grant authority. |
| `slide.tables.add` | api | Add an inspectable table facade with rows, columns, values, cells, rectangular merges, layout JSON, SVG preview, and canonical OpenChestnut plain-text PPTX output. |
| `slideCommentThread.addReply` | api | Append a direct reply to a source-free Office 2021 modern comment thread. Imported reply topology is fixed: existing reply text/status may change, but adding or removing replies fails closed. |
| `slideCommentThread.reopen` | api | Set the modern root comment status back to active while preserving fixed imported identity, anchor, position, and reply topology. |
| `slideCommentThread.resolve` | api | Set the modern root comment status to resolved. Imported export re-proves author/date/anchor/position/topology and source-part hashes before changing only status. |
| `table.merge` | api | Merge one inclusive rectangular table range, retain the upper-left value, clear and lock covered cells, and emit canonical DrawingML merge topology. |

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

#### `compose.text`

Create the same editable paragraph node through the reference-template-compatible children-first text(children, props) helper.

**Schema parameters:**

- `children` (string|string[]|object[]) required — Text or run-like children passed as the first argument.
- `props` (object) — Paragraph name, className, style, sizing, and placement metadata passed as the second argument.

**Schema returns:**

- `node` (object) — Reference-template-compatible alias that returns the same paragraph compose node.

#### `exportPptxWithOpenChestnut`

Export bounded direct slide backgrounds, textbox/rectangle/roundRect/ellipse shapes, rich text and lists, basic fills/lines/shadows, straight/elbow connectors and arrows, embedded pictures with native crop/contain/cover semantics, fixed-grid plain-text tables, recursive native p:grpSp trees, relationship-free rich speaker notes, legacy annotations, Office 2021 modern root/direct-reply threads, source-free bar/line/pie charts, the bounded literal clustered bar+line combo profile with either shared primary axes or a canonical secondary line pair, validated payload-only replacement for eligible imported OLE workbooks, and bounded source-bound text updates for canonical SmartArt document nodes. Recognized imported modern threads allow only existing text/status edits; their identity, author/date metadata, anchor/range, position, topology, relationships, and source hashes remain fixed. Inherited or complex graphs remain preserved and fail closed on unsupported mutation.

**Schema parameters:**

- `presentation` (Presentation) required — Presentation facade within the bounded direct-slide-background/shape/rich-text/picture/fixed-table/connector/recursive-group/plain-text-notes/legacy-comment/Office-2021-modern-comment and literal native-chart boundary. Charts cover bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, 2D bubble, and clustered bar+line combo. A combo supports only primary bars plus all-primary lines or all-secondary lines with the canonical top/right axis pair; formula/external data, advanced plots, irregular combos, and other imported package graphs must remain unchanged.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — PPTX bytes produced by the bundled Open XML SDK WebAssembly codec, including bounded embedded-picture, fixed-grid plain-text-table, and recursive native-group profiles, with codec diagnostics in metadata.

#### `importPptxWithOpenChestnut`

Import PPTX bytes with editable bounded direct slide backgrounds, shapes, rich text, recognized owner-local SlidePart placeholder text, rectangular pictures and native source rectangles, tables, connectors, recursive canonical p:grpSp groups, bar/line/pie charts, the canonical literal clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and a re-proven addable capability for eligible notes-absent slides, unchanged-only legacy comments, fixed-topology modern comment text/status edits, defensive payload access for eligible OLE workbooks, and a source-bound SmartArt plain-node text capability only for the canonical closed four-part one-paragraph/one-run DiagramDataPart profile. Reactions/task fields, nested replies/anchors, connected comment parts, inherited Master/Layout graphs, complex backgrounds/blips/groups, field/link/picture-bullet/layout-bearing notes, irregular combos, ambiguous OLE graphs, rich/multi-run/connected SmartArt, and other unsupported content remain source-bound and read-only.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — PPTX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `presentation` (Presentation) — Imported presentation facade with editable bounded direct slide backgrounds, shapes, rich text, recognized owner-local SlidePart placeholder text, pictures, tables, connectors, recursive canonical groups, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and the clustered bar+line combo profile with either shared primary axes or a secondary line pair. Formula/external data and advanced plot topology remain source-bound. Placeholder identity/geometry/formatting and inherited template graphs remain source-bound; advanced package graphs are read-only except for validated payload-only replacement on eligible OLE workbooks through getEmbeddedWorkbook and replaceEmbeddedWorkbook.

#### `nativeObject.getEmbeddedWorkbook`

Read a defensive FileBlob copy of the XLSX payload from an eligible source-bound top-level OLE object without exposing arbitrary native-part mutation.

**Schema returns:**

- `workbook` (FileBlob) — Defensive XLSX FileBlob copy with source part-path and SHA-256 metadata. Available only for a uniquely bound top-level OLE package relationship.

#### `nativeObject.replaceEmbeddedWorkbook`

Replace only the XLSX payload of an eligible imported top-level OLE object. OpenChestnut validates the new workbook and immutable source binding, preserves the OLE shell, relationships, preview, and all other native parts, and fails closed for malformed or ambiguous graphs.

**Schema parameters:**

- `workbook` (FileBlob|Uint8Array|ArrayBuffer|ArrayBufferView) required — Replacement XLSX bytes, copied defensively and limited to 16 MiB before canonical export validation.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one payload-only replacement on an eligible source-bound top-level OLE object. Export preserves the OLE shell, relationship topology, preview image, and other native parts; invalid XLSX or changed source bindings fail closed.

#### `nativeObject.setDiagramNodeText`

Replace text in one existing source-bound SmartArt document node only when its top-level four-part graph and DiagramDataPart prove the canonical one-paragraph/one-run plain-text profile. Node IDs, graph topology, layout/style/colors, geometry, and every non-data part remain fixed; unsupported diagrams reject without fallback.

**Schema parameters:**

- `nodeId` (string) required — Exact existing SmartArt DiagramDataPart dgm:pt/@modelId from nativeObject.diagramText.nodes. Node creation, removal, ordering, and identity changes are not supported.
- `text` (string) required — Replacement plain text, limited to 32,767 XML-safe characters. Tabs, line feeds, and carriage returns are allowed; other XML control characters and invalid Unicode scalars reject.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — Queues one text-only update for a source-bound top-level SmartArt frame that re-proves the canonical closed dm/lo/qs/cs graph and one direct dgm:t > a:p > a:r > a:t document-node profile. Export may rewrite only its bound DiagramDataPart, preserves node IDs/order, source frame, relationships, layout, quick-style, colors, and all non-data parts, then reimports and re-proves the profile. Rich, multi-run, connected, nested, or otherwise unrecognized SmartArt graphs fail closed without a fallback.

#### `nativeObject.setName`

Native OLE, SmartArt/diagram, contentPart, and media objects imported through OpenChestnut are source-bound and read-only for names; setName rejects instead of mutating the preserved package graph. A separate bounded SmartArt plain-node text capability is exposed only as nativeObject.setDiagramNodeText.

**Schema parameters:**

- `name` (string) required — Requested native-object display name, limited to 1,024 characters. Imported native objects are read-only, so the method rejects.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — No mutation is performed; imported native OLE/diagram/contentPart objects are source-bound and read-only.

#### `nativeObject.setPosition`

Native OLE, SmartArt/diagram, contentPart, and media objects imported through OpenChestnut are source-bound and read-only; setPosition rejects instead of rewriting their geometry or payload graph.

**Schema parameters:**

- `position` (object) required — Requested outer pixel frame. Imported native objects are read-only, so the method rejects.

**Schema returns:**

- `nativeObject` (NativePresentationObject) — No mutation is performed; native geometry and payload graphs remain source-bound and read-only.

#### `Presentation.create`

Create a deck model whose canonical OpenChestnut export supports ordinary slides, bounded direct fade/push transitions, direct solid/style-reference slide backgrounds, shapes, rich text, tables, images, connectors, recursive native p:grpSp groups, plain-text speaker notes, native custom shows with canonical run links, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts, and a bounded literal clustered bar+line combo profile. Combo bars stay on the primary pair; all lines share either that pair or the canonical secondary top/right pair. Formula/external chart data, custom themes, Master/Layout authoring, comments, custom-show topology mutation, advanced plot geometry, mixed line groups, secondary bars, irregular combo graphs, and other package-level features remain outside the source-free PPTX boundary.

**Schema parameters:**

- `slideSize` (object) — Slide width and height in pixels; defaults to 1280x720. On a trusted imported PPTX, changing it updates only the source-bound p:sldSz canvas and never rescales existing coordinates.
- `theme` (object) — Model theme metadata. OpenChestnut 0.2 source-free export requires the default theme; imported themes are read-only.
- `master` (object) — The one canonical source-free Slide Master: name/background, bounded title/body/ctrTitle/subTitle direct-frame placeholders, and bounded textParagraphStyles. Theme overrides are unsupported.
- `masters` (object[]) — Model-level Slide Master definitions. Source-free PPTX authoring accepts exactly one master; imported master graphs remain source-bound and read-only.
- `layouts` (object[]) — Bounded source-free layouts linked to the canonical master. Each uses blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only.
- `commentFormat` (string) — Comment wire family: legacy (default) or modern. Modern selects the bounded Office 2021 author/comments graph; the two families cannot be mixed.

**Schema returns:**

- `presentation` (Presentation) — Editable presentation facade.

#### `presentation.customShows.add`

Define an ordered native p:custShowLst playback route for source-free OpenChestnut export. Text runs may target a show by exact name with optional returnToSlide. Canonical imported shows may change only their name and ordered retained-slide membership; fixed native identity keeps existing run links bound across a rename, while irregular graphs stay opaque.

**Schema parameters:**

- `name` (string) required — Unique custom-show name, compared case-insensitively.
- `slides` (PresentationSlide[]|string[]) required — Ordered non-empty list of slide facades or stable slide IDs from this presentation.
- `nativeId` (number) — Optional preserved unsigned 32-bit p:custShow ID; new IDs are allocated collision-free.

**Schema returns:**

- `customShow` (PresentationCustomShow) — Appended native custom show for source-free PPTX authoring. Imported additions fail closed; use name assignment and setSlides(...) only on an existing canonical show.

#### `presentation.customShows.getItem`

Resolve a source-free or canonical imported custom show by zero-based index, stable facade ID, or exact name.

**Schema parameters:**

- `idOrNameOrIndex` (string|number) required — Stable custom-show ID, exact name, or zero-based collection index.

**Schema returns:**

- `customShow` (PresentationCustomShow|undefined) — Matching custom-show facade or undefined.

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

#### `presentation.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of explicitly used presentation text and bullet font families.

**Schema returns:**

- `families` (string[]) — Explicit font-family inventory; theme tokens such as +mj-lt are excluded.

#### `presentation.inspect`

Emit NDJSON for deck, custom shows, PowerPoint sections, slides, direct slide transitions, textboxes, shapes, grouped shapes, tables, charts, images, and native contentPart/OLE/diagram/media objects with bounded editability, relationship-reference, root-relationship, preserved-part, and eligible embedded-workbook summaries; narrow with search/target anchors and shape fields with include/exclude.

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

- `kind` (string) — Comma-separated deck/theme/layout/slide/transition/textbox/textRange/shape/groupShape/table/chart/image/connector/nativeObject/contentPart/oleObject/diagram/comment/notes/customShow/section kinds.
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

#### `presentation.layout.clearBackground`

Clear a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Clears a direct background on a bounded source-free layout. Imported-layout edits fail closed.

#### `presentation.layout.placeholders.add`

Append a direct-frame title/body/ctrTitle/subTitle text placeholder to a source-free layout. It becomes a native p:ph and must be materialized on each slide through applyLayout/setLayout; object/media/chart/table placeholders remain source-bound.

**Schema parameters:**

- `type` (string) required — title, body, ctrTitle, or subTitle; common aliases centeredTitle and subtitle normalize to native tokens.
- `idx` (number) — Native unsigned placeholder index; index is accepted as an alias.
- `index` (number) — Alias of idx.
- `position` (object) required — Required direct pixel frame { left, top, width, height } for source-free export.
- `text` (string|string[]|object|object[]) — Optional prompt/default text using the bounded presentation text profile.
- `style` (object) — Optional bounded default run/paragraph style.

**Schema returns:**

- `placeholder` (object) — Appended source-free layout placeholder definition. Use slide.applyLayout/setLayout to materialize it on a slide.

#### `presentation.layout.placeholders.summary`

Return a defensive layout-placeholder discovery snapshot with stable IDs, names, native types/indexes, required flags, and direct-frame presence/geometry; editing the snapshot cannot mutate the model.

**Schema returns:**

- `summary` (object) — Fresh defensive snapshot of the layout placeholder collection. It reports ownerId, count, requiredCount, sorted types, and copied items; imported inherited placeholders explicitly report hasDirectPosition: false.

#### `presentation.layout.setBackground`

Set a direct background on a bounded source-free layout. Imported-layout mutation remains source-bound and fails closed.

**Schema parameters:**

- `background` (string|object) required — Direct solid RGB/scheme background or native style reference with index.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Sets a direct background on a bounded source-free layout. Imported-layout edits fail closed.

#### `presentation.layouts.add`

Create one bounded source-free layout under the canonical master. Use blank, title, titleOnly, or obj/titleAndContent plus direct-frame text placeholders; imported layouts remain source-bound and read-only.

**Schema parameters:**

- `name` (string) required — Layout name; passing a name string is also accepted.
- `type` (string) — Source-free type: blank, title, titleOnly, obj, or aliases object/content/titleAndContent. Imported layouts retain their native type read-only.
- `masterId` (string) — Master identity.
- `background` (string|object) — Optional layout background overriding the linked master background.
- `placeholders` (object[]) — Direct-frame title/body/ctrTitle/subTitle source-free text placeholders. Each needs type, idx/index, and position left/top/width/height; object/chart/table/media placeholders are not authored.
- `slideGuides` (object[]) — Imported layouts expose the presentation's read-only native guide definitions. Canonical export preserves them through the source-bound view-properties part.

**Schema returns:**

- `layout` (SlideLayoutTemplate) — Appended bounded source-free layout under the canonical master. Imported layout graphs remain source-bound and read-only.

#### `presentation.layouts.getById`

Resolve a layout by its stable ID without falling back to a same-named or same-typed layout.

**Schema parameters:**

- `id` (string) required — Exact stable layout ID.

**Schema returns:**

- `layout` (SlideLayoutTemplate|undefined) — Matching layout or undefined.

#### `presentation.master`

Access the one canonical source-free Slide Master. It may author a direct background, bounded text styles, and direct-frame title/body/ctrTitle/subTitle placeholders; imported Master graphs remain source-bound and read-only.

**Schema parameters:**

- `id` (string) — Stable master identity used by layouts.
- `name` (string) — Native Slide Master name.
- `background` (string|object) — Solid RGB/scheme background or native background reference with index.
- `theme` (object) — Optional model theme override. Canonical source-free export rejects master-specific theme overrides.
- `placeholders` (object[]) — Source-free direct-frame title/body/ctrTitle/subTitle text placeholders. Each requires type, idx/index, and left/top/width/height; imported placeholders remain source-bound and read-only.
- `textParagraphStyles` (object) — title/body/other level maps (0-8) using the structured paragraph style fields, including embedded or external bulletImage values.
- `slideGuides` (object[]) — Read-only imported PowerPoint guide definitions with horizontal/vertical orientation and raw native position. Source-free authoring and imported mutation are unsupported.

**Schema returns:**

- `master` (PresentationSlideMaster) — One canonical source-free Slide Master or a source-bound imported master. Source-free output supports a direct background, bounded text styles, and direct-frame textual placeholders; imported masters are read-only.

#### `presentation.master.clearBackground`

Clear the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed.

**Schema returns:**

- `master` (PresentationSlideMaster) — Clears the direct background of the one canonical source-free master. Imported-master edits fail closed.

#### `presentation.master.setBackground`

Set the direct background of the one canonical source-free master. Imported-master mutation remains source-bound and fails closed.

**Schema parameters:**

- `background` (string|object) required — Direct solid RGB/scheme background or native style reference with index.

**Schema returns:**

- `master` (PresentationSlideMaster) — Sets the direct background of the one canonical source-free master. Imported-master edits fail closed.

#### `presentation.master.setTheme`

Set a model-level master theme override for preview only. Canonical PPTX export rejects that source-free override; imported-master mutation remains source-bound and fails closed.

**Schema parameters:**

- `theme` (object|null) required — Partial master theme override, or null to inherit presentation.theme.

**Schema returns:**

- `master` (PresentationSlideMaster) — Model-only theme override for preview; canonical export rejects source-free master-specific themes and imported-master edits.

#### `presentation.masters.add`

Append a model-level Slide Master. Source-free PPTX authoring requires exactly one master, so use Presentation.create({ master }) or presentation.master for the canonical profile; multiple masters and imported-master edits fail closed.

**Schema parameters:**

- `id` (string) required — Stable unique master identity used by layouts.
- `name` (string) — Native Slide Master name.
- `background` (string|object) — Solid RGB/scheme background or native background reference with index.
- `theme` (object) — Optional model theme override; source-free master-specific themes are unsupported.
- `placeholders` (object[]) — Direct-frame title/body/ctrTitle/subTitle source-free text placeholders. A second master makes source-free export fail closed.
- `textParagraphStyles` (object) — title/body/other level maps (0-8) using the structured paragraph style fields, including embedded or external bulletImage values.

**Schema returns:**

- `master` (PresentationSlideMaster) — Appended model-level Slide Master. Canonical source-free export accepts exactly one master, so adding another deliberately fails closed.

#### `presentation.masters.getItem`

Resolve a model-level or imported Slide Master by stable ID or name.

**Schema parameters:**

- `idOrName` (string) required — Stable master ID or native master name.

**Schema returns:**

- `master` (PresentationSlideMaster|undefined) — Matching Slide Master or undefined.

#### `presentation.resolve`

Map stable inspect anchor IDs back to facade objects, including custom shows, PowerPoint sections, and slide transitions; imported advanced package objects may be read-only.

**Schema parameters:**

- `id` (string) required — Stable deck, theme, layout, slide, transition, element, custom-show, section, comment, or text-range ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `presentation.sections.add`

Define a native PowerPoint p14:sectionLst entry for source-free OpenChestnut export. Sections together must form the complete ordered slide partition. Canonical imported sections may change only existing names and contiguous boundaries while count, order, stable facade identity, and native GUID stay fixed; irregular graphs remain opaque.

**Schema parameters:**

- `name` (string) required — Unique 1-255-character section name, compared case-insensitively.
- `slides` (PresentationSlide[]|string[]) required — One or more slide facades or stable slide IDs from this presentation. Across all sections, memberships must partition every deck slide exactly once and in deck order.
- `nativeId` (string) — Optional preserved brace-delimited GUID for native p14:section/@id; new source-free sections receive a deterministic GUID.

**Schema returns:**

- `section` (PresentationSection) — Appended native PowerPoint p14:sectionLst entry. Source-free authoring owns the complete ordered slide partition. Canonical imported sections keep count, order, facade identity, and native GUID fixed; only names and contiguous partition boundaries may change. Extension-bearing or irregular section graphs remain opaque and fail closed.

#### `presentation.sections.getItem`

Resolve a source-free or canonical imported PowerPoint section by zero-based index, stable facade ID, or exact name.

**Schema parameters:**

- `idOrNameOrIndex` (string|number) required — Stable section ID, exact name, or zero-based collection index.

**Schema returns:**

- `section` (PresentationSection|undefined) — Matching PowerPoint-section facade or undefined.

#### `presentation.slides.add`

Append an editable core slide with an optional bounded source-free layout, direct fade/push transition, solid/style-reference background, and plain-text speaker notes. A supplied layout is resolved and materialized transactionally; effective imported Layout/Master inheritance is never flattened.

**Schema parameters:**

- `name` (string) — Inspectable slide name.
- `layout` (string|object) — Optional bounded layout name/ID/facade. slides.add resolves it transactionally and materializes its text placeholders; an unknown or cross-presentation layout leaves no slide behind.
- `background` (string|object) — Optional direct slide background: RGB/theme color or { fill, mode: 'solid'|'reference', index? }. Gradient, pattern, image, transform, and effect-bearing backgrounds are preview-only/source-preserved and fail closed on canonical mutation.
- `transition` (object) — Optional direct transition: { effect: 'fade'|'push', direction?: 'left'|'up'|'right'|'down', speed?: 'slow'|'medium'|'fast', advanceOnClick?: boolean, advanceAfterMs?: integer 0..86400000 }. Fade rejects direction; push defaults to left; click defaults to true.
- `notes` (string|PresentationParagraph[]) — Optional speaker notes authored into the canonical PresentationML notes graph. A paragraph has runs plus ordinary direct paragraph/run styling; note-local links, fields, picture bullets, list styles, and body layout are rejected.

**Schema returns:**

- `slide` (Slide) — Appended editable slide. A supplied bounded source-free layout is bound and materialized immediately.

#### `presentation.slides.insert`

Insert a source-free slide after an existing Slide or 0-based index, or at the beginning with after: null. It uses the same transactional layout materialization, bounded direct fade/push transition, and notes/background profile as slides.add; imported additions fail closed, while slide.duplicate and slide.delete each have their own narrow source-preserving OPC profiles.

**Schema parameters:**

- `after` (Slide|number|null) — Existing slide facade or 0-based index to insert after; null inserts first. Omit to append.
- `name` (string) — Inspectable slide name.
- `layout` (string|object) — Optional bounded layout name/ID/facade. The new source-free slide is created and materialized transactionally.
- `background` (string|object) — Optional direct slide background: RGB/theme color or { fill, mode: 'solid'|'reference', index? }.
- `transition` (object) — Optional direct transition with the same bounded fade/push, speed, click, and timer profile as presentation.slides.add.
- `notes` (string|PresentationParagraph[]) — Optional speaker notes authored into the canonical PresentationML notes graph. A paragraph has runs plus ordinary direct paragraph/run styling; note-local links, fields, picture bullets, list styles, and body layout are rejected.

**Schema returns:**

- `slide` (Slide) — Inserted source-free slide. Unknown insertion targets or layouts leave the collection unchanged; imported additions remain fail-closed. See slide.duplicate for the separate bounded source-preserving clone profile.

#### `presentation.slideSize`

Read or set the deck canvas in pixels. On a trusted imported PPTX, a changed size is a deliberately canvas-only source-bound operation: OpenChestnut updates only ppt/presentation.xml p:sldSz, clears an old preset type, and leaves slide, layout, master, chart, and shape coordinates unchanged. It never silently rescales or reflows content; callers must make any layout edits explicitly.

**Schema parameters:**

- `width` (number) required — Finite non-negative canvas width in pixels; a changed imported canvas must resolve to a positive signed 32-bit EMU value.
- `height` (number) required — Finite non-negative canvas height in pixels; a changed imported canvas must resolve to a positive signed 32-bit EMU value.

**Schema returns:**

- `slideSize` ({ width: number, height: number }) — Current deck canvas. A trusted imported PPTX may change only this p:sldSz canvas; existing slide, layout, master, chart, and shape coordinates are preserved exactly, and callers must explicitly recompose any affected layout.

#### `presentation.textRange`

Inspect or resolve stable textRange anchors such as shapeId/text for editable slide text frames.

**Schema parameters:**

- `id` (string) required — Stable shape text-range ID ending in /text.

**Schema returns:**

- `textRange` (TextRange|undefined) — Editable slide text-range facade or undefined.

#### `presentation.theme`

Inspect the model theme and theme inheritance. Custom source-free themes are not authored by OpenChestnut 0.2, and imported themes are source-bound and read-only.

**Schema parameters:**

- `name` (string) — Model theme name. Source-free customization is rejected; imported theme metadata is read-only.
- `colors` (object) — Complete tx1/bg1/tx2/bg2, accent1-accent6, hlink, and folHlink color scheme; dk1/lt1/dk2/lt2 aliases are accepted.
- `fonts` (object) — Major/minor Latin plus optional East-Asian and complex-script font families.
- `textStyles` (object) — Slide Master title/body/other defaults with fontSize, bold, italic, color, fontFamily, and alignment.
- `colorMap` (object) — Slide Master semantic color mapping for bg1/tx1/bg2/tx2, accents, and hyperlinks.

**Schema returns:**

- `theme` (PresentationTheme) — Inspectable model theme; canonical export accepts only the default source-free theme and preserves imported themes read-only.

#### `presentation.validateLayout`

Detect layout QA issues across slides, including off-canvas elements, geometry overlaps, and basic text overflow.

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area in square pixels before reporting.
- `boundsPadding` (number) — Allowed padding outside the slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Layout QA result with ok, issues, ndjson, and truncated.

#### `presentation.verify`

Return QA issues for layout validation, missing master/layout references, placeholder fidelity, chart/data consistency, table shape, image data, and dangling comments.

**Schema parameters:**

- `minOverlapArea` (number) — Minimum overlap area for layout QA.
- `boundsPadding` (number) — Allowed padding outside slide bounds.
- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Presentation semantic/layout QA result.

#### `presentation.view`

Control local editor gridline/guide visibility and inspect imported PowerPoint grid spacing, snap settings, and read-only slide guides. Visibility is local model state; imported viewProps.xml metadata remains source/hash-bound and unchanged by canonical export.

**Schema returns:**

- `view` (PresentationView) — Local gridlinesVisible/guidesVisible state with show/hide/toggle methods, optional imported gridSpacingCxEmu/gridSpacingCyEmu, and serialized hidden guide visibility. Imported slide-guide definitions are exposed read-only through master/layout slideGuides and remain source-bound in PPTX output.

#### `PresentationFile.exportPptx`

Serialize PPTX through the single bundled OpenChestnut codec. Only limits is accepted; legacy codec and lossy-fallback options fail explicitly.

**Schema parameters:**

- `presentation` (Presentation) required — Presentation facade to serialize.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML PPTX package bytes.

#### `PresentationFile.importPptx`

Import PPTX through the single bundled OpenChestnut codec with source-bound opaque preservation, speaker-notes edit/add capability evidence, bounded text-only edits for recognized local SlidePart placeholders and canonical SmartArt plain document nodes, eligible OLE workbook payload access/replacement, and fail-closed unsupported edits.

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `presentation` (Presentation) — Imported presentation facade with editable core objects, bounded text-only replacement for recognized owner-local SlidePart placeholders, recognized direct slide backgrounds, canonical fixed-topology recursive groups, literal bar/line/pie/standard-area/fixed-doughnut/marker-scatter/2D-bubble charts plus the clustered bar+line combo profile with either shared primary axes or a secondary line pair, legacy text-only speaker notes plus fixed-topology relationship-free rich notes and explicit edit/add capability evidence, bounded legacy slide-level comments (unchanged-only), bounded Office 2021 modern root/direct-reply threads (text/status editable), and payload-only replacement for eligible source-bound OLE workbooks. A notes-absent slide can add a canonical NotesSlide only when the source NotesMaster/SlideMaster Theme graph is re-proven safe. Chart formulas/external data and advanced plot topology remain source-bound. Placeholder identity/geometry/formatting and inherited Master/Layout graphs, complex backgrounds/groups, field/link/picture-bullet/layout-bearing notes, mixed line groups, secondary bars, irregular comment anchors/reactions/task fields, themes, other native objects, and unsupported package graphs remain source-bound.

#### `PresentationFile.inspectPptx`

Inspect bounded PPTX parts, content types, relationships, namespace-aware source XML references, legacy notes/comments evidence, and Office 2021 modern author/thread/anchor semantics under decompression budgets.

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

- `package` (object) — PPTX package result with ok, issues, parts, records, bounded NDJSON, and notes/comments semantic validation evidence.

#### `PresentationFile.patchPptx`

Apply path-validated PPTX part patches, including safe slide/master/layout ID lists and slide image/chart DrawingML mutations, and atomically reject dangling package references or invalid notes/comments semantics.

**Schema parameters:**

- `pptx` (FileBlob|Uint8Array) required — PPTX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types, relationships, and PPTX notes/comments semantics atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; PPTX supports slide/master/layout ID lists plus image/chart objects in a slide shape tree.
- `sourceReference` (boolean|object) — Opt-in semantic XML mutation. Image/chart objects require explicit pixel position { left, top, width, height }, validate generated or explicit non-visual objectId, and clean matching slide objects on deletion.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched PPTX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `shape.text.set`

Set plain or structured text with ordered text, field, and line-break inlines; bounded run formatting; character, picture-bullet, or auto-numbered lists; levels, indents, spacing; and external URI, internal-slide, relative-action, or existing custom-show hyperlinks. Missing, opaque, malformed, relationship-bearing, or dangling custom-show targets and unmodeled text graphs fail closed in canonical PPTX export.

**Schema parameters:**

- `text` (string|string[]|object|object[]) required — Plain text, paragraph strings, inline arrays, or paragraph objects. Canonical OpenChestnut export supports ordered text, fields, styled line breaks, bounded run/paragraph formatting, character and picture bullets, auto-numbering, levels, indents, spacing, tab stops, and one absolute uri, target slideId, relative action (nextSlide, previousSlide, firstSlide, lastSlide, endShow), or existing customShow name per link. customShow may include returnToSlide and survives the bounded slide clone as the same relationship-free stable-identity action without adding the clone to show membership; missing, opaque, malformed, relationship-bearing, or dangling targets fail closed.

**Schema returns:**

- `textFrame` (TextFrame) — The same live text frame with normalized paragraphs and a backward-compatible flattened value.

#### `shape.useBackgroundFill`

Read the presence-aware imported PresentationML p:sp useBgFill flag. It affects preview paint but remains source-bound and read-only; source-free authoring or wire mutation fails closed.

**Schema returns:**

- `useBackgroundFill` (boolean|undefined) — True/false only when the native attribute was present; otherwise undefined.

#### `slide.addNotes`

Set speaker notes as text or relationship-free paragraph/run data for inspect, preview, and canonical PPTX output. OpenChestnut authors source-free notes, preserves the legacy text-only edit path, and edits a fixed imported rich paragraph/run topology; fields, hyperlinks, picture bullets, notes-body list styles/layout, and unsafe NotesMaster graphs remain source-bound and fail closed.

**Schema parameters:**

- `text` (string|PresentationParagraph[]) required — Speaker notes text or paragraph/run data. Each structured paragraph follows the presentation text subset; note-local hyperlinks, fields, picture bullets, list styles, and body properties are rejected.

**Schema returns:**

- `notes` (object) — Speaker-notes record. Source-free notes and simple hash-bound imported text remain editable; an imported relationship-free rich body may edit only its fixed paragraph/inline topology. A notes-absent imported slide may add a canonical NotesSlide only when speakerNotes.capability.addable is true; export re-proves that package graph. Fields, hyperlinks, picture bullets, notes-page layout, list styles, and unsafe NotesMaster graphs remain preservation-only.

#### `slide.applyLayout`

Bind a slide to a bounded source-free layout and materialize its effective direct-frame placeholder shapes. Applying the same layout is idempotent; switching a materialized layout fails closed. The resulting p:ph identities and direct frames export natively; imported Layout relationships remain preservation-only.

**Schema parameters:**

- `layout` (string|SlideLayoutTemplate) required — Layout name/ID or layout facade.

**Schema returns:**

- `shapes` (Shape[]) — Binds the slide and materializes effective direct-frame title/body/ctrTitle/subTitle placeholder shapes for native source-free PPTX output. Reapplying the same layout is idempotent; switching an already-materialized layout fails closed.

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

Add a source-free literal bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or clustered bar+line combo chart. Category families use shared literal categories; scatter and bubble use aligned per-series numeric X/Y values, with positive area-based bubble sizes. Supported variants retain title, legend, bounded axes, basic series styling, chart-level data labels, layout JSON, SVG preview, and native ChartPart output across import/edit/re-export. Formula/external data, advanced family geometry, topology changes, and unsupported styling fail closed rather than being flattened.

**Schema parameters:**

- `chartType` (string) — bar, line, pie, standard area, fixed 50%-hole doughnut, marker-only scatter, bounded 2D bubble, or combo for canonical OpenChestnut export. combo is the literal clustered bar+line profile described by series; unsupported or advanced family variants fail closed.
- `title` (string) — Chart title.
- `categories` (string[]) — Shared literal labels required by bar, line, pie, area, doughnut, and combo. Scatter and bubble reject shared categories and use per-series xValues.
- `series` (object[]) required — One or more named series. Category charts require one finite value per category. Scatter and bubble require aligned finite xValues and Y values; bubble additionally requires aligned positive bubbleSizes. Markers are limited to line and scatter, and marker-only scatter rejects a series line in favor of marker.line. For combo, every series declares chartType bar or line; there must be at least one primary bar and one line. Bars cannot be secondary. Lines are either all primary or all axisGroup: secondary; mixed primary/secondary line plots fail closed. Formula sources, point overrides, per-series labels, smooth, trendlines, error bars, and per-series chart types outside combo fail closed.
- `externalData` (object|FileBlob|ArrayBuffer|Uint8Array|string) — Model-only external/embedded workbook metadata. OpenChestnut 0.2 source-free charts require literal categories and values and reject externalData.
- `position` (object) — Pixel left/top/width/height frame.
- `axes` (object) — Basic axis titles, number formats, intervals, bounds, and major units. Category families use a category/value pair; scatter and bubble use two numeric value axes; pie and doughnut reject axes. A combo with all lines axisGroup: secondary may also set axes.secondary.category and axes.secondary.value, written at top/right. Secondary axes are invalid for primary-line combos, mixed line groups, or secondary bars.
- `legend` (object) — Legend options.
- `dataLabels` (boolean|object) — Chart-level showValue/showCategoryName/showSeriesName, circular-only showPercent for pie/doughnut, and a supported bounded position. Per-series overrides are unsupported.
- `styleId` (number) — Model-only chart style metadata; it is not part of the bounded OpenChestnut chart wire.
- `styleIndex` (number) — Model-only alias for styleId.
- `varyColors` (boolean) — Model-only varied-color preference outside the bounded OpenChestnut chart wire.
- `barOptions` (object) — Model-only advanced bar layout options outside the bounded OpenChestnut chart wire.
- `lineOptions` (object) — Model-only advanced line grouping/smoothing options; direct per-series marker formatting remains supported.

**Schema returns:**

- `chart` (ChartElement) — Appended editable native-chart facade.

#### `slide.clearBackground`

Remove the direct slide background so preview and PPTX output inherit from the preserved Layout/Master chain. Unsupported imported background graphs fail closed rather than being flattened or discarded.

**Schema returns:**

- `slide` (Slide) — The same slide with no direct background, inheriting from its preserved Layout/Master chain.

#### `slide.clearTransition`

Remove one canonical direct imported or source-free slide transition. A transition-absent imported slide remains a no-op until an explicit capability-approved add; timing, sound, extension, and opaque-effect graphs remain byte-preserved and reject mutation.

**Schema returns:**

- `slide` (Slide) — The same slide with no direct p:transition. Removing an imported transition requires the same canonical editable source profile as replacement.

#### `slide.comments.addThread`

Create either a bounded legacy PPTX annotation or an Office 2021 modern thread. A comment-free imported presentation may add canonical legacy review comments only when comments.capability.addable is true; existing legacy records remain source-bound and read-only. Modern mode supports a top-level element/text-range/textMatch anchor, one root, direct replies, independent people/timestamps, and active/resolved/closed state; imported modern graphs permit only fixed-topology text/status edits.

**Schema parameters:**

- `target` (undefined|string|Shape|ImageElement|TableElement|ChartElement|ConnectorElement|TextRange|object) — Legacy mode requires undefined. Modern mode accepts a top-level element/text-range ID or facade, { element }, { textRange }, or { textMatch: { element, query, occurrence? } }. Nested group-child and slide-level modern anchors remain unsupported.
- `text` (string) required — Root comment text.
- `author` (string) — Display author. Modern comments may instead provide comments[0].person with brace-delimited GUID id, name, initials, userId, and providerId.
- `position` (object) required — Explicit slide coordinate { x, y, unit?: 'px'|'emu' }. Legacy defaults to px; modern defaults to emu.
- `resolved` (boolean) — Modern root state. Legacy mode requires false.
- `created` (string) — ISO-8601 creation time for the root comment; defaults to the Unix epoch for deterministic output.
- `nativeFormat` (string) — Set modern for explicit Office 2021 authoring; Presentation.create({ commentFormat: 'modern' }) must select the same wire family.
- `comments` (object[]) — Optional root record. Modern records support nativeId/id, authorId/person, text, created, and active/resolved/closed status. Reactions, task fields, extensions, and nested replies fail closed.

**Schema returns:**

- `thread` (SlideCommentThread) — Create a bounded legacy annotation or Office 2021 modern root. A comment-free imported presentation may create canonical legacy parts only after comments.capability.addable preflight; OpenChestnut re-proves the whole source graph, allocates collision-free relationships, and never mixes comment families. Recognized legacy imports remain unchanged-only. Recognized modern imports expose root/direct replies and allow only text/status edits; author/person/date identity, position, target moniker, reply topology, part paths, relationships, and source hashes remain fixed.

#### `slide.comments.capability`

Inspect defensive source-bound comment-family evidence before authoring. A comment-free imported presentation may advertise legacy addability; existing legacy records remain read-only and modern graphs retain their separate fixed-topology edit contract.

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, format, partPresent, addable } evidence. For imported files, addable is true only when the complete presentation has no legacy or Office 2021 comment graph and OpenChestnut can create one canonical legacy CommentAuthorsPart plus slide-local SlideCommentsPart leaves. This is preflight evidence, not mutable write authority; export re-proves the source bytes and fails closed on existing, mixed, connected, or tampered graphs.

#### `slide.compose`

Materialize a clean-room compose tree with row, column, grid, layers, box, paragraph/text, shape, table, chart, image, and rule nodes into editable slide objects.

**Schema parameters:**

- `node` (object) required — Compose tree rooted in row, column, grid, layers, box, paragraph/text, shape, table, chart, image, or rule.
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

#### `slide.delete`

Remove this slide. Source-free decks may remove any non-final slide. An imported PPTX performs a real OPC deletion only for an isolated slide with exactly its layout relationship and no inbound/package-identity references; media, notes, comments, charts, OLE, hyperlinks, custom shows, sections, extensions, and all clone requests fail closed.

**Schema returns:**

- `result` (undefined) — No return value. The slide must belong to a presentation with at least two slides. Source-free export removes it normally. Imported PPTX export deletes the actual SlidePart, its presentation relationship, and its relationship part only after proving that the source slide has only its layout relationship, no inbound relationship, and no custom-show/section/extension or presentation-level identity reference. Media, notes, comments, charts, OLE, hyperlinks, data parts, complex graph delete, and clone requests fail closed.

#### `slide.duplicate`

Clone one original imported PPTX slide only when its unchanged graph contains canonical shapes, canonical inline fixed-grid tables with bounded rectangular merges, recognized closed literal-data charts, eligible top-level embedded-XLSX OLE frames, canonical top-level four-part SmartArt frames, canonical top-level closed InkML content parts, canonical top-level embedded-MP4 media pictures, embedded rectangular images, bounded canonical straight/elbow connectors, and recursively canonical groups containing only the non-native-graph leaf kinds, exactly one layout relationship, picture-bound image relationships, canonical run-level external/internal/relative-action links plus relationship-free custom-show links bound to an existing stable native show ID, and optional closed NotesSlide-to-NotesMaster/back-to-slide plus bounded legacy-comments leaves. Relationship-backed links keep exact IDs and targets; custom-show actions add no relationship and the clone is never inserted into show membership. Every accepted chart frame uniquely consumes one internal relationship to a numbered ChartPart whose child, external, hyperlink, and data relationship sets are empty. Every accepted OLE frame uniquely consumes one internal package relationship to a closed, uniquely inbound XLSX EmbeddedPackagePart and one internal preview ImagePart relationship. Every accepted SmartArt frame owns exactly one internal dm/lo/qs/cs relationship set to closed relationship-free diagram data, layout, quick-style, and colors parts. Every accepted media picture owns one canonical video/media relationship pair to a uniquely inbound, non-empty, relationship-free video/mp4 part plus one poster ImagePart. Every present connector endpoint must resolve to an element in the same copied SlidePart tree. Accepted tables are inline-only and cannot add a fill, link, or another package edge; accepted groups and connectors add no relationship themselves, and every nested picture must consume one exact verified ImagePart relationship. The pending clone resolves connector targets to fresh clone-local elements, while export privately preserves the source-bound endpoint identities. Export creates a distinct SlidePart and presentation relationship, allocates distinct byte-identical ChartPart, EmbeddedPackagePart, four typed diagram parts, and SDK MediaDataPart payloads for the accepted closed leaves, shares the verified layout, immutable ordinary/OLE-preview/media-poster ImageParts, NotesMaster, and presentation-wide CommentAuthorsPart, copies accepted NotesSlide and SlideComments XML byte-for-byte, and repoints only the notes back-reference at the clone while retaining the origin. The clone must remain untouched until export and reimport; its ChartParts, OLE workbook packages, SmartArt parts, InkML parts, and MP4 parts are then independent. Supported chart or OLE-workbook edits on the clone cannot affect the origin; a separately recognized canonical plain-node SmartArt diagram exposes only source-bound node-text replacement, while other SmartArt, InkML, and media remain source-bound/read-only after reimport. Malformed, shared, external, non-XLSX, nested, relationship-bearing, or replacement-pending OLE graphs, nested/noncanonical/connected SmartArt, InkML, or media graphs, malformed/relationship-bearing/dangling custom-show actions, unsupported connector forms or targets, formula/external-data/embedded-workbook/connected/orphan chart graphs, shape-level/hover/unknown/orphan hyperlinks, external or irregular images, and other complex graphs fail closed.

**Schema returns:**

- `slide` (Slide) — A new adjacent Slide. It is available only for an original imported PPTX source slide whose unchanged graph has canonical simple shapes, canonical inline fixed-grid tables, recognized closed literal-data charts, eligible top-level embedded-XLSX OLE frames, canonical top-level four-part SmartArt frames, canonical top-level closed InkML content parts, canonical top-level embedded-MP4 media pictures, embedded rectangular images, bounded canonical straight/elbow connectors, plus recursively canonical p:grpSp groups whose descendants contain only the non-native-graph leaf kinds, exactly one internal layout relationship, image relationships bound only by those pictures, canonical run-level external/internal/relative-action click links, relationship-free custom-show clicks that resolve through the fixed presentation-wide native-ID catalog, and optionally one closed NotesSlide -> NotesMaster/back-to-source-slide leaf plus one bounded legacy SlideCommentsPart leaf. Relationship-backed links retain exact IDs and targets; custom-show actions add no package edge and cloning never changes show membership. Every chart frame must consume one unique internal relationship to a numbered ChartPart with no child, external, hyperlink, or data relationship. Every OLE frame must consume one unique internal package relationship to a closed, uniquely inbound XLSX EmbeddedPackagePart plus one internal preview ImagePart relationship. Every SmartArt frame must contain exactly one canonical dgm:relIds root whose dm/lo/qs/cs relationships bind closed relationship-free diagram data, layout, quick-style, and colors parts. Every media picture must contain the canonical empty media action and consume one video/media relationship pair to the same uniquely inbound, non-empty, relationship-free video/mp4 part plus one poster ImagePart. Every present connector endpoint must resolve to an element in the same copied SlidePart tree. Accepted tables are inline-only: table fills, links, or other package edges remain outside this clone profile. Accepted groups and connectors add no relationship of their own; every nested picture must bind one exact verified ImagePart relationship. The pending clone resolves connector targets to fresh clone-local elements, while export privately restores the source-bound endpoint identities. Export allocates a distinct SlidePart and presentation relationship, byte-copies each accepted chart into a distinct ChartPart, each accepted OLE workbook into a distinct EmbeddedPackagePart, each accepted SmartArt root into four distinct typed diagram parts, each InkML payload into a distinct SDK CustomXmlPart, and each MP4 into a distinct SDK MediaDataPart while sharing the immutable poster. The clone must remain semantically unchanged until its first export and second import; afterward each ChartPart, OLE workbook package, SmartArt part, InkML part, and MP4 part has independent package identity. Supported chart or OLE-workbook edits on the clone do not affect the origin; a separately recognized canonical plain-node SmartArt diagram exposes only source-bound node-text replacement, while other SmartArt, InkML, and media remain source-bound/read-only. Source-free slides, already-cloned slides, same-transaction origin deletion, rich/connected comments, malformed/shared/external/non-XLSX/nested/relationship-bearing/replacement-pending OLE graphs, nested/noncanonical/connected SmartArt, InkML, or media graphs, malformed/relationship-bearing/dangling custom-show actions, unsupported connector forms or targets, formula/external-data/embedded-workbook/connected/orphan charts, shape-level/hover/unknown/orphan hyperlinks, external/data relationships, unbound or irregular images, and every broader graph clone fail closed.

**Notes:**

- The closed native-leaf profile also accepts an unchanged top-level p:contentPart only when it uniquely binds one internal relationship-free application/inkml+xml CustomXmlPart whose document element is ink in the standard InkML namespace. Export preserves the slide-local relationship ID, byte-copies the payload into a distinct SDK-typed clone part, and second import must prove disjoint paths with equal hashes. The object remains opaque and read-only; nested, extension-bearing, ambiguous, mistyped, non-InkML-root, or connected content parts fail closed.
- The embedded-media leaf accepts only one canonical top-level p:pic with the empty ppaction://media sentinel, one video and one media relationship to the same uniquely inbound relationship-free video/mp4 part, and one poster ImagePart. Export preserves both slide-local media IDs, byte-copies the MP4 into a distinct SDK MediaDataPart, and shares only the immutable poster. Second import must prove distinct MP4 paths with equal hashes and the same poster path. Media remains opaque and read-only; native pixel equality validates the poster, not playback, and linked, shared, non-MP4, nested, multi-binding, timing-rich, audio, or connected graphs fail closed.

#### `slide.groups.add`

Author recursive native DrawingML p:grpSp trees with outer off/ext and local chOff/chExt coordinates. The bounded profile supports modeled shapes, connectors, images, tables, charts, and nested groups; canonical imported groups allow fixed-topology semantic edits, while group-level fills/effects, locks, transforms, extensions, or unsupported descendants remain opaque and read-only.

**Schema parameters:**

- `name` (string) — Inspectable group name.
- `position` (object) required — Group frame in parent or slide pixel coordinates.
- `childFrame` (object) — Local child coordinate rectangle mapped through DrawingML chOff/chExt; defaults to the group width/height from 0,0.
- `shapes` (object[]) — Initial child shape/textbox definitions in local coordinates.
- `connectors` (object[]) — Initial child connector definitions in local coordinates.
- `groups` (object[]) — Initial nested group definitions.
- `tables` (object[]) — Initial native DrawingML table definitions in local coordinates.
- `charts` (object[]) — Initial relationship-backed chart definitions in local coordinates.
- `images` (object[]) — Initial relationship-backed picture definitions in local coordinates.
- `children` (object[]) — Ordered mixed child definitions using kind shape, connector, groupShape, table, chart, or image.

**Schema returns:**

- `group` (GroupShape) — Appended recursive grouped-shape facade for resolve, inspect, layout, SVG preview, and native p:grpSp export. Canonical imported groups are source-bound and editable without changing child topology; complex group shells or unsupported descendants are preserved as one opaque read-only object.

#### `slide.images.add`

Add an inspectable image facade with alt text, embedded data, contain/cover/stretch fitting, explicit crop, frame, direct rotation/flips, layout JSON, crop-aware SVG preview, and PPTX output. OpenChestnut maps the bounded rectangular profile to native DrawingML a:srcRect.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `fit` (string) — contain, cover, or stretch. For embedded images, OpenChestnut computes a bounded native a:srcRect from intrinsic dimensions; imported native source rectangles normalize to fit stretch plus explicit crop because PPTX has no fit keyword.
- `crop` (object) — Optional normalized { left, top, right, bottom } source edges in -1..1 with opposing sums below 1. Positive values crop; negative values expand for contain/letterbox semantics. Manual crop is applied before contain/cover fitting.
- `position` (object) — Pixel left/top/width/height frame.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } center transform. OpenChestnut preserves explicit false and safely edits recognized top-level embedded pictures.

**Schema returns:**

- `image` (ImageElement) — Appended editable image facade. OpenChestnut authors/imports embedded PNG/JPEG/GIF/safe-SVG rectangular pictures and permits native source-rectangle add/edit/remove plus same-format byte, name/alt, frame, and direct-transform edits; effects, external sources, complex blips, and non-rectangular geometry remain opaque.

#### `slide.moveTo`

Move this slide to an existing 0-based deck index. On an imported PPTX, OpenChestnut rewrites only the retained source SlidePart order in the presentation slide-ID list; unrelated topology changes and broad graph clones remain fail-closed.

**Schema parameters:**

- `index` (number) required — Existing zero-based destination index. It must be an integer from 0 through presentation.slides.items.length - 1.

**Schema returns:**

- `slide` (Slide) — The same slide at its new collection position. Imported PPTX export rewrites only p:sldIdLst for the retained source SlideParts; unrelated topology changes and broad graph clones fail closed. See slide.duplicate and slide.delete for their separate constrained source-part contracts.

#### `slide.placeholders.getItem`

Resolve a slide placeholder shape by stable ID, name, placeholder type, or numeric index. Imported placeholder.textEditable reports a verified local SlidePart text capability; identity, geometry, formatting, layout binding, and inherited Master/Layout graphs remain source-bound.

**Schema parameters:**

- `idOrNameOrTypeOrIndex` (string|number) required — Placeholder stable ID, display name, type, or numeric idx.

**Schema returns:**

- `shape` (Shape|undefined) — Matching placeholder shape or undefined. Imported shape.placeholder.textEditable is true only when the source binding recognizes the concrete SlidePart's local text body. In that case text.set(...) preserves native formatting/topology while replacing characters; use text.replace(...) for an in-run edit. The capability is re-proved from source on export and cannot be granted by mutating the model flag. Identity, geometry, formatting, and layout binding remain read-only.

#### `slide.setBackground`

Set a direct slide background to a six-digit RGB/theme color solid fill or a native style reference. Recognized imported direct backgrounds are hash-bound and editable; inherited Layout/Master backgrounds remain inherited.

**Schema parameters:**

- `background` (string|object) required — Direct RGB/theme color or { fill, mode: 'solid'|'reference', index? }; reference index must be an unsigned 32-bit integer.

**Schema returns:**

- `slide` (Slide) — The same slide with a normalized direct background; canonical PPTX export never flattens inherited Layout/Master backgrounds.

#### `slide.setLayout`

Alias of slide.applyLayout(layout): bind and materialize a bounded source-free layout for native PPTX export.

**Schema parameters:**

- `layout` (string|SlideLayoutTemplate) required — Layout name/ID or layout facade.

**Schema returns:**

- `slide` (Slide) — Alias of applyLayout that returns the slide.

#### `slide.setTransition`

Set a direct p:transition to bounded fade or directional push behavior with slow/medium/fast speed plus click/timer advancement. Source-free slides may author it; imported slides may replace one canonical existing direct transition or add one only when transition.capability.addable is true. Timing, sound, extension, opaque-effect, and every other source graph fail closed.

**Schema parameters:**

- `transition` (object) required — { effect: 'fade'|'push', direction?: 'left'|'up'|'right'|'down', speed?: 'slow'|'medium'|'fast', advanceOnClick?: boolean, advanceAfterMs?: integer 0..86400000 }. Fade rejects direction; push defaults to left; speed defaults to medium; click defaults to true.

**Schema returns:**

- `slide` (Slide) — The same slide with a normalized direct p:transition. Source-free slides may author it. An imported slide may replace exactly one canonical direct fade/push transition, or add one only when transition.capability.addable proves the root contains only p:cSld plus optional p:clrMapOvr and has no transition, timing, or extension leaf. Opaque source graphs are not reconstructed.

#### `slide.shapes.add`

Add a shape/textbox with preset or bounded literal custom geometry, position, optional center-based rotation/flips, fill, line, text, and DrawingML text-body layout.

**Schema parameters:**

- `name` (string) — Inspectable shape name.
- `geometry` (string) — rect, ellipse, roundRect, textbox, or custom. Custom requires customPaths.
- `customPaths` (object[]) — For geometry custom, 1-64 literal DrawingML paths with positive integer width/height and bounded moveTo, lineTo, cubicBezTo, and close commands. Guides, handles, connection sites, arcs, quadratic curves, text rectangles, and path-specific paint overrides are not authored.
- `position` (object) — Pixel left/top/width/height frame.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } center transform. Rotation is bounded to -360 through 360 degrees and flip booleans retain explicit false. OpenChestnut authors/imports this direct DrawingML transform on supported shapes; complex or unknown native transform graphs remain read-only.
- `text` (string|string[]|object|object[]) — Plain text or structured paragraphs accepted by shape.text.set, including ordered text/field/line-break inlines, paragraph tab stops, styles, and relationship-backed hyperlinks.
- `textBodyProperties` (object) — DrawingML text-frame layout: pixel insets; anchor/wrap/AutoFit; -360..360 degree rotation; horizontal/vertical/vertical270 text; horizontal/vertical overflow; 1-16 columns with pixel spacing and RTL flow; and upright text.
- `fill` (string|object) — Shape fill.
- `line` (object) — Line color, width, dash, and arrow metadata.
- `placeholder` (object) — Optional layout placeholder metadata.

**Schema returns:**

- `shape` (Shape) — Appended editable shape/textbox.

#### `slide.speakerNotes.capability`

Return defensive sourceBound, partPresent, editable, and addable evidence. addable identifies an imported notes-absent slide whose source NotesMaster/SlideMaster Theme graph can safely receive a canonical NotesSlide. Export independently re-proves the package graph, so mutating model or wire data cannot grant authority.

**Schema returns:**

- `capability` (object) — Defensive { sourceBound, partPresent, editable, addable } evidence. addable is true only for an imported notes-absent slide whose presentation graph is safely extensible. It is Agent preflight evidence, not mutable write authority; OpenChestnut independently re-proves the source package before export.

#### `slide.tables.add`

Add an inspectable table facade with rows, columns, values, cells, rectangular merges, layout JSON, SVG preview, and canonical OpenChestnut plain-text PPTX output.

**Schema parameters:**

- `values` (unknown[][]) required — Table cell value matrix.
- `name` (string) — Inspectable table name.
- `position` (object) — Pixel left/top/width/height frame.
- `style` (object) — Table/cell fill, margins, borders, and text style.
- `styleOptions` (object) — Optional headerRow and bandedRows booleans plus model-rendering font options. OpenChestnut authors the two native flags, but keeps them immutable after source-bound import.

**Schema returns:**

- `table` (TableElement) — Appended editable table facade. OpenChestnut accepts a non-empty rectangular 1-256-column by 1-2048-row plain-text grid with non-overlapping rectangular merges; recognized imports may change name, complete frame, and visible origin/unmerged cell text without changing merge topology or native style flags.

#### `slideCommentThread.addReply`

Append a direct reply to a source-free Office 2021 modern comment thread. Imported reply topology is fixed: existing reply text/status may change, but adding or removing replies fails closed.

**Schema parameters:**

- `text` (string) required — Direct reply text.
- `author` (string) — Reply display author.
- `person` (object) — Modern author identity with id/name/initials/userId/providerId.
- `created` (string) — ISO-8601 timestamp.
- `status` (string) — active, resolved, or closed; defaults to active.

**Schema returns:**

- `thread` (SlideCommentThread) — Append one direct source-free modern reply and return the thread. Imported topology changes fail closed.

#### `slideCommentThread.reopen`

Set the modern root comment status back to active while preserving fixed imported identity, anchor, position, and reply topology.

**Schema returns:**

- `thread` (SlideCommentThread) — Set resolved=false and the modern root status to active. Legacy comments cannot encode this state.

#### `slideCommentThread.resolve`

Set the modern root comment status to resolved. Imported export re-proves author/date/anchor/position/topology and source-part hashes before changing only status.

**Schema returns:**

- `thread` (SlideCommentThread) — Set resolved=true and the modern root status to resolved. Legacy comments cannot encode this state.

#### `table.merge`

Merge one inclusive rectangular table range, retain the upper-left value, clear and lock covered cells, and emit canonical DrawingML merge topology.

**Schema parameters:**

- `range` (object) required — Inclusive zero-based { startRow, endRow, startColumn, endColumn } rectangle. It must span at least two in-bounds cells and cannot overlap an existing merge.

**Schema returns:**

- `table` (TableElement) — The same table after preserving the upper-left value, clearing covered values, and making covered cells read-only. Imported merge topology remains source-bound and cannot be changed.

## shared

| Name | Kind | Summary |
| --- | --- | --- |
| `clearOfficeFontDesignMetrics` | api | Clear process-level and scoped Office font design metrics. |
| `createCanvasRenderer` | api | Create an optional node-canvas renderer adapter from open-office-artifact-tool/renderers/canvas for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG or JPEG. |
| `createLibreOfficeRenderer` | api | Create a LibreOffice CLI renderer adapter from open-office-artifact-tool/renderers/libreoffice for DOCX/XLSX/PPTX/HTML/PDF FileBlob conversion, typically to PDF. |
| `createNativeOfficeRenderer` | api | Create a native Office renderer adapter from open-office-artifact-tool/native/office-bridge that calls a JSON stdin/stdout sidecar command with timeout, temp-file isolation, cleanup, and structured errors. |
| `createPlaywrightRenderer` | api | Create an optional Playwright renderer adapter from open-office-artifact-tool/renderers/playwright for deterministic SVG/HTML to PNG, WebP, JPEG, or PDF conversion with network blocked by default. |
| `createPopplerRenderer` | api | Create a Poppler CLI renderer adapter from open-office-artifact-tool/renderers/poppler for application/pdf FileBlob page rasterization to PNG, PPM, or TIFF. |
| `createSharpRenderer` | api | Create an optional sharp renderer adapter from open-office-artifact-tool/renderers/sharp for SVG/PNG/JPEG/WebP FileBlob raster conversion to PNG, WebP, or JPEG. |
| `registerScopedOfficeFontDesignMetrics` | api | Register a last-in-first-resolved scoped font design-metric collection and return an idempotent disposer. |
| `renderArtifact` | api | Render an artifact through its render/export method, attach normalized FileBlob metadata, and optionally pass SVG output through a caller-provided renderer adapter for PNG/WebP/JPEG/PDF output. |
| `renderFileWithNativeOffice` | api | Render or convert a DOCX/XLSX/PPTX/PDF FileBlob through a configured native Office bridge command, returning a FileBlob for PDF/PNG/WebP or other requested output. |
| `resolveOfficeFontDesignMetrics` | api | Resolve the requested primary family, style, and nearest numeric weight from scoped then process-level font design metrics without silently skipping to later family fallbacks. |
| `setOfficeFontDesignMetrics` | api | Replace the process-level Office font design-metric registry with normalized public metric records used by deterministic layout integrations. |
| `skiaPaintBaselineCompensationPx` | api | Return the signed subpixel residual between a finite paint baseline and its nearest integer pixel, or zero for non-finite input. |
| `verifyArtifact` | api | Run an artifact's verify() method and return a bounded NDJSON QA report. |
| `visualQaArtifact` | api | Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, optionally register small translations, and return a configurable aligned PNG diff heatmap. |

### shared details

#### `clearOfficeFontDesignMetrics`

Clear process-level and scoped Office font design metrics.

**Schema returns:**

- `result` (undefined) — All registered metrics are removed synchronously.

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

#### `registerScopedOfficeFontDesignMetrics`

Register a last-in-first-resolved scoped font design-metric collection and return an idempotent disposer.

**Schema parameters:**

- `entries` (object[]) required — Iterable normalized font design-metric candidates.

**Schema returns:**

- `dispose` (function) — Idempotently removes only this scoped registration.

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

#### `resolveOfficeFontDesignMetrics`

Resolve the requested primary family, style, and nearest numeric weight from scoped then process-level font design metrics without silently skipping to later family fallbacks.

**Schema parameters:**

- `request` (object) required — { family: string[], weight?, style? }; the first family is the explicit lookup target.

**Schema returns:**

- `metric` (object|undefined) — A defensive normalized metric record or undefined.

#### `setOfficeFontDesignMetrics`

Replace the process-level Office font design-metric registry with normalized public metric records used by deterministic layout integrations.

**Schema parameters:**

- `entries` (object[]) required — Iterable records with family, weight, unitsPerEm, ascent, non-negative descent, and optional lineGap/style/width.

**Schema returns:**

- `result` (undefined) — Registry replacement is synchronous.

#### `skiaPaintBaselineCompensationPx`

Return the signed subpixel residual between a finite paint baseline and its nearest integer pixel, or zero for non-finite input.

**Schema parameters:**

- `value` (number) required — Baseline coordinate in CSS pixels.

**Schema returns:**

- `compensation` (number) — A finite residual in the interval [-0.5, 0.5).

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

Render an artifact, compare PNG/JPEG/WebP/PPM decoded pixels against a baseline render, optionally register small translations, and return a configurable aligned PNG diff heatmap.

**Examples:**

- await visualQaArtifact(document, { baseline, pixelDiff: true, minBytes: 100 })

**Options:**

- baseline/expected/baselineBlob
- pixelDiff
- diffImage
- diffPalette
- diffAlignment
- pixelRegistration
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
- `pixelRegistration` (boolean|number|object) — Optionally search a bounded baseline translation (up to 8 pixels) before comparison; records sampled and exact before/after metrics plus ignored edge pixels.
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
| `exportXlsxWithOpenChestnut` | api | Export the bounded Workbook model through the bundled C# Open XML SDK WebAssembly codec: cells, formulas, styles, merges, dimensions, freezes, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, standard Office 2010 line/column/stacked sparklines, and canonical one-variable or two-variable What-If data tables. Imported QueryTables permit only source-bound one-way refresh hardening through table.setQueryRefreshPolicy; connections, commands, fields, sorts, topology, dynamic-array topology, pivots, and unsupported extension graphs are preservation-only or fail closed. |
| `fx.ABS` | formula | Return the absolute value of a number. |
| `fx.AND` | formula | Return TRUE when all conditions are true. |
| `fx.AVERAGE` | formula | Average numeric values across arguments and ranges in the clean-room formula engine. |
| `fx.AVERAGEIF` | formula | Average values whose corresponding entries match case-insensitive comparison or wildcard criteria. |
| `fx.AVERAGEIFS` | formula | Average values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.CEILING` | formula | Round a number up to the nearest significance. |
| `fx.CHOOSECOLS` | formula | Select and reorder one or more 1-based or negative column indexes from an array. |
| `fx.CHOOSEROWS` | formula | Select and reorder one or more 1-based or negative row indexes from an array. |
| `fx.CONCAT` | formula | Concatenate text values and ranges. |
| `fx.COUNT` | formula | Count numeric values across arguments and ranges. |
| `fx.COUNTA` | formula | Count non-empty values across arguments and ranges, including text, logical values, errors, and empty-text formula results. |
| `fx.COUNTBLANK` | formula | Count blank cells and formula results that are empty text in one range. |
| `fx.COUNTIF` | formula | Count values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcard semantics. |
| `fx.COUNTIFS` | formula | Count rows where multiple criteria ranges of the same size match case-insensitive comparison or wildcard criteria. |
| `fx.CUMIPMT` | formula | Calculate cumulative interest paid across a bounded inclusive range of constant-payment loan periods. |
| `fx.CUMPRINC` | formula | Calculate cumulative principal paid across a bounded inclusive range of constant-payment loan periods. |
| `fx.DATE` | formula | Return an Excel serial in the workbook's 1900 or 1904 date system, with overflow and 1900 serial-60 compatibility. |
| `fx.DATEVALUE` | formula | Convert deterministic ISO or English month-name date text to a serial in the workbook's 1900 or 1904 date system; ambiguous locale-numeric dates return #VALUE!. |
| `fx.DAY` | formula | Return the day component of a serial in the workbook's date system, including 1900 compatibility serial 60. |
| `fx.DAYS` | formula | Return the whole-day difference between two Excel date serials. |
| `fx.DB` | formula | Calculate one fixed-declining-balance depreciation period with an optional first-year month count. |
| `fx.DDB` | formula | Calculate one double-declining-balance depreciation period with an optional positive factor. |
| `fx.DROP` | formula | Drop rows and optional columns from the start or end of an array and spill the remainder. |
| `fx.EDATE` | formula | Shift a serial date by whole months and clamp the day to the target month end. |
| `fx.EOMONTH` | formula | Return the final date serial of a month offset from a start date. |
| `fx.EXPAND` | formula | Expand an array to requested row and column dimensions with optional padding. |
| `fx.FILTER` | formula | Filter rows from a source range with a boolean or comparison include array and spill the matching rows. |
| `fx.FIND` | formula | Return the 1-based position of a case-sensitive literal text sequence. |
| `fx.FLOOR` | formula | Round a number down to the nearest significance. |
| `fx.FV` | formula | Calculate the future value of a finite constant-payment stream from rate, term, payment, optional present value, and payment timing. |
| `fx.HLOOKUP` | formula | Look up one scalar in the first row of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key row and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-row index returns #REF!. |
| `fx.HOUR` | formula | Return the 0 through 23 hour component from a nonnegative serial or supported time text. |
| `fx.HSTACK` | formula | Append arrays horizontally, padding shorter arrays with #N/A to the maximum row count. |
| `fx.IF` | formula | Return one value when a condition is true and another when false. |
| `fx.IFERROR` | formula | Return a fallback value when an expression evaluates to a formula error. |
| `fx.IFNA` | formula | Return a fallback only when an expression evaluates to #N/A; preserve every other result or error. |
| `fx.IFS` | formula | Evaluate condition/value pairs in order and return the first matching value, or #N/A when no condition matches. |
| `fx.INDEX` | formula | Select one value from a nonempty rectangular range of at most 10,000 cells with host-compatible row and optional column selectors, preserving an error-valued selector such as a failed MATCH. Only the documented 2- or 3-argument array/range form is modeled; missing or extra selectors and oversized ranges return #VALUE!, while a missing or out-of-range source cell returns #REF!. |
| `fx.INT` | formula | Round a number down to the nearest integer. |
| `fx.IPMT` | formula | Calculate the interest component of one constant-payment loan period from finite rate, period, term, present value, optional future value, and payment-timing inputs. |
| `fx.IRR` | formula | Return a bounded-convergence periodic return rate for a finite cash-flow vector. |
| `fx.ISBLANK` | formula | Return TRUE when a referenced value is empty. |
| `fx.ISERR` | formula | Return TRUE for recognized formula errors other than #N/A. |
| `fx.ISERROR` | formula | Return TRUE when a value is any recognized formula error. |
| `fx.ISNA` | formula | Return TRUE only when a value is the #N/A error. |
| `fx.ISNUMBER` | formula | Return TRUE when a value is numeric. |
| `fx.ISTEXT` | formula | Return TRUE when a value is text and not a formula error. |
| `fx.LARGE` | formula | Return the k-th largest numeric value in an array or range. |
| `fx.LEFT` | formula | Return characters from the start of a text value. |
| `fx.LEN` | formula | Return the length of a text value. |
| `fx.LOWER` | formula | Convert text to lowercase. |
| `fx.MATCH` | formula | Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells. Exact 0 matching is wildcard-aware; default/1 approximate matching requires a proven ascending homogeneous numeric or text vector and returns the greatest matching-or-lower key, while -1 requires proven descending order and returns the smallest matching-or-higher key. Two-dimensional, oversized, mixed, unordered, or invalid-mode inputs return #VALUE!. |
| `fx.MAX` | formula | Return the maximum numeric value across arguments and ranges. |
| `fx.MAXIFS` | formula | Return the largest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.MEDIAN` | formula | Return the middle numeric value, or the average of the two middle values, across arguments and ranges. |
| `fx.MID` | formula | Return characters from the middle of a text value. |
| `fx.MIN` | formula | Return the minimum numeric value across arguments and ranges. |
| `fx.MINIFS` | formula | Return the smallest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.MINUTE` | formula | Return the 0 through 59 minute component from a nonnegative serial or supported time text. |
| `fx.MIRR` | formula | Calculate a modified periodic internal rate of return using distinct finance and reinvestment rates for a finite cash-flow vector. |
| `fx.MODE.SNGL` | formula | Return the most frequently occurring numeric value, or #N/A when no value repeats. |
| `fx.MONTH` | formula | Return the month component of a serial in the workbook's 1900 or 1904 date system. |
| `fx.NA` | formula | Return the #N/A error value to mark unavailable data explicitly. |
| `fx.NETWORKDAYS` | formula | Count Monday-through-Friday dates inclusively between two serial dates, excluding optional holidays. |
| `fx.NETWORKDAYS.INTL` | formula | Count inclusive workdays with a numbered or Monday-first seven-character custom weekend and optional holidays. |
| `fx.NOT` | formula | Reverse the truth value of a condition. |
| `fx.NPER` | formula | Solve the finite payment-period count from rate, payment, present value, optional future value, and payment timing. |
| `fx.NPV` | formula | Discount a finite periodic cash-flow vector beginning one period after the present value date. |
| `fx.OR` | formula | Return TRUE when any condition is true. |
| `fx.PMT` | formula | Calculate a constant-period loan payment from finite rate, term, present value, optional future value, and payment-timing inputs. |
| `fx.PPMT` | formula | Calculate the principal component of one constant-payment loan period using the same bounded inputs as IPMT. |
| `fx.PV` | formula | Calculate the present value of a finite constant-payment stream from rate, term, payment, optional future value, and payment timing. |
| `fx.RANK.EQ` | formula | Return a number's equal rank in a numeric range, descending by default or ascending when order is nonzero. |
| `fx.RATE` | formula | Solve a bounded periodic interest rate from an integer payment term, payment, present value, optional future value, payment timing, and optional guess. |
| `fx.RIGHT` | formula | Return characters from the end of a text value. |
| `fx.ROUND` | formula | Round a numeric value to decimal places or, with negative digits, positions left of the decimal point. |
| `fx.ROUNDDOWN` | formula | Round a numeric value toward zero at the requested positive or negative digit position. |
| `fx.ROUNDUP` | formula | Round a numeric value away from zero at the requested positive or negative digit position. |
| `fx.SEARCH` | formula | Return the 1-based position of case-insensitive text, supporting Excel ?, *, and ~ wildcard syntax. |
| `fx.SECOND` | formula | Return the 0 through 59 second component from a nonnegative serial or supported time text. |
| `fx.SEQUENCE` | formula | Return a dynamic array sequence that spills into neighboring cells in the clean-room formula engine. |
| `fx.SLN` | formula | Calculate straight-line depreciation from cost, salvage value, and useful life. |
| `fx.SMALL` | formula | Return the k-th smallest numeric value in an array or range. |
| `fx.SORT` | formula | Sort a range by a 1-based column index and spill the sorted rows. |
| `fx.SUM` | formula | Sum numeric values across arguments and ranges. |
| `fx.SUMIF` | formula | Sum corresponding values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcards. |
| `fx.SUMIFS` | formula | Sum values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria. |
| `fx.SUMPRODUCT` | formula | Multiply corresponding numeric values in equally sized arrays and return the sum of those products; bounded same-shape direct-range predicate factors support comparisons, unary signs, and scalar arithmetic within SUMPRODUCT. |
| `fx.SWITCH` | formula | Match an expression against ordered value/result pairs and return an optional default or #N/A when no value matches. |
| `fx.TAKE` | formula | Take rows and optional columns from the start or end of an array and spill the result. |
| `fx.TEXT` | formula | Format an Excel serial date as text with the bounded yyyy, yy, m/mm/mmm/mmmm, and d/dd token profile and literal separators. |
| `fx.TEXTJOIN` | formula | Join text values with a delimiter and optional empty-value skipping. |
| `fx.TIME` | formula | Return a time fraction from hour, minute, and second values from 0 through 32767, carrying overflow and wrapping at 24 hours. |
| `fx.TIMEVALUE` | formula | Convert deterministic 12-hour or 24-hour time text, optionally following date text, to a fraction of one day. |
| `fx.TOCOL` | formula | Flatten an array into one spilled column, optionally ignoring blanks or errors and scanning by column. |
| `fx.TOROW` | formula | Flatten an array into one spilled row, optionally ignoring blanks or errors and scanning by column. |
| `fx.TRANSPOSE` | formula | Transpose a source range into a spilled dynamic array with spillRange/spillValues inspect metadata. |
| `fx.TRIM` | formula | Trim leading/trailing whitespace and collapse internal whitespace. |
| `fx.UNIQUE` | formula | Return unique rows from a range as a spilled dynamic array. |
| `fx.UPPER` | formula | Convert text to uppercase. |
| `fx.VALUE` | formula | Convert deterministic ASCII numeric text with optional grouping, scientific notation, accounting parentheses, or percent suffix to a number. |
| `fx.VLOOKUP` | formula | Look up one scalar in the first column of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key column and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-column index returns #REF!. |
| `fx.VSTACK` | formula | Append arrays vertically, padding narrower arrays with #N/A to the maximum column count. |
| `fx.WEEKDAY` | formula | Return a weekday number for Excel return types 1, 2, 3, and 11 through 17. |
| `fx.WORKDAY` | formula | Move forward or backward by working days while skipping weekends and optional holidays. |
| `fx.WORKDAY.INTL` | formula | Move by workdays using a numbered or Monday-first seven-character custom weekend and optional holidays. |
| `fx.WRAPCOLS` | formula | Wrap a one-dimensional vector into columns of a requested height, padding the final column when needed. |
| `fx.WRAPROWS` | formula | Wrap a one-dimensional vector into rows of a requested width, padding the final row when needed. |
| `fx.XIRR` | formula | Return a bounded-convergence annualized return rate for date-aligned finite cash flows using a 365-day year. |
| `fx.XLOOKUP` | formula | Look up one scalar in same-shaped one-dimensional row or column vectors of 1 through 10,000 cells; exact, next-smaller, next-larger, wildcard, and first/last linear search modes are modeled, while binary-search modes and mismatched or two-dimensional ranges fail as #VALUE!. |
| `fx.XMATCH` | formula | Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells, with exact, next-smaller, next-larger, wildcard, and forward or reverse linear search modes; two-dimensional, oversized, and binary-search inputs fail as #VALUE!. |
| `fx.XNPV` | formula | Discount date-aligned finite cash flows by actual day offsets from the first date using a 365-day year. |
| `fx.YEAR` | formula | Return the year component of a serial in the workbook's 1900 or 1904 date system. |
| `importXlsxWithOpenChestnut` | api | Import XLSX bytes through OpenChestnut with editable core cells, formulas, styles, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, and bounded numeric-X/Y/positive-Size bubble charts. Imported data-table topology is source-bound and read-only. A recognized source-bound QueryTable can only disable automatic refresh through table.setQueryRefreshPolicy; connections, commands, fields, sorts, topology, non-marker scatter styles, noncanonical bubble profiles, nested/branched replies, mentions, dynamic-array topology, pivots, non-reversible sparkline graphs, and other advanced package content remain source-bound and read-only. |
| `invokeOpenChestnut` | api | Advanced experimental byte-boundary API for invoking the public OpenChestnut codec protocol with generated wire-message objects. |
| `openChestnutStatus` | api | Lazily initialize the bundled OpenChestnut WebAssembly runtime and report its protocol, assembly, and integrity manifest. |
| `range.clear` | api | Clear range contents, formats, or both without silently changing validations, dimensions, or other package graphs. |
| `range.conditionalFormats.add` | api | Add a conditional formatting rule; cellIs/expression/containsText/colorScale plus standard dataBar/iconSet rules cross the public model and OpenChestnut, with computedStyle inspect records, layout JSON visuals, SVG preview, and native XLSX rendering. |
| `range.copyFrom` | api | Copy values, formulas, or complete cells from an equally sized or evenly tiling source range with relative A1 translation. |
| `range.copyTo` | api | Copy this range to an equally sized or evenly tiled destination range. |
| `range.dataValidation` | api | Assign a list, whole, decimal, date, time, text-length, or custom-formula validation rule to a range, including bounded input prompts, error alerts, blank policy, and intuitive list-arrow visibility; use sheet.dataValidations.add({ range, rule }) for the collection form. |
| `range.displayFormulas` | api | Read displayed A1 formulas, including the anchor formula projected across non-editable dynamic-array or legacy-array result cells. |
| `range.fillDown` | api | Copy top-row contents and formatting down the range while translating relative A1 formula references. |
| `range.fillRight` | api | Copy left-column contents and formatting right across the range while translating relative A1 formula references. |
| `range.format` | api | Assign cell styles, symbolic theme/tint/indexed colors, patterned fills, native dimensions, pixel sizing, and hidden axes through a live range format facade. |
| `range.format.autofitColumns` | api | Measure displayed range values deterministically and set native best-fit widths on each selected column. |
| `range.format.autofitRows` | api | Measure explicit/wrapped range text deterministically and set native custom heights on each selected row. |
| `range.formulaInfos` | api | Read per-cell stored/projected formula metadata with editability, spill/array source, anchor, and reference evidence. |
| `range.formulasR1C1` | api | Read or assign R1C1 formulas relative to each target cell while storing canonical A1 formulas. |
| `range.getCell` | api | Select one zero-based cell relative to the current range. |
| `range.getColumn` | api | Select one zero-based column relative to the current range. |
| `range.getCurrentRegion` | api | Expand to the contiguous data region bounded by fully blank rows and columns. |
| `range.getRangeByIndexes` | api | Select a bounded zero-based subrange relative to the current range. |
| `range.getRow` | api | Select one zero-based row relative to the current range. |
| `range.merge` | api | Merge the target range as one region or as separate row-wise regions when across=true. |
| `range.offset` | api | Return an equally sized range shifted by zero-based row and column offsets, rejecting worksheet overflow. |
| `range.resize` | api | Return a range at the same upper-left cell with explicit positive row and column counts. |
| `range.setNumberFormat` | api | Assign one number format or an evenly tiling matrix of Excel-invariant number-format codes. |
| `range.unmerge` | api | Remove merged regions intersecting the target range. |
| `range.write` | api | Write a mixed matrix or one explicit values/formulas/formulasR1C1 payload from the range anchor and return the actual written range. |
| `range.writeValues` | api | Write a one- or two-dimensional value matrix from the range anchor. |
| `sheet.charts.add` | api | Create an inspectable worksheet chart from a range or config; setData(range) infers category series, scatter per-series numeric xValues/xFormula plus y values/formula, or one exact X/Y/positive-Size bubble series. series.fill sets an explicit #RRGGBB solid color, series.line sets bounded RGB color/dash/width (series.stroke is an alias), line/scatter markers set direct symbol/size/RGB fill/bounded outline semantics, lineOptions controls standard/stacked/percent-stacked grouping, smooth interpolation, and direct vary-colors behavior, dataLabels controls plot-level value/category/series-name visibility and bounded position, and xAxis/yAxis configure primary titles, formats, intervals, and linear value bounds. Marker-only scatter rejects series.line/stroke and writes an explicit native no-fill series outline; use marker.line for marker borders. Bubble charts use two numeric axes and reject ambiguous range shortcuts or nonpositive sizes. |
| `sheet.dataTables.__getDefinitions` | api | Return defensive inspectable definitions for the worksheet's canonical What-If data tables, including result range, native anchor, inputs, orientation, and display formula. |
| `sheet.dataTables.add` | api | Create a canonical native Excel What-If data table from a rectangular formula/input grid and one row input, one column input, or both. Excel or another compatible host calculates the result values; the JavaScript evaluator does not simulate TABLE. |
| `sheet.images.add` | api | Create an inspectable worksheet image from a data URL, URI, or prompt with one-cell, two-cell, or absolute pixel geometry plus optional percentage crop, bounded grayscale/luminance/opacity effects, rotation, and horizontal/vertical flips. |
| `sheet.pivotTables.add` | api | Create a native bounded XLSX PivotTable with derived cached output, cache records, and exact axis-item filters, or use richer model-only grouping/calculation/date-filter semantics for inspect and preview. Recognized imports are hash-bound and read-only. |
| `sheet.sparklineGroups.add` | api | Create standard Office 2010 line/column/stacked sparkline groups for inspect, SVG preview, and OpenChestnut XLSX export. Source-free groups use reversible one-dimensional target/source mappings; recognized imported groups support fixed-topology semantic edits while unsupported native graphs remain source-bound. |
| `sheet.tables.add` | api | Create an ordinary worksheet table over an A1 range with headers, columns, totals metadata, style, and bounded filtering/sorting. QueryTable bindings cannot be authored; recognized imported bindings expose only table.setQueryRefreshPolicy for one-way automatic-refresh hardening, while all other QueryTable edits fail closed. |
| `SpreadsheetFile.exportCsv` | api | Export one worksheet or range as UTF-8 CSV, using calculated values unless formula output is explicitly requested. |
| `SpreadsheetFile.exportDelimited` | api | Serialize one workbook sheet/range as bounded CSV/TSV text with calculated-value defaults and RFC-style quoting. |
| `SpreadsheetFile.exportTsv` | api | Export one worksheet or range as UTF-8 tab-separated text with RFC-style quoting where needed. |
| `SpreadsheetFile.exportXlsx` | api | Serialize a Workbook facade through the single bundled OpenChestnut codec. |
| `SpreadsheetFile.importCsv` | api | Import UTF-8 CSV bytes into an editable Workbook through the bounded delimited parser. |
| `SpreadsheetFile.importDelimited` | api | Parse bounded RFC-style CSV/TSV bytes into an editable Workbook, including quoted delimiters, escaped quotes, and embedded newlines. |
| `SpreadsheetFile.importTsv` | api | Import UTF-8 tab-separated bytes into an editable Workbook through the bounded delimited parser. |
| `SpreadsheetFile.importXlsx` | api | Load XLSX through the single bundled OpenChestnut codec into an editable Workbook facade. |
| `SpreadsheetFile.inspectDelimited` | api | Inspect bounded CSV/TSV bytes as file/row records with dimensions, delimiter, quoting, and formula-like cell evidence. |
| `SpreadsheetFile.inspectXlsx` | api | Inspect bounded XLSX parts, content types, relationships, and namespace-aware source XML r:id/r:embed/r:link references under decompression budgets. |
| `SpreadsheetFile.patchXlsx` | api | Apply path-validated XLSX part patches, build worksheet/table/drawing/image/chart/pivot source references, and atomically reject dangling content types or relationships. |
| `table.setQueryRefreshPolicy` | api | On one recognized imported QueryTable, monotonically disable automatic refresh without changing its connection, command, fields, sort, refresh history, or topology. |
| `thread.addReply` | api | Append a direct reply to an Office threaded-comment root with independent author/person/date/done metadata. Nested or branched reply graphs and mentions fail closed. |
| `workbook.comments.addThread` | api | Create one root Office threaded comment per thread with GUID/person metadata, date, and resolved state; attach bounded direct replies with thread.addReply(). |
| `workbook.connections` | api | Inspect bounded non-secret metadata for imported database connections. Connections are source-bound and read-only: authoring, removal, or mutation makes canonical XLSX export fail closed. |
| `Workbook.create` | api | Create an empty workbook with an explicit date system and optional native SpreadsheetML theme colors. |
| `workbook.definedNames.add` | api | Create a workbook or sheet-scoped defined name over an A1 range; exported as native workbook.xml definedName and usable in formulas such as SUM(RevenueData). |
| `workbook.fontFamilies` | api | Return a fresh sorted, case-insensitively deduplicated list of workbook default and explicit cell font families. |
| `workbook.formulaGraph` | api | Return a bounded dependency graph of formula nodes, edges, dependents, cycles, formula errors, and syntax-input/reference-budget refusals for workbook QA. |
| `workbook.inspect` | api | Emit bounded NDJSON records for workbook, connections, sheets, worksheet protections, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude. |
| `workbook.layoutJson` | api | Return workbook/worksheet layout JSON with cell, table, chart, image, sparkline, rule bounding boxes, and target/search context slicing. |
| `workbook.recalculate` | api | Recalculate bounded workbook formulas and dynamic-array spills, with dependency edges, cycles, errors, and syntax-input/reference-budget refusals. |
| `workbook.render` | api | Return a lightweight SVG preview for a sheet/range or layout JSON when called with { format: 'layout' }. |
| `workbook.resolve` | api | Resolve stable workbook, source-bound connection, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs. |
| `workbook.setCalculation` | api | Set bounded workbook-level SpreadsheetML calculation mode, on-save/full-recalculation flags, iterative-calculation limits, and full-precision policy. |
| `workbook.setDateSystem` | api | Select the Excel 1900 or 1904 serial-date system for formula calculation and native workbookPr export. |
| `workbook.sharedArrayFormulas` | formula | Import and export bounded shared and legacy-array formula metadata. XLDAPR dynamic-array anchors are inspectable after import but source-bound and read-only; creating, detaching, or editing their topology makes XLSX export fail closed. |
| `workbook.spillReferences` | formula | Use a direct or defined-name A1# reference to consume only an anchor's current, unblocked dynamic spill matrix. Supported range consumers and a direct re-spill read the verified matrix; scalar/general-vector coercion returns #VALUE!, non-spilling anchors return #REF!, and graph/trace record one spillReference edge to the anchor. |
| `workbook.structuredReferences` | formula | Evaluate Excel table references including sections, column ranges/unions, space intersections, escaped special-character headers, unqualified calculated-column references, and @/#This Row context while expanding exact table-cell precedents. |
| `workbook.trace` | api | Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references and syntax-input/reference-budget refusals flagged. |
| `workbook.verify` | api | Return bounded QA issues for source-bound connections, sheets, formulas (including syntax-input and reference-budget refusals), tables, charts, and comments. |
| `workbook.windows` | api | Access the ordered workbook-window collection; window 0 is the primary view used by legacy worksheet-selection APIs. |
| `workbook.windows.add` | api | Append an additional workbook window with its own active worksheet and selected tab group. |
| `workbook.worksheets.add` | api | Append an editable visible, hidden, or very-hidden worksheet with a stable name and ID. |
| `workbook.worksheets.getSelectedWorksheets` | api | Return the visible worksheet-tab group selected in the primary workbook window, in workbook order. |
| `workbook.worksheets.setActiveWorksheet` | api | Select the visible worksheet opened by default and used by workbook operations that omit an explicit sheet. |
| `workbook.worksheets.setSelectedWorksheets` | api | Select one or more visible worksheet tabs in the primary workbook window while retaining exactly one active worksheet. |
| `workbookWindow.getActiveWorksheet` | api | Return the visible active worksheet for one workbook window. |
| `workbookWindow.getSelectedWorksheets` | api | Return one window's visible selected worksheet tabs in workbook order. |
| `workbookWindow.setActiveWorksheet` | api | Set one window's active worksheet and collapse that window's selected tab group to it. |
| `workbookWindow.setSelectedWorksheets` | api | Set one window's non-empty visible selected tab group, which must include its active worksheet. |
| `worksheet.freezePanes.freezeColumns` | api | Freeze a leading column count in the worksheet view while preserving any frozen rows. |
| `worksheet.freezePanes.freezeRows` | api | Freeze a leading row count in the worksheet view while preserving any frozen columns. |
| `worksheet.freezePanes.unfreeze` | api | Remove all frozen worksheet panes and restore a single scrollable view. |
| `worksheet.getRange` | api | Select an A1 range for values, formulas, formatting, merge, fill, and copy operations. |
| `worksheet.getUsedRange` | api | Return the worksheet used rectangle, optionally excluding formatting-only cells with valuesOnly=true. |
| `worksheet.mergeCells` | api | Merge an A1 range as one region or merge each row separately with across=true, retaining only upper-left content. |
| `worksheet.protection` | api | Author, inspect, edit, or remove one passwordless worksheet editing restriction with an intuitive allowed-operation list. Cell locked/hidden styles become effective only while protection is active. This is not encryption or access control; password/hash variants remain source-owned and fail closed on replacement. |
| `worksheet.sortState` | api | Get or set bounded worksheet-level row/column sorting; columnSort=true uses unique single-row conditions across the sort range. |
| `worksheet.unmergeCells` | api | Remove every merged region intersecting an A1 range without discarding the retained upper-left content. |
| `worksheet.visibility` | api | Read or assign native worksheet visibility as visible, hidden, or veryHidden; at least one sheet must remain visible. |

### workbook details

#### `exportXlsxWithOpenChestnut`

Export the bounded Workbook model through the bundled C# Open XML SDK WebAssembly codec: cells, formulas, styles, merges, dimensions, freezes, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, bounded numeric-X/Y/positive-Size bubble charts, standard Office 2010 line/column/stacked sparklines, and canonical one-variable or two-variable What-If data tables. Imported QueryTables permit only source-bound one-way refresh hardening through table.setQueryRefreshPolicy; connections, commands, fields, sorts, topology, dynamic-array topology, pivots, and unsupported extension graphs are preservation-only or fail closed.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade within the core cell/formula/style/merge/dimension/freeze/ordinary-table/image/validation/conditional-format/root-plus-direct-reply-comment/bar-line-pie-chart/standard-sparkline boundary. A recognized imported QueryTable may only receive one-way automatic-refresh hardening through table.setQueryRefreshPolicy; imported connections, commands, fields, sorts, topology, nested reply graphs, mentions, dynamic-array topology, pivots, non-reversible sparkline graphs, and other advanced package graphs must remain unchanged or fail closed.
- `recalculate` (boolean) — Recalculate formulas before serialization; defaults to true.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — XLSX bytes produced by the bundled Open XML SDK WebAssembly codec, with codec diagnostics in metadata.

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

Average values whose corresponding entries match case-insensitive comparison or wildcard criteria.

**Examples:**

- =AVERAGEIF(A1:A10,"East*",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =AVERAGEIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.AVERAGEIFS`

Average values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =AVERAGEIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

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

#### `fx.COUNTA`

Count non-empty values across arguments and ranges, including text, logical values, errors, and empty-text formula results.

**Examples:**

- =COUNTA(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTA(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTBLANK`

Count blank cells and formula results that are empty text in one range.

**Examples:**

- =COUNTBLANK(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTBLANK(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIF`

Count values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcard semantics.

**Examples:**

- =COUNTIF(A1:A10,"East*")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.COUNTIFS`

Count rows where multiple criteria ranges of the same size match case-insensitive comparison or wildcard criteria.

**Examples:**

- =COUNTIFS(A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =COUNTIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.CUMIPMT`

Calculate cumulative interest paid across a bounded inclusive range of constant-payment loan periods.

**Examples:**

- =CUMIPMT(B1,B2,B3,1,12,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CUMIPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- All six arguments are required. The bounded evaluator requires positive rate and present value, payment type 0 or 1, and integer start/end periods ordered from 1 through the term; the ending period is capped at 9,999. Invalid inputs return #VALUE! or #NUM! rather than coercing a range.

#### `fx.CUMPRINC`

Calculate cumulative principal paid across a bounded inclusive range of constant-payment loan periods.

**Examples:**

- =CUMPRINC(B1,B2,B3,1,12,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =CUMPRINC(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- All six arguments are required. The bounded evaluator shares CUMIPMT's positive-rate, positive-present-value, integer-period, bounded-end, and payment-timing contract; it returns the signed principal cash flow.

#### `fx.DATE`

Return an Excel serial in the workbook's 1900 or 1904 date system, with overflow and 1900 serial-60 compatibility.

**Examples:**

- =DATE(2026,7,12)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DATEVALUE`

Convert deterministic ISO or English month-name date text to a serial in the workbook's 1900 or 1904 date system; ambiguous locale-numeric dates return #VALUE!.

**Examples:**

- =DATEVALUE("2026-07-13")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DATEVALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DAY`

Return the day component of a serial in the workbook's date system, including 1900 compatibility serial 60.

**Examples:**

- =DAY(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DAYS`

Return the whole-day difference between two Excel date serials.

**Examples:**

- =DAYS(B1,A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DAYS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.DB`

Calculate one fixed-declining-balance depreciation period with an optional first-year month count.

**Examples:**

- =DB(B1,B2,B3,A2)
- =DB(B1,B2,B3,A2,6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DB(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires nonnegative cost and salvage, salvage no greater than cost, integer life and period from 1 through 9,999, and an integer month from 1 through 12. A partial first year permits one prorated final period; the native three-decimal declining rate is not silently switched to straight-line.

#### `fx.DDB`

Calculate one double-declining-balance depreciation period with an optional positive factor.

**Examples:**

- =DDB(B1,B2,B3,A2)
- =DDB(B1,B2,B3,A2,1.5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DDB(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires nonnegative cost and salvage, salvage no greater than cost, and integer life and period from 1 through 9,999. The factor defaults to 2, must be positive, and depreciation is capped at the remaining amount above salvage without a silent straight-line switch.

#### `fx.DROP`

Drop rows and optional columns from the start or end of an array and spill the remainder.

**Examples:**

- =DROP(A2:C10,1,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =DROP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.EDATE`

Shift a serial date by whole months and clamp the day to the target month end.

**Examples:**

- =EDATE(A1,3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EDATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.EOMONTH`

Return the final date serial of a month offset from a start date.

**Examples:**

- =EOMONTH(A1,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =EOMONTH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.FIND`

Return the 1-based position of a case-sensitive literal text sequence.

**Examples:**

- =FIND("Review",A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FIND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FLOOR`

Round a number down to the nearest significance.

**Examples:**

- =FLOOR(A1,5)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FLOOR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.FV`

Calculate the future value of a finite constant-payment stream from rate, term, payment, optional present value, and payment timing.

**Examples:**

- =FV(B1,B2,B3)
- =FV(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =FV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive finite term, and payment type 0 or 1. It uses the same cash-flow equation as PMT and PV, including the zero-rate case.

#### `fx.HLOOKUP`

Look up one scalar in the first row of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key row and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-row index returns #REF!.

**Examples:**

- =HLOOKUP("Revenue",A1:D4,3,FALSE)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.HOUR`

Return the 0 through 23 hour component from a nonnegative serial or supported time text.

**Examples:**

- =HOUR(TIMEVALUE("6:45 PM"))

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =HOUR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.IFS`

Evaluate condition/value pairs in order and return the first matching value, or #N/A when no condition matches.

**Examples:**

- =IFS(A1>=90,"A",A1>=80,"B",TRUE,"C")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.INDEX`

Select one value from a nonempty rectangular range of at most 10,000 cells with host-compatible row and optional column selectors, preserving an error-valued selector such as a failed MATCH. Only the documented 2- or 3-argument array/range form is modeled; missing or extra selectors and oversized ranges return #VALUE!, while a missing or out-of-range source cell returns #REF!.

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

#### `fx.IPMT`

Calculate the interest component of one constant-payment loan period from finite rate, period, term, present value, optional future value, and payment-timing inputs.

**Examples:**

- =IPMT(B1,A2,B2,B3)
- =IPMT(B1,A2,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive term, an integer period from 1 through the term, and payment type 0 or 1. Period-one interest is zero for payment type 1; invalid inputs return #VALUE! or #NUM!.

#### `fx.IRR`

Return a bounded-convergence periodic return rate for a finite cash-flow vector.

**Examples:**

- =IRR(B2:B8)
- =IRR(B2:B8,0.15)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =IRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Cash flows must contain both a positive and a negative finite number. The optional finite guess defaults to 0.1; no converged valid root or an invalid rate returns #NUM! rather than a guessed value.

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

#### `fx.LARGE`

Return the k-th largest numeric value in an array or range.

**Examples:**

- =LARGE(A1:A10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =LARGE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells. Exact 0 matching is wildcard-aware; default/1 approximate matching requires a proven ascending homogeneous numeric or text vector and returns the greatest matching-or-lower key, while -1 requires proven descending order and returns the smallest matching-or-higher key. Two-dimensional, oversized, mixed, unordered, or invalid-mode inputs return #VALUE!.

**Examples:**

- =MATCH("Beta*",A2:A10,0)

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

#### `fx.MAXIFS`

Return the largest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =MAXIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MAXIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MEDIAN`

Return the middle numeric value, or the average of the two middle values, across arguments and ranges.

**Examples:**

- =MEDIAN(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MEDIAN(...).
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

#### `fx.MINIFS`

Return the smallest numeric value where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =MINIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MINIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MINUTE`

Return the 0 through 59 minute component from a nonnegative serial or supported time text.

**Examples:**

- =MINUTE(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MINUTE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MIRR`

Calculate a modified periodic internal rate of return using distinct finance and reinvestment rates for a finite cash-flow vector.

**Examples:**

- =MIRR(B2:B6,B7,B8)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MIRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator accepts 2 through 10,000 finite cash flows containing both signs. Finance and reinvestment rates must be greater than -1; negative flows are discounted at the finance rate, positive flows compound at the reinvestment rate, and invalid profiles return #VALUE! or #NUM! rather than choosing an implied rate.

#### `fx.MODE.SNGL`

Return the most frequently occurring numeric value, or #N/A when no value repeats.

**Examples:**

- =MODE.SNGL(A1:A10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MODE.SNGL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.MONTH`

Return the month component of a serial in the workbook's 1900 or 1904 date system.

**Examples:**

- =MONTH(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =MONTH(...).
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

#### `fx.NETWORKDAYS`

Count Monday-through-Friday dates inclusively between two serial dates, excluding optional holidays.

**Examples:**

- =NETWORKDAYS(A1,B1,Holidays)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NETWORKDAYS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.NETWORKDAYS.INTL`

Count inclusive workdays with a numbered or Monday-first seven-character custom weekend and optional holidays.

**Examples:**

- =NETWORKDAYS.INTL(A1,B1,7,Holidays)
- =NETWORKDAYS.INTL(A1,B1,"0000011")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NETWORKDAYS.INTL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.NOT`

Reverse the truth value of a condition.

**Examples:**

- =NOT(A1>0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NOT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.NPER`

Solve the finite payment-period count from rate, payment, present value, optional future value, and payment timing.

**Examples:**

- =NPER(B1,B2,B3)
- =NPER(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NPER(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1 and payment type 0 or 1. It returns a closed-form finite period count, which may be zero or negative for the supplied cash-flow signs; a zero payment at zero rate or an invalid real solution returns #NUM!.

#### `fx.NPV`

Discount a finite periodic cash-flow vector beginning one period after the present value date.

**Examples:**

- =NPV(B1,B2:B6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =NPV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Rate must be finite and greater than -1. The bounded evaluator accepts at most 10,000 finite numeric cash flows and returns #VALUE! or #NUM! rather than coercing malformed inputs.

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

Calculate a constant-period loan payment from finite rate, term, present value, optional future value, and payment-timing inputs.

**Examples:**

- =PMT(B1,B2,B3)
- =PMT(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive term, and payment type 0 or 1. Invalid numeric inputs return #VALUE! or #NUM!.

#### `fx.PPMT`

Calculate the principal component of one constant-payment loan period using the same bounded inputs as IPMT.

**Examples:**

- =PPMT(B1,A2,B2,B3)
- =PPMT(B1,A2,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PPMT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- For every supported period, PMT equals IPMT plus PPMT. The evaluator rejects an out-of-range or non-integer period and invalid payment timing with #NUM! rather than coercing them.

#### `fx.PV`

Calculate the present value of a finite constant-payment stream from rate, term, payment, optional future value, and payment timing.

**Examples:**

- =PV(B1,B2,B3)
- =PV(B1,B2,B3,B4,1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =PV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The bounded evaluator requires rate > -1, a positive finite term, and payment type 0 or 1. It preserves standard cash-flow signs and returns #VALUE! or #NUM! for invalid inputs rather than coercing them.

#### `fx.RANK.EQ`

Return a number's equal rank in a numeric range, descending by default or ascending when order is nonzero.

**Examples:**

- =RANK.EQ(A1,A1:A10,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RANK.EQ(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.RATE`

Solve a bounded periodic interest rate from an integer payment term, payment, present value, optional future value, payment timing, and optional guess.

**Examples:**

- =RATE(B1,B2,B3)
- =RATE(B1,B2,B3,B4,1,0.1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =RATE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The evaluator accepts an integer term from 1 through 9,999, type 0 or 1, and a finite guess greater than -1 (default 0.1). It finds a converged rate greater than -1 nearest the guess or returns #NUM! rather than inventing a rate.

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

Round a numeric value to decimal places or, with negative digits, positions left of the decimal point.

**Examples:**

- =ROUND(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUND(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUNDDOWN`

Round a numeric value toward zero at the requested positive or negative digit position.

**Examples:**

- =ROUNDDOWN(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUNDDOWN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.ROUNDUP`

Round a numeric value away from zero at the requested positive or negative digit position.

**Examples:**

- =ROUNDUP(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =ROUNDUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SEARCH`

Return the 1-based position of case-insensitive text, supporting Excel ?, *, and ~ wildcard syntax.

**Examples:**

- =SEARCH("review",A1)
- =SEARCH("Re*W",A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SEARCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SECOND`

Return the 0 through 59 second component from a nonnegative serial or supported time text.

**Examples:**

- =SECOND(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SECOND(...).
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

#### `fx.SLN`

Calculate straight-line depreciation from cost, salvage value, and useful life.

**Examples:**

- =SLN(B1,B2,B3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SLN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- The evaluator accepts finite numeric inputs and returns #DIV/0! for zero life rather than coercing a rate. This is the direct per-period expense, not a declining-balance schedule.

#### `fx.SMALL`

Return the k-th smallest numeric value in an array or range.

**Examples:**

- =SMALL(A1:A10,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SMALL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

Sum corresponding values using case-insensitive numeric/text criteria and Excel ?, *, and ~ wildcards.

**Examples:**

- =SUMIF(A1:A10,"East*",B1:B10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIF(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMIFS`

Sum values where all supplied criteria ranges have the same size and match case-insensitive comparison or wildcard criteria.

**Examples:**

- =SUMIFS(C1:C10,A1:A10,"East*",B1:B10,">=10")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMIFS(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SUMPRODUCT`

Multiply corresponding numeric values in equally sized arrays and return the sum of those products; bounded same-shape direct-range predicate factors support comparisons, unary signs, and scalar arithmetic within SUMPRODUCT.

**Examples:**

- =SUMPRODUCT(A1:A10,B1:B10)
- =SUMPRODUCT(C1:C10,--(A1:A10="Open"))

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SUMPRODUCT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.SWITCH`

Match an expression against ordered value/result pairs and return an optional default or #N/A when no value matches.

**Examples:**

- =SWITCH(A1,"East",1,"West",2,0)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =SWITCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (boolean) — Calculated cell value or an Excel-style formula error string.

#### `fx.TAKE`

Take rows and optional columns from the start or end of an array and spill the result.

**Examples:**

- =TAKE(A2:C10,3,-2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TAKE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown[][]) — Spilled two-dimensional formula result.

#### `fx.TEXT`

Format an Excel serial date as text with the bounded yyyy, yy, m/mm/mmm/mmmm, and d/dd token profile and literal separators.

**Examples:**

- =TEXT(DATE(2026,7,12),"yyyymmdd")
- =TEXT(A1,"mmm yyyy")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXT(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TEXTJOIN`

Join text values with a delimiter and optional empty-value skipping.

**Examples:**

- =TEXTJOIN("/",TRUE,A1:A3)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TEXTJOIN(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (string) — Calculated cell value or an Excel-style formula error string.

#### `fx.TIME`

Return a time fraction from hour, minute, and second values from 0 through 32767, carrying overflow and wrapping at 24 hours.

**Examples:**

- =TIME(16,48,10)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TIME(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.TIMEVALUE`

Convert deterministic 12-hour or 24-hour time text, optionally following date text, to a fraction of one day.

**Examples:**

- =TIMEVALUE("6:45 PM")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =TIMEVALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.VALUE`

Convert deterministic ASCII numeric text with optional grouping, scientific notation, accounting parentheses, or percent suffix to a number.

**Examples:**

- =VALUE("1,234.50")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =VALUE(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.VLOOKUP`

Look up one scalar in the first column of a nonempty rectangular range of at most 10,000 cells; FALSE/0 performs an exact, wildcard-aware lookup, while TRUE/1 or omission requires a proven ascending homogeneous numeric or text key column and returns the greatest matching-or-lower key. Invalid table/mode/index inputs and unproven ordering return #VALUE!, while an out-of-range return-column index returns #REF!.

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

#### `fx.WEEKDAY`

Return a weekday number for Excel return types 1, 2, 3, and 11 through 17.

**Examples:**

- =WEEKDAY(A1,2)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WEEKDAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WORKDAY`

Move forward or backward by working days while skipping weekends and optional holidays.

**Examples:**

- =WORKDAY(A1,10,Holidays)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WORKDAY(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.WORKDAY.INTL`

Move by workdays using a numbered or Monday-first seven-character custom weekend and optional holidays.

**Examples:**

- =WORKDAY.INTL(A1,10,11,Holidays)
- =WORKDAY.INTL(A1,10,"0000011")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =WORKDAY.INTL(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

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

#### `fx.XIRR`

Return a bounded-convergence annualized return rate for date-aligned finite cash flows using a 365-day year.

**Examples:**

- =XIRR(B2:B8,C2:C8)
- =XIRR(B2:B8,C2:C8,0.15)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XIRR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Values and dates must have the same nonzero count, dates must be valid, and cash flows must contain both signs. The optional finite guess defaults to 0.1; invalid or unconverged cases return #VALUE! or #NUM!.

#### `fx.XLOOKUP`

Look up one scalar in same-shaped one-dimensional row or column vectors of 1 through 10,000 cells; exact, next-smaller, next-larger, wildcard, and first/last linear search modes are modeled, while binary-search modes and mismatched or two-dimensional ranges fail as #VALUE!.

**Examples:**

- =XLOOKUP("Gamma",A2:A4,B2:B4,"missing")

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XLOOKUP(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (unknown) — Calculated cell value or an Excel-style formula error string.

#### `fx.XMATCH`

Return a 1-based lookup position in one row or column vector of 1 through 10,000 cells, with exact, next-smaller, next-larger, wildcard, and forward or reverse linear search modes; two-dimensional, oversized, and binary-search inputs fail as #VALUE!.

**Examples:**

- =XMATCH("Beta*",A2:A10,2,-1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XMATCH(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `fx.XNPV`

Discount date-aligned finite cash flows by actual day offsets from the first date using a 365-day year.

**Examples:**

- =XNPV(B1,B2:B6,C2:C6)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =XNPV(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

**Notes:**

- Values and dates must have the same nonzero count; each date must be valid in the workbook date system. Rate must be greater than -1 and the vector is bounded to 10,000 entries.

#### `fx.YEAR`

Return the year component of a serial in the workbook's 1900 or 1904 date system.

**Examples:**

- =YEAR(A1)

**Schema parameters:**

- `formula` (string) required — Excel-style cell formula beginning with =YEAR(...).
- `arguments` (unknown[]) required — Function arguments may contain literals, cell references, ranges, arrays, or nested formulas as supported by the clean-room evaluator.

**Schema returns:**

- `value` (number) — Calculated cell value or an Excel-style formula error string.

#### `importXlsxWithOpenChestnut`

Import XLSX bytes through OpenChestnut with editable core cells, formulas, styles, ordinary tables, PNG/JPEG pictures, validation, conditional formatting, threaded-comment roots with direct replies, bar/line/pie/area/doughnut charts, marker-only numeric-X/Y scatter charts, and bounded numeric-X/Y/positive-Size bubble charts. Imported data-table topology is source-bound and read-only. A recognized source-bound QueryTable can only disable automatic refresh through table.setQueryRefreshPolicy; connections, commands, fields, sorts, topology, non-marker scatter styles, noncanonical bubble profiles, nested/branched replies, mentions, dynamic-array topology, pivots, non-reversible sparkline graphs, and other advanced package content remain source-bound and read-only.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|ArrayBuffer) required — XLSX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `workbook` (Workbook) — Imported bounded workbook facade with editable core objects, canonical Office 2010 sparkline groups, and source/opaque package evidence. A recognized QueryTable permits only table.setQueryRefreshPolicy automatic-refresh hardening; connections, commands, fields, sorts, QueryTable topology, dynamic-array topology, pivots, non-reversible sparkline graphs, and unsupported package graphs are exposed only for inspection or preserved unchanged.

#### `invokeOpenChestnut`

Advanced experimental byte-boundary API for invoking the public OpenChestnut codec protocol with generated wire-message objects.

**Schema parameters:**

- `request` (object) required — Generated public CodecRequest wire-message initializer. Prefer the typed XLSX helpers unless implementing codec infrastructure.

**Schema returns:**

- `response` (object) — Decoded public CodecResponse wire message; structured codec failures throw OpenChestnutCodecError.

#### `openChestnutStatus`

Lazily initialize the bundled OpenChestnut WebAssembly runtime and report its protocol, assembly, and integrity manifest.

**Schema returns:**

- `status` (object) — Bundled OpenChestnut runtime status with protocolVersion, assemblyName, and integrity manifest.

#### `range.clear`

Clear range contents, formats, or both without silently changing validations, dimensions, or other package graphs.

**Schema parameters:**

- `applyTo` (string) — contents, formats, or all (default).

**Schema returns:**

- `result` (undefined) — No return value. Formula/spill topology is detached when contents are cleared.

#### `range.conditionalFormats.add`

Add a conditional formatting rule; cellIs/expression/containsText/colorScale plus standard dataBar/iconSet rules cross the public model and OpenChestnut, with computedStyle inspect records, layout JSON visuals, SVG preview, and native XLSX rendering.

**Examples:**

- range.conditionalFormats.add('cellIs', { operator: 'greaterThan', formula: 10, format: { fill: 'green' } })
- range.conditionalFormats.add('dataBar', { color: '#2563eb', thresholds: ['min', 'max'], showValue: true })
- range.conditionalFormats.add('iconSet', { iconSet: '3TrafficLights1', thresholds: [0, '50%', '80%'], reverse: false })
- range.conditionalFormats.addColorScale({ colors: ['#fee2e2', '#fef3c7', '#22c55e'] })

**Schema parameters:**

- `ruleType` (string) required — cellIs, expression, containsText, colorScale, dataBar, or iconSet.
- `formula` (string|number) — Rule formula or scalar threshold. Omit for containsText; the range facade derives the required relative SEARCH formula.
- `text` (string) — Required search text for containsText rules.
- `operator` (string) — Comparison operator for cellIs rules.
- `format` (object) — Style patch applied when the rule matches.
- `colors` (string[]) — Two or three colors for colorScale rules.
- `color` (string|object) — RGB or symbolic Spreadsheet color for a standard gradient dataBar.
- `thresholds` (Array<string|number|object>) — Typed min/max/num/percent/percentile cfvo thresholds: exactly two for dataBar and one per icon for iconSet.
- `iconSet` (string) — One of the 17 base SpreadsheetML icon-set names. Office 2010 x14-only 3Triangles, 3Stars, and 5Boxes fail closed.
- `showValue` (boolean) — Show the formatted cell value beside the data bar or icon; false renders only the visual.
- `reverse` (boolean) — Reverse a built-in icon set's visual order.
- `gradient` (boolean) — Standard data bars are gradient-filled. false requires x14 and fails closed in this profile.

**Schema returns:**

- `conditionalFormat` (object) — Inspectable conditional-format rule with stable id.

#### `range.copyFrom`

Copy values, formulas, or complete cells from an equally sized or evenly tiling source range with relative A1 translation.

**Schema parameters:**

- `sourceRange` (Range) required — Source range whose row/column dimensions must evenly tile the destination.
- `mode` (string) — values, formulas, or all (default). Relative A1 formulas translate per destination cell.

**Schema returns:**

- `result` (undefined) — No return value; the destination range is updated transactionally in memory.

#### `range.copyTo`

Copy this range to an equally sized or evenly tiled destination range.

**Schema parameters:**

- `destinationRange` (Range) required — Destination range evenly tiled by this source range.
- `mode` (string) — values, formulas, or all (default).

**Schema returns:**

- `result` (undefined) — No return value; equivalent to destinationRange.copyFrom(sourceRange, mode).

#### `range.dataValidation`

Assign a list, whole, decimal, date, time, text-length, or custom-formula validation rule to a range, including bounded input prompts, error alerts, blank policy, and intuitive list-arrow visibility; use sheet.dataValidations.add({ range, rule }) for the collection form.

**Schema parameters:**

- `type` (string) required — Validation type: list, whole, decimal, date, time, textLength, or custom.
- `values` (unknown[]) — One through 256 non-empty, comma-free, control-safe inline list values whose quoted SpreadsheetML formula is at most 255 characters; list rules may use formula1 instead.
- `formula1` (string|number) — Primary validation formula/value.
- `formula2` (string|number) — Secondary formula/value for between rules.
- `operator` (string) — between, notBetween, equal, notEqual, lessThan, lessThanOrEqual, greaterThan, or greaterThanOrEqual.
- `allowBlank` (boolean) — Whether blank cells pass validation. Omission keeps the source-free compatibility default true.
- `showInputMessage` (boolean) — Show the bounded prompt when the cell is selected.
- `promptTitle` (string) — Input-prompt title, at most 32 characters.
- `prompt` (string) — Input-prompt message, at most 255 characters.
- `showErrorMessage` (boolean) — Show an error alert when entered data fails the rule.
- `errorTitle` (string) — Error-alert title, at most 32 characters.
- `error` (string) — Error-alert message, at most 255 characters.
- `errorStyle` (string) — stop, warning, or information.
- `showDropdown` (boolean) — For list rules, true means the in-cell arrow is visible. This deliberately hides SpreadsheetML's inverted showDropDown encoding.

**Schema returns:**

- `validation` (object) — Inspectable/editable bounded data-validation rule anchored to one contiguous range. Imported unsupported extension or multi-area graphs remain source-bound and unchanged.

#### `range.displayFormulas`

Read displayed A1 formulas, including the anchor formula projected across non-editable dynamic-array or legacy-array result cells.

**Schema returns:**

- `formulas` (string[][]) — A1 display-formula matrix, projecting spill/array anchors into non-editable result cells.

#### `range.fillDown`

Copy top-row contents and formatting down the range while translating relative A1 formula references.

**Schema returns:**

- `range` (Range) — The same range after top-row contents/formats are filled down with relative formula translation.

#### `range.fillRight`

Copy left-column contents and formatting right across the range while translating relative A1 formula references.

**Schema returns:**

- `range` (Range) — The same range after left-column contents/formats are filled right with relative formula translation.

#### `range.format`

Assign cell styles, symbolic theme/tint/indexed colors, patterned fills, native dimensions, pixel sizing, and hidden axes through a live range format facade.

**Examples:**

- sheet.getRange('A1:D1').format = { fill: '#0f172a', font: { bold: true }, columnWidth: 18, rowHeight: 24 }
- sheet.getRange('A1:D20').format.columnWidthPx = 120

**Schema parameters:**

- `fill` (string|object) — Solid color or { patternType, foreground, background }; colors accept RGB strings or { theme|indexed|auto, tint } references.
- `font` (object) — Font properties: bold, italic, underline, strike, color, size, and name. Color accepts RGB or symbolic SpreadsheetML references.
- `numberFormat` (string) — Excel number format code.
- `alignment` (object) — horizontal, vertical, wrapText, textRotation, indent, shrinkToFit, and readingOrder options.
- `border` (object) — A shared { style, color } border or per-edge records; colors accept RGB or theme/tint/indexed/auto references.
- `protection` (object) — Cell locked and hidden flags preserved through SpreadsheetML style records.
- `columnWidth` (number) — Column width in Excel character units for every column intersecting the range.
- `columnWidthPx` (number) — Column width in CSS pixels, converted with the public SpreadsheetML maximum-digit-width formula.
- `rowHeight` (number) — Row height in points for every row intersecting the range.
- `rowHeightPx` (number) — Row height in CSS pixels, converted at 96 DPI.
- `columnHidden` (boolean) — Hide or show every column intersecting the range.
- `rowHidden` (boolean) — Hide or show every row intersecting the range.

**Schema returns:**

- `range` (Range) — The formatted range facade.

#### `range.format.autofitColumns`

Measure displayed range values deterministically and set native best-fit widths on each selected column.

**Schema returns:**

- `range` (Range) — The same range after deterministic native best-fit column widths are applied.

#### `range.format.autofitRows`

Measure explicit/wrapped range text deterministically and set native custom heights on each selected row.

**Schema returns:**

- `range` (Range) — The same range after deterministic custom row heights are applied.

#### `range.formulaInfos`

Read per-cell stored/projected formula metadata with editability, spill/array source, anchor, and reference evidence.

**Schema returns:**

- `formulaInfos` (Array<Array<object|null>>) — Stored or projected per-cell formula evidence with kind, display, editability, source, anchor, and ref where applicable.

#### `range.formulasR1C1`

Read or assign R1C1 formulas relative to each target cell while storing canonical A1 formulas.

**Schema parameters:**

- `formulas` (string[][]) — R1C1 formulas relative to each target cell; blank strings clear formulas.

**Schema returns:**

- `formulas` (string[][]) — R1C1 formula matrix; stored formulas remain canonical A1.

#### `range.getCell`

Select one zero-based cell relative to the current range.

**Schema parameters:**

- `row` (number) required — Zero-based row offset within the current range.
- `column` (number) required — Zero-based column offset within the current range.

**Schema returns:**

- `range` (Range) — One-cell relative range.

#### `range.getColumn`

Select one zero-based column relative to the current range.

**Schema parameters:**

- `column` (number) required — Zero-based column offset within the current range.

**Schema returns:**

- `range` (Range) — One-column relative range spanning the current rows.

#### `range.getCurrentRegion`

Expand to the contiguous data region bounded by fully blank rows and columns.

**Schema returns:**

- `range` (Range) — Contiguous region bounded by fully blank rows and columns.

#### `range.getRangeByIndexes`

Select a bounded zero-based subrange relative to the current range.

**Schema parameters:**

- `startRow` (number) required — Zero-based row offset within the current range.
- `startColumn` (number) required — Zero-based column offset within the current range.
- `rowCount` (number) required — Positive subrange row count.
- `columnCount` (number) required — Positive subrange column count.

**Schema returns:**

- `range` (Range) — Bounded relative subrange.

#### `range.getRow`

Select one zero-based row relative to the current range.

**Schema parameters:**

- `row` (number) required — Zero-based row offset within the current range.

**Schema returns:**

- `range` (Range) — One-row relative range spanning the current columns.

#### `range.merge`

Merge the target range as one region or as separate row-wise regions when across=true.

**Schema parameters:**

- `across` (boolean) — Merge each target row independently when true.

**Schema returns:**

- `range` (Range) — The same range after merge creation.

#### `range.offset`

Return an equally sized range shifted by zero-based row and column offsets, rejecting worksheet overflow.

**Schema parameters:**

- `rowOffset` (number) required — Signed row offset.
- `columnOffset` (number) required — Signed column offset.

**Schema returns:**

- `range` (Range) — Equally sized shifted range within XLSX bounds.

#### `range.resize`

Return a range at the same upper-left cell with explicit positive row and column counts.

**Schema parameters:**

- `rowCount` (number) required — Positive output row count.
- `columnCount` (number) required — Positive output column count.

**Schema returns:**

- `range` (Range) — Resized range with the same upper-left cell.

#### `range.setNumberFormat`

Assign one number format or an evenly tiling matrix of Excel-invariant number-format codes.

**Schema parameters:**

- `format` (string|string[][]) required — Excel-invariant number-format code or an evenly tiling format matrix.

**Schema returns:**

- `range` (Range) — The same range after number-format assignment.

#### `range.unmerge`

Remove merged regions intersecting the target range.

**Schema returns:**

- `range` (Range) — The same range after intersecting merges are removed.

#### `range.write`

Write a mixed matrix or one explicit values/formulas/formulasR1C1 payload from the range anchor and return the actual written range.

**Schema parameters:**

- `value` (unknown[][]|unknown[]|object) required — Mixed values/formulas matrix, or exactly one of { values }, { formulas }, or { formulasR1C1 }.

**Schema returns:**

- `range` (Range) — Actual rectangular range written from the receiver's upper-left cell.

#### `range.writeValues`

Write a one- or two-dimensional value matrix from the range anchor.

**Schema parameters:**

- `values` (unknown[][]|unknown[]) required — One row or a rectangular value matrix written from the range anchor.

**Schema returns:**

- `result` (undefined) — No return value; inspect the target range after writing.

#### `sheet.charts.add`

Create an inspectable worksheet chart from a range or config; setData(range) infers category series, scatter per-series numeric xValues/xFormula plus y values/formula, or one exact X/Y/positive-Size bubble series. series.fill sets an explicit #RRGGBB solid color, series.line sets bounded RGB color/dash/width (series.stroke is an alias), line/scatter markers set direct symbol/size/RGB fill/bounded outline semantics, lineOptions controls standard/stacked/percent-stacked grouping, smooth interpolation, and direct vary-colors behavior, dataLabels controls plot-level value/category/series-name visibility and bounded position, and xAxis/yAxis configure primary titles, formats, intervals, and linear value bounds. Marker-only scatter rejects series.line/stroke and writes an explicit native no-fill series outline; use marker.line for marker borders. Bubble charts use two numeric axes and reject ambiguous range shortcuts or nonpositive sizes.

**Schema parameters:**

- `chartType` (string) required — Canonical OpenChestnut XLSX chart type: bar, line, pie, area, doughnut, scatter, or bubble. Other model names fail closed on export.
- `source` (Range|object) — Source range or explicit chart config.
- `title` (string) — Chart title.
- `titleTextStyle` (object) — Optional chart-title style with fontSize from 1 through 4000 points.
- `lineOptions` (object) — Line-chart-only { grouping?, smooth?, varyColors? }. grouping is standard, stacked, or percentStacked; omission authors the standard default. smooth preserves explicit false as native c:smooth val=0. varyColors=true authors direct c:varyColors val=1; false or omission removes that optional node.
- `dataLabels` (boolean|object) — Optional plot-level labels. A boolean controls showValue; an object accepts boolean showValue/showCategoryName, optional presence-aware showSeriesName, and position: bestFit, bottom, center, insideBase, insideEnd, left, outsideEnd, right, or top. Per-series/per-point labels, number formats, and label text styles remain outside this bounded profile.
- `categories` (string[]) — Explicit shared categories for category charts. Scatter and bubble require this to stay empty and use per-series numeric xValues.
- `series` (object[]) — Explicit series definitions with name, optional numeric values/formula, category-chart categoryFormula, scatter/bubble numeric xValues/xFormula, and bubble-only positive bubbleSizes/bubbleSizeFormula exactly aligned with X/Y point counts. Optional #RRGGBB solid fill and optional line { fill, style, width } are supported; line/scatter marker { symbol, size, fill, line } remains marker-only. When internal range formulas are present, inspect/render/OpenChestnut export resolve live category or numeric X/Y/Size caches from those cells. line.fill and marker.fill are #RRGGBB; both line objects use style solid, dashed, dotted, dash-dot, or dash-dot-dot and width 0 through 1584 points. Marker-only scatter rejects the series-level line/stroke aliases and uses marker.line only for marker borders. bubble3D, negative bubbles, custom scale, and non-area sizing are source-bound/read-only. marker.symbol is none, dot, circle, square, diamond, triangle, x, star, plus, or dash; marker.size is an integer from 2 through 72. stroke { color, style, weight } is a series-line compatibility alias and must not conflict with line.
- `xAxis` (object) — Primary text category axis with title.text, textStyle.fontSize, numberFormatCode, and tickLabelInterval; scatter and bubble instead use a numeric value axis with min, max, and majorUnit. Pie and doughnut charts have no axes.
- `yAxis` (object) — Primary numeric value axis with title.text, tick-label textStyle.fontSize, numberFormatCode, min, max, and majorUnit; tickLabelInterval is accepted as a compatibility alias for majorUnit. Pie and doughnut charts have no axes.
- `position` (object) — Pixel chart frame.

**Schema returns:**

- `chart` (WorksheetChart) — Editable worksheet chart facade.

#### `sheet.dataTables.__getDefinitions`

Return defensive inspectable definitions for the worksheet's canonical What-If data tables, including result range, native anchor, inputs, orientation, and display formula.

**Schema returns:**

- `definitions` (object[]) — Fresh defensive copies with zero-based result-range bounds, formulaRef, anchor, normalized rowInput/columnInput, rowOriented, twoVariable, and displayFormula. Mutating the returned objects does not mutate the worksheet.

#### `sheet.dataTables.add`

Create a canonical native Excel What-If data table from a rectangular formula/input grid and one row input, one column input, or both. Excel or another compatible host calculates the result values; the JavaScript evaluator does not simulate TABLE.

**Schema parameters:**

- `range` (string|Range) required — A rectangular A1 range at least 2x2. Its top-left cell must contain the formula to evaluate; its first row and first column contain substitution values, and the remaining rectangle is the native result range.
- `rowInput` (string) — Optional same-sheet single-cell A1 input reference. With no columnInput this authors a row-oriented one-variable table.
- `columnInput` (string) — Optional same-sheet single-cell A1 input reference. With no rowInput this authors a column-oriented one-variable table; provide both inputs for a two-variable table.

**Schema returns:**

- `result` (undefined) — No return value. Source-free canonical tables are authored as native t=dataTable formulas. Imported topology, input bindings, and orientation remain source-bound and read-only; unsupported or overlapping graphs fail closed without fallback.

#### `sheet.images.add`

Create an inspectable worksheet image from a data URL, URI, or prompt with one-cell, two-cell, or absolute pixel geometry plus optional percentage crop, bounded grayscale/luminance/opacity effects, rotation, and horizontal/vertical flips.

**Schema parameters:**

- `dataUrl` (string) — Embedded image data URL.
- `uri` (string) — External image URI metadata.
- `prompt` (string) — Generation/source prompt metadata.
- `alt` (string) — Alternative text.
- `anchor` (object) — One-cell { from, extent }, two-cell { type:'twoCell', from, to, editAs? }, or page-relative { type:'absolute', position:{leftPx,topPx}, extent } geometry. Cell markers use 0-based row/col plus optional rowOffsetPx/colOffsetPx; editAs is twoCell, oneCell, or absolute.
- `crop` (object) — Optional { leftPercent, topPercent, rightPercent, bottomPercent } source rectangle. Each signed offset is -100 through 100; opposing sums must remain below 100. Positive values inset and negative values outset.
- `effects` (object) — Optional { grayscale, brightnessPercent, contrastPercent, opacityPercent } profile. Brightness/contrast are -100 through 100; opacity is 0 through 100. OpenChestnut maps it to bounded DrawingML blip effects.
- `transform` (object) — Optional { rotationDegrees, flipHorizontal, flipVertical } picture transform. Rotation is -360 through 360 degrees at DrawingML 1/60000-degree precision; flip booleans preserve explicit false values.
- `fit` (string) — contain or cover intent.

**Schema returns:**

- `image` (WorksheetImage) — Editable worksheet image facade.

#### `sheet.pivotTables.add`

Create a native bounded XLSX PivotTable with derived cached output, cache records, and exact axis-item filters, or use richer model-only grouping/calculation/date-filter semantics for inspect and preview. Recognized imports are hash-bound and read-only.

**Schema parameters:**

- `name` (string) — Stable pivot name.
- `sourceRange` (string|Range) required — Source data range.
- `targetRange` (string|Range) required — Destination anchor/range.
- `rowFields` (string[]) — Ordered row field names. Native source-free OpenChestnut authoring accepts 1 through 8 fields in a tabular, no-automatic-subtotal profile.
- `columnFields` (string[]) — Column field names. Native source-free OpenChestnut authoring currently accepts zero or one.
- `valueFields` (object[]) — One through 32 value-field definitions. Each names a source field and sum/count/average/min/max aggregation; repeated source fields with different aggregations are allowed.
- `rowGrandTotals` (boolean) — Add a native grand-total column and derived cached values when a column field is present.
- `columnGrandTotals` (boolean) — Add a native grand-total row and derived cached values.
- `groupFields` (object[]) — Derived group fields with unique name/sourceField. Calendar/time groupBy values years/quarters/months/days/hours/minutes/seconds form OOXML base/par hierarchies and accept bounded groupInterval values; range uses numeric startNum/endNum/groupInterval buckets; discrete uses named groups of source items.
- `calculatedFields` (object[]) — Calculated value fields over grouped source-field sums with arithmetic, percent, concatenation, comparisons, string/boolean constants, 12 bounded numeric functions, AND/OR/NOT, lazy IF/IFERROR/IFNA, NA, ISERROR/ISNUMBER/ISTEXT, Excel Compatibility Version 2 surrogate-aware LEN/LEFT/RIGHT/MID, LOWER/UPPER/ASCII-space TRIM, and workbook-date-system-aware DATE/YEAR/MONTH/DAY/EDATE/EOMONTH/DAYS/WEEKDAY/TIME/HOUR/MINUTE/SECOND/NETWORKDAYS/WORKDAY/NETWORKDAYS.INTL/WORKDAY.INTL. Business-day functions accept standard or international weekend rules and one optional scalar holiday. Accepts [Field] or quoted field references; cell references, holiday arrays/ranges, calculated-field chaining, and non-whitelisted functions are rejected.
- `filters` (object|object[]) — Axis filters. Exact include/exclude lists of 1 through 1024 string, finite-number, boolean, or null items on the existing row/column axis are native; they author standard hidden PivotField items and remain semantic across host normalization. Absolute dateEqual/dateNotEqual/dateOlderThan/dateOlderThanOrEqual/dateNewerThan/dateNewerThanOrEqual/dateBetween/dateNotBetween filters with whole-day ISO dates by default or useWholeDay=false plus ISO date-time/Date thresholds at UTC-second precision, and relative UTC types yesterday/today/tomorrow, last/this/next week/month/quarter/year, and yearToDate remain model/preview-only. Relative filters remain whole-day, accept optional deterministic asOf, and use Monday-start ISO weeks.
- `refreshPolicy` (object) — OOXML cache policy: refreshOnLoad, saveData, enableRefresh, invalid, missingItemsLimit, refreshedBy, and refreshedDateIso.

**Schema returns:**

- `pivot` (WorksheetPivotTable) — Native XLSX authoring is bounded to 1 through 8 tabular row fields without automatic subtotals, optional one column field, 1 through 32 sum/count/average/min/max value fields, and exact include/exclude filters on those axes. Multiple values use the canonical x=-2 data-layout axis. Cached output is a derived projection; grouping, calculated fields, date/conditional filters, compact/subtotal-bearing multi-row layouts, and other richer profiles remain model-only and fail closed on native export. Recognized imports expose semantics but keep config, source data, cached output, and topology read-only.

#### `sheet.sparklineGroups.add`

Create standard Office 2010 line/column/stacked sparkline groups for inspect, SVG preview, and OpenChestnut XLSX export. Source-free groups use reversible one-dimensional target/source mappings; recognized imported groups support fixed-topology semantic edits while unsupported native graphs remain source-bound.

**Schema parameters:**

- `type` (string) — line, column, or stacked.
- `targetRange` (string|Range) required — One-dimensional destination range. Each target cell receives one native sparkline.
- `sourceData` (string|Range) required — One-dimensional source for one target, or a reversible rectangle whose rows/columns map exactly to the target cells.
- `dateAxisRange` (string|Range) — Optional one-dimensional date axis with one entry per sparkline point.
- `seriesColor` (string|object) — RGB or native theme/indexed/automatic series color.
- `negativeColor` (string|object) — Optional negative-value color.
- `axisColor` (string|object) — Optional horizontal-axis color.
- `markersColor` (string|object) — Optional ordinary-marker color.
- `firstMarkerColor` (string|object) — Optional first-point marker color.
- `lastMarkerColor` (string|object) — Optional last-point marker color.
- `highMarkerColor` (string|object) — Optional high-point marker color.
- `lowMarkerColor` (string|object) — Optional low-point marker color.
- `lineWeight` (number) — Positive line weight in points; defaults to 1.
- `displayHidden` (boolean) — Whether hidden source cells contribute to the native sparkline.
- `displayEmptyCellsAs` (string|number) — span/connect, gap, zero, or compatible numeric value 1, 2, or 3.
- `markers` (object) — Optional show/high/low/first/last/negative marker booleans.
- `axis` (object) — Optional manualMin/manualMax, minMode/maxMode (individual/group/custom or 0/1/2), showAxis, and rightToLeft settings.

**Schema returns:**

- `sparkline` (SparklineGroup) — Editable standard Office 2010 x14 sparkline group for inspect/layout/SVG preview and OpenChestnut XLSX I/O. Source-free groups use the documented reversible mapping; imported canonical groups are source-bound and permit property edits without topology changes. Unsupported native sparkline graphs remain opaque and unchanged.

#### `sheet.tables.add`

Create an ordinary worksheet table over an A1 range with headers, columns, totals metadata, style, and bounded filtering/sorting. QueryTable bindings cannot be authored; recognized imported bindings expose only table.setQueryRefreshPolicy for one-way automatic-refresh hardening, while all other QueryTable edits fail closed.

**Schema parameters:**

- `range` (string|Range) required — A1 range or range facade.
- `hasHeaders` (boolean) — Whether the first row contains headers.
- `name` (string) — Stable Excel table name.
- `style` (string) — Table style name.
- `columnNames` (string[]) — Compatibility projection of table-column names.
- `columnDefinitions` (object[]) — Rich columns with name, calculatedColumnFormula/array, and totalsRowFunction/label/formula/array metadata.
- `filters` (object[]) — Zero-based table-column exact-value/blank, grouped-date/calendar, one/two-criterion custom, dynamic type/threshold, top/bottom item/percent, standard icon-set, or stable cell-fill/font-color AutoFilters; color filters use kind='color', target='cell'|'font', and color without exposing dxfId.
- `sortState` (object) — Bounded row-oriented value/icon/color-sort state with reference, caseSensitive, optional sortMethod ('none'|'pinYin'|'stroke'), and ordered single-column conditions; value conditions may carry customList. Table AutoFilter sorts reject columnSort per SpreadsheetML.
- `showTotals` (boolean) — Expose the totals row required by totals metadata.

**Schema returns:**

- `table` (WorksheetTable) — Editable ordinary worksheet table facade. QueryTable bindings are import-only and read-only except that a recognized imported binding permits table.setQueryRefreshPolicy to harden automatic-refresh switches; all connection, command, field, sort, refresh-history, topology, and other QueryTable changes fail closed.

#### `SpreadsheetFile.exportCsv`

Export one worksheet or range as UTF-8 CSV, using calculated values unless formula output is explicitly requested.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 CSV FileBlob.

#### `SpreadsheetFile.exportDelimited`

Serialize one workbook sheet/range as bounded CSV/TSV text with calculated-value defaults and RFC-style quoting.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range; defaults to the used range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present; defaults to false.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 CSV/TSV FileBlob with row/column metadata.

#### `SpreadsheetFile.exportTsv`

Export one worksheet or range as UTF-8 tab-separated text with RFC-style quoting where needed.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to serialize.
- `sheetName` (string) — Worksheet name; defaults to the first sheet.
- `range` (string) — Optional A1 range.
- `formulas` (boolean) — Emit formulas instead of calculated values where present.
- `lineEnding` (string) — LF or CRLF output; defaults to CRLF.
- `includeBom` (boolean) — Prefix a UTF-8 BOM; defaults to false.
- `maxBytes` (number) — Maximum encoded output bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum exported rows; defaults to 100000.
- `maxColumns` (number) — Maximum exported columns; defaults to 16384.

**Schema returns:**

- `blob` (FileBlob) — UTF-8 TSV FileBlob.

#### `SpreadsheetFile.exportXlsx`

Serialize a Workbook facade through the single bundled OpenChestnut codec.

**Schema parameters:**

- `workbook` (Workbook) required — Workbook facade to recalculate and serialize.
- `recalculate` (boolean) — Recalculate formulas before serialization; defaults to true.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `blob` (FileBlob) — Native OOXML XLSX package bytes.

#### `SpreadsheetFile.importCsv`

Import UTF-8 CSV bytes into an editable Workbook through the bounded delimited parser.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 CSV text or bytes.
- `sheetName` (string) — Imported worksheet name.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importDelimited`

Parse bounded RFC-style CSV/TSV bytes into an editable Workbook, including quoted delimiters, escaped quotes, and embedded newlines.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 delimited text or bytes.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `sheetName` (string) — Imported worksheet name; defaults to Sheet1.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importTsv`

Import UTF-8 tab-separated bytes into an editable Workbook through the bounded delimited parser.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 TSV text or bytes.
- `sheetName` (string) — Imported worksheet name.
- `coerceTypes` (boolean) — Convert unquoted boolean/numeric-looking cells; defaults to false.
- `maxBytes` (number) — Maximum encoded input bytes; defaults to 10 MiB.
- `maxRows` (number) — Maximum parsed rows; defaults to 100000.
- `maxColumns` (number) — Maximum parsed columns per row; defaults to 16384.

**Schema returns:**

- `workbook` (Workbook) — Imported editable workbook facade.

#### `SpreadsheetFile.importXlsx`

Load XLSX through the single bundled OpenChestnut codec into an editable Workbook facade.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `limits` (object) — Optional maxInputBytes, maxUncompressedBytes, maxParts, maxSheets, maxCells, and maxCompressionRatio codec budgets.

**Schema returns:**

- `workbook` (Workbook) — Imported workbook facade with editable core cells, formulas, styles, ordinary tables, images, basic charts, validation, conditional formatting, threaded-comment roots/direct replies, and canonical Office 2010 sparkline groups. A recognized imported QueryTable may only use table.setQueryRefreshPolicy to disable automatic refresh; connections, commands, fields, sorts, topology, nested reply graphs, mentions, dynamic arrays, pivots, non-reversible sparkline graphs, and unsupported package graphs remain source-bound and read-only.

#### `SpreadsheetFile.inspectDelimited`

Inspect bounded CSV/TSV bytes as file/row records with dimensions, delimiter, quoting, and formula-like cell evidence.

**Schema parameters:**

- `input` (FileBlob|Uint8Array|string) required — UTF-8 CSV/TSV text or bytes.
- `delimiter` (string) — Single field delimiter; defaults to comma.
- `maxBytes` (number) — Maximum encoded input bytes.
- `maxRows` (number) — Maximum parsed rows.
- `maxColumns` (number) — Maximum parsed columns per row.
- `maxPreviewRows` (number) — Maximum row records in bounded output; defaults to 20.
- `maxChars` (number) — Maximum bounded NDJSON output size.

**Schema returns:**

- `inspection` (object) — Delimited-file summary, bounded row records, and NDJSON evidence.

#### `SpreadsheetFile.inspectXlsx`

Inspect bounded XLSX parts, content types, relationships, and namespace-aware source XML r:id/r:embed/r:link references under decompression budgets.

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

Apply path-validated XLSX part patches, build worksheet/table/drawing/image/chart/pivot source references, and atomically reject dangling content types or relationships.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) required — XLSX package bytes.
- `patches` (array|object) required — Safe part edits with text, xml, json, bytes, content, remove, or delete.
- `maxPatchBytes` (number) — Maximum bytes per replacement part.
- `maxParts` (number) — Maximum resulting package part count.
- `syncContentTypes` (boolean) — Synchronize inferred or explicit content-type declarations; defaults to true.
- `syncRelationships` (boolean) — Remove relationships to deleted parts and apply relationship recipes; defaults to true.
- `syncSourceReferences` (boolean) — Apply opt-in standard sourceReference XML mutations for supported semantic recipes; defaults to true.
- `validateResult` (boolean) — Validate final content types and relationships atomically; defaults to true. Set false only for deliberate invalid-package fixtures.
- `recipe` (string|object) — Standard OOXML part recipe with optional source/id/target and sourceReference fields; XLSX supports worksheet/table lists, pivot cache/record bindings, typed pivotTable relationships, and explicit-anchor drawing/image/chart nodes.
- `sourceReference` (boolean|object) — Opt-in source XML mutation. Image/chart objects require explicit anchor geometry; pivotCacheDefinition requires a unique cacheId; pivotCacheRecords binds the cache root to its records relationship.
- `relationship` (object) — Per-patch source/id/type/target/targetMode relationship recipe; explicit ID collisions require replaceExisting:true. relationships accepts an array.

**Schema returns:**

- `blob` (FileBlob) — Patched XLSX FileBlob with part/relationship/content-type/source-reference update counts and validation metadata.

#### `table.setQueryRefreshPolicy`

On one recognized imported QueryTable, monotonically disable automatic refresh without changing its connection, command, fields, sort, refresh history, or topology.

**Examples:**

- table.setQueryRefreshPolicy({ disableRefresh: true, backgroundRefresh: false, refreshOnLoad: false })

**Schema parameters:**

- `policy` (object) required — One or more of exactly { disableRefresh: true, backgroundRefresh: false, firstBackgroundRefresh: false, refreshOnLoad: false }. Unknown keys, unsafe values, an empty object, QueryTable authoring, and non-source-bound mutations fail closed.

**Schema returns:**

- `queryTable` (object) — The same recognized imported QueryTable projection after a one-way automatic-refresh hardening request. The source connection, command/credential metadata, fields, deleted-field history, refresh-local sort, topology, and unknown XML remain immutable and are re-proved before export.

**Notes:**

- This is a source-bound safety operation, not QueryTable authoring or general editing. Each supplied field has exactly one permitted value: disableRefresh: true; backgroundRefresh, firstBackgroundRefresh, and refreshOnLoad: false. Export proves the original query part, immutable connection part, and normalized XML residual before reparsing the result. Commands, credentials, connection bindings, fields, deleted-field history, refresh-local sort state, unknown XML, and every other root attribute remain immutable; unsupported or altered source graphs fail closed.

#### `thread.addReply`

Append a direct reply to an Office threaded-comment root with independent author/person/date/done metadata. Nested or branched reply graphs and mentions fail closed.

**Schema parameters:**

- `text` (string) required — Direct reply text.
- `author` (string) — Reply author; defaults to comments.setSelf or the root author.
- `id` (string) — Optional brace-delimited native comment GUID; otherwise OpenChestnut derives one deterministically.
- `personId` (string) — Optional brace-delimited native person GUID.
- `person` (object) — Optional displayName/userId/providerId identity record.
- `date` (string) — Optional ISO-8601 reply timestamp.
- `done` (boolean) — Optional native reply done state.

**Schema returns:**

- `thread` (CommentThread) — The same thread with one appended direct reply. Setting parentId to another reply, adding mentions, or creating a branched/nested graph makes canonical export fail closed.

#### `workbook.comments.addThread`

Create one root Office threaded comment per thread with GUID/person metadata, date, and resolved state; attach bounded direct replies with thread.addReply().

**Schema parameters:**

- `target` (Range|object) required — Target single-cell range or cell descriptor.
- `text` (string) required — Initial comment text.
- `author` (string) — Root comment author; defaults to comments.setSelf identity.
- `id` (string) — Optional stable model thread ID.
- `comment` (object) — Optional native root metadata: brace-delimited GUID id/personId, person record, ISO date, and done state.
- `resolved` (boolean) — Initial thread resolution state.

**Schema returns:**

- `thread` (CommentThread) — Attached Office threaded-comment root. Direct replies added through addReply cross canonical OpenChestnut export/import; nested/branched replies and mentions fail closed.

#### `workbook.connections`

Inspect bounded non-secret metadata for imported database connections. Connections are source-bound and read-only: authoring, removal, or mutation makes canonical XLSX export fail closed.

**Schema returns:**

- `connections` (object[]) — Recognized imported connection roots exposed for inspection only. All fields, count, and order are source-bound and read-only; provider strings, commands, credentials, source paths, children, extensions, and unsupported types remain hidden and preserved.

#### `Workbook.create`

Create an empty workbook with an explicit date system and optional native SpreadsheetML theme colors.

**Schema parameters:**

- `dateSystem` (string) — Excel serial-date system: '1900' (default) or '1904'.
- `date1904` (boolean) — Boolean alias for dateSystem; true selects the 1904 system.
- `theme` (object) — Theme name and dk1/lt1/dk2/lt2, accent1-accent6, hlink, and folHlink colors written to xl/theme/theme1.xml.
- `calculation` (object) — Optional bounded workbook calcPr policy; omitted means no authored calculation-properties element.

**Schema returns:**

- `workbook` (Workbook) — Empty editable workbook facade with a normalized date system.

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
- hidden

**Schema parameters:**

- `name` (string) required — Defined name.
- `refersTo` (string) required — Sheet-qualified A1 reference.
- `scope` (string) — Optional worksheet scope.
- `comment` (string) — Optional description.
- `hidden` (boolean) — Optional native hidden flag; explicit false is preserved.

**Schema returns:**

- `definedName` (DefinedName) — Created or updated defined-name facade.

**Returns:**

DefinedName facade with id/name/refersTo/scope

#### `workbook.fontFamilies`

Return a fresh sorted, case-insensitively deduplicated list of workbook default and explicit cell font families.

**Schema returns:**

- `families` (string[]) — Font-family inventory; mutating the returned array does not mutate the workbook.

#### `workbook.formulaGraph`

Return a bounded dependency graph of formula nodes, edges, dependents, cycles, formula errors, and syntax-input/reference-budget refusals for workbook QA.

**Schema parameters:**

- `recalculate` (boolean) — Recalculate before reading the graph; defaults to true.
- `maxChars` (number) — Maximum bounded NDJSON graph-record size.

**Schema returns:**

- `graph` (object) — Bounded formula nodes, edges, cycles, errors, syntax-input/reference-budget refusals, and NDJSON.

#### `workbook.inspect`

Emit bounded NDJSON records for workbook, connections, sheets, worksheet protections, tables, formulas, matches, comments, validations, conditional formats, and drawings; narrow with search/target anchors and shape fields with include/exclude.

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

- `kind` (string) — Comma-separated record kinds such as connection, formula, table, style, computedStyle, chart, image.
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

Recalculate bounded workbook formulas and dynamic-array spills, with dependency edges, cycles, errors, and syntax-input/reference-budget refusals.

**Schema returns:**

- `graph` (object) — Updated bounded formula dependency graph including cycles, errors, and syntax-input/reference-budget refusals.

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

Resolve stable workbook, source-bound connection, worksheet, table, pivot, chart, image, sparkline, rule, comment, and defined-name IDs.

**Schema parameters:**

- `id` (string) required — Stable workbook, sheet, table, pivot, chart, image, sparkline, rule, comment, or defined-name ID.

**Schema returns:**

- `object` (object|undefined) — Resolved editable facade/record or undefined.

#### `workbook.setCalculation`

Set bounded workbook-level SpreadsheetML calculation mode, on-save/full-recalculation flags, iterative-calculation limits, and full-precision policy.

**Examples:**

- workbook.setCalculation({ mode: 'automatic', fullCalculationOnLoad: true, forceFullCalculation: true })
- workbook.setCalculation({ mode: 'manual', iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 } })

**Options:**

- mode
- calculateOnSave
- fullCalculationOnLoad
- forceFullCalculation
- iteration
- fullPrecision

**Schema parameters:**

- `mode` (string) — automatic, automaticExceptTables, or manual.
- `calculateOnSave` (boolean) — Request calculation when a host application saves the workbook.
- `fullCalculationOnLoad` (boolean) — Request a full calculation when a host application opens the workbook.
- `forceFullCalculation` (boolean) — Force full rather than dependency-only recalculation.
- `iteration` (object) — Optional { enabled, maxIterations, maxChange } circular-calculation policy.
- `fullPrecision` (boolean) — Calculate using stored values rather than displayed precision when true.

**Schema returns:**

- `workbook` (Workbook) — The same workbook with a bounded native workbook.xml calcPr policy.

**Returns:**

Workbook facade with bounded native calcPr policy

#### `workbook.setDateSystem`

Select the Excel 1900 or 1904 serial-date system for formula calculation and native workbookPr export.

**Schema parameters:**

- `dateSystem` (string|boolean) required — '1900' or false for the 1900 system; '1904' or true for the 1904 system.

**Schema returns:**

- `workbook` (Workbook) — The same workbook after changing its formula and OOXML date-system context.

#### `workbook.sharedArrayFormulas`

Import and export bounded shared and legacy-array formula metadata. XLDAPR dynamic-array anchors are inspectable after import but source-bound and read-only; creating, detaching, or editing their topology makes XLSX export fail closed.

**Schema parameters:**

- `xlsx` (FileBlob|Uint8Array) — XLSX bytes containing shared, legacy-array, or XLDAPR dynamic-array formula records.
- `formula` (string) — Shared or legacy-array formula expression. Imported dynamic-array expressions are read-only.
- `ref` (string) — Shared group or legacy-array range; imported dynamic spill ranges are inspection-only.

**Schema returns:**

- `metadata` (object) — Shared/legacy metadata is bounded and editable. dynamicArrayRef/spill metadata is import-preserved read-only; authoring, detaching, or editing it makes canonical XLSX export fail closed.

#### `workbook.spillReferences`

Use a direct or defined-name A1# reference to consume only an anchor's current, unblocked dynamic spill matrix. Supported range consumers and a direct re-spill read the verified matrix; scalar/general-vector coercion returns #VALUE!, non-spilling anchors return #REF!, and graph/trace record one spillReference edge to the anchor.

**Examples:**

- =SUM(A1#)
- =MATCH(12,'Source Data'!A1#,0)
- =FILTER(A1#,A1#>10)
- =CurrentSpill

**Schema parameters:**

- `formula` (string) required — Formula containing a direct or defined-name A1# dynamic spill reference.
- `anchor` (string) — Optional explanatory A1 anchor such as Source Data!A1; the actual formula must carry #.

**Schema returns:**

- `value` (unknown|unknown[][]|#REF!|#VALUE!|#SPILL!) — Current model-calculation spill value. Only the documented range consumers and direct re-spill profile accept A1#; scalar/general-vector coercion is #VALUE!, a non-spilling anchor is #REF!, and imported dynamic-array package topology remains source-bound.

**Notes:**

- A1# is a model-calculation range reference, not source-free XLSX dynamic-array topology authoring. The evaluator recalculates the formula anchor before reading it, verifies a current rectangular spill of at most 10,000 cells, charges each read against the 20,000-cell formula total, and preserves one spillReference dependency edge to the anchor. A blocked/error anchor propagates its current error; an ordinary scalar/non-spilling anchor is #REF!.

#### `workbook.structuredReferences`

Evaluate Excel table references including sections, column ranges/unions, space intersections, escaped special-character headers, unqualified calculated-column references, and @/#This Row context while expanding exact table-cell precedents.

**Examples:**

- =SUM(TableName[Column])
- =SUM(TableName[[#Data],[First]:[Last]])
- =SUM(TableName[[First]:[Second]] TableName[[Second]:[Third]])
- =[Revenue]-[Cost]
- =TasksTable[@Revenue]
- =SUM(TasksTable[[#This Row],[Revenue]:[Cost]])
- =TasksTable['#Items]
- =TasksTable[Bracket'[Value']]

**Schema parameters:**

- `formula` (string) required — Formula containing an Excel table structured reference.
- `table` (string) — Worksheet table name; omitted only for a calculated-column reference inside that table.
- `selector` (string) required — Column, escaped special-character header, section, current-row, range, union, or space-intersection selector.

**Schema returns:**

- `value` (unknown) — Calculated scalar/array value with stable table-cell precedents.

**Notes:**

- Supports #Headers/#Data/#All/#Totals/#This Row and @, unqualified current-row references inside tables, contiguous column ranges, comma-separated column unions, space intersections over common cells, and apostrophe escaping for [, ], #, ', and @ in column headers. Disjoint intersections return #NULL!; current-row references outside the referenced table return #VALUE!.

#### `workbook.trace`

Return a formula precedent tree and bounded NDJSON trace for a target cell, with circular references and syntax-input/reference-budget refusals flagged.

**Schema parameters:**

- `reference` (string|Range) required — Target A1 reference, optionally sheet-qualified, or range facade.
- `maxDepth` (number) — Maximum precedent recursion depth; defaults to 8.
- `maxChars` (number) — Maximum bounded NDJSON trace size.

**Schema returns:**

- `trace` (object) — Precedent tree plus bounded flat NDJSON trace; oversized syntax or sources are reported rather than walked.

#### `workbook.verify`

Return bounded QA issues for source-bound connections, sheets, formulas (including syntax-input and reference-budget refusals), tables, charts, and comments.

**Schema parameters:**

- `maxChars` (number) — Maximum bounded NDJSON issue output size.

**Schema returns:**

- `report` (object) — Workbook formula/structure/drawing/rule QA result, including syntax-input and reference-budget refusals.

#### `workbook.windows`

Access the ordered workbook-window collection; window 0 is the primary view used by legacy worksheet-selection APIs.

**Schema returns:**

- `windows` (WorkbookWindowCollection) — Ordered windows. Index 0 is the primary window; additional windows retain independent active and selected worksheet state.

#### `workbook.windows.add`

Append an additional workbook window with its own active worksheet and selected tab group.

**Schema parameters:**

- `activeWorksheet` (string|number|Worksheet) — Visible worksheet name, zero-based worksheet index, or worksheet object. Defaults to the primary window's active worksheet.
- `selectedWorksheets` (Array<string|number|Worksheet>) — Optional non-empty unique visible selection for the new window.

**Schema returns:**

- `window` (WorkbookWindow) — Appended workbook window. Source-free XLSX export authors a matching workbookView and one sheetView per worksheet.

#### `workbook.worksheets.add`

Append an editable visible, hidden, or very-hidden worksheet with a stable name and ID.

**Schema parameters:**

- `name` (string) — Unique worksheet name; defaults to SheetN.
- `visibility` (string) — visible (default), hidden, or veryHidden.

**Schema returns:**

- `worksheet` (Worksheet) — Appended editable worksheet with bounded native visibility.

#### `workbook.worksheets.getSelectedWorksheets`

Return the visible worksheet-tab group selected in the primary workbook window, in workbook order.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected visible worksheet tabs in workbook order, always including the active worksheet.

#### `workbook.worksheets.setActiveWorksheet`

Select the visible worksheet opened by default and used by workbook operations that omit an explicit sheet.

**Schema parameters:**

- `worksheet` (string|number|Worksheet) required — Visible worksheet name, zero-based collection index, or worksheet object from this workbook.

**Schema returns:**

- `worksheet` (Worksheet) — Selected visible worksheet. XLSX export writes its zero-based position to workbookView activeTab and collapses the primary tab selection to that worksheet.

#### `workbook.worksheets.setSelectedWorksheets`

Select one or more visible worksheet tabs in the primary workbook window while retaining exactly one active worksheet.

**Schema parameters:**

- `worksheets` (Array<string|number|Worksheet>) required — Non-empty unique list of visible worksheet names, zero-based indexes, or worksheet objects. If the current active worksheet is omitted, the first requested worksheet becomes active.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected worksheet tabs in workbook order; native XLSX export writes sheetView tabSelected for workbookViewId 0.

#### `workbookWindow.getActiveWorksheet`

Return the visible active worksheet for one workbook window.

**Schema returns:**

- `worksheet` (Worksheet) — Visible active worksheet for this window.

#### `workbookWindow.getSelectedWorksheets`

Return one window's visible selected worksheet tabs in workbook order.

**Schema returns:**

- `worksheets` (Worksheet[]) — Visible selected worksheet tabs for this window in workbook order, always including its active worksheet.

#### `workbookWindow.setActiveWorksheet`

Set one window's active worksheet and collapse that window's selected tab group to it.

**Schema parameters:**

- `worksheet` (string|number|Worksheet) required — Visible worksheet resolved within the owning workbook.

**Schema returns:**

- `worksheet` (Worksheet) — Selected worksheet; the window's selected group is collapsed to this worksheet.

#### `workbookWindow.setSelectedWorksheets`

Set one window's non-empty visible selected tab group, which must include its active worksheet.

**Schema parameters:**

- `worksheets` (Array<string|number|Worksheet>) required — Non-empty unique visible selection. If the current active worksheet is omitted, the first requested worksheet becomes active.

**Schema returns:**

- `worksheets` (Worksheet[]) — Selected worksheet tabs for this window in workbook order.

#### `worksheet.freezePanes.freezeColumns`

Freeze a leading column count in the worksheet view while preserving any frozen rows.

**Schema parameters:**

- `columnCount` (number) required — Integer number of leading columns to freeze; zero clears only the column freeze.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade with rows, columns, topLeftCell, activePane, and frozen state.

#### `worksheet.freezePanes.freezeRows`

Freeze a leading row count in the worksheet view while preserving any frozen columns.

**Schema parameters:**

- `rowCount` (number) required — Integer number of leading rows to freeze; zero clears only the row freeze.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade with rows, columns, topLeftCell, activePane, and frozen state.

#### `worksheet.freezePanes.unfreeze`

Remove all frozen worksheet panes and restore a single scrollable view.

**Schema returns:**

- `freezePanes` (object) — Worksheet frozen-pane facade reset to zero frozen rows and columns.

#### `worksheet.getRange`

Select an A1 range for values, formulas, formatting, merge, fill, and copy operations.

**Schema parameters:**

- `address` (string) required — A1 cell or range address such as A1:D10.

**Schema returns:**

- `range` (Range) — Editable range facade for values, formulas, formatting, and rules.

#### `worksheet.getUsedRange`

Return the worksheet used rectangle, optionally excluding formatting-only cells with valuesOnly=true.

**Schema parameters:**

- `valuesOnly` (boolean) — When true, exclude cells represented only by formatting or other non-value state.

**Schema returns:**

- `range` (Range) — Used worksheet rectangle, or A1 for an empty worksheet.

#### `worksheet.mergeCells`

Merge an A1 range as one region or merge each row separately with across=true, retaining only upper-left content.

**Schema parameters:**

- `address` (string|Range) required — A1 range to merge.
- `across` (boolean) — Merge each row as a separate region instead of one rectangular region.

**Schema returns:**

- `worksheet` (Worksheet) — The same worksheet with native merged-range state.

#### `worksheet.protection`

Author, inspect, edit, or remove one passwordless worksheet editing restriction with an intuitive allowed-operation list. Cell locked/hidden styles become effective only while protection is active. This is not encryption or access control; password/hash variants remain source-owned and fail closed on replacement.

**Schema parameters:**

- `enabled` (boolean) — Protection is active when present. Assign null, false, or { enabled: false } to remove a recognized passwordless restriction.
- `allow` (string[]) — Allowed operations: selectLockedCells, selectUnlockedCells, formatCells, formatColumns, formatRows, insertColumns, insertRows, insertHyperlinks, deleteColumns, deleteRows, sort, autoFilter, pivotTables, editObjects, or editScenarios. Omission allows selection of locked and unlocked cells only.

**Schema returns:**

- `protection` (object|undefined) — Passwordless worksheet editing restriction. OpenChestnut contains SpreadsheetML's inverted lock flags and source binding; password/hash/extension profiles are preserved opaquely and semantic replacement fails closed. This is not encryption, authentication, or access control.

#### `worksheet.sortState`

Get or set bounded worksheet-level row/column sorting; columnSort=true uses unique single-row conditions across the sort range.

**Schema parameters:**

- `reference` (string) required — Whole worksheet range whose rows or columns are sorted.
- `caseSensitive` (boolean) — Whether text comparisons are case-sensitive.
- `sortMethod` ('none'|'pinYin'|'stroke') — Optional locale-specific SpreadsheetML method; omission remains distinct from explicit 'none'.
- `columnSort` (boolean) — Optional presence-aware direction. true sorts columns left-to-right; false explicitly selects ordinary row sorting.
- `conditions` (object[]) required — Ordered unique single rows when columnSort=true, otherwise unique single columns; value conditions may add customList and icon/color selectors reuse the table-sort shape.

**Schema returns:**

- `sortState` (object) — Bounded worksheet-level sort state. QueryTable refresh sorts may be inspected after import but remain immutable; the only QueryTable edit is root automatic-refresh hardening through table.setQueryRefreshPolicy.

#### `worksheet.unmergeCells`

Remove every merged region intersecting an A1 range without discarding the retained upper-left content.

**Schema parameters:**

- `address` (string|Range) required — A1 range whose intersecting merged regions should be removed.

**Schema returns:**

- `worksheet` (Worksheet) — The same worksheet after intersecting merges are removed.

#### `worksheet.visibility`

Read or assign native worksheet visibility as visible, hidden, or veryHidden; at least one sheet must remain visible.

**Schema parameters:**

- `visibility` (string) required — visible, hidden, or veryHidden.

**Schema returns:**

- `visibility` (string) — Normalized worksheet visibility; workbook verification/export rejects an all-hidden workbook.

