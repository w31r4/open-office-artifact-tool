---
name: open-office-spreadsheets
description: Create, edit, inspect, render, and verify XLSX/CSV spreadsheet artifacts with open-office-artifact-tool.
---

# Open Office Spreadsheets

Use this project skill for standalone `.xlsx`, `.csv`, and `.tsv` artifact work. It is the clean-room workflow for `open-office-artifact-tool`; it does not control a live Microsoft Excel session.

## Contract

- Use `open-office-artifact-tool` public exports. Do not import reference implementation internals.
- Prefer one writable `.mjs` builder that can be rerun after focused edits.
- Write rectangular value/formula matrices in blocks and keep derived values as formulas.
- Preserve the style and formulas of an imported workbook unless the user asks for a redesign.
- Inspect `displayValue` in layout evidence when number/date formats matter; raw `value` remains the calculation value while model/native previews should show the formatted text.
- Keep inspect output bounded and save large QA evidence to files instead of printing it.
- A workbook is deliverable only after durable XLSX export/import, semantic verification, and visual review pass.

## Authoring workflow

1. Create or import the workbook.
2. Inspect the relevant sheets, ranges, formulas, styles, and drawings.
3. Apply values, formulas, formatting, tables, validations, comments, and drawings through public facade APIs. Use `workbook.worksheets.setActiveWorksheet(...)` for the primary window's sheet opened by default. If multiple visible tabs should remain selected, call `setSelectedWorksheets([...])` afterwards; the group must include the active sheet. Use `workbook.windows.add({ activeWorksheet, selectedWorksheets })` only when the artifact intentionally needs another native workbook window with independent tab state.
4. Recalculate and spot-check important results and formula traces.
5. Export XLSX, inspect its bounded native package records, import the exported file again, and run the project verifier.
6. Inspect the preview at full size. Fix formula errors, clipping, unreadable formatting, and broken objects before delivery.
7. Render every user-facing worksheet. Once the previews are approved, save raster baselines and require pixel comparisons on later exports.
8. For final delivery, render the real XLSX/CSV/TSV through LibreOffice to PDF and Poppler page PNGs; compare every native page and its page count.

CSV and TSV deliberately flatten one worksheet/range and cannot preserve styles, drawings, validations, comments, or workbook formulas. Use XLSX for editable fidelity. Delimited export writes calculated values by default; set `formulas: true` only when formula text is explicitly required.

XLSX import follows worksheet and drawing relationships instead of assuming conventional filenames. Metadata-free embedded images and basic bar/line/pie charts are restored from native SpreadsheetDrawingML, including one-cell, two-cell, and absolute anchors, chart caches, and range formulas. Always inspect and verify imported drawings; advanced chart styling, axes, data labels, and grouped shapes remain fidelity work.

Office 2019 threaded comments are relationship-driven too. Set `workbook.comments.setSelf(...)`, create a root with `workbook.comments.addThread(...)`, and use `thread.addReply(text, config)` for replies. The exporter writes schema-valid brace-delimited GUIDs, a workbook-owned person registry, worksheet-owned threaded-comment parts, `parentId`, `dT`, and per-comment `done` state. Metadata-free import follows arbitrary workbook/worksheet relationship targets, accepts alternate namespace prefixes, preserves native comment/person identity and dates, and keeps multiple independent roots on the same cell separate. Low-level relocation uses a `person` recipe sourced from `xl/workbook.xml` plus a `threadedComments` recipe sourced from the owning worksheet. Package QA must reject external/wrong-source/orphan/multiple parts, invalid or duplicate GUIDs, missing people/parents, cross-cell parents, and cycles.

For bounded low-level package edits, `SpreadsheetFile.patchXlsx(...)` can add or remove worksheet/table references and complete DrawingML chains. Create a `drawing` recipe from the worksheet first, then use `image` or `chart` recipes sourced from that drawing with an explicit `sourceReference.anchor`; the patcher owns content types, relationships, namespace prefixes, non-visual object IDs, source nodes, deletion cleanup, and final package validation.

Worksheet-backed native PivotTables are also relationship-driven. Metadata-free import follows workbook `pivotCaches`, cache definitions, and the worksheet's typed pivotTable relationships at arbitrary valid targets to rebuild editable `sheet.pivotTables` facades. Use `rowFields` plus `columnFields` for cross-tabs, and restrict axis items with exactly one `include` or `exclude` array per filtered field. Use `groupFields: [{ name, sourceField, groupBy }]` for `years`, `quarters`, `months`, `days`, `hours`, `minutes`, or `seconds` fields from ISO, `Date`, or Excel-serial values; order fine-to-coarse levels through the derived names and use `range.groupInterval` for bounded Day/time buckets. Use `groupBy: "range"` with `range: { startNum, endNum, groupInterval }` for numeric buckets; omit boundaries to derive them from numeric source values, or set `autoStart`/`autoEnd` explicitly. Use `groupBy: "discrete"` with `groups: [{ name, items }]` for named manual groups; a typed source item may belong to at most one group and unlisted items remain individually visible. The exporter owns the standard cache `fieldGroup base/par`, `rangePr`, `discretePr`, and `groupItems` structures and preserves full timestamps. Date filters support all eight absolute comparison/range types, using whole-day ISO dates by default. Set `useWholeDay: false` with ISO date-time or `Date` thresholds for UTC-second precision; export writes the public x14 `pivotFilter useWholeDay="0"` extension and metadata-free import accepts any namespace prefix. Relative filters cover `yesterday`/`today`/`tomorrow`, last/this/next week/month/quarter/year, and `yearToDate`; they remain whole-day. Relative calculation is UTC with Monday-start ISO weeks; set filter `asOf` for deterministic workflows. Because native OOXML stores only the dynamic type, pass `SpreadsheetFile.importXlsx(..., { relativeDateAsOf })` when metadata-free QA must reproduce an anchored result. The runnable formula-summary fixture exercises precise absolute filtering through export, inspect, import, modeled verification, and native package rendering. LibreOffice currently renders but does not apply the x14 predicate in its Pivot view, so do not claim precise native semantics without Microsoft Excel validation. `calculatedFields` evaluates grouped source-field sums through arithmetic, comparisons, strings/booleans, 12 bounded numeric functions, `AND`/`OR`/`NOT`, lazy `IF`/`IFERROR`/`IFNA`, `NA`, `ISERROR`/`ISNUMBER`/`ISTEXT`, Excel Compatibility Version 2 surrogate-aware `LEN`/`LEFT`/`RIGHT`/`MID`, `LOWER`/`UPPER`, ASCII-space-only `TRIM`, and workbook-date-system-aware `DATE`/`YEAR`/`MONTH`/`DAY`; reference fields as `[Revenue]` or `'Revenue'`, then use the calculated name from `valueFields`. The date functions inherit the owning workbook's `dateSystem`, retain the 1900 serial-60 leap-day quirk, and map `1904-01-01` to serial `0` only in a 1904 workbook. Lazy functions validate every branch but evaluate only the selected branch, so unused division or `#N/A` errors do not leak. Cell references, calculated-field chaining, unknown fields, non-whitelisted functions, invalid argument counts, invalid date systems, and oversized formulas fail explicitly. Imported formulas outside this subset remain available for re-export but produce `#NAME?` and a `pivotCalculatedFieldUnsupported` verification issue, which must not be silently waived. Treat `refreshPolicy` as an OOXML contract: review `refreshOnLoad`, `saveData`, `enableRefresh`, stale-cache `invalid`, retained-item limits, and refresh provenance rather than assuming a save always embeds cache records. With `saveData: false`, the exported package deliberately omits the cache-record part and its relationship. Low-level creation uses `pivotCacheRecords` sourced from its cache definition with `sourceReference: true`, `pivotCacheDefinition` sourced from the workbook with a unique `sourceReference.cacheId`, and a `pivotTable` recipe sourced from the target worksheet without an extra sourceReference node. The cache-definition patch should also add a relationship sourced from the PivotTable part back to that cache definition, matching OpenXML part ownership. Inspect, resolve, layout, verify, package validation, and native rendering must all pass afterward.

The Pivot calculated-field whitelist currently contains 45 scalar functions. Its deterministic date subset also includes `EDATE`/`EOMONTH` month shifting, `DAYS(end, start)`, and `WEEKDAY` return types `1`, `2`, `3`, and `11` through `17`. `TIME` builds a wrapped clock fraction from truncated `0..32767` hour/minute/second inputs, while `HOUR`/`MINUTE`/`SECOND` extract components from numeric serials or strict time text. `NETWORKDAYS`/`WORKDAY` use the standard Saturday/Sunday weekend; their `.INTL` forms accept Excel weekend codes `1..7`/`11..17` or an exact seven-character Monday-first `0`/`1` pattern. Counts are endpoint-inclusive and preserve direction, offsets are truncated, and this bounded scalar AST accepts at most one optional holiday date rather than arrays or ranges. Month offsets are truncated, short target months clamp to month end, and invalid serials, time inputs, weekend patterns, or return types produce explicit Excel errors. The `formula-summary` fixture exercises these functions under a 1904 workbook and verifies metadata-free second export.

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create({ theme: { name: "Report Theme", colors: { accent1: "#0F766E", accent2: "#7C3AED" } } });
const sheet = workbook.worksheets.add("Summary", { visibility: "visible" });
const staging = workbook.worksheets.add("Staging", { visibility: "veryHidden" });
sheet.getRange("A1:C3").values = [
  ["Month", "Revenue", "Cost"],
  ["Jan", 100, 60],
  ["Feb", 120, 70],
];
sheet.getRange("D1").values = [["Margin"]];
sheet.getRange("D2").formulas = [["=(B2-C2)/B2"]];
sheet.getRange("D2:D3").fillDown();
sheet.getRange("A1:D1").format = {
  fill: { patternType: "lightTrellis", foreground: { theme: 4, tint: 0.4 }, background: "#F8FAFC" },
  font: { bold: true, color: { theme: 4, tint: -0.25 } },
  alignment: { horizontal: "center" },
};
sheet.getRange("B2:C3").format = { numberFormat: "$#,##0" };
sheet.getRange("D2:D3").format = { numberFormat: "0.0%" };
sheet.getRange("A1:A3").format.columnWidth = 12;
sheet.getRange("B1:D3").format.columnWidthPx = 96;
sheet.getRange("A1:D1").format.rowHeight = 24;
sheet.getRange("A5:D5").values = [["Quarter summary", null, null, null]];
sheet.getRange("A5:D5").merge();
sheet.freezePanes.freezeRows(1);
workbook.comments.setSelf({ displayName: "Spreadsheet Agent" });
const review = workbook.comments.addThread({ cell: sheet.getRange("D2") }, "Verify the modeled margin.");
review.addReply("Margin evidence is approved.", { author: "QA Reviewer" }).resolve();
workbook.recalculate();

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save("output.xlsx");

const csv = await SpreadsheetFile.exportCsv(workbook, { sheetName: "Summary", range: "A1:D20" });
await csv.save("output.csv");
const importedTsv = await SpreadsheetFile.importTsv("Name\tValue\nRevenue\t120", { coerceTypes: true });
```

For workbooks that use Excel's 1904 serial-date system, create the workbook with `Workbook.create({ dateSystem: "1904" })` or call `workbook.setDateSystem("1904")`. The setting drives date formulas and native `workbookPr date1904` import/export; do not manually offset stored dates by 1,462 days.

For dates or times imported as text, use `DATEVALUE`, `TIMEVALUE`, and `VALUE` before arithmetic. The clean-room evaluator deliberately accepts unambiguous `yyyy-mm-dd`, English month-name dates with a four-digit year, 12/24-hour times, and ASCII numbers with grouping/percent/accounting syntax. It rejects ambiguous numeric dates such as `2/3/2026` instead of inheriting the host machine locale. `TIME`, `HOUR`, `MINUTE`, and `SECOND` use Excel day fractions, honor 0–32767 overflow bounds, and remain deterministic under both workbook date systems. The formula-summary fixture exercises these conversions through model calculation, XLSX roundtrip, all-sheet Playwright rendering, and LibreOffice/Poppler native QA.

Use `NETWORKDAYS.INTL` and `WORKDAY.INTL` when weekends are not Saturday/Sunday. The weekend argument may be an Excel weekend number (`1`–`7` or `11`–`17`) or a seven-character Monday-first mask such as `"0000011"`; keep holidays as serial-date cells/ranges in the same workbook date system.

Conditional aggregation through `COUNTIF(S)`, `SUMIF(S)`, and `AVERAGEIF(S)` uses case-insensitive Excel criteria. `?` matches one text character, `*` matches a text sequence, and `~` escapes `?`, `*`, or `~`; for example, `=COUNTIF(A1:A20,"East*")` and `=COUNTIF(A1:A20,"A~*literal")`. All ranges supplied to an `IFS` variant must have the same dimensions or calculation returns `#VALUE!`.

Inside a worksheet table, calculated-column formulas may use unqualified current-row references such as `=[Revenue]-[Cost]`, or explicit `@`/`#This Row` references such as `=Sales[@Revenue]`. A single space between two or more fully qualified table references is Excel's intersection operator: `=SUM(Sales[[Region]:[Revenue]] Sales[[Revenue]:[Cost]])` evaluates only their common Revenue cells, while disjoint references return `#NULL!`. The `structured-intersection` fixture exercises calculation, exact trace precedents, package formula preservation, and Playwright QA. The locally installed LibreOfficeDev 26.8 alpha displays `Err:509` when directly rendering the structured formula, but rewrites it to an A1-range intersection during ODS roundtrip; this project restores that rewritten formula and value, while native Excel execution remains a Windows validation item. Escape `[`, `]`, `#`, `'`, and `@` in special-character column headers with a leading apostrophe, for example `Sales['#Items]` or `Sales[Bracket'[Value']]`. The calculator, formula graph, trace output, and metadata-free native XLSX import share the same parser, so verify the exact effective precedents rather than only the cached value.

Use `range.fillDown()` or `range.fillRight()` to copy the leading row/column with Excel-style relative formula translation. Use `range.merge()` for one rectangular merge, `range.merge(true)` for separate row-wise merges, and `range.unmerge()` to remove intersecting merged regions. Fill operations reject merged intersections so a workflow cannot silently overwrite ambiguous merged-cell contents.

For scrollable reports, use `sheet.freezePanes.freezeRows(count)` and `freezeColumns(count)`; the two axes compose, and `unfreeze()` clears both. These methods write native SpreadsheetML `sheetViews/pane` state and are preserved when importing third-party XLSX files.

`range.format` is live: assign `columnWidth`/`rowHeight` in Excel character/point units, or `columnWidthPx`/`rowHeightPx` in pixels. `columnHidden` and `rowHidden` preserve hidden axes. Borders accept either one `{ style, color }` record for all four sides or independent `left`/`right`/`top`/`bottom`/`diagonal` records. Colors may be RGB strings or `{ theme, tint }`, `{ indexed }`, or `{ auto: true }`; non-RGB imports retain the symbolic reference plus deterministic `resolved` RGB. Pattern fills use `{ patternType, foreground, background }` and must survive metadata-free import, second export, model rendering, and native LibreOffice rendering. Use `autofitColumns()`/`autofitRows()` only on the smallest intended range, then cap unusually large results before delivery.

## Verification commands

Verify any exported workbook and write bounded inspect, verify, layout, visual-QA, and preview evidence:

```sh
node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input output.xlsx \
  --output-dir tmp/spreadsheet-qa \
  --sheet Summary \
  --range A1:D20 \
  --render-format png
```

`--render-format svg` is dependency-light. `png`, `webp`, `jpeg`, and `pdf` use the optional Playwright renderer and require its Chromium runtime.

The same verifier accepts delimited files and writes bounded row evidence instead of XLSX package-part evidence:

```sh
node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input output.csv \
  --input-format csv \
  --sheet Sheet1 \
  --coerce-types true \
  --output-dir tmp/csv-qa
```

Run the checked-in clean-room fixture end to end:

```sh
node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/formula-summary.json \
  --render-format png --all-sheets true --native-render required \
  --output-dir tmp/spreadsheet-skill-fixture
```

This fixture intentionally runs `JavaScript export → OpenChestnut import → source-preserving OpenChestnut export` before inspect/render/verify. Its advanced theme, styles, table, pivot, chart, image, comments, formulas, and native pages therefore exercise both codec boundaries.

The source-built Open XML SDK WebAssembly codec is available for the bounded workbook features: primitive/cached-formula cells, native shared-formula groups, legacy array-formula anchors, the complete 12-slot workbook theme, complete static `range.format` profiles (font, pattern fill, nine border edges, alignment, protection, and built-in/custom number formats), ordinary workbook- or sheet-scoped range defined names with optional comment/hidden state, worksheet tables with calculated-column/totals-row, exact/grouped-date/custom/dynamic/Top10/icon/color AutoFilter, custom-list/value/icon/color sorting with optional `none`/`pinYin`/`stroke` method, source-bound QueryTable root plus refresh/field/deleted-field/sort semantics, source-bound database connection root metadata, embedded PNG/JPEG one-cell pictures, date-system selection, sheets, merges, dimensions, gridlines, and frozen panes. The checked-in fixture performs OpenChestnut authoring, import, and source-preserving second export, so its custom theme, defined names, static styles, shared/array topology, custom/percentage formats, native sorting/color selectors, embedded picture, DrawingPart/ImagePart relationships, and `tableParts`/TableDefinitionPart relationship cross the public wire twice before QA:

```sh
node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/open-chestnut-basic.json \
  --codec open-chestnut \
  --native-render required \
  --output-dir tmp/open-chestnut-spreadsheet-fixture
```

For this first OpenChestnut picture slice, use `sheet.images.add({ name, alt, dataUrl, anchor })` with a base64 `image/png` or `image/jpeg` data URL and `{ from: { row, col, rowOffsetPx?, colOffsetPx? }, extent: { widthPx, heightPx } }`. The adapter content-addresses the bytes; OpenChestnut owns the worksheet Drawing relationship, DrawingPart, picture relationship, and ImagePart. A recognized imported picture may edit `name`, `alt`, and the one-cell anchor in place. Imported image add/remove, byte replacement, external URI/prompt sources, two-cell/absolute anchors, crop/effect editing, charts, shapes, and groups remain fail-closed or opaque. Relationship IDs are package locators, never durable image IDs.

An existing QueryTable is a source-bound graph, not a source-free table option. The dedicated fixture first creates an ordinary workbook, injects a standards-based synthetic ConnectionsPart/QueryTablePart as external-source input, then uses only OpenChestnut to import, disable refresh, edit QueryTable root/refresh/field policy and bounded connection-root metadata in place, export, and roundtrip before package/model/Playwright QA:

```sh
node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/open-chestnut-query-table.json \
  --codec open-chestnut \
  --native-render off \
  --output-dir tmp/open-chestnut-query-table-fixture
```

Imported recognized tables expose `table.queryTable` with `name`, `connectionId`, presence-aware root policy, optional `growShrinkType`/`autoFormatId`, and `refresh`. The refresh object exposes presence-aware `preserveSortFilterLayout`, `fieldIdWrapped`, `headersInLastRefresh`, `minimumVersion`, `nextId`, unbound-column counters, `fields`, optional `deletedFieldNames`, and optional refresh-local `sortState`. A field exposes its immutable positive `id` and optional immutable `tableColumnId` plus editable `name`, `dataBound`, `rowNumbers`, `fillFormulas`, and `clipped`; field count, order, IDs, and table bindings cannot change. Deleted-field names can change in place, but count/order cannot. Refresh-local sort uses the same locale/custom-list/value/icon/color shape described below and may switch between ordinary single-column conditions and `columnSort: true` single-row conditions; its range, case sensitivity, method, direction, and selectors can change while presence and condition count remain source topology. Formula-fill fields must set `dataBound: false`, while clipped fields must set `dataBound: true`. OpenChestnut requires the connection ID to exist in the hash-bound source ConnectionsPart. Recognized database/type-5 entries appear in `workbook.connections` with immutable `connectionId`, `type`, and `refreshedVersion`; `name`, optional `description`, `keepAlive`, `intervalMinutes` (`0..32767`), `background`, `refreshOnLoad`, and `saveData` may be edited in place. Provider strings, commands, credentials, source paths, children, extensions, unknown attributes, and unsupported connection types stay hidden and preserved. Connection add/remove/reorder, type/version/ID mutation, source-free creation, and refresh execution fail closed. An unsupported deleted-field or sort branch remains independently hidden while safe root/field edits continue; an unrecognized complete refresh remains hidden while a separately recognized root can still be edited. Do not set `queryTable` or `connections` on a newly authored workbook: both OpenChestnut and the JavaScript codec reject source-free fabrication, and the JavaScript codec also rejects export of an imported query graph it cannot preserve. The fixture deliberately sets `disableRefresh: true` and keeps native office rendering off so QA does not invoke an external-data consumer; Playwright still verifies the modeled render.

For QA fixtures only, a range operation may declare `formulaMetadata: { kind: "shared", sharedIndex, reference }` on a complete shared group or `{ kind: "array", reference }` on its anchor. This mirrors already-imported native metadata; normal agent editing still uses `range.formulas`. Editing one shared member through that public range API detaches the complete group to ordinary formulas.

A fixture table may set `range`, `name`, `hasHeaders`, `showTotals`, `showBandedColumns`, `style`, `columnDefinitions`, `filters`, and `sortState`. Each column definition accepts `name`, leading-`=` `calculatedColumnFormula`, `calculatedColumnFormulaArray`, `totalsRowFunction`, `totalsRowLabel`, leading-`=` `totalsRowFormula`, and `totalsRowFormulaArray`; worksheet cell formulas remain a separate modeled source of cached/evaluated values. Filter profiles are `{ columnIndex, kind: "values", values, includeBlank }`, `{ columnIndex, kind: "values", values: [], includeBlank, calendarType, dateGroups: [{ grouping, year, month?, day?, hour?, minute?, second? }] }`, `{ columnIndex, kind: "custom", matchAll, criteria: [{ operator, value }] }`, `{ columnIndex, kind: "dynamic", type, value?, maxValue? }`, `{ columnIndex, kind: "top10", top, percent, value, filterValue? }`, `{ columnIndex, kind: "icon", iconSet, iconId? }`, or `{ columnIndex, kind: "color", target: "cell"|"font", color }`. Exact values and grouped dates cannot coexist in one `values` profile; date groups require the complete field hierarchy through their grouping level. Custom filters accept one or two criteria and the operators `equal`, `notEqual`, `lessThan`, `lessThanOrEqual`, `greaterThan`, or `greaterThanOrEqual`.

Bounded sorting uses `{ reference, caseSensitive, sortMethod?, columnSort?, conditions: [{ reference, descending, customList?, kind?, iconSet?, iconId?, target?, color? }] }`; `sortMethod`, when present, is `"none"`, `"pinYin"`, or `"stroke"`. Omit `columnSort` for ordinary row sorting, where each ordered condition is a unique single column with the sort range's row span. Set `worksheet.sortState.columnSort: true` or a QueryTable refresh sort's `columnSort: true` for left-to-right column sorting, where each condition is a unique single row with the sort range's column span; explicit `false` remains distinct from omission. SpreadsheetML does not allow `columnSort` inside a table AutoFilter, so table `sortState` rejects either explicit boolean value. Omit `kind` for a value sort, which may carry a nonempty control-free `customList` of at most 32,767 characters; use `kind: "icon"` for an icon selector, or `kind: "color"` with `target` and `color` for a stable color selector. A custom list cannot be combined with icon/color selectors. A sheet fixture may set top-level `sortState`; imported recognized tables may edit their row-sort fields while retaining part/relationship identity. A top-level fixture `definedNames` array accepts `{ id?, name, refersTo, scope?, comment?, hidden? }`; OpenChestnut authors only ordinary sheet-qualified A1 ranges. A top-level `calculation` object accepts `mode` (`automatic`, `automaticExceptTables`, or `manual`), presence-aware save/full-recalculation/full-precision flags, and optional `iteration: { enabled, maxIterations, maxChange }`; it writes workbook `calcPr` policy but does not itself execute Excel. Imported recognized names and calculation properties may edit their bounded fields in place while topology and source bindings stay fixed. Formula/constant, reserved, macro, and extended names plus R1C1, concurrent, completed-state, or extended calculation profiles remain hidden and preserved. Unsupported geometry remains hidden in the hash-bound source package, multi-property differential styles and unsupported extensions remain read-only, and imported table add/remove or QueryTable graph topology change fails closed. Direct WebAssembly authoring also remains fail-closed for QueryTable/external-connection creation, gradient fills, HSL/preset/scheme theme colors, theme font/format-scheme editing, named/cell style inheritance, formula/constant or advanced defined names, pivots, charts/groups/shapes and non-one-cell image profiles, comments, validations, conditional formatting, data-table formulas, and dynamic-array extension semantics. On imported workbooks a cached-value-only edit leaves `<f>` untouched, while a modeled theme edit preserves the imported font/format schemes and other residual XML, and a modeled static-style change clones the current XF and its referenced resources before changing only owned fields; unknown XF/font/fill/border/alignment/protection content and formula XML remain source-preserved. All other unsupported package content stays behind the bounded, hash-bound source package and opaque-graph checks. Use the default `javascript` codec to create advanced workbooks only when no source-bound QueryTable graph must survive; never discard or fabricate preservation state merely to make an export pass.

The table fixture adapter also accepts explicit `showFilterButton`, `showFirstColumn`, `showLastColumn`, `showRowStripes`, compatibility `columnNames`, bounded `filters`, and bounded `sortState` forms above when a QA case needs the complete modeled table profile.

Create and later compare an approved sheet/range baseline:

```sh
node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input output.xlsx --sheet Summary --range A1:D20 \
  --render-format png --all-sheets true \
  --native-render required \
  --baseline-dir tmp/spreadsheet-baselines \
  --write-baseline true

node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input output.xlsx --sheet Summary --range A1:D20 \
  --render-format png --all-sheets true \
  --native-render required \
  --baseline-dir tmp/spreadsheet-baselines
```

## QA gates

- Inspect the key table/formula range with bounded `maxChars`.
- Run `verifyArtifact(workbook)` and resolve every error-level issue.
- Trace important result cells when the calculation chain is non-trivial.
- Export and re-import the XLSX before final verification.
- Treat worksheet visibility as native document state: use only `visible`, `hidden`, or `veryHidden`, keep at least one visible sheet, and verify the selected delivery sheet remains active in every intended window. OpenChestnut preserves imported workbook views and fails closed if an edit hides any active or selected tab; it does not silently select another sheet. Source-bound multi-window editing requires a complete matching `sheetView` graph and cannot add, remove, or reorder windows.
- For theme/pattern styling, inspect `xl/theme/theme1.xml` and `xl/styles.xml`, then confirm the imported style retains `theme`/`tint`/`indexed` and `patternType` instead of only a flattened RGB approximation.
- Save `SpreadsheetFile.inspectXlsx()` evidence so native package shape, content types, `.rels` targets, source XML `r:id`/`r:embed`/`r:link` references, and decompression budgets are part of QA; require zero relationship-reference issues.
- For threaded comments, require zero package semantic issues, inspect both `xl/persons` and worksheet threaded-comment relationship targets, remove clean-room metadata once in tests, and verify native IDs, people, dates, parent links, resolution state, and distinct same-cell roots survive import and a second export.
- For CSV/TSV, save `SpreadsheetFile.inspectDelimited()` evidence and review delimiter, dimensions, quoting, BOM, and formula-like cell counts; remember that delimited delivery cannot preserve workbook-only features.
- Produce a layout record and visual preview for every user-facing sheet with `--all-sheets true`; the selected `--range` narrows only the primary sheet.
- Use Playwright raster output for deterministic previews; use LibreOffice or Microsoft Office when native application fidelity is the acceptance criterion.
- Raster baselines use `visualQaArtifact(..., { pixelDiff: true })`; baseline filenames are stable per worksheet. Supplying `--baseline-dir` is fail-closed: initialize it with `--write-baseline true`, because missing sheet images and missing, empty, or non-contiguously numbered native-page sets are rejected.
- When pixels change, the QA output records a `.diff.png` heatmap path for focused agent review.
- Use `--diff-alignment center|top-left|strict`, `--diff-color '#ff1848'`, and `--diff-unchanged-color '#334155'` to make dimension changes and review palettes explicit.
- For same-size renders with a known platform jitter, opt in to bounded registration with `--registration-offset 2`; use `--registration-improvement 0.1` to require at least 10% sampled mismatch improvement. QA records the chosen baseline translation and ignored edge pixels.
- `--native-render required` converts the input XLSX/CSV/TSV with LibreOffice and rasterizes every PDF page with Poppler; page-count or pixel changes fail QA.
- Deliver only the requested workbook unless the user asks for QA intermediates.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: https://github.com/w31r4/open-office-artifact-tool/blob/main/docs/coverage.md
- Fixture runner: `scripts/run-fixture.mjs`
- Generic workbook verifier: `scripts/verify-workbook.mjs`
