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
3. Apply values, formulas, formatting, tables, validations, comments, and drawings through public facade APIs.
4. Recalculate and spot-check important results and formula traces.
5. Export XLSX, inspect its bounded native package records, import the exported file again, and run the project verifier.
6. Inspect the preview at full size. Fix formula errors, clipping, unreadable formatting, and broken objects before delivery.
7. Render every user-facing worksheet. Once the previews are approved, save raster baselines and require pixel comparisons on later exports.
8. For final delivery, render the real XLSX/CSV/TSV through LibreOffice to PDF and Poppler page PNGs; compare every native page and its page count.

CSV and TSV deliberately flatten one worksheet/range and cannot preserve styles, drawings, validations, comments, or workbook formulas. Use XLSX for editable fidelity. Delimited export writes calculated values by default; set `formulas: true` only when formula text is explicitly required.

XLSX import follows worksheet and drawing relationships instead of assuming conventional filenames. Metadata-free embedded images and basic bar/line/pie charts are restored from native SpreadsheetDrawingML, including one-cell, two-cell, and absolute anchors, chart caches, and range formulas. Always inspect and verify imported drawings; advanced chart styling, axes, data labels, and grouped shapes remain fidelity work.

Office 2019 threaded comments are relationship-driven too. Set `workbook.comments.setSelf(...)`, create a root with `workbook.comments.addThread(...)`, and use `thread.addReply(text, config)` for replies. The exporter writes schema-valid brace-delimited GUIDs, a workbook-owned person registry, worksheet-owned threaded-comment parts, `parentId`, `dT`, and per-comment `done` state. Metadata-free import follows arbitrary workbook/worksheet relationship targets, accepts alternate namespace prefixes, preserves native comment/person identity and dates, and keeps multiple independent roots on the same cell separate. Low-level relocation uses a `person` recipe sourced from `xl/workbook.xml` plus a `threadedComments` recipe sourced from the owning worksheet. Package QA must reject external/wrong-source/orphan/multiple parts, invalid or duplicate GUIDs, missing people/parents, cross-cell parents, and cycles.

For bounded low-level package edits, `SpreadsheetFile.patchXlsx(...)` can add or remove worksheet/table references and complete DrawingML chains. Create a `drawing` recipe from the worksheet first, then use `image` or `chart` recipes sourced from that drawing with an explicit `sourceReference.anchor`; the patcher owns content types, relationships, namespace prefixes, non-visual object IDs, source nodes, deletion cleanup, and final package validation.

Worksheet-backed native PivotTables are also relationship-driven. Metadata-free import follows workbook `pivotCaches`, cache definitions, and the worksheet's typed pivotTable relationships at arbitrary valid targets to rebuild editable `sheet.pivotTables` facades. Use `rowFields` plus `columnFields` for cross-tabs, and restrict axis items with exactly one `include` or `exclude` array per filtered field. Use `groupFields: [{ name, sourceField, groupBy }]` for `years`, `quarters`, `months`, `days`, `hours`, `minutes`, or `seconds` fields from ISO, `Date`, or Excel-serial values; order fine-to-coarse levels through the derived names and use `range.groupInterval` for bounded Day/time buckets. Use `groupBy: "range"` with `range: { startNum, endNum, groupInterval }` for numeric buckets; omit boundaries to derive them from numeric source values, or set `autoStart`/`autoEnd` explicitly. Use `groupBy: "discrete"` with `groups: [{ name, items }]` for named manual groups; a typed source item may belong to at most one group and unlisted items remain individually visible. The exporter owns the standard cache `fieldGroup base/par`, `rangePr`, `discretePr`, and `groupItems` structures and preserves full timestamps. Date filters support all eight absolute comparison/range types, using whole-day ISO dates by default. Set `useWholeDay: false` with ISO date-time or `Date` thresholds for UTC-second precision; export writes the public x14 `pivotFilter useWholeDay="0"` extension and metadata-free import accepts any namespace prefix. Relative filters cover `yesterday`/`today`/`tomorrow`, last/this/next week/month/quarter/year, and `yearToDate`; they remain whole-day. Relative calculation is UTC with Monday-start ISO weeks; set filter `asOf` for deterministic workflows. Because native OOXML stores only the dynamic type, pass `SpreadsheetFile.importXlsx(..., { relativeDateAsOf })` when metadata-free QA must reproduce an anchored result. The runnable formula-summary fixture exercises precise absolute filtering through export, inspect, import, modeled verification, and native package rendering. LibreOffice currently renders but does not apply the x14 predicate in its Pivot view, so do not claim precise native semantics without Microsoft Excel validation. `calculatedFields` evaluates grouped source-field sums through arithmetic, comparisons, strings/booleans, 12 bounded numeric functions, `AND`/`OR`/`NOT`, lazy `IF`/`IFERROR`/`IFNA`, `NA`, and `ISERROR`/`ISNUMBER`/`ISTEXT`; reference fields as `[Revenue]` or `'Revenue'`, then use the calculated name from `valueFields`. Lazy functions validate every branch but evaluate only the selected branch, so unused division or `#N/A` errors do not leak. Cell references, calculated-field chaining, unknown fields, non-whitelisted functions, invalid argument counts, and oversized formulas fail explicitly. Imported formulas outside this subset remain available for re-export but produce `#NAME?` and a `pivotCalculatedFieldUnsupported` verification issue, which must not be silently waived. Treat `refreshPolicy` as an OOXML contract: review `refreshOnLoad`, `saveData`, `enableRefresh`, stale-cache `invalid`, retained-item limits, and refresh provenance rather than assuming a save always embeds cache records. With `saveData: false`, the exported package deliberately omits the cache-record part and its relationship. Low-level creation uses `pivotCacheRecords` sourced from its cache definition with `sourceReference: true`, `pivotCacheDefinition` sourced from the workbook with a unique `sourceReference.cacheId`, and a `pivotTable` recipe sourced from the target worksheet without an extra sourceReference node. The cache-definition patch should also add a relationship sourced from the PivotTable part back to that cache definition, matching OpenXML part ownership. Inspect, resolve, layout, verify, package validation, and native rendering must all pass afterward.

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create({ theme: { name: "Report Theme", colors: { accent1: "#0F766E", accent2: "#7C3AED" } } });
const sheet = workbook.worksheets.add("Summary");
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
  --output-dir tmp/spreadsheet-skill-fixture
```

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
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Generic workbook verifier: `scripts/verify-workbook.mjs`
