---
name: open-office-spreadsheets
description: Create, edit, inspect, render, and verify XLSX/CSV spreadsheet artifacts with open-office-artifact-tool.
---

# Open Office Spreadsheets

Use this project skill for standalone `.xlsx`, `.csv`, and `.tsv` artifact work. It does not control a live Microsoft Excel session.

XLSX has one codec path. `SpreadsheetFile.importXlsx(...)` and `SpreadsheetFile.exportXlsx(...)` always use the bundled OpenChestnut C# WebAssembly codec. There is no JavaScript Office codec, selector, or fallback. Import accepts only `limits`; export accepts only `limits` and `recalculate`. Passing legacy options such as `codec`, `allowLossy`, `preferNative`, or `relativeDateAsOf` is an error.

## Contract

- Use the public exports from `open-office-artifact-tool`; do not import codec internals.
- Prefer one rerunnable `.mjs` builder and write rectangular value/formula matrices in blocks.
- Keep derived values as formulas, call `workbook.recalculate()`, and inspect important cached results.
- Preserve imported formulas, formatting, and source-bound package content unless the user asks to change them.
- A workbook is deliverable only after XLSX export, package inspection, canonical re-import, semantic verification, and visual review pass.
- `SpreadsheetFile.inspectXlsx(...)` and `patchXlsx(...)` are explicit low-level package tools. The facade never invokes them as a codec fallback.

## Canonical XLSX boundary

OpenChestnut authors and roundtrips the mainstream workflow:

- string, number, boolean, date-serial, blank, and formula cells with cached values;
- number formats, fonts, fills, borders, alignment, protection, row/column dimensions, hidden axes, merges, sheet visibility, active/selected sheets, and freeze panes;
- ordinary defined names, calculation policy, worksheet tables, and bounded value/custom/date/dynamic/top/icon/color filters and sorts;
- embedded PNG/JPEG pictures with bounded one-cell, two-cell, or absolute anchors;
- one-plot bar, line, and pie charts with deterministic cached categories/values;
- list/whole/decimal/date/time/text-length/custom data validation;
- `cellIs`, `expression`, `containsText`, and two/three-color-scale conditional formatting;
- root-only Office 2019 threaded comments with one author and resolution state per thread.

JavaScript `Date` values are converted to the workbook's Excel serial system before crossing the wire. Use `Workbook.create({ dateSystem: "1900" })` or `"1904"` deliberately; do not manually add or subtract 1,462 days.

Imported objects outside this boundary remain in the hash-bound source package. An unchanged source-bound workbook preserves them byte-for-byte where required. Adding or editing an unsupported object fails explicitly; opaque content without its trusted source snapshot also fails. Do not attempt to work around this with a package reconstruction or alternate export path.

Examples of unsupported creation/editing include PivotTables, threaded-comment replies, dynamic-array extension semantics, data-table formulas, grouped shapes, combo or multi-plot charts, external QueryTable creation, and unknown style/chart/comment graphs. A recognized imported QueryTable may be edited only within its fixed source topology.

## Authoring workflow

1. Create a `Workbook` or import an XLSX through `SpreadsheetFile.importXlsx(input, { limits? })`.
2. Inspect the relevant sheets, ranges, formulas, styles, tables, drawings, validations, conditional formats, and comments.
3. Apply only changes inside the canonical boundary.
4. Recalculate and spot-check important values and formula traces.
5. Export with `SpreadsheetFile.exportXlsx(workbook, { limits?, recalculate? })`.
6. Inspect the native package, import the exported XLSX again, and run `verifyArtifact`.
7. Render every user-facing worksheet and review the previews at full size.
8. When native fidelity matters, render the real XLSX through LibreOffice to PDF and Poppler page PNGs and compare every page.

CSV and TSV flatten one worksheet or range. They cannot preserve styles, drawings, validations, comments, or workbook formulas. Delimited export writes calculated values by default; use `formulas: true` only when formula text is the requested deliverable.

## Supported authoring example

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create({ dateSystem: "1900" });
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:D4").values = [
  ["Month", "Revenue", "Cost", "Status"],
  [new Date("2026-01-01T00:00:00Z"), 100, 60, "Done"],
  [new Date("2026-02-01T00:00:00Z"), 120, 70, "In progress"],
  [new Date("2026-03-01T00:00:00Z"), 150, 90, "Planned"],
];
sheet.getRange("E1").values = [["Margin"]];
sheet.getRange("E2").formulas = [["=(B2-C2)/B2"]];
sheet.getRange("E2:E4").fillDown();
sheet.getRange("A1:E1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
};
sheet.getRange("A2:A4").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("B2:C4").format.numberFormat = "$#,##0";
sheet.getRange("E2:E4").format.numberFormat = "0.0%";
sheet.getRange("D2:D4").dataValidation = {
  rule: { type: "list", values: ["Planned", "In progress", "Done"] },
};
sheet.getRange("E2:E4").conditionalFormats.add("cellIs", {
  operator: "greaterThan",
  formula: 0.4,
  format: { fill: "#DCFCE7" },
});
sheet.tables.add("A1:E4", true, "SummaryTable").style = "TableStyleMedium4";
sheet.freezePanes.freezeRows(1);
sheet.charts.add("line", sheet.getRange("A1:B4")).setPosition("G1", "L12");
workbook.comments.addThread(
  { cell: sheet.getRange("E2") },
  "Check the margin before publishing.",
  { author: "Reviewer", date: "2026-07-16T09:00:00.000Z" },
);
workbook.recalculate();

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save("output.xlsx");

const roundtrip = await SpreadsheetFile.importXlsx(output);
const second = await SpreadsheetFile.exportXlsx(roundtrip, { recalculate: false });
await second.save("output-roundtrip.xlsx");
```

Do not call `thread.addReply(...)` in an XLSX authoring workflow. Replies are an imported opaque profile and cannot be newly created or edited through the canonical model.

## Tables, pictures, and charts

Use worksheet tables for rectangular data with unique non-empty column names. Keep formulas leading-`=` and keep filter/sort references inside the table range. Table topology in a recognized source-bound QueryTable graph is fixed.

Add images with `sheet.images.add({ name, alt, dataUrl, anchor })`. `dataUrl` must be base64 PNG or JPEG. Use one-cell `{ from, extent }`, two-cell `{ type: "twoCell", from, to }`, or absolute `{ type: "absolute", position, extent }` anchors. Imported image replacement must retain the original PNG/JPEG MIME type and fixed topology.

Use `sheet.charts.add("bar" | "line" | "pie", sourceRangeOrConfig)`. Keep one plot, stable chart type, finite cached values, and fixed imported series/point topology. Advanced native chart graphs remain opaque and read-only.

## Verification commands

Verify any exported workbook and write bounded inspect, package, semantic, layout, preview, and visual-QA evidence:

```sh
node skills/spreadsheets/scripts/verify-workbook.mjs \
  --input output.xlsx \
  --output-dir tmp/spreadsheet-qa \
  --sheet Summary \
  --range A1:E20 \
  --render-format png
```

Run each checked-in canonical fixture through create, import, optional fixed-topology edit, second export, inspect, render, and verify:

```sh
node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/formula-summary.json \
  --render-format png --all-sheets true \
  --native-render required \
  --output-dir tmp/spreadsheet-skill-fixture

node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/open-chestnut-basic.json \
  --native-render required \
  --output-dir tmp/open-chestnut-spreadsheet-fixture

node skills/spreadsheets/scripts/run-fixture.mjs \
  --fixture skills/spreadsheets/fixtures/open-chestnut-query-table.json \
  --native-render off \
  --output-dir tmp/open-chestnut-query-table-fixture
```

`--render-format svg` is dependency-light. Raster formats require the Playwright Chromium runtime. `--native-render required` requires LibreOffice, `pdfinfo`, and `pdftoppm`.

## QA gates

- Inspect the key table/formula range with bounded `maxChars`.
- Resolve every error-level `verifyArtifact(workbook)` issue.
- Export and canonical-reimport the XLSX before final verification.
- Require zero package relationship-reference issues.
- For validation and conditional formatting, inspect both the modeled rule and the worksheet XML after roundtrip.
- For threaded comments, require one root per modeled thread, valid native GUID/person/date metadata, and no reply topology.
- Produce a preview for every user-facing sheet; use `--all-sheets true` for multi-sheet delivery.
- When a baseline is supplied, missing sheets/pages and pixel changes fail closed.
- Deliver only the requested workbook unless the user asks for QA intermediates.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Workbook verifier: `scripts/verify-workbook.mjs`
