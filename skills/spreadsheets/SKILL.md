---
name: open-office-spreadsheets
description: Create, edit, inspect, render, and verify XLSX/CSV spreadsheet artifacts with open-office-artifact-tool.
---

# Open Office Spreadsheets

Use this project skill for standalone `.xlsx`, `.csv`, and `.tsv` artifact work. It is the clean-room workflow for `open-office-artifact-tool`; it does not control a live Microsoft Excel session.

## Contract

- Use `open-office-artifact-tool` public exports. Do not import the reference package or its compiled internals.
- Prefer one writable `.mjs` builder that can be rerun after focused edits.
- Write rectangular value/formula matrices in blocks and keep derived values as formulas.
- Preserve the style and formulas of an imported workbook unless the user asks for a redesign.
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

```js
import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Summary");
sheet.getRange("A1:C3").values = [
  ["Month", "Revenue", "Cost"],
  ["Jan", 100, 60],
  ["Feb", 120, 70],
];
sheet.getRange("D1").values = [["Margin"]];
sheet.getRange("D2:D3").formulas = [["=(B2-C2)/B2"], ["=(B3-C3)/B3"]];
sheet.getRange("A1:D1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
  alignment: { horizontal: "center" },
};
sheet.getRange("B2:C3").format = { numberFormat: "$#,##0" };
sheet.getRange("D2:D3").format = { numberFormat: "0.0%" };
workbook.recalculate();

const output = await SpreadsheetFile.exportXlsx(workbook);
await output.save("output.xlsx");

const csv = await SpreadsheetFile.exportCsv(workbook, { sheetName: "Summary", range: "A1:D20" });
await csv.save("output.csv");
const importedTsv = await SpreadsheetFile.importTsv("Name\tValue\nRevenue\t120", { coerceTypes: true });
```

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
- Save `SpreadsheetFile.inspectXlsx()` evidence so the native package shape, content types, and decompression budgets are part of QA.
- For CSV/TSV, save `SpreadsheetFile.inspectDelimited()` evidence and review delimiter, dimensions, quoting, BOM, and formula-like cell counts; remember that delimited delivery cannot preserve workbook-only features.
- Produce a layout record and visual preview for every user-facing sheet with `--all-sheets true`; the selected `--range` narrows only the primary sheet.
- Use Playwright raster output for deterministic previews; use LibreOffice or Microsoft Office when native application fidelity is the acceptance criterion.
- Raster baselines use `visualQaArtifact(..., { pixelDiff: true })`; baseline filenames are stable per worksheet.
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
