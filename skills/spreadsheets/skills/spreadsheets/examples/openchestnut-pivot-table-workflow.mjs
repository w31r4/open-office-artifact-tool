import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildPivotTableWorkbook() {
  const workbook = Workbook.create();
  const data = workbook.worksheets.add("Data");
  data.getRange("A1:E13").write([
    ["Region", "Channel", "Product", "Revenue", "Units"],
    ["East", "Direct", "Alpha", 70, 7],
    ["East", "Direct", "Beta", 50, 5],
    ["East", "Partner", "Alpha", 45, 4],
    ["East", "Partner", "Beta", 35, 4],
    ["West", "Direct", "Alpha", 90, 9],
    ["West", "Direct", "Beta", 60, 6],
    ["West", "Partner", "Alpha", 55, 5],
    ["West", "Partner", "Beta", 35, 4],
    ["North", "Direct", "Alpha", 60, 6],
    ["North", "Direct", "Beta", 50, 5],
    ["North", "Partner", "Alpha", 40, 4],
    ["North", "Partner", "Beta", 30, 3],
  ]);
  data.getRange("A1:E1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("D2:D13").setNumberFormat("$#,##0");
  data.getRange("E2:E13").setNumberFormat("#,##0");
  data.getRange("A1:E13").format.autofitColumns();
  data.freezePanes.freezeRows(1);
  data.showGridLines = false;

  const summary = workbook.worksheets.add("Pivot Summary");
  summary.getRange("A1:H6").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:H1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A", size: 9 } };
  summary.getRange("A1:H1").format.wrapText = true;
  summary.getRange("A1:H1").format.rowHeightPx = 42;
  summary.getRange("A6:H6").format = { fill: "#E2E8F0", font: { bold: true, color: "#0F172A" } };
  for (const range of ["C2:C6", "E2:E6", "G2:G6"]) summary.getRange(range).setNumberFormat("$#,##0");
  for (const range of ["D2:D6", "F2:F6", "H2:H6"]) summary.getRange(range).setNumberFormat("#,##0");
  // Keep the eight-column summary within the same 440 px width budget that was
  // stable on GitHub's Linux LibreOffice/metric stack. The smaller header font
  // and explicit wrap height preserve the long value labels without clipping.
  summary.getRange("A1:A6").format.columnWidthPx = 58;
  summary.getRange("B1:B6").format.columnWidthPx = 54;
  summary.getRange("C1:F6").format.columnWidthPx = 46;
  summary.getRange("G1:G6").format.columnWidthPx = 78;
  summary.getRange("H1:H6").format.columnWidthPx = 66;
  summary.showGridLines = false;
  summary.pivotTables.add({
    name: "Revenue and units by region",
    sourceRange: "Data!A1:E13",
    targetRange: "A1",
    rowFields: ["Region", "Channel"],
    columnFields: ["Product"],
    valueFields: [
      { field: "Revenue", summarizeBy: "sum", name: "Revenue" },
      { field: "Units", summarizeBy: "sum", name: "Units" },
    ],
    filters: [{ field: "Region", exclude: ["North"] }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: true, saveData: true, enableRefresh: true },
  });
  return workbook;
}

export async function createPivotTableWorkbook(outputPath) {
  const workbook = buildPivotTableWorkbook();
  const summary = workbook.worksheets.getItem("Pivot Summary");
  const pivot = summary.pivotTables.items[0];
  assert.deepEqual(pivot.computedValues(), [
    ["Region", "Channel", "Alpha — Revenue", "Alpha — Units", "Beta — Revenue", "Beta — Units", "Grand Total — Revenue", "Grand Total — Units"],
    ["East", "Direct", 70, 7, 50, 5, 120, 12],
    ["East", "Partner", 45, 4, 35, 4, 80, 8],
    ["West", "Direct", 90, 9, 60, 6, 150, 15],
    ["West", "Partner", 55, 5, 35, 4, 90, 9],
    ["Grand Total", "", 260, 25, 180, 19, 440, 44],
  ]);

  const inspection = workbook.inspect({ kind: "sheet,pivotTable,style", sheetName: summary.name, range: "A1:H6", maxChars: 16_000 });
  assert.match(inspection.ndjson, /"kind":"pivotTable"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: summary.name, range: "A1:H6", format: "svg" });
  assert.match(await preview.text(), /Revenue and units by region/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedPivot = imported.worksheets.getItem(summary.name).pivotTables.items[0];
  assert.deepEqual(importedPivot.computedValues(), pivot.computedValues());
  assert.deepEqual(importedPivot.filters, [{ field: "Region", exclude: ["North"] }]);
  assert.equal(importedPivot.refreshPolicy.saveData, true);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  assert.deepEqual(roundTrip.worksheets.getItem(summary.name).pivotTables.items[0].computedValues(), pivot.computedValues());

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, preview };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-pivot-table-workflow.xlsx");
  const result = await createPivotTableWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
