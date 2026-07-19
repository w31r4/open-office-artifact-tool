import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildPivotTableWorkbook() {
  const workbook = Workbook.create();
  const data = workbook.worksheets.add("Data");
  data.getRange("A1:D7").write([
    ["Region", "Channel", "Revenue", "Units"],
    ["East", "Direct", 120, 12],
    ["East", "Partner", 80, 8],
    ["West", "Direct", 150, 15],
    ["West", "Partner", 90, 9],
    ["North", "Direct", 110, 11],
    ["North", "Partner", 70, 7],
  ]);
  data.getRange("A1:D1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("C2:C7").setNumberFormat("$#,##0");
  data.getRange("D2:D7").setNumberFormat("#,##0");
  data.getRange("A1:D7").format.autofitColumns();
  data.freezePanes.freezeRows(1);
  data.showGridLines = false;

  const summary = workbook.worksheets.add("Pivot Summary");
  summary.getRange("A1:G5").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:G1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A" } };
  summary.getRange("A1:G1").format.wrapText = true;
  summary.getRange("A1:G1").format.rowHeightPx = 34;
  summary.getRange("A5:G5").format = { fill: "#E2E8F0", font: { bold: true, color: "#0F172A" } };
  for (const range of ["B2:B5", "D2:D5", "F2:F5"]) summary.getRange(range).setNumberFormat("$#,##0");
  for (const range of ["C2:C5", "E2:E5", "G2:G5"]) summary.getRange(range).setNumberFormat("#,##0");
  summary.getRange("A1:A5").format.columnWidthPx = 112;
  summary.getRange("B1:E5").format.columnWidthPx = 80;
  summary.getRange("F1:G5").format.columnWidthPx = 96;
  summary.showGridLines = false;
  summary.pivotTables.add({
    name: "Revenue and units by region",
    sourceRange: "Data!A1:D7",
    targetRange: "A1",
    rowFields: ["Region"],
    columnFields: ["Channel"],
    valueFields: [
      { field: "Revenue", summarizeBy: "sum", name: "Revenue" },
      { field: "Units", summarizeBy: "sum", name: "Units" },
    ],
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
    ["Region", "Direct — Revenue", "Direct — Units", "Partner — Revenue", "Partner — Units", "Grand Total — Revenue", "Grand Total — Units"],
    ["East", 120, 12, 80, 8, 200, 20],
    ["West", 150, 15, 90, 9, 240, 24],
    ["North", 110, 11, 70, 7, 180, 18],
    ["Grand Total", 380, 38, 240, 24, 620, 62],
  ]);

  const inspection = workbook.inspect({ kind: "sheet,pivotTable,style", sheetName: summary.name, range: "A1:G5", maxChars: 16_000 });
  assert.match(inspection.ndjson, /"kind":"pivotTable"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: summary.name, range: "A1:G5", format: "svg" });
  assert.match(await preview.text(), /Revenue and units by region/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedPivot = imported.worksheets.getItem(summary.name).pivotTables.items[0];
  assert.deepEqual(importedPivot.computedValues(), pivot.computedValues());
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
