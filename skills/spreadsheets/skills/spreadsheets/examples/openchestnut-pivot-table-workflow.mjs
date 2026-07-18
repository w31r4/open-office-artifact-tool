import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildPivotTableWorkbook() {
  const workbook = Workbook.create();
  const data = workbook.worksheets.add("Data");
  data.getRange("A1:C7").write([
    ["Region", "Channel", "Revenue"],
    ["East", "Direct", 120],
    ["East", "Partner", 80],
    ["West", "Direct", 150],
    ["West", "Partner", 90],
    ["North", "Direct", 110],
    ["North", "Partner", 70],
  ]);
  data.getRange("A1:C1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("C2:C7").setNumberFormat("$#,##0");
  data.getRange("A1:C7").format.autofitColumns();
  data.freezePanes.freezeRows(1);
  data.showGridLines = false;

  const summary = workbook.worksheets.add("Pivot Summary");
  summary.getRange("A1:D5").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:D1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A" } };
  summary.getRange("A5:D5").format = { fill: "#E2E8F0", font: { bold: true, color: "#0F172A" } };
  summary.getRange("B2:D5").setNumberFormat("$#,##0");
  summary.getRange("A1:A5").format.columnWidthPx = 112;
  summary.getRange("B1:C5").format.columnWidthPx = 88;
  summary.getRange("D1:D5").format.columnWidthPx = 112;
  summary.showGridLines = false;
  summary.pivotTables.add({
    name: "Revenue by region",
    sourceRange: "Data!A1:C7",
    targetRange: "A1",
    rowFields: ["Region"],
    columnFields: ["Channel"],
    valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue" }],
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
    ["Region", "Direct", "Partner", "Grand Total"],
    ["East", 120, 80, 200],
    ["West", 150, 90, 240],
    ["North", 110, 70, 180],
    ["Grand Total", 380, 240, 620],
  ]);

  const inspection = workbook.inspect({ kind: "sheet,pivotTable,style", sheetName: summary.name, range: "A1:D5", maxChars: 16_000 });
  assert.match(inspection.ndjson, /"kind":"pivotTable"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: summary.name, range: "A1:D5", format: "svg" });
  assert.match(await preview.text(), /Revenue by region/);

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
