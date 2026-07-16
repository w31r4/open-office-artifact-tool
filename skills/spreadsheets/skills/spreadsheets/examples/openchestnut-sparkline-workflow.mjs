import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildSparklineWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Operating Trends");
  sheet.getRange("A1:H4").write([
    ["Metric", "Jan", "Feb", "Mar", "Apr", "May", "Trend", "Variance"],
    ["Revenue", 72, 81, 75, 96, 110, null, null],
    ["Margin", 40, 44, 38, 51, 58, null, null],
    ["Variance", -12, 8, -4, 15, 21, null, null],
  ]);
  sheet.getRange("A1:H1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A2:A4").format = { font: { bold: true } };
  sheet.getRange("B2:F4").setNumberFormat("0.0");
  sheet.getRange("A1:H4").format.autofitColumns();
  sheet.getRange("G2:H4").format.columnWidth = 108;
  sheet.getRange("G2:H4").format.rowHeight = 32;
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;

  sheet.sparklineGroups.add({
    type: "line",
    targetRange: "G2:G4",
    sourceData: "B2:F4",
    dateAxisRange: "B1:F1",
    seriesColor: "#0EA5E9",
    negativeColor: "#DC2626",
    markers: { show: true, high: true, low: true, negative: true },
    highMarkerColor: "#22C55E",
    lowMarkerColor: "#EF4444",
    axis: { showAxis: true, manualMin: -20, manualMax: 120, minMode: 2, maxMode: 2 },
    lineWeight: 2,
  });
  sheet.getRange("H2:H4").sparklines.add("column", sheet.getRange("B2:F4"), {
    seriesColor: { theme: 4, tint: 0.1 },
    negativeColor: "#B91C1C",
  });
  return workbook;
}

export async function createSparklineWorkbook(outputPath) {
  const workbook = buildSparklineWorkbook();
  const sheet = workbook.worksheets.getItem("Operating Trends");
  assert.equal(sheet.sparklineGroups.items.length, 2);
  assert.equal(sheet.sparklineGroups.items[0].sparklineCount, 3);

  const inspection = workbook.inspect({ kind: "sheet,sparkline,style", sheetName: sheet.name, range: "A1:H4", maxChars: 16_000 });
  assert.match(inspection.ndjson, /"kind":"sparkline"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: sheet.name, range: "A1:H4", format: "svg" });
  assert.match(await preview.text(), /<polyline/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem(sheet.name);
  assert.deepEqual(importedSheet.sparklineGroups.items.map((group) => group.type), ["line", "column"]);
  importedSheet.sparklineGroups.items[0].seriesColor = "#F97316";
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  assert.equal(roundTrip.worksheets.getItem(sheet.name).sparklineGroups.items[0].seriesColor, "#F97316");

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, preview };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-sparkline-workflow.xlsx");
  const result = await createSparklineWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
