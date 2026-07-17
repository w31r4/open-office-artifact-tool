import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildBubbleWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Opportunity Analysis");
  sheet.getRange("A1:C6").values = [
    ["Active customers", "Annual revenue ($k)", "Pipeline size ($k)"],
    [10, 42, 4],
    [18, 61, 8],
    [25, 73, 12],
    [34, 91, 18],
    [45, 118, 27],
  ];
  sheet.getRange("A1:C1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A2:C6").format.numberFormat = "0";
  // Keep the source table and chart inside one printable native-render page.
  sheet.getRange("A1:A6").format.columnWidthPx = 112;
  sheet.getRange("B1:B6").format.columnWidthPx = 126;
  sheet.getRange("C1:C6").format.columnWidthPx = 126;
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);

  // The range shortcut is deliberately unambiguous: exactly X | Y | Size.
  // Use explicit series config when a bubble chart needs multiple series.
  const chart = sheet.charts.add("bubble", sheet.getRange("A1:C6"));
  chart.name = "Revenue opportunity";
  chart.title = "Larger pipelines cluster in higher-value accounts";
  chart.titleTextStyle.fontSize = 13;
  chart.xAxis = { title: { text: "Active customers" }, min: 0, max: 50, majorUnit: 10, numberFormatCode: "0" };
  chart.yAxis = { title: { text: "Annual revenue ($k)" }, min: 0, max: 130, majorUnit: 25, numberFormatCode: "$0" };
  chart.series.items[0].fill = "#0EA5E9";
  chart.series.items[0].line = { fill: "#075985", width: 1 };
  chart.setPosition("A8", "F21");
  return workbook;
}

export async function createBubbleWorkbook(outputPath) {
  const workbook = buildBubbleWorkbook();
  const sheet = workbook.worksheets.getItem("Opportunity Analysis");
  const chart = sheet.charts.items[0];
  assert.deepEqual(chart.series.items[0].xValues, [10, 18, 25, 34, 45]);
  assert.deepEqual(chart.series.items[0].bubbleSizes, [4, 8, 12, 18, 27]);
  assert.equal(chart.xAxis.axisType, "valueAxis");

  const inspection = workbook.inspect({ kind: "sheet,drawing", sheetName: sheet.name, maxChars: 12_000 });
  assert.match(inspection.ndjson, /"chartType":"bubble"/);
  assert.match(inspection.ndjson, /"bubbleSizes":\[4,8,12,18,27\]/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = chart.toSvg();
  assert.match(previewSvg, /data-bubble-size="27"/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem(sheet.name);
  const importedChart = importedSheet.charts.items[0];
  assert.equal(importedChart.type, "bubble");
  assert.deepEqual(importedChart.categories, []);
  assert.equal(importedChart.series.items[0].xFormula, "'Opportunity Analysis'!A2:A6");
  assert.equal(importedChart.series.items[0].formula, "'Opportunity Analysis'!B2:B6");
  assert.equal(importedChart.series.items[0].bubbleSizeFormula, "'Opportunity Analysis'!C2:C6");

  importedSheet.getRange("A3:C3").values = [[20, 66, 10]];
  importedChart.series.items[0].xValues[1] = 20;
  importedChart.series.items[0].values[1] = 66;
  importedChart.series.items[0].bubbleSizes[1] = 10;
  importedChart.title = "Edited account opportunity";
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  const roundTripChart = roundTrip.worksheets.getItem(sheet.name).charts.items[0];
  assert.equal(roundTripChart.title, "Edited account opportunity");
  assert.deepEqual(roundTripChart.series.items[0].xValues, [10, 20, 25, 34, 45]);
  assert.deepEqual(roundTripChart.series.items[0].values, [42, 66, 73, 91, 118]);
  assert.deepEqual(roundTripChart.series.items[0].bubbleSizes, [4, 10, 12, 18, 27]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-bubble-chart-workflow.xlsx");
  const result = await createBubbleWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
