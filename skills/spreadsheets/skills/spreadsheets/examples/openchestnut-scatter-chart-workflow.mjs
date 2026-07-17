import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildScatterWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Relationship Analysis");
  sheet.getRange("A1:C6").values = [
    ["Ad spend ($k)", "Qualified leads", "Won deals"],
    [10, 24, 4],
    [18, 39, 7],
    [25, 52, 9],
    [34, 71, 13],
    [45, 91, 18],
  ];
  sheet.getRange("A1:C1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A2:A6").format.numberFormat = "$0";
  sheet.getRange("A1:C6").format.autofitColumns();
  sheet.showGridLines = false;
  sheet.freezePanes.freezeRows(1);

  const chart = sheet.charts.add("scatter", sheet.getRange("A1:C6"));
  chart.name = "Pipeline relationship";
  chart.title = "Pipeline outcomes rise with ad spend";
  chart.titleTextStyle.fontSize = 13;
  chart.xAxis = { title: { text: "Ad spend ($k)" }, min: 0, max: 50, majorUnit: 10, numberFormatCode: "$0" };
  chart.yAxis = { title: { text: "Outcome count" }, min: 0, max: 100, majorUnit: 20, numberFormatCode: "0" };
  chart.series.items[0].marker = { symbol: "circle", size: 8, fill: "#0EA5E9", line: { fill: "#075985", width: 1 } };
  chart.series.items[1].marker = { symbol: "diamond", size: 8, fill: "#F97316", line: { fill: "#9A3412", width: 1 } };
  chart.setPosition("E2", "L20");
  return workbook;
}

export async function createScatterWorkbook(outputPath) {
  const workbook = buildScatterWorkbook();
  const sheet = workbook.worksheets.getItem("Relationship Analysis");
  const chart = sheet.charts.items[0];
  assert.deepEqual(chart.series.items.map((series) => series.xValues), [[10, 18, 25, 34, 45], [10, 18, 25, 34, 45]]);
  assert.equal(chart.xAxis.axisType, "valueAxis");

  const inspection = workbook.inspect({ kind: "sheet,drawing", sheetName: sheet.name, maxChars: 12_000 });
  assert.match(inspection.ndjson, /"chartType":"scatter"/);
  assert.match(inspection.ndjson, /"xValues":\[10,18,25,34,45\]/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = chart.toSvg();
  assert.match(previewSvg, /<circle/);
  assert.match(previewSvg, /<polygon/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem(sheet.name);
  const importedChart = importedSheet.charts.items[0];
  assert.equal(importedChart.type, "scatter");
  assert.deepEqual(importedChart.categories, []);
  assert.equal(importedChart.series.items[0].xFormula, "'Relationship Analysis'!A2:A6");
  assert.equal(importedChart.series.items[1].formula, "'Relationship Analysis'!C2:C6");

  importedSheet.getRange("A3:C3").values = [[20, 43, 8]];
  for (const series of importedChart.series.items) series.xValues[1] = 20;
  importedChart.series.items[0].values[1] = 43;
  importedChart.series.items[1].values[1] = 8;
  importedChart.title = "Edited pipeline relationship";
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  const roundTripChart = roundTrip.worksheets.getItem(sheet.name).charts.items[0];
  assert.equal(roundTripChart.title, "Edited pipeline relationship");
  assert.deepEqual(roundTripChart.series.items[0].xValues, [10, 20, 25, 34, 45]);
  assert.deepEqual(roundTripChart.series.items[1].values, [4, 8, 9, 13, 18]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-scatter-chart-workflow.xlsx");
  const result = await createScatterWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
