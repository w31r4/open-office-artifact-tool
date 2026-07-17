import assert from "node:assert/strict";
import JSZip from "jszip";

import { SpreadsheetFile, Workbook, verifyArtifact } from "../src/index.mjs";

const PNG_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC";

const workbook = Workbook.create({ dateSystem: "1900" });
const sheet = workbook.worksheets.add("Summary");
sheet.showGridLines = false;
sheet.getRange("A1:E4").values = [
  ["Month", "Revenue", "Cost", "Status", "Date"],
  ["Jan", 100, 60, "Done", new Date("2026-01-15T00:00:00.000Z")],
  ["Feb", 120, 70, "Review", new Date("2026-02-15T00:00:00.000Z")],
  ["Mar", 150, 90, "Planned", new Date("2026-03-15T00:00:00.000Z")],
];
sheet.getRange("F1:F4").values = [["Margin"], [null], [null], [null]];
sheet.getRange("F2:F4").formulas = [
  ["=(B2-C2)/B2"],
  ["=(B3-C3)/B3"],
  ["=(B4-C4)/B4"],
];
sheet.getRange("A1:F1").format = {
  fill: "#0F766E",
  font: { bold: true, color: "#FFFFFF" },
  border: { bottom: { style: "double", color: "#38BDF8" } },
  alignment: { horizontal: "center", vertical: "center" },
};
sheet.getRange("B2:C4").format = { numberFormat: "$#,##0.00" };
sheet.getRange("E2:E4").format = { numberFormat: "yyyy-mm-dd" };
sheet.getRange("F2:F4").format = {
  numberFormat: "0.0%",
  protection: { locked: false, hidden: true },
};
sheet.getRange("A1:A6").format.columnWidthPx = 96;
sheet.getRange("B1:F6").format.columnWidthPx = 84;
sheet.getRange("A1:F1").format.rowHeightPx = 28;
sheet.getRange("A8:F8").format.rowHidden = true;
sheet.getRange("A6:F6").values = [["Quarter summary", null, null, null, null, null]];
sheet.getRange("A6:F6").merge();
sheet.freezePanes.freezeRows(1);
sheet.freezePanes.freezeColumns(1);

const table = sheet.tables.add("A1:F4", true, "SummaryTable");
table.style = "TableStyleMedium4";
table.showRowStripes = true;

sheet.getRange("D2:D4").dataValidation = {
  rule: { type: "list", values: ["Planned", "Review", "Done"] },
};
sheet.dataValidations.add({
  range: "B2:B4",
  rule: { type: "whole", operator: "between", formula1: "0", formula2: "1000" },
});
sheet.getRange("F2:F4").conditionalFormats.add("cellIs", {
  operator: "greaterThan",
  formula: "0.4",
  format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
});
sheet.getRange("B2:B4").conditionalFormats.add("expression", {
  formula: "B2>C2",
  format: { fill: "#DBEAFE" },
});
sheet.getRange("D2:D4").conditionalFormats.add("containsText", {
  text: "Review",
  formula: "NOT(ISERROR(SEARCH(\"Review\",D2)))",
  format: { fill: "#FEF3C7" },
});
sheet.getRange("C2:C4").conditionalFormats.addColorScale({
  colors: ["#FEE2E2", "#FEF3C7", "#22C55E"],
});

const bar = sheet.charts.add("bar", sheet.getRange("A1:C4"));
bar.name = "Bar chart";
bar.title = "Revenue and cost";
bar.setPosition("H1", "L10");
const line = sheet.charts.add("line", sheet.getRange("A1:C4"));
line.name = "Line chart";
line.title = "Revenue trend";
line.setPosition("M1", "Q10");
const pie = sheet.charts.add("pie", sheet.getRange("A1:B4"));
pie.name = "Pie chart";
pie.title = "Revenue share";
pie.setPosition("H12", "L22");
const area = sheet.charts.add("area", sheet.getRange("A1:C4"));
area.name = "Area chart";
area.title = "Revenue and cost area";
area.setPosition("M12", "Q22");
const doughnut = sheet.charts.add("doughnut", sheet.getRange("A1:B4"));
doughnut.name = "Doughnut chart";
doughnut.title = "Revenue mix";
doughnut.setPosition("H24", "L34");
sheet.getRange("S1:T4").values = [["Units", "Price"], [10, 30], [20, 55], [30, 88]];
const scatter = sheet.charts.add("scatter", sheet.getRange("S1:T4"));
scatter.name = "Scatter chart";
scatter.title = "Price relationship";
scatter.xAxis = { title: { text: "Units" }, min: 0, max: 40, majorUnit: 10, numberFormatCode: "0" };
scatter.yAxis = { title: { text: "Price" }, min: 0, max: 100, majorUnit: 20, numberFormatCode: "$0" };
scatter.series.items[0].marker = { symbol: "diamond", size: 8, fill: "#38BDF8" };
scatter.setPosition("M24", "Q34");
sheet.images.add({
  name: "Status mark",
  alt: "Green status marker",
  dataUrl: PNG_DATA_URL,
  anchor: { from: { row: 6, col: 0, rowOffsetPx: 4, colOffsetPx: 4 }, extent: { widthPx: 64, heightPx: 48 } },
});

workbook.comments.setSelf({ displayName: "Spreadsheet Agent" });
const marginReviewThread = workbook.comments.addThread(
  { cell: sheet.getRange("F2") },
  "Check the calculated margin.",
  {
    id: "margin-review",
    author: "Reviewer",
    resolved: true,
    comment: {
      id: "{11111111-1111-4111-8111-111111111111}",
      personId: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      date: "2026-07-16T09:00:00.000Z",
      person: {
        id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
        displayName: "Reviewer",
        userId: "reviewer@example.com",
        providerId: "None",
      },
    },
  },
);
marginReviewThread.addReply("Confirmed against the source workbook.", {
  id: "{22222222-2222-4222-8222-222222222222}",
  personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
  author: "Lead reviewer",
  date: "2026-07-16T09:30:00.000Z",
  person: {
    id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    displayName: "Lead reviewer",
    userId: "lead@example.com",
    providerId: "None",
  },
  done: true,
});

workbook.recalculate();
assert.deepEqual(sheet.getRange("F2:F4").values, [[0.4], [0.4166666666666667], [0.4]]);
assert.equal(sheet.getRange("E2").values[0][0] instanceof Date, true);
const modelInspect = workbook.inspect({
  kind: "workbook,sheet,table,formula,style,drawing,dataValidation,conditionalFormat,thread",
  sheetName: "Summary",
  range: "A1:Q34",
  maxChars: 32_000,
});
assert.match(modelInspect.ndjson, /"name":"SummaryTable"/);
assert.match(modelInspect.ndjson, /"formula":"=\(B2-C2\)\/B2"/);
assert.match(modelInspect.ndjson, /"drawingType":"chart"/);
assert.match(modelInspect.ndjson, /"kind":"dataValidation"/);
assert.match(modelInspect.ndjson, /"kind":"conditionalFormat"/);
assert.match(modelInspect.ndjson, /Check the calculated margin/);
const modelVerification = verifyArtifact(workbook);
assert.equal(modelVerification.ok, true, modelVerification.ndjson);

const firstXlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(firstXlsx.metadata.codec, "open-chestnut");
const packageInspect = await SpreadsheetFile.inspectXlsx(firstXlsx, { maxChars: 32_000 });
assert.equal(packageInspect.ok, true, packageInspect.ndjson);
assert.equal(packageInspect.records[0].semanticIssues, 0);

const firstZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
assert.ok(firstZip.file("xl/workbook.xml"));
assert.ok(firstZip.file("xl/worksheets/sheet1.xml"));
assert.ok(firstZip.file("xl/tables/table1.xml"));
assert.ok(firstZip.file("xl/drawings/drawing1.xml"));
const firstChartPaths = Object.keys(firstZip.files).filter((name) => /\/charts\/chart\d+\.xml$/i.test(name));
assert.equal(firstChartPaths.length, 6);
const firstChartXml = await Promise.all(firstChartPaths.map((name) => firstZip.file(name).async("text")));
assert.equal(firstChartXml.filter((xml) => /<c:scatterChart>/.test(xml)).length, 1);
assert.match(firstChartXml.find((xml) => /<c:scatterChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>/);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/media\//i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(firstZip.file("customXml/open-office-artifact.json"), null);
const firstWorksheetXml = await firstZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(firstWorksheetXml, /<x:mergeCell ref="A6:F6"/);
assert.match(firstWorksheetXml, /<x:dataValidations count="2">/);
assert.equal((firstWorksheetXml.match(/<x:conditionalFormatting\b/g) || []).length, 4);
const firstThreadedPart = Object.keys(firstZip.files).find((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
const firstThreadedXml = await firstZip.file(firstThreadedPart).async("text");
assert.match(firstThreadedXml, /parentId="\{11111111-1111-4111-8111-111111111111\}"/);
assert.match(firstThreadedXml, /Confirmed against the source workbook\./);

const imported = await SpreadsheetFile.importXlsx(firstXlsx);
const importedSheet = imported.worksheets.getItem("Summary");
assert.ok(importedSheet);
assert.deepEqual(importedSheet.getRange("F2:F4").values, [[0.4], [0.4166666666666667], [0.4]]);
assert.deepEqual(importedSheet.getRange("F2:F4").formulas, [
  ["=(B2-C2)/B2"],
  ["=(B3-C3)/B3"],
  ["=(B4-C4)/B4"],
]);
assert.equal(typeof importedSheet.getRange("E2").values[0][0], "number");
assert.equal(importedSheet.getRange("E2").format.numberFormat, "yyyy-mm-dd");
assert.equal(importedSheet.getRange("A1").format.fill, "#0F766E");
assert.equal(importedSheet.getRange("A1").format.font.bold, true);
assert.equal(importedSheet.getRange("A1").format.border.bottom.style, "double");
assert.deepEqual(importedSheet.getRange("F2").format.protection, { locked: false, hidden: true });
assert.deepEqual(importedSheet.mergedRanges, ["A6:F6"]);
assert.ok(Math.abs(importedSheet.getRange("A1:A6").format.columnWidthPx - 96) <= 1);
assert.ok(Math.abs(importedSheet.getRange("A1:F1").format.rowHeightPx - 28) < 0.01);
assert.equal(importedSheet.getRange("A8:F8").format.rowHidden, true);
assert.deepEqual(importedSheet.freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
assert.equal(importedSheet.tables.items[0].name, "SummaryTable");
assert.equal(importedSheet.images.items[0].alt, "Green status marker");
assert.deepEqual(importedSheet.charts.items.map((chart) => chart.type), ["bar", "line", "pie", "area", "doughnut", "scatter"]);
assert.match(importedSheet.charts.items[3].toSvg(), /data-series-index="0"/);
assert.match(importedSheet.charts.items[4].toSvg(), /data-point-index="0"/);
const importedScatter = importedSheet.charts.items[5];
assert.deepEqual(importedScatter.categories, []);
assert.deepEqual(importedScatter.series.items[0].xValues, [10, 20, 30]);
assert.deepEqual(importedScatter.series.items[0].values, [30, 55, 88]);
assert.equal(importedScatter.series.items[0].xFormula, "'Summary'!S2:S4");
assert.equal(importedScatter.xAxis.axisType, "valueAxis");
assert.match(importedScatter.toSvg(), /<polygon[^>]+38BDF8/i);
assert.deepEqual(importedSheet.dataValidations.items.map((item) => item.rule.type), ["list", "whole"]);
assert.deepEqual(importedSheet.conditionalFormattings.items.map((item) => item.ruleType), ["cellIs", "expression", "containsText", "colorScale"]);
assert.equal(imported.comments.threads.length, 1);
assert.equal(imported.comments.threads[0].comments.length, 2);
assert.equal(imported.comments.threads[0].comments[0].text, "Check the calculated margin.");
assert.equal(imported.comments.threads[0].comments[1].text, "Confirmed against the source workbook.");
assert.equal(imported.comments.threads[0].resolved, true);

importedSheet.getRange("B2").values = [[110]];
importedSheet.getRange("E2").values = [[new Date("2026-01-20T00:00:00.000Z")]];
importedSheet.getRange("F2").formulas = [["=(B2-C2)/B2"]];
importedSheet.getRange("F2").format.fill = "#BBF7D0";
importedSheet.getRange("A1:A6").format.columnWidthPx = 104;
importedSheet.getRange("A1:F1").format.rowHeightPx = 30;
importedSheet.freezePanes.unfreeze();
importedSheet.freezePanes.freezeRows(2);
importedSheet.freezePanes.freezeColumns(1);
importedSheet.tables.items[0].style = "TableStyleMedium9";
importedSheet.images.items[0].alt = "Edited green status marker";
importedSheet.charts.items[1].title = "Edited revenue trend";
importedScatter.title = "Edited price relationship";
importedScatter.series.items[0].xValues[1] = 22;
importedScatter.series.items[0].values[1] = 60;
const listValidation = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
listValidation.rule.values.push("Blocked");
const marginConditional = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs");
marginConditional.formula = "0.45";
marginConditional.format.fill = "#BBF7D0";
const importedThread = imported.comments.threads[0];
importedThread.comments[0].text = "Margin reviewed after edit.";
importedThread.comments[1].text = "Reply reviewed after edit.";
importedThread.reopen();
imported.recalculate();
assert.equal(importedSheet.getRange("F2").values[0][0], 50 / 110);

const secondXlsx = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
const second = await SpreadsheetFile.importXlsx(secondXlsx);
const secondSheet = second.worksheets.getItem("Summary");
assert.equal(secondSheet.getRange("B2").values[0][0], 110);
assert.equal(secondSheet.getRange("F2").values[0][0], 50 / 110);
assert.equal(secondSheet.getRange("F2").format.fill, "#BBF7D0");
assert.ok(Math.abs(secondSheet.getRange("A1:A6").format.columnWidthPx - 104) <= 1);
assert.ok(Math.abs(secondSheet.getRange("A1:F1").format.rowHeightPx - 30) < 0.01);
assert.deepEqual(secondSheet.freezePanes.toJSON(), { rows: 2, columns: 1, frozen: true, topLeftCell: "B3", activePane: "bottomRight" });
assert.equal(secondSheet.tables.items[0].style, "TableStyleMedium9");
assert.equal(secondSheet.images.items[0].alt, "Edited green status marker");
assert.equal(secondSheet.charts.items[1].title, "Edited revenue trend");
assert.equal(secondSheet.charts.items[5].title, "Edited price relationship");
assert.deepEqual(secondSheet.charts.items[5].series.items[0].xValues, [10, 22, 30]);
assert.deepEqual(secondSheet.charts.items[5].series.items[0].values, [30, 60, 88]);
assert.deepEqual(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.values, ["Planned", "Review", "Done", "Blocked"]);
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs").formula, "0.45");
assert.equal(second.comments.threads[0].comments.length, 2);
assert.equal(second.comments.threads[0].comments[0].text, "Margin reviewed after edit.");
assert.equal(second.comments.threads[0].comments[1].text, "Reply reviewed after edit.");
assert.equal(second.comments.threads[0].resolved, false);

const secondInspect = second.inspect({ kind: "workbook,sheet,table,formula,style,drawing,dataValidation,conditionalFormat,thread", maxChars: 32_000 });
assert.match(secondInspect.ndjson, /Edited revenue trend/);
assert.match(secondInspect.ndjson, /Margin reviewed after edit/);
const secondVerification = verifyArtifact(second);
assert.equal(secondVerification.ok, true, secondVerification.ndjson);
const secondPackageInspect = await SpreadsheetFile.inspectXlsx(secondXlsx, { maxChars: 32_000 });
assert.equal(secondPackageInspect.ok, true, secondPackageInspect.ndjson);
assert.equal(secondPackageInspect.records[0].semanticIssues, 0);

function chartBoundaryWorkbook(type) {
  const candidate = Workbook.create();
  const candidateSheet = candidate.worksheets.add("Chart boundary");
  candidateSheet.getRange("A1:B3").values = [["Quarter", "Revenue"], ["Q1", 40], ["Q2", 60]];
  return { candidate, chart: candidateSheet.charts.add(type, candidateSheet.getRange("A1:B3")) };
}

const invalidDoughnutAxes = chartBoundaryWorkbook("doughnut");
invalidDoughnutAxes.chart.xAxis = {};
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidDoughnutAxes.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /doughnut charts cannot carry.*axes/i.test(error.message),
);

const invalidAreaLineOptions = chartBoundaryWorkbook("area");
invalidAreaLineOptions.chart.lineOptions = { grouping: "standard" };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidAreaLineOptions.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /lineOptions require a line chart/i.test(error.message),
);

const invalidScatter = chartBoundaryWorkbook("scatter");
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(invalidScatter.candidate),
  (error) => error?.code === "invalid_spreadsheet_chart" && /xValue.*finite/i.test(error.message),
);

const scatterWithSeriesLine = Workbook.create();
const scatterWithSeriesLineSheet = scatterWithSeriesLine.worksheets.add("Scatter line boundary");
scatterWithSeriesLineSheet.getRange("A1:B3").values = [["X", "Y"], [1, 2], [2, 4]];
const scatterWithSeriesLineChart = scatterWithSeriesLineSheet.charts.add("scatter", scatterWithSeriesLineSheet.getRange("A1:B3"));
scatterWithSeriesLineChart.series.items[0].line = { fill: "#2563EB", width: 2 };
assert.equal(scatterWithSeriesLine.verify().ok, false);
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(scatterWithSeriesLine),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /marker-only scatter.*marker\.line/i.test(error.message),
);

const unsupportedBubble = chartBoundaryWorkbook("bubble");
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(unsupportedBubble.candidate),
  (error) => error?.code === "unsupported_spreadsheet_chart" && /bar, line, pie, area, doughnut, or scatter/i.test(error.message),
);

const connectionWorkbook = Workbook.create({
  connections: [{ connectionId: 1, name: "Source-free connection", type: 1, refreshedVersion: 1 }],
});
connectionWorkbook.worksheets.add("Main").getRange("A1").values = [["No connection authoring"]];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(connectionWorkbook),
  (error) => error?.code === "unsupported_workbook_features" && /source-free workbook connections/i.test(error.message),
);

const queryWorkbook = Workbook.create();
const querySheet = queryWorkbook.worksheets.add("Main");
querySheet.getRange("A1:B2").values = [["Key", "Value"], ["A", 1]];
const queryTable = querySheet.tables.add("A1:B2", true, "QueryTable");
queryTable.queryTable = { name: "Source-free query", connectionId: 1 };
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(queryWorkbook),
  (error) => error?.code === "unsupported_query_table_edit",
);

const dynamicWorkbook = Workbook.create();
const dynamicSheet = dynamicWorkbook.worksheets.add("Main");
const dynamicCell = dynamicSheet.store.get("A1");
dynamicCell.formula = "=SEQUENCE(2)";
dynamicCell.formulaType = "dynamicArray";
dynamicCell.dynamicArrayRef = "A1:A2";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(dynamicWorkbook),
  (error) => error?.code === "unsupported_workbook_features" && /source-free dynamic array/i.test(error.message),
);

const importedWithoutSourceSnapshot = await SpreadsheetFile.importXlsx(firstXlsx);
const workbookState = importedWithoutSourceSnapshot[Symbol.for("open-office-artifact-tool.open-chestnut-state")];
workbookState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

console.log("spreadsheet tests passed");
