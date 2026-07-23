import assert from "node:assert/strict";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile, Workbook, verifyArtifact } from "../src/index.mjs";
import { formatSpreadsheetDisplayValue } from "../src/spreadsheet/ooxml-styles.mjs";

const PNG_DATA_URL = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC";

assert.equal(formatSpreadsheetDisplayValue(-8884.878867834168, { numberFormat: "$#,##0;[Red]($#,##0);-" }), "($8,885)");

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
sheet.protection = {
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
};
const copiedProtection = sheet.protection;
copiedProtection.allow.push("formatCells");
assert.deepEqual(sheet.protection, {
  enabled: true,
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
});
assert.throws(() => { sheet.protection = { allow: ["unknownOperation"] }; }, /unsupported worksheet protection operation/i);
assert.throws(() => { sheet.protection = { password: "secret" }; }, /unsupported field.*password.*intentionally not accepted/i);
assert.throws(() => { sheet.protection = { enabled: false, allow: ["sort"] }; }, /disabled worksheet protection cannot declare allowed operations/i);

const table = sheet.tables.add("A1:F4", true, "SummaryTable");
table.style = "TableStyleMedium4";
table.showRowStripes = true;

sheet.getRange("D2:D4").dataValidation = {
  rule: {
    type: "list",
    values: ["Planned", "Review", "Done"],
    allowBlank: false,
    showInputMessage: true,
    promptTitle: "Choose a status",
    prompt: "Use one approved workflow state.",
    showErrorMessage: true,
    errorTitle: "Invalid status",
    error: "Choose a value from the list.",
    errorStyle: "warning",
    showDropdown: true,
  },
};
sheet.dataValidations.add({
  range: "B2:B4",
  rule: { type: "whole", operator: "between", formula1: "0", formula2: "1000" },
});
sheet.dataValidations.add({
  range: "C2:C4",
  rule: {
    type: "custom",
    formula1: "=C2<=B2",
    allowBlank: true,
    showErrorMessage: true,
    errorTitle: "Cost exceeds revenue",
    error: "Enter a cost no greater than revenue.",
    errorStyle: "stop",
  },
});
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "whole", formula1: "0", showDropdown: true } }),
  /showDropdown is valid only for list rules/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["North, East"] } }),
  /inline values must be non-empty and cannot contain commas or control characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: [""] } }),
  /inline values must be non-empty/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready\u007F"] } }),
  /control characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "custom", formula1: "=A2<>\"\"", errorStyle: "retry" } }),
  /errorStyle must be one of stop, warning, information/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready"], promptTitle: "x".repeat(33) } }),
  /promptTitle must be at most 32 characters/i,
);
assert.throws(
  () => sheet.dataValidations.add({ range: "A2:A4", rule: { type: "list", values: ["Ready"], imeMode: "fullAlpha" } }),
  /unsupported field: imeMode/i,
);
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
sheet.getRange("B2:B4").conditionalFormats.add("dataBar", {
  color: "#2563EB",
  thresholds: ["min", "max"],
  showValue: false,
});
sheet.getRange("C2:C4").conditionalFormats.add("iconSet", {
  iconSet: "3Arrows",
  thresholds: [0, "50%", { type: "percent", value: 80 }],
  reverse: true,
});
assert.throws(
  () => sheet.getRange("B2:B4").conditionalFormats.add("dataBar", { gradient: false }),
  /gradient=false requires the x14 solid-data-bar extension/i,
);
assert.throws(
  () => sheet.getRange("C2:C4").conditionalFormats.add("iconSet", { iconSet: "3Triangles" }),
  /requires the x14 extension namespace/i,
);
assert.throws(
  () => sheet.getRange("C2:C4").conditionalFormats.add("iconSet", { iconSet: "3Arrows", showValue: "false" }),
  /showValue must be a boolean/i,
);

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
doughnut.dataLabels = { showCategoryName: true, showPercent: true, position: "outsideEnd" };
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
  kind: "workbook,sheet,worksheetProtection,table,formula,style,computedStyle,drawing,dataValidation,conditionalFormat,thread",
  sheetName: "Summary",
  range: "A1:Q34",
  maxChars: 32_000,
});
assert.match(modelInspect.ndjson, /"name":"SummaryTable"/);
assert.match(modelInspect.ndjson, /"formula":"=\(B2-C2\)\/B2"/);
assert.match(modelInspect.ndjson, /"drawingType":"chart"/);
assert.match(modelInspect.ndjson, /"kind":"dataValidation"/);
assert.match(modelInspect.ndjson, /"kind":"conditionalFormat"/);
assert.match(modelInspect.ndjson, /"kind":"worksheetProtection"/);
assert.match(modelInspect.ndjson, /"kind":"dataBar"/);
assert.match(modelInspect.ndjson, /"kind":"iconSet"/);
assert.match(modelInspect.ndjson, /Check the calculated margin/);
const modelLayout = sheet.layoutJson({ range: "B2:C4" });
assert.equal(modelLayout.cells.find((cell) => cell.address === "B2").conditionalFormats.find((rule) => rule.ruleType === "dataBar").visual.showValue, false);
assert.equal(modelLayout.cells.find((cell) => cell.address === "C4").conditionalFormats.find((rule) => rule.ruleType === "iconSet").visual.reverse, true);
const modelSvg = sheet.toSvg();
assert.match(modelSvg, /id="data-bar-1-1-0"/);
assert.match(modelSvg, /[▼➜▲]/);
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
assert.match(firstChartXml.find((xml) => /<c:doughnutChart>/.test(xml)), /<c:showPercent val="1"\s*\/>/);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/media\//i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(Object.keys(firstZip.files).filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name)).length, 1);
assert.equal(firstZip.file("customXml/open-office-artifact.json"), null);
const firstWorksheetXml = await firstZip.file("xl/worksheets/sheet1.xml").async("text");
const firstProtectionXml = firstWorksheetXml.match(/<x:sheetProtection\b[^>]*\/>/)?.[0] || "";
assert.match(firstProtectionXml, /sheet="1"/);
assert.match(firstProtectionXml, /selectLockedCells="0"/);
assert.match(firstProtectionXml, /selectUnlockedCells="0"/);
assert.match(firstProtectionXml, /sort="0"/);
assert.match(firstProtectionXml, /autoFilter="0"/);
assert.match(firstProtectionXml, /formatCells="1"/);
assert.doesNotMatch(firstProtectionXml, /password=|algorithmName=|hashValue=|saltValue=|spinCount=/);
assert.match(firstWorksheetXml, /<x:mergeCell ref="A6:F6"/);
assert.match(firstWorksheetXml, /<x:dataValidations count="3">/);
assert.match(firstWorksheetXml, /type="list" errorStyle="warning" allowBlank="0" showDropDown="0" showInputMessage="1" showErrorMessage="1" errorTitle="Invalid status" error="Choose a value from the list\." promptTitle="Choose a status" prompt="Use one approved workflow state\." sqref="D2:D4"/);
assert.match(firstWorksheetXml, /type="custom" errorStyle="stop" allowBlank="1" showErrorMessage="1" errorTitle="Cost exceeds revenue" error="Enter a cost no greater than revenue\." sqref="C2:C4"/);
assert.equal((firstWorksheetXml.match(/<x:conditionalFormatting\b/g) || []).length, 6);
assert.match(firstWorksheetXml, /<x:dataBar showValue="0">[\s\S]*?<x:cfvo type="min"\s*\/>[\s\S]*?<x:cfvo type="max"\s*\/>[\s\S]*?<x:color rgb="FF2563EB"\s*\/>/);
assert.match(firstWorksheetXml, /<x:iconSet iconSet="3Arrows" reverse="1">[\s\S]*?<x:cfvo type="num" val="0"\s*\/>[\s\S]*?<x:cfvo type="percent" val="50"\s*\/>[\s\S]*?<x:cfvo type="percent" val="80"\s*\/>/);
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
assert.deepEqual(importedSheet.protection, {
  enabled: true,
  allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
});
assert.equal(importedSheet.tables.items[0].name, "SummaryTable");
assert.equal(importedSheet.images.items[0].alt, "Green status marker");
assert.deepEqual(importedSheet.charts.items.map((chart) => chart.type), ["bar", "line", "pie", "area", "doughnut", "scatter"]);
assert.match(importedSheet.charts.items[3].toSvg(), /data-series-index="0"/);
assert.match(importedSheet.charts.items[4].toSvg(), /data-point-index="0"/);
assert.equal(importedSheet.charts.items[4].dataLabels.showPercent, true);
assert.match(importedSheet.charts.items[4].toSvg(), /data-chart-label-index="0"[^>]*>[^<]*%/);
const importedScatter = importedSheet.charts.items[5];
assert.deepEqual(importedScatter.categories, []);
assert.deepEqual(importedScatter.series.items[0].xValues, [10, 20, 30]);
assert.deepEqual(importedScatter.series.items[0].values, [30, 55, 88]);
assert.equal(importedScatter.series.items[0].xFormula, "'Summary'!S2:S4");
assert.equal(importedScatter.xAxis.axisType, "valueAxis");
assert.match(importedScatter.toSvg(), /<polygon[^>]+38BDF8/i);
assert.deepEqual(importedSheet.dataValidations.items.map((item) => item.rule.type), ["list", "whole", "custom"]);
const importedListValidation = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
assert.deepEqual(importedListValidation.rule, {
  type: "list",
  values: ["Planned", "Review", "Done"],
  allowBlank: false,
  showInputMessage: true,
  promptTitle: "Choose a status",
  prompt: "Use one approved workflow state.",
  showErrorMessage: true,
  errorTitle: "Invalid status",
  error: "Choose a value from the list.",
  errorStyle: "warning",
  showDropdown: true,
});
assert.deepEqual(importedSheet.dataValidations.items.find((item) => item.rule.type === "custom").rule, {
  type: "custom",
  formula1: "=C2<=B2",
  allowBlank: true,
  showErrorMessage: true,
  errorTitle: "Cost exceeds revenue",
  error: "Enter a cost no greater than revenue.",
  errorStyle: "stop",
});
assert.deepEqual(importedSheet.conditionalFormattings.items.map((item) => item.ruleType), ["cellIs", "expression", "containsText", "colorScale", "dataBar", "iconSet"]);
const importedDataBar = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar");
assert.equal(importedDataBar.color, "#2563EB");
assert.equal(importedDataBar.showValue, false);
assert.deepEqual(importedDataBar.thresholds, [{ type: "min" }, { type: "max" }]);
const importedIconSet = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet");
assert.equal(importedIconSet.iconSet, "3Arrows");
assert.equal(importedIconSet.reverse, true);
assert.deepEqual(importedIconSet.thresholds, [{ type: "num", value: 0 }, { type: "percent", value: 50 }, { type: "percent", value: 80 }]);
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
importedSheet.protection = { allow: ["selectUnlockedCells", "formatCells"] };
importedSheet.tables.items[0].style = "TableStyleMedium9";
importedSheet.images.items[0].alt = "Edited green status marker";
importedSheet.charts.items[1].title = "Edited revenue trend";
importedScatter.title = "Edited price relationship";
importedScatter.series.items[0].xValues[1] = 22;
importedScatter.series.items[0].values[1] = 60;
const listValidation = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
listValidation.rule.values.push("Blocked");
listValidation.rule.prompt = "Pick the current workflow state.";
listValidation.rule.errorStyle = "information";
listValidation.rule.showDropdown = false;
const marginConditional = importedSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs");
marginConditional.formula = "0.45";
marginConditional.format.fill = "#BBF7D0";
importedDataBar.color = "#0EA5E9";
importedDataBar.showValue = true;
importedIconSet.thresholds[1].value = 60;
importedIconSet.reverse = false;
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
assert.deepEqual(secondSheet.protection, { enabled: true, allow: ["selectUnlockedCells", "formatCells"] });
assert.equal(secondSheet.tables.items[0].style, "TableStyleMedium9");
assert.equal(secondSheet.images.items[0].alt, "Edited green status marker");
assert.equal(secondSheet.charts.items[1].title, "Edited revenue trend");
assert.equal(secondSheet.charts.items[4].dataLabels.showPercent, true);
assert.equal(secondSheet.charts.items[5].title, "Edited price relationship");
assert.deepEqual(secondSheet.charts.items[5].series.items[0].xValues, [10, 22, 30]);
assert.deepEqual(secondSheet.charts.items[5].series.items[0].values, [30, 60, 88]);
assert.deepEqual(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.values, ["Planned", "Review", "Done", "Blocked"]);
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.prompt, "Pick the current workflow state.");
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.errorStyle, "information");
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "list").rule.showDropdown, false);
assert.equal(secondSheet.dataValidations.items.find((item) => item.rule.type === "custom").rule.formula1, "=C2<=B2");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "cellIs").formula, "0.45");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar").color, "#0EA5E9");
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "dataBar").showValue, true);
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet").thresholds[1].value, 60);
assert.equal(secondSheet.conditionalFormattings.items.find((item) => item.ruleType === "iconSet").reverse, false);
assert.equal(second.comments.threads[0].comments.length, 2);
assert.equal(second.comments.threads[0].comments[0].text, "Margin reviewed after edit.");
assert.equal(second.comments.threads[0].comments[1].text, "Reply reviewed after edit.");
assert.equal(second.comments.threads[0].resolved, false);

secondSheet.protection = null;
const protectionRemoved = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(second, { recalculate: false }));
assert.equal(protectionRemoved.worksheets.getItem("Summary").protection, undefined);

const secondInspect = second.inspect({ kind: "workbook,sheet,table,formula,style,drawing,dataValidation,conditionalFormat,thread", maxChars: 32_000 });
assert.match(secondInspect.ndjson, /Edited revenue trend/);
assert.match(secondInspect.ndjson, /Margin reviewed after edit/);
const secondVerification = verifyArtifact(second);
assert.equal(secondVerification.ok, true, secondVerification.ndjson);
const secondPackageInspect = await SpreadsheetFile.inspectXlsx(secondXlsx, { maxChars: 32_000 });
assert.equal(secondPackageInspect.ok, true, secondPackageInspect.ndjson);
assert.equal(secondPackageInspect.records[0].semanticIssues, 0);

const bubbleWorkbook = Workbook.create();
const bubbleSheet = bubbleWorkbook.worksheets.add("Opportunities");
bubbleSheet.getRange("A1:C4").values = [
  ["Customers", "Revenue", "Pipeline"],
  [10, 42, 4],
  [20, 68, 9],
  [30, 85, 16],
];
const bubble = bubbleSheet.charts.add("bubble", bubbleSheet.getRange("A1:C4"));
bubble.name = "Opportunity bubble";
bubble.title = "Revenue opportunity";
bubble.xAxis = { title: { text: "Customers" }, min: 0, max: 40, majorUnit: 10, numberFormatCode: "0" };
bubble.yAxis = { title: { text: "Revenue" }, min: 0, max: 100, majorUnit: 20, numberFormatCode: "$0" };
bubble.series.items[0].fill = "#0EA5E9";
bubble.series.items[0].line = { fill: "#0369A1", width: 1.5 };
bubble.setPosition("E2", "L18");
assert.deepEqual(bubble.categories, []);
assert.deepEqual(bubble.series.items[0].xValues, [10, 20, 30]);
assert.deepEqual(bubble.series.items[0].values, [42, 68, 85]);
assert.deepEqual(bubble.series.items[0].bubbleSizes, [4, 9, 16]);
assert.equal(bubble.series.items[0].xFormula, "'Opportunities'!A2:A4");
assert.equal(bubble.series.items[0].formula, "'Opportunities'!B2:B4");
assert.equal(bubble.series.items[0].bubbleSizeFormula, "'Opportunities'!C2:C4");
assert.match(bubble.toSvg(), /data-bubble-size="16"/);
assert.equal(bubbleWorkbook.verify().ok, true);
const bubbleXlsx = await SpreadsheetFile.exportXlsx(bubbleWorkbook);
const bubbleZip = await JSZip.loadAsync(new Uint8Array(await bubbleXlsx.arrayBuffer()));
const bubbleChartPath = Object.keys(bubbleZip.files).find((name) => /\/charts\/chart\d+\.xml$/i.test(name));
assert.ok(bubbleChartPath);
const bubbleXml = await bubbleZip.file(bubbleChartPath).async("text");
assert.match(bubbleXml, /<c:bubbleChart>/);
assert.match(bubbleXml, /<c:xVal>[\s\S]*<c:yVal>[\s\S]*<c:bubbleSize>/);
assert.equal((bubbleXml.match(/<c:valAx>/g) || []).length, 2);
assert.doesNotMatch(bubbleXml, /<c:cat>/);
const importedBubbleWorkbook = await SpreadsheetFile.importXlsx(bubbleXlsx);
const importedBubble = importedBubbleWorkbook.worksheets.getItem("Opportunities").charts.items[0];
assert.equal(importedBubble.type, "bubble");
assert.equal(importedBubble.xAxis.axisType, "valueAxis");
assert.deepEqual(importedBubble.series.items[0].bubbleSizes, [4, 9, 16]);
importedBubble.title = "Edited revenue opportunity";
importedBubble.series.items[0].xValues[1] = 22;
importedBubble.series.items[0].values[1] = 70;
importedBubble.series.items[0].bubbleSizes[1] = 12;
const editedBubbleWorkbook = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(importedBubbleWorkbook));
const editedBubble = editedBubbleWorkbook.worksheets.getItem("Opportunities").charts.items[0];
assert.equal(editedBubble.title, "Edited revenue opportunity");
assert.deepEqual(editedBubble.series.items[0].xValues, [10, 22, 30]);
assert.deepEqual(editedBubble.series.items[0].values, [42, 70, 85]);
assert.deepEqual(editedBubble.series.items[0].bubbleSizes, [4, 12, 16]);

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

const bubbleShortcutBoundary = Workbook.create();
const bubbleShortcutBoundarySheet = bubbleShortcutBoundary.worksheets.add("Bubble boundary");
bubbleShortcutBoundarySheet.getRange("A1:D3").values = [["X", "Y", "Size", "Unexpected"], [1, 2, 3, 4], [2, 4, 6, 8]];
assert.throws(
  () => bubbleShortcutBoundarySheet.charts.add("bubble", bubbleShortcutBoundarySheet.getRange("A1:D3")),
  /requires exactly three columns ordered X \| Y \| Size/i,
);
const nonPositiveBubble = Workbook.create();
const nonPositiveBubbleSheet = nonPositiveBubble.worksheets.add("Non-positive bubble");
nonPositiveBubbleSheet.getRange("A1:C3").values = [["X", "Y", "Size"], [1, 2, 0], [2, 4, 6]];
assert.throws(
  () => nonPositiveBubbleSheet.charts.add("bubble", nonPositiveBubbleSheet.getRange("A1:C3")),
  /Size value.*finite and positive/i,
);
const mismatchedBubble = Workbook.create();
const mismatchedBubbleSheet = mismatchedBubble.worksheets.add("Mismatched bubble");
const mismatchedBubbleChart = mismatchedBubbleSheet.charts.add("bubble", {
  name: "Mismatched bubble",
  series: [{ name: "Series", xValues: [1, 2], values: [2, 4], bubbleSizes: [5] }],
});
assert.equal(mismatchedBubbleChart.type, "bubble");
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(mismatchedBubble),
  (error) => error?.code === "invalid_spreadsheet_chart" && /bubbleSizes.*y values/i.test(error.message),
);

const pivotWorkbook = Workbook.create();
const pivotData = pivotWorkbook.worksheets.add("Data");
pivotData.getRange("A1:C5").values = [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["East", "B", 20],
  ["West", "A", 30],
  ["West", "B", 40],
];
const pivotSummary = pivotWorkbook.worksheets.add("Summary");
pivotSummary.getRange("A1:D4").format = { fill: "#ECFEFF" };
const nativePivot = pivotSummary.pivotTables.add({
  name: "Sales by region",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", summarizeBy: "sum" }],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(nativePivot.computedValues(), [
  ["Region", "A", "B", "Grand Total"],
  ["East", 10, 20, 30],
  ["West", 30, 40, 70],
  ["Grand Total", 40, 60, 100],
]);
const nativePivotXlsx = await SpreadsheetFile.exportXlsx(pivotWorkbook);
const nativePivotZip = await JSZip.loadAsync(new Uint8Array(await nativePivotXlsx.arrayBuffer()));
const nativePivotPart = Object.keys(nativePivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const nativePivotCache = Object.keys(nativePivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
const nativePivotRecords = Object.keys(nativePivotZip.files).find((name) => /pivotCache\/pivotCacheRecords.*\.xml$/i.test(name));
assert.ok(nativePivotPart);
assert.ok(nativePivotCache);
assert.ok(nativePivotRecords);
assert.match(await nativePivotZip.file(nativePivotPart).async("text"), /name="Sales by region"[\s\S]*location ref="A1:D4"[\s\S]*subtotal="sum"/);
assert.match(await nativePivotZip.file(nativePivotCache).async("text"), /worksheetSource ref="A1:C5" sheet="Data"/);
assert.match(await nativePivotZip.file(nativePivotRecords).async("text"), /count="4"/);

const importedPivotWorkbook = await SpreadsheetFile.importXlsx(nativePivotXlsx);
const importedPivot = importedPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.equal(importedPivot.name, "Sales by region");
assert.deepEqual(importedPivot.rowFields, ["Region"]);
assert.deepEqual(importedPivot.columnFields, ["Product"]);
assert.deepEqual(importedPivot.valueFields, [{ field: "Sales", summarizeBy: "sum", name: "Sum of Sales" }]);
assert.deepEqual(importedPivot.computedValues(), nativePivot.computedValues());
assert.equal(importedPivotWorkbook.worksheets.getItem("Summary").getRange("A1").format.fill, "#ECFEFF");
const secondPivotXlsx = await SpreadsheetFile.exportXlsx(importedPivotWorkbook);
const secondPivotZip = await JSZip.loadAsync(new Uint8Array(await secondPivotXlsx.arrayBuffer()));
assert.equal(await secondPivotZip.file(nativePivotPart).async("text"), await nativePivotZip.file(nativePivotPart).async("text"));
assert.equal(await secondPivotZip.file(nativePivotCache).async("text"), await nativePivotZip.file(nativePivotCache).async("text"));
assert.equal(await secondPivotZip.file(nativePivotRecords).async("text"), await nativePivotZip.file(nativePivotRecords).async("text"));

const filteredPivotWorkbook = Workbook.create();
const filteredPivotData = filteredPivotWorkbook.worksheets.add("Data");
filteredPivotData.getRange("A1:C5").values = [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["East", "B", 20],
  ["West", "A", 30],
  ["West", "B", 40],
];
const filteredPivotSummary = filteredPivotWorkbook.worksheets.add("Summary");
filteredPivotSummary.getRange("A1:C3").format = { fill: "#FFF7ED" };
const filteredPivot = filteredPivotSummary.pivotTables.add({
  name: "Filtered sales",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", summarizeBy: "sum" }],
  filters: [
    { field: "Region", include: ["East"] },
    { field: "Product", exclude: ["B"] },
  ],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(filteredPivot.computedValues(), [
  ["Region", "A", "Grand Total"],
  ["East", 10, 10],
  ["Grand Total", 10, 10],
]);
const filteredPivotXlsx = await SpreadsheetFile.exportXlsx(filteredPivotWorkbook);
const filteredPivotZip = await JSZip.loadAsync(new Uint8Array(await filteredPivotXlsx.arrayBuffer()));
const filteredPivotPart = Object.keys(filteredPivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const filteredPivotXml = await filteredPivotZip.file(filteredPivotPart).async("text");
assert.match(filteredPivotXml, /location ref="A1:C3"/);
assert.match(filteredPivotXml, /includeNewItemsInFilter="0"[\s\S]*<item x="1" h="1"/);
assert.match(filteredPivotXml, /includeNewItemsInFilter="1"[\s\S]*<item x="1" h="1"/);
const importedFilteredPivotWorkbook = await SpreadsheetFile.importXlsx(filteredPivotXlsx);
const importedFilteredPivot = importedFilteredPivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedFilteredPivot.filters, [
  { field: "Region", include: ["East"] },
  { field: "Product", exclude: ["B"] },
]);
assert.deepEqual(importedFilteredPivot.computedValues(), filteredPivot.computedValues());
assert.equal(importedFilteredPivotWorkbook.worksheets.getItem("Summary").getRange("C3").format.fill, "#FFF7ED");
const secondFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(importedFilteredPivotWorkbook);
const secondFilteredPivotZip = await JSZip.loadAsync(new Uint8Array(await secondFilteredPivotXlsx.arrayBuffer()));
assert.equal(await secondFilteredPivotZip.file(filteredPivotPart).async("text"), filteredPivotXml);
importedFilteredPivot.filters[0].include[0] = "West";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /read-only/i.test(error.message),
);

const typedFilteredPivotWorkbook = Workbook.create();
const typedFilteredPivotData = typedFilteredPivotWorkbook.worksheets.add("Data");
typedFilteredPivotData.getRange("A1:B5").values = [["Key", "Sales"], [1, 10], [true, 20], [null, 30], ["Other", 40]];
const typedFilteredPivotSummary = typedFilteredPivotWorkbook.worksheets.add("Summary");
const typedFilteredPivot = typedFilteredPivotSummary.pivotTables.add({
  name: "Typed filter items",
  sourceRange: "Data!A1:B5",
  targetRange: "A1",
  rowFields: ["Key"],
  valueFields: [{ field: "Sales", name: "Sales" }],
  filters: [{ field: "Key", include: [true, null] }],
  columnGrandTotals: true,
});
assert.deepEqual(typedFilteredPivot.computedValues(), [
  ["Key", "Sales"],
  [true, 20],
  [null, 30],
  ["Grand Total", 50],
]);
const typedFilteredPivotXlsx = await SpreadsheetFile.exportXlsx(typedFilteredPivotWorkbook);
const importedTypedFilteredPivot = (await SpreadsheetFile.importXlsx(typedFilteredPivotXlsx)).worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedTypedFilteredPivot.filters, [{ field: "Key", include: [true, null] }]);
assert.deepEqual(importedTypedFilteredPivot.computedValues(), typedFilteredPivot.computedValues());

const emptyFilteredPivotWorkbook = Workbook.create();
const emptyFilteredPivotSheet = emptyFilteredPivotWorkbook.worksheets.add("Data");
emptyFilteredPivotSheet.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
emptyFilteredPivotSheet.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Region"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Region", exclude: ["East", "West"] }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(emptyFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /hide every source row/i.test(error.message),
);

const dateFilteredPivotWorkbook = Workbook.create();
const dateFilteredPivotSheet = dateFilteredPivotWorkbook.worksheets.add("Data");
dateFilteredPivotSheet.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
dateFilteredPivotSheet.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Region"],
  valueFields: [{ field: "Sales" }],
  filters: [{ field: "Region", type: "today", asOf: "2026-07-19" }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(dateFilteredPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_filter" && /exact include\/exclude/i.test(error.message),
);

const multiValuePivotWorkbook = Workbook.create();
const multiValuePivotData = multiValuePivotWorkbook.worksheets.add("Data");
multiValuePivotData.getRange("A1:D5").values = [
  ["Region", "Product", "Sales", "Units"],
  ["East", "A", 10, 2],
  ["East", "B", 20, 4],
  ["West", "A", 30, 6],
  ["West", "B", 40, 8],
];
const multiValuePivotSummary = multiValuePivotWorkbook.worksheets.add("Summary");
multiValuePivotSummary.getRange("A1:G4").format = { fill: "#F0FDFA" };
const multiValuePivot = multiValuePivotSummary.pivotTables.add({
  name: "Revenue and units by region",
  sourceRange: "Data!A1:D5",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [
    { field: "Sales", summarizeBy: "sum", name: "Revenue" },
    { field: "Units", summarizeBy: "average", name: "Average units" },
  ],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
assert.deepEqual(multiValuePivot.computedValues(), [
  ["Region", "A — Revenue", "A — Average units", "B — Revenue", "B — Average units", "Grand Total — Revenue", "Grand Total — Average units"],
  ["East", 10, 2, 20, 4, 30, 3],
  ["West", 30, 6, 40, 8, 70, 7],
  ["Grand Total", 40, 4, 60, 6, 100, 5],
]);
const multiValuePivotXlsx = await SpreadsheetFile.exportXlsx(multiValuePivotWorkbook);
const multiValuePivotZip = await JSZip.loadAsync(new Uint8Array(await multiValuePivotXlsx.arrayBuffer()));
const multiValuePivotPart = Object.keys(multiValuePivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const multiValuePivotCache = Object.keys(multiValuePivotZip.files).find((name) => /pivotCache\/pivotCacheDefinition.*\.xml$/i.test(name));
const multiValuePivotRecords = Object.keys(multiValuePivotZip.files).find((name) => /pivotCache\/pivotCacheRecords.*\.xml$/i.test(name));
const multiValuePivotXml = await multiValuePivotZip.file(multiValuePivotPart).async("text");
assert.match(multiValuePivotXml, /location ref="A1:G4"/);
assert.match(multiValuePivotXml, /colFields count="2">[\s\S]*field x="1"[\s\S]*field x="-2"/);
assert.match(multiValuePivotXml, /dataFields count="2">[\s\S]*name="Revenue"[^>]*fld="2"[^>]*subtotal="sum"[\s\S]*name="Average units"[^>]*fld="3"[^>]*subtotal="average"/);
assert.match(multiValuePivotXml, /colItems count="6">[\s\S]*<i i="1">/);

const importedMultiValuePivotWorkbook = await SpreadsheetFile.importXlsx(multiValuePivotXlsx);
const importedMultiValuePivot = importedMultiValuePivotWorkbook.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(importedMultiValuePivot.valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "average", name: "Average units" },
]);
assert.deepEqual(importedMultiValuePivot.computedValues(), multiValuePivot.computedValues());
assert.equal(importedMultiValuePivotWorkbook.worksheets.getItem("Summary").getRange("G4").format.fill, "#F0FDFA");
const secondMultiValuePivotXlsx = await SpreadsheetFile.exportXlsx(importedMultiValuePivotWorkbook);
const secondMultiValuePivotZip = await JSZip.loadAsync(new Uint8Array(await secondMultiValuePivotXlsx.arrayBuffer()));
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotPart).async("text"), multiValuePivotXml);
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotCache).async("text"), await multiValuePivotZip.file(multiValuePivotCache).async("text"));
assert.equal(await secondMultiValuePivotZip.file(multiValuePivotRecords).async("text"), await multiValuePivotZip.file(multiValuePivotRecords).async("text"));

const hostNormalizedMultiValueZip = await JSZip.loadAsync(new Uint8Array(await multiValuePivotXlsx.arrayBuffer()));
const hostNormalizedMultiValueXml = multiValuePivotXml
  .replace(/<rowItems\b[\s\S]*?<\/rowItems>/, "")
  .replace(/<colItems\b[\s\S]*?<\/colItems>/, "");
assert.notEqual(hostNormalizedMultiValueXml, multiValuePivotXml);
hostNormalizedMultiValueZip.file(multiValuePivotPart, hostNormalizedMultiValueXml);
const hostNormalizedMultiValueBytes = await hostNormalizedMultiValueZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const hostNormalizedMultiValueFile = new FileBlob(hostNormalizedMultiValueBytes, {
  type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  name: "host-normalized-multi-value-pivot.xlsx",
});
const importedHostNormalizedMultiValue = await SpreadsheetFile.importXlsx(hostNormalizedMultiValueFile);
const hostNormalizedPivot = importedHostNormalizedMultiValue.worksheets.getItem("Summary").pivotTables.items[0];
assert.deepEqual(hostNormalizedPivot.valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "average", name: "Average units" },
]);
const preservedHostNormalizedMultiValue = await SpreadsheetFile.exportXlsx(importedHostNormalizedMultiValue);
const preservedHostNormalizedMultiValueZip = await JSZip.loadAsync(new Uint8Array(await preservedHostNormalizedMultiValue.arrayBuffer()));
assert.equal(await preservedHostNormalizedMultiValueZip.file(multiValuePivotPart).async("text"), hostNormalizedMultiValueXml);

const noColumnMultiValueWorkbook = Workbook.create();
const noColumnMultiValueData = noColumnMultiValueWorkbook.worksheets.add("Data");
noColumnMultiValueData.getRange("A1:C5").values = [
  ["Region", "Sales", "Units"],
  ["East", 10, 2],
  ["East", 20, 4],
  ["West", 30, 6],
  ["West", 40, 8],
];
const noColumnMultiValueSummary = noColumnMultiValueWorkbook.worksheets.add("Summary");
const noColumnMultiValuePivot = noColumnMultiValueSummary.pivotTables.add({
  name: "Regional metrics",
  sourceRange: "Data!A1:C5",
  targetRange: "A1",
  rowFields: ["Region"],
  valueFields: [
    { field: "Sales", summarizeBy: "sum", name: "Revenue" },
    { field: "Units", summarizeBy: "count", name: "Unit records" },
  ],
  columnGrandTotals: true,
});
assert.deepEqual(noColumnMultiValuePivot.computedValues(), [
  ["Region", "Revenue", "Unit records"],
  ["East", 30, 2],
  ["West", 70, 2],
  ["Grand Total", 100, 4],
]);
const noColumnMultiValueXlsx = await SpreadsheetFile.exportXlsx(noColumnMultiValueWorkbook);
const noColumnMultiValueZip = await JSZip.loadAsync(new Uint8Array(await noColumnMultiValueXlsx.arrayBuffer()));
const noColumnMultiValuePart = Object.keys(noColumnMultiValueZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const noColumnMultiValueXml = await noColumnMultiValueZip.file(noColumnMultiValuePart).async("text");
assert.match(noColumnMultiValueXml, /location ref="A1:C4"/);
assert.match(noColumnMultiValueXml, /colFields count="1">[\s\S]*field x="-2"/);
assert.match(noColumnMultiValueXml, /colItems count="2">[\s\S]*<i i="1">/);
const importedNoColumnMultiValue = await SpreadsheetFile.importXlsx(noColumnMultiValueXlsx);
assert.deepEqual(importedNoColumnMultiValue.worksheets.getItem("Summary").pivotTables.items[0].valueFields, [
  { field: "Sales", summarizeBy: "sum", name: "Revenue" },
  { field: "Units", summarizeBy: "count", name: "Unit records" },
]);

const editedImportedPivot = await SpreadsheetFile.importXlsx(nativePivotXlsx);
editedImportedPivot.worksheets.getItem("Data").getRange("C2").values = [[11]];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedImportedPivot),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /source data.*read-only/i.test(error.message),
);
const editedPivotOutput = await SpreadsheetFile.importXlsx(nativePivotXlsx);
editedPivotOutput.worksheets.getItem("Summary").getRange("B2").values = [[11]];
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(editedPivotOutput),
  (error) => error?.code === "unsupported_spreadsheet_pivot_edit" && /cached output.*read-only/i.test(error.message),
);
const multiRowPivotWorkbook = Workbook.create();
const multiRowPivotSheet = multiRowPivotWorkbook.worksheets.add("Data");
multiRowPivotSheet.getRange("A1:C3").values = [["Region", "Product", "Sales"], ["East", "A", 10], ["West", "B", 20]];
const multiRowPivot = multiRowPivotSheet.pivotTables.add({
  name: "Sales by region and product",
  sourceRange: "A1:C3",
  targetRange: "E1",
  rowFields: ["Region", "Product"],
  valueFields: [{ field: "Sales", name: "Sales" }],
  columnGrandTotals: true,
});
assert.deepEqual(multiRowPivot.computedValues(), [
  ["Region", "Product", "Sales"],
  ["East", "A", 10],
  ["West", "B", 20],
  ["Grand Total", "", 30],
]);
const multiRowPivotXlsx = await SpreadsheetFile.exportXlsx(multiRowPivotWorkbook);
const multiRowPivotZip = await JSZip.loadAsync(new Uint8Array(await multiRowPivotXlsx.arrayBuffer()));
const multiRowPivotPart = Object.keys(multiRowPivotZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
const multiRowPivotXml = await multiRowPivotZip.file(multiRowPivotPart).async("text");
assert.match(multiRowPivotXml, /location ref="E1:G4"[^>]*firstDataCol="2"/);
assert.match(multiRowPivotXml, /rowFields count="2">[\s\S]*field x="0"[\s\S]*field x="1"/);
assert.match(multiRowPivotXml, /pivotField[^>]*axis="axisRow"[^>]*compact="0"[^>]*defaultSubtotal="0"/);
const importedMultiRowPivotWorkbook = await SpreadsheetFile.importXlsx(multiRowPivotXlsx);
const importedMultiRowPivot = importedMultiRowPivotWorkbook.worksheets.getItem("Data").pivotTables.items[0];
assert.deepEqual(importedMultiRowPivot.rowFields, ["Region", "Product"]);
assert.deepEqual(importedMultiRowPivot.computedValues(), multiRowPivot.computedValues());

const overBudgetRowPivotWorkbook = Workbook.create();
const overBudgetRowPivotSheet = overBudgetRowPivotWorkbook.worksheets.add("Data");
const overBudgetRowHeaders = [...Array.from({ length: 9 }, (_, index) => `Axis ${index + 1}`), "Sales"];
overBudgetRowPivotSheet.getRange("A1:J2").values = [overBudgetRowHeaders, [...Array.from({ length: 9 }, (_, index) => `Value ${index + 1}`), 10]];
overBudgetRowPivotSheet.pivotTables.add({
  sourceRange: "A1:J2",
  targetRange: "L1",
  rowFields: overBudgetRowHeaders.slice(0, 9),
  valueFields: [{ field: "Sales" }],
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(overBudgetRowPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_profile" && /1 through 8 row fields/i.test(error.message),
);
const overBudgetPivotWorkbook = Workbook.create();
const overBudgetPivotSheet = overBudgetPivotWorkbook.worksheets.add("Data");
overBudgetPivotSheet.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
overBudgetPivotSheet.pivotTables.add({
  sourceRange: "A1:B3",
  targetRange: "D1",
  rowFields: ["Region"],
  valueFields: Array.from({ length: 33 }, (_, index) => ({ field: "Sales", summarizeBy: "sum", name: `Revenue ${index + 1}` })),
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(overBudgetPivotWorkbook),
  (error) => error?.code === "unsupported_spreadsheet_pivot_profile" && /1 through 32 value fields/i.test(error.message),
);
const collidingPivotWorkbook = Workbook.create();
const collidingPivotSheet = collidingPivotWorkbook.worksheets.add("Data");
collidingPivotSheet.getRange("A1:C3").values = [["Region", "Product", "Sales"], ["East", "A", 10], ["West", "B", 20]];
collidingPivotSheet.getRange("E1").values = [["occupied"]];
collidingPivotSheet.pivotTables.add({ sourceRange: "A1:C3", targetRange: "E1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(collidingPivotWorkbook),
  (error) => error?.code === "spreadsheet_pivot_output_collision" && /overlaps existing worksheet cell E1/i.test(error.message),
);
const collidingMultiValuePivotWorkbook = Workbook.create();
const collidingMultiValuePivotData = collidingMultiValuePivotWorkbook.worksheets.add("Data");
collidingMultiValuePivotData.getRange("A1:D3").values = [
  ["Region", "Product", "Sales", "Units"],
  ["East", "A", 10, 2],
  ["West", "B", 20, 4],
];
const collidingMultiValuePivotSummary = collidingMultiValuePivotWorkbook.worksheets.add("Summary");
collidingMultiValuePivotSummary.getRange("G4").values = [["occupied widened edge"]];
collidingMultiValuePivotSummary.pivotTables.add({
  sourceRange: "Data!A1:D3",
  targetRange: "A1",
  rowFields: ["Region"],
  columnFields: ["Product"],
  valueFields: [{ field: "Sales", name: "Revenue" }, { field: "Units", name: "Units" }],
  rowGrandTotals: true,
  columnGrandTotals: true,
});
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(collidingMultiValuePivotWorkbook),
  (error) => error?.code === "spreadsheet_pivot_output_collision" && /overlaps existing worksheet cell G4/i.test(error.message),
);
const duplicatePivotWorkbook = Workbook.create();
const duplicatePivotData = duplicatePivotWorkbook.worksheets.add("Data");
duplicatePivotData.getRange("A1:B3").values = [["Region", "Sales"], ["East", 10], ["West", 20]];
const duplicatePivotSummaryA = duplicatePivotWorkbook.worksheets.add("Summary A");
const duplicatePivotSummaryB = duplicatePivotWorkbook.worksheets.add("Summary B");
duplicatePivotSummaryA.pivotTables.add({ name: "Sales Pivot", sourceRange: "Data!A1:B3", targetRange: "A1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
duplicatePivotSummaryB.pivotTables.add({ name: "sales pivot", sourceRange: "Data!A1:B3", targetRange: "A1", rowFields: ["Region"], valueFields: [{ field: "Sales" }] });
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(duplicatePivotWorkbook),
  (error) => error?.code === "invalid_spreadsheet_pivot" && /name Sales Pivot must be unique across the workbook/i.test(error.message),
);

for (const [summarizeBy, expected] of [
  ["sum", 60],
  ["count", 3],
  ["average", 20],
  ["min", 10],
  ["max", 30],
]) {
  const aggregationWorkbook = Workbook.create();
  const aggregationSheet = aggregationWorkbook.worksheets.add("Data");
  aggregationSheet.getRange("A1:C4").values = [
    ["Region", "Product", "Sales"],
    ["East", "A", 10],
    ["East", "A", 20],
    ["East", "A", 30],
  ];
  const aggregationSummary = aggregationWorkbook.worksheets.add("Summary");
  const aggregationPivot = aggregationSummary.pivotTables.add({
    name: `${summarizeBy} Sales`,
    sourceRange: "Data!A1:C4",
    targetRange: "A1",
    rowFields: ["Region"],
    columnFields: ["Product"],
    valueFields: [{ field: "Sales", summarizeBy }],
  });
  assert.equal(aggregationPivot.computedValues()[1][1], expected);
  const aggregationXlsx = await SpreadsheetFile.exportXlsx(aggregationWorkbook);
  const aggregationZip = await JSZip.loadAsync(new Uint8Array(await aggregationXlsx.arrayBuffer()));
  const aggregationPivotPart = Object.keys(aggregationZip.files).find((name) => /xl\/pivotTables\/pivotTable.*\.xml$/i.test(name));
  assert.match(await aggregationZip.file(aggregationPivotPart).async("text"), new RegExp(`subtotal="${summarizeBy}"`));
  const aggregationImported = await SpreadsheetFile.importXlsx(aggregationXlsx);
  const importedAggregationPivot = aggregationImported.worksheets.getItem("Summary").pivotTables.items[0];
  assert.equal(importedAggregationPivot.valueFields[0].summarizeBy, summarizeBy);
  assert.equal(importedAggregationPivot.computedValues()[1][1], expected);
}

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

const ifsFormulaWorkbook = Workbook.create();
const ifsFormulaSheet = ifsFormulaWorkbook.worksheets.add("Criteria");
ifsFormulaSheet.getRange("A1:C6").values = [
  ["Region", "Amount", "Status"],
  ["East", 14, "Yes"],
  ["East", 7, "Yes"],
  ["West", 3, "Yes"],
  ["East", "n/a", "Yes"],
  ["West", 20, "No"],
];
ifsFormulaSheet.getRange("E1:E4").formulas = [
  ["=MINIFS(B2:B6,A2:A6,\"East\",C2:C6,\"Yes\")"],
  ["=MAXIFS(B2:B6,A2:A6,\"East\",C2:C6,\"Yes\")"],
  ["=MINIFS(B2:B6,A2:A5,\"East\")"],
  ["=MAXIFS(B2:B6,A2:A6,\"North\")"],
];
assert.deepEqual(ifsFormulaSheet.getRange("E1:E4").values, [[7], [14], ["#VALUE!"], [0]]);
ifsFormulaSheet.getRange("G1:G7").formulas = [
  ["=IFS(FALSE,\"wrong\",TRUE,\"selected\")"],
  ["=IFS(FALSE,\"no match\")"],
  ["=IFS(TRUE,\"short circuit\",TRUE,1/0)"],
  ["=SWITCH(A2,\"West\",1,\"East\",2,0)"],
  ["=SWITCH(\"North\",\"West\",1,\"East\",2,0)"],
  ["=SWITCH(\"North\",\"West\",1)"],
  ["=IFS(TRUE)"],
];
assert.deepEqual(ifsFormulaSheet.getRange("G1:G7").values, [["selected"], ["#N/A"], ["short circuit"], [2], [0], ["#N/A"], ["#VALUE!"]]);
const ifsFormulaXlsx = await SpreadsheetFile.exportXlsx(ifsFormulaWorkbook);
const importedIfsFormulaWorkbook = await SpreadsheetFile.importXlsx(ifsFormulaXlsx);
assert.deepEqual(importedIfsFormulaWorkbook.worksheets.getItem("Criteria").getRange("E1:E4").formulas, ifsFormulaSheet.getRange("E1:E4").formulas);
assert.deepEqual(importedIfsFormulaWorkbook.worksheets.getItem("Criteria").getRange("G1:G7").formulas, ifsFormulaSheet.getRange("G1:G7").formulas);

const templateFormulaWorkbook = Workbook.create();
const templateFormulaSheet = templateFormulaWorkbook.worksheets.add("Template formulas");
templateFormulaSheet.getRange("A1:A7").values = [[1], [null], [false], ["text"], [0], ["#DIV/0!"], [null]];
templateFormulaSheet.getRange("A7").formulas = [["=IF(FALSE,1,\"\")"]];
templateFormulaSheet.getRange("B1:B5").formulas = [
  ["=COUNTA(A1:A7)"],
  ["=COUNTBLANK(A1:A7)"],
  ["=COUNTBLANK(A1:A2)"],
  ["=COUNTBLANK(A1)"],
  ["=COUNTBLANK(A1:A2,A3)"],
];
assert.deepEqual(templateFormulaSheet.getRange("B1:B5").values, [[6], [2], [1], [0], ["#VALUE!"]]);
templateFormulaSheet.getRange("C1:C6").formulas = [
  ["=TEXT(DATE(2026,7,12),\"yyyymmdd\")"],
  ["=TEXT(DATE(2026,7,12),\"mmm yyyy\")"],
  ["=TEXT(DATE(2026,7,12),\"mmmm yyyy\")"],
  ["=TEXT(DATE(2026,7,2),\"yyyy-mm-dd\")"],
  ["=TEXT(60,\"yyyy-mm-dd\")"],
  ["=TEXT(DATE(2026,7,12),\"0.00\")"],
];
assert.deepEqual(templateFormulaSheet.getRange("C1:C6").values, [["20260712"], ["Jul 2026"], ["July 2026"], ["2026-07-02"], ["1900-02-29"], ["#VALUE!"]]);
const templateFormula1904Workbook = Workbook.create({ dateSystem: "1904" });
const templateFormula1904Sheet = templateFormula1904Workbook.worksheets.add("Date system");
templateFormula1904Sheet.getRange("A1").formulas = [["=TEXT(DATE(1904,1,1),\"yyyy-mm-dd\")"]];
assert.deepEqual(templateFormula1904Sheet.getRange("A1").values, [["1904-01-01"]]);
const templateFormulaXlsx = await SpreadsheetFile.exportXlsx(templateFormulaWorkbook);
const importedTemplateFormulaWorkbook = await SpreadsheetFile.importXlsx(templateFormulaXlsx);
assert.deepEqual(importedTemplateFormulaWorkbook.worksheets.getItem("Template formulas").getRange("B1:B5").formulas, templateFormulaSheet.getRange("B1:B5").formulas);
assert.deepEqual(importedTemplateFormulaWorkbook.worksheets.getItem("Template formulas").getRange("C1:C6").formulas, templateFormulaSheet.getRange("C1:C6").formulas);

const textPositionWorkbook = Workbook.create();
const textPositionSheet = textPositionWorkbook.worksheets.add("Text position");
textPositionSheet.getRange("A1:A5").values = [
  ["Quarterly Review"],
  ["quarterly review"],
  ["A*B"],
  ["A?B"],
  ["Launch 🚀 Review"],
];
textPositionSheet.getRange("B1:B10").formulas = [
  ["=SEARCH(\"review\",A1)"],
  ["=FIND(\"Review\",A2)"],
  ["=FIND(\"review\",A2)"],
  ["=SEARCH(\"Re*W\",A1)"],
  ["=SEARCH(\"~*\",A3)"],
  ["=SEARCH(\"~?\",A4)"],
  ["=SEARCH(\"🚀\",A5)"],
  ["=FIND(\"*\",A3)"],
  ["=SEARCH(\"Review\",A1,12)"],
  ["=SEARCH(\"R\",A1,99)"],
];
assert.deepEqual(textPositionSheet.getRange("B1:B10").values, [[11], ["#VALUE!"], [11], [11], [2], [2], [8], [2], ["#VALUE!"], ["#VALUE!"]]);
textPositionSheet.getRange("A1:A5").conditionalFormats.add("expression", {
  formula: "NOT(ISERROR(SEARCH(\"review\",A1)))",
  format: { fill: "#FEF3C7" },
});
const textPositionStyles = textPositionWorkbook.inspect({ kind: "computedStyle", sheetName: "Text position", range: "A1:A5" }).ndjson
  .trim()
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.deepEqual(textPositionStyles.map((record) => [record.address, record.style.fill]), [["A1", "#FEF3C7"], ["A2", "#FEF3C7"], ["A5", "#FEF3C7"]]);
textPositionSheet.getRange("C1:C3").values = [["Quarterly Review"], ["quarterly review"], ["Draft"]];
textPositionSheet.getRange("C1:C3").conditionalFormats.add("containsText", {
  text: "Review",
  format: { fill: "#DBEAFE" },
});
const containsTextStyles = textPositionWorkbook.inspect({ kind: "computedStyle", sheetName: "Text position", range: "C1:C3" }).ndjson
  .trim()
  .split("\n")
  .filter(Boolean)
  .map((line) => JSON.parse(line));
assert.deepEqual(containsTextStyles.map((record) => [record.address, record.style.fill]), [["C1", "#DBEAFE"], ["C2", "#DBEAFE"]]);

const importedWithoutSourceSnapshot = await SpreadsheetFile.importXlsx(firstXlsx);
const workbookState = importedWithoutSourceSnapshot[Symbol.for("open-office-artifact-tool.open-chestnut-state")];
workbookState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

console.log("spreadsheet tests passed");
