import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";
import { parsePivotCacheDefinition } from "../src/spreadsheet/ooxml-pivots.mjs";
import { evaluatePivotFormula } from "../src/spreadsheet/pivots.mjs";

assert.equal(evaluatePivotFormula("=2+3*4", {}, []), 14);
assert.equal(evaluatePivotFormula("=('Revenue'-'Cost')/2", { Revenue: 15, Cost: 9 }), 3);
assert.equal(evaluatePivotFormula("='Owner''s Revenue'*10%", { "Owner's Revenue": 80 }), 8);
assert.equal(evaluatePivotFormula("=Revenue/0", { Revenue: 10 }), "#DIV/0!");
assert.equal(evaluatePivotFormula("=-(Revenue/0)%", { Revenue: 10 }), "#DIV/0!");
const groupedCacheContract = parsePivotCacheDefinition(`<pivotCacheDefinition><cacheSource><worksheetSource ref="A1:A2" sheet="S"/></cacheSource><cacheFields count="2"><cacheField name="Source"><sharedItems count="1"><s v="A"/></sharedItems></cacheField><cacheField name="Grouped" databaseField="0"><fieldGroup base="0"/></cacheField></cacheFields></pivotCacheDefinition>`);
assert.deepEqual(groupedCacheContract.sourceFields, ["Source"]);
assert.deepEqual(groupedCacheContract.calculatedFields, []);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
sheet.getRange("A1:D1").format = { fill: "#0f172a", font: { bold: true, color: "#ffffff", name: "Aptos", size: 12 }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: { style: "thin", color: "#334155" } };
sheet.getRange("B2:C3").setFormat({ numberFormat: "#,##0", fill: "sky-100" });
workbook.recalculate();
assert.equal(sheet.freezePanes.freezeRows(1), sheet.freezePanes);
assert.equal(sheet.freezePanes.freezeColumns(2), sheet.freezePanes);
assert.deepEqual(sheet.freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
sheet.freezePanes.freezeRows(0);
assert.deepEqual(sheet.freezePanes.toJSON(), { rows: 0, columns: 2, frozen: true, topLeftCell: "C1", activePane: "topRight" });
sheet.freezePanes.unfreeze().freezeRows(1).freezeColumns(2);
sheet.showGridLines = false;
assert.throws(() => sheet.freezePanes.freezeRows(-1), /integer from 0/);
assert.throws(() => sheet.freezePanes.freezeColumns(1.5), /integer from 0/);
assert.match(workbook.help("worksheet.freezePanes").ndjson, /Freeze a leading row count/);
const sheetViewInspect = workbook.inspect({ kind: "sheet" }).ndjson;
assert.match(sheetViewInspect, /"showGridLines":false/);
assert.match(sheetViewInspect, /"freezePanes":\{"rows":1,"columns":2,"frozen":true,"topLeftCell":"C2","activePane":"bottomRight"\}/);
assert.deepEqual(sheet.layoutJson().view.freezePanes, { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const invalidFreezeBook = Workbook.create();
const invalidFreezeSheet = invalidFreezeBook.worksheets.add("InvalidFreeze");
invalidFreezeSheet.getRange("A1").values = [[1]];
invalidFreezeSheet.freezePanes._rows = -1;
assert.ok(invalidFreezeBook.verify().issues.some((issue) => issue.type === "invalidFrozenRows"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidFreezeBook), /frozen row count must be an integer/);
const unfrozenBook = Workbook.create();
const unfrozenSheet = unfrozenBook.worksheets.add("Unfrozen");
unfrozenSheet.getRange("A1").values = [[1]];
unfrozenSheet.freezePanes.freezeRows(1).freezeColumns(1).unfreeze();
const unfrozenZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(unfrozenBook)).arrayBuffer()));
assert.doesNotMatch(await unfrozenZip.file("xl/worksheets/sheet1.xml").async("text"), /<sheetViews>/);

const dimensionBook = Workbook.create();
const dimensionSheet = dimensionBook.worksheets.add("Dimensions");
dimensionSheet.getRange("A1:C3").values = [["ID", "Long descriptive heading", "Hidden"], [1, "first\nsecond\nthird", "x"], [2, "short", "y"]];
dimensionSheet.getRange("A1:A3").format.columnWidth = 18;
dimensionSheet.getRange("B1:B3").format.columnWidthPx = 120;
dimensionSheet.getRange("A1:C1").format.rowHeight = 24;
dimensionSheet.getRange("A2:C2").format.rowHeightPx = 30;
dimensionSheet.getRange("C1:C3").format.columnHidden = true;
dimensionSheet.getRange("A3:C3").format.rowHidden = true;
dimensionSheet.getRange("A1:C1").format.fill = "#123456";
assert.equal(dimensionSheet.getRange("A1").format.fill, "#123456");
assert.equal(dimensionSheet.getRange("A1:A3").format.columnWidth, 18);
assert.ok(Math.abs(dimensionSheet.getRange("B1:B3").format.columnWidthPx - 120) <= 1);
assert.equal(dimensionSheet.getRange("A1:C1").format.rowHeight, 24);
assert.ok(Math.abs(dimensionSheet.getRange("A2:C2").format.rowHeightPx - 30) < 0.01);
assert.equal(dimensionSheet.getRange("C1:C3").format.columnHidden, true);
assert.equal(dimensionSheet.getRange("A3:C3").format.rowHidden, true);
assert.throws(() => { dimensionSheet.getRange("A1").format.columnWidth = 0; }, /column width must be greater than 0/);
assert.throws(() => { dimensionSheet.getRange("A1").format.rowHeightPx = -1; }, /row height must be greater than 0/);
dimensionSheet.getRange("A1:B3").format.autofitColumns();
dimensionSheet.getRange("A1:B3").format.autofitRows();
assert.ok(dimensionSheet.getRange("B1:B3").format.columnWidthPx > dimensionSheet.getRange("A1:A3").format.columnWidthPx);
assert.ok(dimensionSheet.getRange("A2:C2").format.rowHeight > 15);
const dimensionInspect = dimensionBook.inspect({ kind: "dimension" }).ndjson;
assert.match(dimensionInspect, /"kind":"column"/);
assert.match(dimensionInspect, /"bestFit":true/);
assert.match(dimensionInspect, /"column":"C"[\s\S]*"hidden":true/);
assert.match(dimensionInspect, /"kind":"row"[\s\S]*"row":3[\s\S]*"hidden":true/);
const dimensionLayout = dimensionSheet.layoutJson();
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "C1").bbox[2], 0);
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "A3").bbox[3], 0);
assert.equal(dimensionLayout.cells.find((cell) => cell.address === "C1").hidden, true);
const dimensionXlsx = await SpreadsheetFile.exportXlsx(dimensionBook);
const dimensionZip = await JSZip.loadAsync(new Uint8Array(await dimensionXlsx.arrayBuffer()));
const dimensionXml = await dimensionZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(dimensionXml, /<cols>[\s\S]*<col min="1" max="1" width="[^"]+" customWidth="1" bestFit="1"\/>/);
assert.match(dimensionXml, /<col min="3" max="3" hidden="1"\/>/);
assert.match(dimensionXml, /<row r="1" ht="[^"]+" customHeight="1">/);
assert.match(dimensionXml, /<row r="3"[^>]*hidden="1"[^>]*>/);
const dimensionRoundtripSheet = (await SpreadsheetFile.importXlsx(dimensionXlsx)).worksheets.getItem("Dimensions");
assert.equal(dimensionRoundtripSheet.getRange("C1:C3").format.columnHidden, true);
assert.equal(dimensionRoundtripSheet.getRange("A3:C3").format.rowHidden, true);
assert.ok(dimensionRoundtripSheet.getRange("B1:B3").format.columnWidthPx > dimensionRoundtripSheet.getRange("A1:A3").format.columnWidthPx);
const thirdPartyDimensionXml = dimensionXml
  .replace(/<cols>[\s\S]*?<\/cols>/, '<cols><col min="2" max="3" width="20" customWidth="1" hidden="1" bestFit="1"/></cols>')
  .replace('</sheetData>', '<row r="4" ht="33" customHeight="1" hidden="1"/></sheetData>');
dimensionZip.file("xl/worksheets/sheet1.xml", thirdPartyDimensionXml);
const thirdPartyDimensionBook = await SpreadsheetFile.importXlsx(new FileBlob(await dimensionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: dimensionXlsx.type }));
const thirdPartyDimensionSheet = thirdPartyDimensionBook.worksheets.getItem("Dimensions");
assert.equal(thirdPartyDimensionSheet.getRange("B1:C3").format.columnHidden, true);
assert.equal(thirdPartyDimensionSheet.getRange("B1:C3").format.columnWidthPx, 140);
assert.equal(thirdPartyDimensionSheet.getRange("A4:C4").format.rowHeight, 33);
assert.equal(thirdPartyDimensionSheet.getRange("A4:C4").format.rowHidden, true);
const thirdPartyDimensionRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(thirdPartyDimensionBook)).arrayBuffer()));
assert.match(await thirdPartyDimensionRoundtripZip.file("xl/worksheets/sheet1.xml").async("text"), /<col min="2" max="3" width="20" customWidth="1" hidden="1" bestFit="1"\/>/);
const invalidDimensionBook = Workbook.create();
const invalidDimensionSheet = invalidDimensionBook.worksheets.add("InvalidDimensions");
invalidDimensionSheet.getRange("A1").values = [[1]];
invalidDimensionSheet.columnDimensions.set(0, { width: 300 });
assert.ok(invalidDimensionBook.verify().issues.some((issue) => issue.type === "invalidColumnWidth"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidDimensionBook), /column 1 width must be greater than 0/);

const mergeFillBook = Workbook.create();
const mergeSheet = mergeFillBook.worksheets.add("Merged");
mergeSheet.getRange("A1:F2").values = [["Header", "discard", "discard", "North", "discard", "discard"], ["Left", "discard", "discard", "South", "discard", "discard"]];
const mergedHeaderRange = mergeSheet.getRange("A1:C1");
assert.equal(mergedHeaderRange.merge(), mergedHeaderRange);
assert.deepEqual(mergeSheet.getRange("A1:C1").values, [["Header", null, null]]);
mergeSheet.getRange("D1:F2").merge(true);
assert.deepEqual(mergeSheet.mergedRanges, ["A1:C1", "D1:F1", "D2:F2"]);
assert.deepEqual(mergeSheet.getRange("D1:F2").values, [["North", null, null], ["South", null, null]]);
assert.throws(() => mergeSheet.getRange("B1:D1").merge(), /overlaps existing merged range/);
mergeSheet.getRange("A1:B1").unmerge();
assert.deepEqual(mergeSheet.mergedRanges, ["D1:F1", "D2:F2"]);
mergeSheet.mergeCells("A1:C1");
mergeSheet.mergeCells("J10:K11");
mergeSheet.getRange("M1").merge();
assert.equal(mergeSheet.mergedRanges.includes("M1:M1"), false);
assert.equal(mergeSheet.usedBounds().right, 10);
assert.equal(mergeSheet.usedBounds().bottom, 10);
const mergeInspect = mergeFillBook.inspect({ kind: "sheet,merge" }).ndjson;
assert.match(mergeInspect, /"mergedRanges":4/);
assert.match(mergeInspect, /"kind":"mergedCell"[\s\S]*"range":"J10:K11"/);
const mergeLayout = mergeSheet.layoutJson();
assert.equal(mergeLayout.merges.length, 4);
assert.equal(mergeLayout.cells.find((cell) => cell.address === "A1").colSpan, 3);
assert.equal(mergeLayout.cells.find((cell) => cell.address === "B1").mergedParent, "A1");
assert.equal(mergeLayout.cells.find((cell) => cell.address === "B1").bbox[2], 0);
assert.equal((mergeSheet.toSvg().match(/>Header<\/text>/g) || []).length, 1);

const dataSheet = mergeFillBook.worksheets.add("Data Sheet");
dataSheet.getRange("A1:A3").values = [[1], [2], [3]];
const fillSheet = mergeFillBook.worksheets.add("Fill");
fillSheet.getRange("A1:B3").values = [[10, 2], [20, 3], [30, 4]];
fillSheet.getRange("C1").formulas = [["=A1+$B$1+B$1+$A1+'Data Sheet'!A1"]];
fillSheet.getRange("C1").format = { fill: "#abcdef", font: { bold: true } };
const filledFormulaRange = fillSheet.getRange("C1:C3");
assert.equal(filledFormulaRange.fillDown(), filledFormulaRange);
assert.deepEqual(fillSheet.getRange("C1:C3").formulas.flat(), ["=A1+$B$1+B$1+$A1+'Data Sheet'!A1", "=A2+$B$1+B$1+$A2+'Data Sheet'!A2", "=A3+$B$1+B$1+$A3+'Data Sheet'!A3"]);
assert.deepEqual(fillSheet.getRange("C1:C3").values.flat(), [25, 46, 67]);
assert.equal(fillSheet.getRange("C3").format.fill, "#abcdef");
assert.equal(fillSheet.getRange("C3").format.font.bold, true);
fillSheet.getRange("D1").formulas = [["=IF(A1=\"A1\",A1,0)"]];
fillSheet.getRange("D1:D3").fillDown();
assert.equal(fillSheet.getRange("D2").formulas[0][0], '=IF(A2="A1",A2,0)');
fillSheet.getRange("A5").formulas = [["=A1+$A1+A$1+$A$1"]];
fillSheet.getRange("A5").format = { fill: "#fedcba" };
fillSheet.getRange("A5:C5").fillRight();
assert.deepEqual(fillSheet.getRange("A5:C5").formulas.flat(), ["=A1+$A1+A$1+$A$1", "=B1+$A1+B$1+$A$1", "=C1+$A1+C$1+$A$1"]);
assert.equal(fillSheet.getRange("C5").format.fill, "#fedcba");
fillSheet.getRange("A7").values = [["repeat"]];
fillSheet.getRange("A7:C7").fillRight();
assert.deepEqual(fillSheet.getRange("A7:C7").values, [["repeat", "repeat", "repeat"]]);
fillSheet.getRange("E1:F1").merge();
assert.throws(() => fillSheet.getRange("E1:E3").fillDown(), /intersects merged cells/);

const mergeFillXlsx = await SpreadsheetFile.exportXlsx(mergeFillBook);
const mergeFillZip = await JSZip.loadAsync(new Uint8Array(await mergeFillXlsx.arrayBuffer()));
const mergeSheetXml = await mergeFillZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(mergeSheetXml, /<mergeCells count="4"><mergeCell ref="A1:C1"\/><mergeCell ref="D1:F1"\/><mergeCell ref="D2:F2"\/><mergeCell ref="J10:K11"\/><\/mergeCells>/);
const mergeFillRoundtrip = await SpreadsheetFile.importXlsx(mergeFillXlsx);
assert.deepEqual(mergeFillRoundtrip.worksheets.getItem("Merged").mergedRanges, ["A1:C1", "D1:F1", "D2:F2", "J10:K11"]);
assert.deepEqual(mergeFillRoundtrip.worksheets.getItem("Fill").getRange("C1:C3").formulas, fillSheet.getRange("C1:C3").formulas);
assert.equal(mergeFillRoundtrip.worksheets.getItem("Fill").getRange("C3").format.fill, "#ABCDEF");
const overlappingMergeXml = mergeSheetXml.replace(/<mergeCells[\s\S]*?<\/mergeCells>/, '<mergeCells count="2"><mergeCell ref="A1:C1"/><mergeCell ref="B1:D1"/></mergeCells>');
mergeFillZip.file("xl/worksheets/sheet1.xml", overlappingMergeXml);
const overlappingMergeBook = await SpreadsheetFile.importXlsx(new FileBlob(await mergeFillZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: mergeFillXlsx.type }));
assert.ok(overlappingMergeBook.verify().issues.some((issue) => issue.type === "overlappingMergedRanges"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(overlappingMergeBook), /merged range B1:D1 overlaps A1:C1/);

const inspect = workbook.inspect({ kind: "workbook,sheet,table,formula", range: "A1:C3", include: "values,formulas" });
assert.match(inspect.ndjson, /"value":5|"values":/);
assert.match(inspect.ndjson, /"formula":"=A2\+B2"/);
const styleInspect = workbook.inspect({ kind: "style", range: "A1:D3", maxChars: 8000 }).ndjson;
assert.match(styleInspect, /"kind":"style"/);
assert.match(styleInspect, /"numberFormat":"#,##0"/);
assert.match(styleInspect, /"fill":"sky-100"/);
assert.match(styleInspect, /"alignment":\{"horizontal":"center"/);
assert.match(styleInspect, /"border":\{"style":"thin"/);
assert.match(workbook.help("range.format").ndjson, /cell style/);
assert.match(workbook.help("fx.PMT").ndjson, /fx.PMT/);
assert.match(workbook.help("range.dataValidation").ndjson, /validation rule/);
assert.match(workbook.help("range.conditionalFormats.add").ndjson, /conditional formatting/);
assert.match(workbook.help("workbook.comments.addThread").ndjson, /threaded comments/);
const statusRange = sheet.getRange("D2:D3");
statusRange.values = [["Not Started"], ["In Progress"]];
statusRange.dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Done"] } };
const cf = sheet.getRange("C2:C3").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 10, format: { fill: "green" } });
const customCf = sheet.getRange("A2:B3").conditionalFormats.addCustom("=A2<B2", { fill: "sky-100" });
workbook.comments.setSelf({ displayName: "Analyst" });
const thread = workbook.comments.addThread({ cell: sheet.getRange("C2") }, "Formula checks revenue sum.");
thread.addReply("Reviewed by model.").resolve();

const tasksTable = sheet.tables.add("A1:D3", true, "TasksTable");
tasksTable.style = "TableStyleMedium4";
tasksTable.showBandedColumns = true;
tasksTable.rows.add(null, [["8", "13", "21", "Done"]]);
sheet.getRange("E2").formulas = [["=SUM(TasksTable[Sum])"]];
workbook.recalculate();
assert.equal(sheet.getRange("E2").values[0][0], 38);
assert.deepEqual(tasksTable.getDataRows()[0].slice(0, 3), [2, 3, 5]);
assert.equal(tasksTable.getHeaderRowRange().values[0][0], "A");

const chartSource = sheet.getRange("F1:G4");
chartSource.values = [["Month", "Revenue"], ["Jan", 100], ["Feb", 120], ["Mar", 130]];
const colorScaleCf = chartSource.getRange("G2:G4").conditionalFormats.addColorScale({ colors: ["#fee2e2", "#fef3c7", "#22c55e"] });
const revenueName = workbook.definedNames.add("RevenueData", "Sheet1!G2:G4", { comment: "Revenue data body" });
sheet.getRange("E3").formulas = [["=SUM(RevenueData)"]];
workbook.recalculate();
assert.equal(sheet.getRange("E3").values[0][0], 350);
const chartFromRange = sheet.charts.add("line", chartSource);
chartFromRange.title = "Revenue Trend";
chartFromRange.hasLegend = false;
chartFromRange.setPosition("I1", "M10");
const chartFromConfig = sheet.charts.add("bar", { name: "ScoresChart", title: "Scores", categories: ["A", "B"], series: [{ name: "Score", values: [9, 7] }], position: { left: 40, top: 220, width: 240, height: 160 } });
const revenuePivot = sheet.pivotTables.add({
  name: "RevenuePivot",
  sourceRange: "F1:G4",
  targetRange: "N1:O4",
  rows: ["Month"],
  values: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue sum" }],
});
assert.deepEqual(revenuePivot.computedValues(), [["Month", "Revenue sum"], ["Jan", 100], ["Feb", 120], ["Mar", 130]]);
sheet.getRange("P1:T7").values = [
  ["Region", "Quarter", "Product", "Revenue", "Cost"],
  ["East", "Q1", "Core", 10, 6],
  ["East", "Q2", "Core", 20, 12],
  ["West", "Q1", "Core", 30, 20],
  ["West", "Q2", "Legacy", 40, 25],
  ["East", "Q1", "Legacy", 5, 3],
  ["West", "Q1", "Legacy", 7, 4],
];
const regionalPivot = sheet.pivotTables.add({
  name: "RegionalPivot",
  sourceRange: "P1:T7",
  targetRange: "V1:X4",
  rowFields: ["Region"],
  columnFields: ["Quarter"],
  valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue total" }, { field: "Profit", name: "Profit total" }],
  calculatedFields: [{ name: "Profit", formula: "=([Revenue] - [Cost]) * 100%" }],
  filters: { Quarter: { include: ["Q1"] } },
  refreshPolicy: { refreshOnLoad: false, saveData: true, enableRefresh: false, invalid: true, missingItemsLimit: 3, refreshedBy: "QA Agent", refreshedDateIso: "2026-07-12T00:00:00Z" },
});
assert.deepEqual(regionalPivot.computedValues(), [["Region", "Q1 — Revenue total", "Q1 — Profit total"], ["East", 15, 6], ["West", 37, 13]]);
assert.deepEqual(regionalPivot.calculatedFields, [{ name: "Profit", formula: "=('Revenue'-'Cost')*100%", numFmtId: 0, references: ["Revenue", "Cost"] }]);
assert.deepEqual(regionalPivot.filters, [{ field: "Quarter", include: ["Q1"] }]);
assert.equal(regionalPivot.refreshPolicy.saveData, true);
assert.match(regionalPivot.inspectRecord().columnFields.join(","), /Quarter/);
assert.match(JSON.stringify(regionalPivot.layoutJson()), /refreshPolicy/);
assert.match(JSON.stringify(regionalPivot.layoutJson()), /calculatedFields/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Missing"] }), /not present in the source headers/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], columnFields: ["Region"] }), /both a row and column field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Quarter: ["Q1"] } }), /must also be a row or column field/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Region: { include: ["East"], exclude: ["West"] } } }), /exactly one of include or exclude/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], filters: { Region: ["North"] } }), /unknown item North/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], refreshPolicy: { saveData: "yes" } }), /saveData must be a boolean/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:S7", targetRange: "T8", rowFields: ["Region"], refreshPolicy: { refreshedDateIso: "today" } }), /XML date-time string/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Revenue", formula: "Cost" }] }), /must be unique/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue - Missing" }] }), /unknown source field Missing/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue +" }] }), /ended unexpectedly/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: "Revenue-Cost", numFmtId: -1 }] }), /non-negative integer/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: [{ name: "Profit", formula: `Revenue+${"1+".repeat(2050)}1` }] }), /4096 characters/);
assert.throws(() => sheet.pivotTables.add({ sourceRange: "P1:T7", targetRange: "V8", rowFields: ["Region"], calculatedFields: Array.from({ length: 129 }, (_, index) => ({ name: `C${index}`, formula: "Revenue" })) }), /exceeds 128 fields/);
assert.equal(workbook.resolve(revenuePivot.id).name, "RevenuePivot");
assert.match(workbook.help("sheet.pivotTables.add").ndjson, /pivot table facade/);
const image = sheet.images.add({
  name: "LogoImage",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  alt: "Logo placeholder",
  anchor: { from: { row: 8, col: 0 }, extent: { widthPx: 120, heightPx: 80 } },
});
const sparklineGroup = sheet.sparklineGroups.add({
  type: "line",
  targetRange: "H2:H2",
  sourceData: chartSource.getRange("G2:G4"),
  seriesColor: "#0284c7",
  markers: { show: true },
});
const sparklineAlias = sheet.getRange("H3:H3").sparklines.add("column", chartSource.getRange("G2:G4"), { seriesColor: "#f97316" });
assert.equal(image.alt, "Logo placeholder");
assert.deepEqual(sparklineGroup.values(), [100, 120, 130]);
assert.equal(sparklineAlias.type, "column");
assert.match(workbook.help("sheet.images.add").ndjson, /worksheet image/);
assert.match(workbook.help("sheet.sparklineGroups.add").ndjson, /sparklines/);
assert.equal(chartFromRange.series.items[0].name, "Revenue");
assert.match(chartFromRange.series.items[0].formula, /G2:G4/);
assert.equal(chartFromConfig.series.getItemAt(0).values[1], 7);

const metadataInspect = workbook.inspect({ kind: "dataValidation,conditionalFormat,thread,table,pivotTable,drawing", maxChars: 16000 }).ndjson;
assert.match(metadataInspect, /"kind":"dataValidation"/);
assert.match(metadataInspect, /"type":"list"/);
assert.match(metadataInspect, /"kind":"conditionalFormat"/);
assert.match(metadataInspect, /"ruleType":"cellIs"/);
assert.match(metadataInspect, /"ruleType":"expression"/);
assert.match(metadataInspect, /"ruleType":"colorScale"/);
assert.match(metadataInspect, /fee2e2/);
assert.match(metadataInspect, /"kind":"thread"/);
assert.match(metadataInspect, /Formula checks revenue sum/);
assert.match(metadataInspect, /TasksTable/);
assert.match(metadataInspect, /RevenuePivot/);
assert.match(metadataInspect, /Revenue sum/);
assert.match(metadataInspect, /Revenue Trend/);
assert.match(metadataInspect, /ScoresChart/);
assert.match(metadataInspect, /LogoImage/);
assert.match(metadataInspect, /"kind":"sparkline"/);
const definedNameInspect = workbook.inspect({ kind: "definedName", target: revenueName.id, maxChars: 4000 }).ndjson;
assert.match(definedNameInspect, /RevenueData/);
assert.match(definedNameInspect, /Sheet1!G2:G4/);
assert.equal(workbook.resolve(revenueName.id).name, "RevenueData");
assert.equal(workbook.resolve("RevenueData").refersTo, "Sheet1!G2:G4");
assert.match(workbook.help("workbook.definedNames.add").ndjson, /defined name/);
const targetedImageInspect = workbook.inspect({ kind: "drawing", target: image.id, maxChars: 4000 }).ndjson;
assert.match(targetedImageInspect, /LogoImage/);
assert.doesNotMatch(targetedImageInspect, /Revenue Trend/);
const targetedFormulaInspect = workbook.inspect({ kind: "formula", target: "Sheet1!C2", maxChars: 4000 }).ndjson;
assert.match(targetedFormulaInspect, /=A2\+B2/);
assert.doesNotMatch(targetedFormulaInspect, /=A3\+B3/);
const formulaContextInspect = workbook.inspect({ kind: "formula", target: "Sheet1!C2", after: 1, maxChars: 4000 }).ndjson;
assert.match(formulaContextInspect, /=A2\+B2/);
assert.match(formulaContextInspect, /=A3\+B3/);
const shapedFormulaInspect = workbook.inspect({ kind: "formula", target: "Sheet1!E2", include: "formula,value,precedents", exclude: "dependents", maxChars: 4000 }).ndjson;
assert.match(shapedFormulaInspect, /TasksTable\[Sum\]/);
assert.match(shapedFormulaInspect, /"value":38/);
assert.match(shapedFormulaInspect, /"precedents"/);
assert.doesNotMatch(shapedFormulaInspect, /"dependents"/);
assert.equal(workbook.resolve(thread.id).resolved, true);
assert.equal(workbook.resolve(cf.id).operator, "greaterThan");
assert.equal(workbook.resolve(customCf.id).formula, "=A2<B2");
const computedCellIsStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!C3", maxChars: 4000 }).ndjson;
assert.match(computedCellIsStyle, new RegExp(cf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(computedCellIsStyle, /"fill":"green"/);
const computedExpressionStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!A2", maxChars: 4000 }).ndjson;
assert.match(computedExpressionStyle, new RegExp(customCf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(computedExpressionStyle, /"fill":"sky-100"/);
const colorScaleStyle = workbook.inspect({ kind: "computedStyle", target: "Sheet1!G4", maxChars: 4000 }).ndjson;
assert.match(colorScaleStyle, new RegExp(colorScaleCf.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(colorScaleStyle, /"fill":"#22c55e"/);
assert.equal(workbook.resolve(tasksTable.id).name, "TasksTable");
assert.equal(workbook.resolve(chartFromRange.id).title, "Revenue Trend");
assert.equal(workbook.resolve(image.id).alt, "Logo placeholder");
assert.equal(workbook.resolve(sparklineGroup.id).type, "line");
assert.equal(workbook.resolve(sparklineAlias.id).type, "column");
const trace = workbook.trace("Sheet1!C2");
assert.equal(trace.tree.address, "C2");
assert.equal(trace.tree.value, 5);
assert.deepEqual(trace.tree.precedents.map((node) => node.address), ["A2", "B2"]);
assert.match(trace.ndjson, /"precedents":\["Sheet1!A2","Sheet1!B2"\]/);
const structuredTrace = workbook.trace("Sheet1!E2");
assert.equal(structuredTrace.tree.value, 38);
assert.deepEqual(structuredTrace.tree.precedents.map((node) => node.address), ["C2", "C3", "C4"]);
assert.match(workbook.inspect({ kind: "formula", target: "Sheet1!E2", maxChars: 8000 }).ndjson, /TasksTable\[Sum\]/);
assert.match(workbook.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson, /RevenueData/);
assert.match(workbook.inspect({ kind: "formulaNode", target: "Sheet1!E3", maxChars: 12000 }).ndjson, /Sheet1!G4/);
assert.match(workbook.inspect({ kind: "formulaNode", target: "Sheet1!E2", maxChars: 12000 }).ndjson, /Sheet1!C4/);
assert.match(workbook.help("workbook.structuredReferences").ndjson, /TableName\[Column\]/);
assert.match(workbook.help("workbook.sharedArrayFormulas").ndjson, /shared formulas/);
assert.match(workbook.help("workbook.trace").ndjson, /precedent tree/);
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);

const structuredBook = Workbook.create();
const structuredSheet = structuredBook.worksheets.add("Structured");
structuredSheet.getRange("A1:C4").values = [["Region", "Revenue", "Cost"], ["East", 10, 4], ["West", 5, 2], ["Total", 15, 6]];
const structuredTable = structuredSheet.tables.add({ range: "A1:C4", name: "SalesTable", showTotals: true });
structuredTable.showTotals = true;
structuredSheet.getRange("D1:E4").formulas = [
  ["=TEXTJOIN(\"|\",TRUE,SalesTable[#Headers])", "=TEXTJOIN(\"|\",TRUE,SalesTable[[#Headers],[Region]:[Cost]])"],
  ["=SUM(SalesTable[[#Data],[Revenue]])", "=SUM(SalesTable[[#Data],[Revenue]:[Cost]])"],
  ["=SUM(SalesTable[[#Totals],[Revenue]])", "=TEXTJOIN(\"|\",TRUE,SalesTable[[#Data],[Region],[Cost]])"],
  ["=SUM(SalesTable[[#All],[Revenue]])", "=SUM(SalesTable[[#Totals],[Revenue]:[Cost]])"],
];
structuredBook.recalculate();
assert.deepEqual(structuredSheet.getRange("D1:E4").values, [["Region|Revenue|Cost", "Region|Revenue|Cost"], [15, 21], [15, "East|4|West|2"], [30, 21]]);
const structuredNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!D2", maxChars: 12000 }).ndjson;
assert.match(structuredNode, /SalesTable\[\[#Data\],\[Revenue\]\]/);
assert.match(structuredNode, /Structured!B2/);
assert.match(structuredNode, /Structured!B3/);
assert.doesNotMatch(structuredNode, /Structured!B4/);
const structuredRangeNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!E2", maxChars: 12000 }).ndjson;
assert.match(structuredRangeNode, /SalesTable\[\[#Data\],\[Revenue\]:\[Cost\]\]/);
assert.match(structuredRangeNode, /Structured!B2/);
assert.match(structuredRangeNode, /Structured!C3/);
assert.doesNotMatch(structuredRangeNode, /Structured!A2/);
const structuredUnionNode = structuredBook.inspect({ kind: "formulaNode", target: "Structured!E3", maxChars: 12000 }).ndjson;
assert.match(structuredUnionNode, /SalesTable\[\[#Data\],\[Region\],\[Cost\]\]/);
assert.match(structuredUnionNode, /Structured!A2/);
assert.match(structuredUnionNode, /Structured!C3/);
assert.doesNotMatch(structuredUnionNode, /Structured!B2/);
assert.match(structuredBook.help("workbook.structuredReferences").ndjson, /\[First\]:\[Last\]/);

const escapedStructuredBook = Workbook.create();
const escapedStructuredSheet = escapedStructuredBook.worksheets.add("Escaped");
escapedStructuredSheet.getRange("A1:G4").values = [
  ["Revenue", "Cost", "#Items", "@Rate", "Bracket[Value]", "Owner's", "Net"],
  [10, 4, 2, 0.1, 7, 1, null],
  [5, 2, 3, 0.2, 8, 2, null],
  [12, 6, 1, 0.3, 9, 3, null],
];
escapedStructuredSheet.tables.add({ range: "A1:G4", name: "Specials" });
escapedStructuredSheet.getRange("G2:G4").formulas = [
  ["=[Revenue]-[Cost]+['#Items]+[Owner''s]"],
  ["=Specials[@Revenue]-Specials[@Cost]+Specials[@['@Rate]]"],
  ["=SUM(Specials[[#This Row],[Revenue]:[Bracket'[Value']]])"],
];
escapedStructuredSheet.getRange("H2").formulas = [["=Specials[@Revenue]"]];
escapedStructuredBook.recalculate();
assert.deepEqual(escapedStructuredSheet.getRange("G2:G4").values.flat(), [9, 3.2, 28.3]);
assert.equal(escapedStructuredSheet.getRange("H2").values[0][0], "#VALUE!");
const escapedStructuredG2 = escapedStructuredBook.inspect({ kind: "formulaNode", target: "Escaped!G2", maxChars: 12000 }).ndjson;
assert.match(escapedStructuredG2, /Escaped!A2/);
assert.match(escapedStructuredG2, /Escaped!B2/);
assert.match(escapedStructuredG2, /Escaped!C2/);
assert.match(escapedStructuredG2, /Escaped!F2/);
assert.doesNotMatch(escapedStructuredG2, /Escaped!A3/);
const escapedStructuredTrace = escapedStructuredBook.trace("Escaped!G4");
assert.deepEqual(escapedStructuredTrace.tree.precedents.map((node) => node.address), ["A4", "B4", "C4", "D4", "E4"]);
const escapedStructuredXlsx = await SpreadsheetFile.exportXlsx(escapedStructuredBook);
const escapedStructuredZip = await JSZip.loadAsync(new Uint8Array(await escapedStructuredXlsx.arrayBuffer()));
assert.match(await escapedStructuredZip.file("xl/worksheets/sheet1.xml").async("text"), /Specials\[\[#This Row\],\[Revenue\]:\[Bracket'\[Value'\]\]\]/);
escapedStructuredZip.file("xl/custom/native-table.xml", await escapedStructuredZip.file("xl/tables/table1.xml").async("text"));
escapedStructuredZip.remove("xl/tables/table1.xml");
escapedStructuredZip.file("xl/worksheets/_rels/sheet1.xml.rels", (await escapedStructuredZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text")).replace("../tables/table1.xml", "../custom/native-table.xml"));
escapedStructuredZip.file("[Content_Types].xml", (await escapedStructuredZip.file("[Content_Types].xml").async("text")).replace("/xl/tables/table1.xml", "/xl/custom/native-table.xml"));
escapedStructuredZip.remove("customXml/open-office-artifact.json");
const escapedStructuredNative = await SpreadsheetFile.importXlsx(new FileBlob(await escapedStructuredZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: escapedStructuredXlsx.type }));
assert.deepEqual(escapedStructuredNative.worksheets.getItem("Escaped").getRange("G2:G4").values.flat(), [9, 3.2, 28.3]);
assert.deepEqual(escapedStructuredNative.worksheets.getItem("Escaped").tables.getItemOrNullObject("Specials").columnNames, ["Revenue", "Cost", "#Items", "@Rate", "Bracket[Value]", "Owner's", "Net"]);
assert.deepEqual(escapedStructuredNative.trace("Escaped!G3").tree.precedents.map((node) => node.address), ["A3", "B3", "D3"]);
assert.match(escapedStructuredNative.help("workbook.structuredReferences").ndjson, /#This Row/);

const graphBook = Workbook.create();
const inputsSheet = graphBook.worksheets.add("Inputs");
inputsSheet.getRange("A1:A3").values = [[2], [4], [6]];
const calcSheet = graphBook.worksheets.add("Calc");
calcSheet.getRange("A1:A5").formulas = [["=SUM(Inputs!A1:A3)"], ["=AVERAGE(Inputs!A1:A3)"], ["=MIN(Inputs!A1:A3)"], ["=MAX(Inputs!A1:A3)"], ["=COUNT(Inputs!A1:A3)"]];
graphBook.recalculate();
assert.deepEqual(calcSheet.getRange("A1:A5").values.flat(), [12, 4, 2, 6, 3]);
const formulaGraph = graphBook.formulaGraph();
assert.equal(formulaGraph.nodes.length, 5);
assert.ok(formulaGraph.edges.some((edge) => edge.from === "Calc!A1" && edge.to === "Inputs!A3"));
assert.match(graphBook.inspect({ kind: "formulaNode,formulaEdge", maxChars: 12000 }).ndjson, /"kind":"formulaEdge"/);
assert.match(graphBook.inspect({ kind: "formula", maxChars: 12000 }).ndjson, /"precedents":\["Inputs!A1","Inputs!A2","Inputs!A3"\]/);

const catalogBook = Workbook.create();
const catalogSheet = catalogBook.worksheets.add("Catalog");
catalogSheet.getRange("A1:D4").values = [["Name", "Score", "Region", "Code"], ["Alpha", 10, "East", "AX-100"], ["Beta", 15, "West", "BX-200"], ["Gamma", 20, "East", "CX-300"]];
catalogSheet.getRange("F1:F39").formulas = [
  ["=IF(B2>9,\"ok\",\"bad\")"],
  ["=ROUND(1.234,2)"],
  ["=COUNTIF(C2:C4,\"East\")"],
  ["=SUMIF(C2:C4,\"East\",B2:B4)"],
  ["=VLOOKUP(\"Beta\",A2:B4,2,FALSE)"],
  ["=XLOOKUP(\"Gamma\",A2:A4,B2:B4,\"missing\")"],
  ["=CONCAT(LEFT(A2,2),\"-\",RIGHT(D2,3))"],
  ["=TEXTJOIN(\"/\",TRUE,A2:A3)"],
  ["=LEN(TRIM(\"  padded  \"))"],
  ["=UPPER(LOWER(\"MiXeD\"))"],
  ["=AND(B2>5,B3>10)"],
  ["=OR(B2>50,B3>10)"],
  ["=NOT(B2<5)"],
  ["=ABS(-4)"],
  ["=CEILING(7,3)"],
  ["=FLOOR(7,3)"],
  ["=INDEX(A2:D4,2,4)"],
  ["=MATCH(\"Gamma\",A2:A4,0)"],
  ["=INDEX(B2:B4,MATCH(\"Beta\",A2:A4,0),1)"],
  ["=MATCH(16,B2:B4,1)"],
  ["=COUNTIFS(C2:C4,\"East\",B2:B4,\">=10\")"],
  ["=SUMIFS(B2:B4,C2:C4,\"East\",B2:B4,\">10\")"],
  ["=SUMPRODUCT(B2:B4,B2:B4)"],
  ["=HLOOKUP(\"Region\",A1:D4,3,FALSE)"],
  ["=IFERROR(XLOOKUP(\"Missing\",A2:A4,B2:B4),\"not found\")"],
  ["=ISNUMBER(B2)"],
  ["=ISTEXT(A2)"],
  ["=ISBLANK(E2)"],
  ["=ISERROR(INDEX(A1:A1,5))"],
  ["=AVERAGEIF(C2:C4,\"East\",B2:B4)"],
  ["=AVERAGEIFS(B2:B4,C2:C4,\"East\",B2:B4,\">10\")"],
  ["=XMATCH(\"Beta\",A2:A4)"],
  ["=XMATCH(\"A*\",A2:A4,2)"],
  ["=XMATCH(\"East\",C2:C4,0,-1)"],
  ["=XMATCH(16,B2:B4,-1)"],
  ["=XMATCH(16,B2:B4,1)"],
  ["=XMATCH(\"Missing\",A2:A4)"],
  ["=XMATCH(15,B2:B4,0,2)"],
  ["=XMATCH(15,B2:B4,0,7)"],
];
catalogBook.recalculate();
assert.deepEqual(catalogSheet.getRange("F1:F39").values.flat(), ["ok", 1.23, 2, 30, 15, 20, "Al-100", "Alpha/Beta", 6, "MIXED", true, true, true, 4, 9, 6, "BX-200", 3, 15, 2, 2, 20, 725, "West", "not found", true, true, true, true, 15, 20, 2, 1, 3, 2, 3, "#N/A", 2, "#VALUE!"]);
assert.match(catalogBook.help("fx.XLOOKUP").ndjson, /lookup/);
assert.match(catalogBook.help("fx.INDEX").ndjson, /1-based row/);
assert.match(catalogBook.help("fx.MATCH").ndjson, /1-based position/);
assert.match(catalogBook.help("fx.XMATCH").ndjson, /reverse search/);
assert.match(catalogBook.help("fx.COUNTIFS").ndjson, /multiple criteria/);
assert.match(catalogBook.help("fx.SUMIFS").ndjson, /all supplied criteria/);
assert.match(catalogBook.help("fx.SUMPRODUCT").ndjson, /corresponding numeric values/);
assert.match(catalogBook.help("fx.HLOOKUP").ndjson, /first row/);
assert.match(catalogBook.help("fx.IFERROR").ndjson, /fallback value/);
assert.match(catalogBook.help("fx.ISERROR").ndjson, /recognized formula error/);
assert.match(catalogBook.help("fx.AVERAGEIFS").ndjson, /all supplied criteria/);
assert.match(catalogBook.help("fx.TEXTJOIN").ndjson, /delimiter/);
assert.match(catalogBook.inspect({ kind: "formula", maxChars: 20000 }).ndjson, /XLOOKUP/);

const criteriaBook = Workbook.create();
const criteriaSheet = criteriaBook.worksheets.add("Criteria");
criteriaSheet.getRange("A1:D8").values = [
  ["Alpha", 1, "East", 1],
  ["ALPINE", 2, "EAST", "#N/A"],
  ["Beta", 3, "West", 3],
  ["A*literal", 4, "East", 4],
  ["A?literal", 5, "East", 5],
  ["", 6, "East", 6],
  [10, 7, "East", 7],
  ["10", 8, "East", 8],
];
criteriaSheet.getRange("F1:F15").formulas = [
  ["=COUNTIF(A1:A8,\"alp*\")"],
  ["=COUNTIF(A1:A8,\"A~*literal\")"],
  ["=COUNTIF(A1:A8,\"A~?literal\")"],
  ["=COUNTIF(A1:A8,\"*\")"],
  ["=COUNTIF(A1:A8,\"10\")"],
  ["=SUMIF(A1:A8,\"alp*\",B1:B8)"],
  ["=SUMIFS(B1:B8,A1:A8,\"a*\",C1:C8,\"east\")"],
  ["=AVERAGEIF(A1:A8,\"alp*\",B1:B8)"],
  ["=COUNTIFS(A1:A2,\"*\",C1:C3,\"east\")"],
  ["=SUMIFS(B1:B2,A1:A3,\"*\")"],
  ["=SUMIF(C1:C8,\"east\",D1:D8)"],
  ["=COUNTIFS(A1:B2,\"*\",A1:D1,\"*\")"],
  ["=SUMIFS(D1:D8,C1:C8,\"east\")"],
  ["=AVERAGEIF(C1:C8,\"east\",D1:D8)"],
  ["=AVERAGEIFS(D1:D8,C1:C8,\"east\")"],
];
criteriaBook.recalculate();
assert.deepEqual(criteriaSheet.getRange("F1:F15").values.flat(), [2, 1, 1, 6, 2, 3, 12, 1.5, "#VALUE!", "#VALUE!", "#N/A", "#VALUE!", "#N/A", "#N/A", "#N/A"]);
assert.match(criteriaBook.help("fx.SUMIFS").ndjson, /wildcard/);
const criteriaRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(criteriaBook));
assert.deepEqual(criteriaRoundtrip.worksheets.getItem("Criteria").getRange("F1:F15").values, criteriaSheet.getRange("F1:F15").values);

const dateBook = Workbook.create();
const dateSheet = dateBook.worksheets.add("Dates");
dateSheet.getRange("A1:A2").values = [[45292], [45299]];
dateSheet.getRange("B1:B32").formulas = [
  ["=DATE(1900,2,29)"],
  ["=YEAR(60)"],
  ["=MONTH(60)"],
  ["=DAY(60)"],
  ["=DATE(2024,2,29)"],
  ["=DATE(2024,13,1)"],
  ["=DATE(2024,1,0)"],
  ["=EDATE(DATE(2024,1,31),1)"],
  ["=EOMONTH(DATE(2024,2,10),0)"],
  ["=DAYS(DATE(2024,3,1),DATE(2024,2,28))"],
  ["=WEEKDAY(DATE(2024,1,1),1)"],
  ["=WEEKDAY(DATE(2024,1,1),2)"],
  ["=WEEKDAY(60,1)"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10))"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,10),A1)"],
  ["=NETWORKDAYS(DATE(2024,1,10),DATE(2024,1,1),A1)"],
  ["=WORKDAY(DATE(2024,1,5),1)"],
  ["=WORKDAY(DATE(2024,1,5),1,A2)"],
  ["=WORKDAY(DATE(2024,1,8),-1)"],
  ["=DATE(\"bad\",1,1)"],
  ["=YEAR(-1)"],
  ["=WEEKDAY(DATE(2024,1,1),9)"],
  ["=DATE(0,1,1)"],
  ["=EDATE(60,1)"],
  ["=EDATE(60,-1)"],
  ["=EOMONTH(60,0)"],
  ["=DAYS(61,59)"],
  ["=WEEKDAY(61,1)"],
  ["=WEEKDAY(DATE(2024,1,1),3)"],
  ["=WEEKDAY(DATE(2024,1,1),12)"],
  ["=NETWORKDAYS(DATE(2024,1,6),DATE(2024,1,7))"],
  ["=NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,2),\"bad\")"],
];
dateBook.recalculate();
assert.deepEqual(dateSheet.getRange("B1:B32").values.flat(), [60, 1900, 2, 29, 45351, 45658, 45291, 45351, 45351, 2, 2, 1, 5, 8, 7, -7, 45299, 45300, 45296, "#VALUE!", "#NUM!", "#NUM!", 1, 89, 29, 60, 2, 5, 0, 7, 0, "#VALUE!"]);
assert.match(dateBook.help("fx.DATE").ndjson, /1900 serial-60 compatibility/);
assert.match(dateBook.help("fx.NETWORKDAYS").ndjson, /optional holidays/);
const dateRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(dateBook));
assert.deepEqual(dateRoundtrip.worksheets.getItem("Dates").getRange("B1:B32").values, dateSheet.getRange("B1:B32").values);

const date1904Book = Workbook.create({ dateSystem: "1904" });
assert.equal(date1904Book.dateSystem, "1904");
assert.equal(Workbook.create({ date1904: true }).dateSystem, "1904");
assert.equal(Workbook.create().setDateSystem(true).dateSystem, "1904");
assert.throws(() => Workbook.create({ dateSystem: "unix" }), /expected 1900 or 1904/);
const date1904Sheet = date1904Book.worksheets.add("Dates1904");
date1904Sheet.getRange("A1:A16").formulas = [
  ["=DATE(1904,1,1)"],
  ["=DATE(1904,1,2)"],
  ["=DATE(2024,2,29)"],
  ["=YEAR(0)"],
  ["=MONTH(0)"],
  ["=DAY(0)"],
  ["=EDATE(0,1)"],
  ["=EOMONTH(0,0)"],
  ["=DAYS(DATE(2024,3,1),DATE(2024,2,28))"],
  ["=WEEKDAY(0,1)"],
  ["=WEEKDAY(0,2)"],
  ["=NETWORKDAYS(DATE(1904,1,1),DATE(1904,1,10))"],
  ["=WORKDAY(0,1)"],
  ["=DATE(1900,1,1)"],
  ["=DATE(9999,12,31)"],
  ["=DATE(1904,2,29)"],
];
date1904Book.recalculate();
assert.deepEqual(date1904Sheet.getRange("A1:A16").values.flat(), [0, 1, 43889, 1904, 1, 1, 31, 30, 2, 6, 5, 6, 3, "#NUM!", 2957003, 59]);
const date1904Inspect = date1904Book.inspect({ kind: "workbook" });
assert.match(date1904Inspect.ndjson, /"dateSystem":"1904"/);
assert.match(date1904Inspect.ndjson, /"date1904":true/);
assert.match(date1904Book.help("workbook.setDateSystem").ndjson, /workbookPr/);
const date1904Xlsx = await SpreadsheetFile.exportXlsx(date1904Book);
const date1904Zip = await JSZip.loadAsync(new Uint8Array(await date1904Xlsx.arrayBuffer()));
const date1904WorkbookXml = await date1904Zip.file("xl/workbook.xml").async("text");
assert.match(date1904WorkbookXml, /<workbookPr date1904="1"\/>/);
assert.match(await date1904Zip.file("customXml/open-office-artifact.json").async("text"), /"dateSystem": "1904"/);
const date1904Roundtrip = await SpreadsheetFile.importXlsx(date1904Xlsx);
assert.equal(date1904Roundtrip.dateSystem, "1904");
assert.deepEqual(date1904Roundtrip.worksheets.getItem("Dates1904").getRange("A1:A16").values, date1904Sheet.getRange("A1:A16").values);
const date1904TrueXlsx = await SpreadsheetFile.patchXlsx(date1904Xlsx, [{ path: "xl/workbook.xml", xml: date1904WorkbookXml.replace('date1904="1"', 'date1904="true"') }]);
assert.equal((await SpreadsheetFile.importXlsx(date1904TrueXlsx)).dateSystem, "1904");
const date1900PatchedXlsx = await SpreadsheetFile.patchXlsx(date1904Xlsx, [{ path: "xl/workbook.xml", xml: date1904WorkbookXml.replace('date1904="1"', 'date1904="0"') }]);
const date1900PatchedBook = await SpreadsheetFile.importXlsx(date1900PatchedXlsx);
assert.equal(date1900PatchedBook.dateSystem, "1900");
assert.equal(date1900PatchedBook.worksheets.getItem("Dates1904").getRange("A3").values[0][0], 45351);
const invalidDateSystemBook = Workbook.create();
invalidDateSystemBook.worksheets.add("Invalid").getRange("A1").values = [[1]];
invalidDateSystemBook.dateSystem = "invalid";
assert.ok(invalidDateSystemBook.verify().issues.some((issue) => issue.type === "invalidDateSystem"));
await assert.rejects(() => SpreadsheetFile.exportXlsx(invalidDateSystemBook), /expected 1900 or 1904/);

const intlDateBook = Workbook.create();
const intlDateSheet = intlDateBook.worksheets.add("IntlDates");
intlDateSheet.getRange("A1:A2").values = [[45293], [-1]];
intlDateSheet.getRange("B1:B26").formulas = [
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7))"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),2)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),11)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0000011\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0010001\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"1111111\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,7),DATE(2024,1,1),1)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1,A1)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"bad\")"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),0)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1.5)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),1,A2)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),5)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,2)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),3,\"0010001\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,1,A1)"],
  ["=WORKDAY.INTL(DATE(2024,1,8),-5)"],
  ["=WORKDAY.INTL(DATE(2024,1,6),0)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,\"1111111\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,8)"],
  ["=WORKDAY.INTL(DATE(2024,1,1),1,1,A2)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),17)"],
  ["=NETWORKDAYS.INTL(DATE(2006,1,1),DATE(2006,1,31))"],
  ["=WORKDAY.INTL(DATE(2012,1,1),30,17)"],
  ["=NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,7),\"0000000\")"],
  ["=WORKDAY.INTL(DATE(2024,1,1),7,\"0000000\")"],
];
intlDateBook.recalculate();
assert.deepEqual(intlDateSheet.getRange("B1:B26").values.flat(), [5, 5, 6, 5, 5, 0, -5, 4, "#VALUE!", "#NUM!", "#NUM!", "#NUM!", 45299, 45293, 45296, 45294, 45292, 45297, "#VALUE!", "#NUM!", "#NUM!", 6, 22, 40944, 7, 45299]);
assert.match(intlDateBook.help("fx.NETWORKDAYS.INTL").ndjson, /seven-character custom weekend/);
assert.match(intlDateBook.help("fx.WORKDAY.INTL").ndjson, /numbered or Monday-first/);

const intl1904Book = Workbook.create({ dateSystem: "1904" });
const intl1904Sheet = intl1904Book.worksheets.add("Intl1904");
intl1904Sheet.getRange("A1:A3").formulas = [
  ["=NETWORKDAYS.INTL(0,6,7)"],
  ["=WORKDAY.INTL(0,1,7)"],
  ["=WORKDAY.INTL(0,1,11)"],
];
intl1904Book.recalculate();
assert.deepEqual(intl1904Sheet.getRange("A1:A3").values.flat(), [5, 2, 1]);
const intl1904Roundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(intl1904Book));
assert.equal(intl1904Roundtrip.dateSystem, "1904");
assert.deepEqual(intl1904Roundtrip.worksheets.getItem("Intl1904").getRange("A1:A3").values, [[5], [2], [1]]);

const formulaEdgeBook = Workbook.create();
const formulaEdgeSheet = formulaEdgeBook.worksheets.add("FormulaEdges");
formulaEdgeSheet.getRange("A1:B3").values = [[1, 10], [2, 20], [3, 30]];
formulaEdgeSheet.getRange("D1:D15").formulas = [
  ["=SUMPRODUCT(A1:B2,A1:A4)"],
  ["=AVERAGEIF(A1:A3,\">10\",B1:B3)"],
  ["=IFERROR(AVERAGEIF(A1:A3,\">10\",B1:B3),\"empty\")"],
  ["=HLOOKUP(1,A1:B3,5,FALSE)"],
  ["=ISERROR(#N/A)"],
  ["=ISTEXT(\"\")"],
  ["=ISBLANK(\"\")"],
  ["=NOT(TRUE)"],
  ["=NA()"],
  ["=IFNA(NA(),\"missing\")"],
  ["=IFNA(#VALUE!,\"missing\")"],
  ["=ISNA(NA())"],
  ["=ISNA(#REF!)"],
  ["=ISERR(#REF!)"],
  ["=ISERR(NA())"],
];
formulaEdgeBook.recalculate();
assert.deepEqual(formulaEdgeSheet.getRange("D1:D15").values.flat(), ["#VALUE!", "#DIV/0!", "empty", "#REF!", true, true, false, false, "#N/A", "missing", "#VALUE!", true, false, true, false]);

const statisticsBook = Workbook.create();
const statisticsSheet = statisticsBook.worksheets.add("Statistics");
statisticsSheet.getRange("A1:A6").values = [[10], [20], [20], [40], ["text"], [null]];
statisticsSheet.getRange("D1:D15").formulas = [
  ["=MEDIAN(A1:A6)"],
  ["=MODE.SNGL(A1:A6)"],
  ["=LARGE(A1:A6,2)"],
  ["=SMALL(A1:A6,2)"],
  ["=RANK.EQ(20,A1:A6)"],
  ["=RANK.EQ(20,A1:A6,1)"],
  ["=ROUND(1250,-2)"],
  ["=ROUND(-1250,-2)"],
  ["=ROUNDUP(12.341,2)"],
  ["=ROUNDUP(-1250,-2)"],
  ["=ROUNDDOWN(12.349,2)"],
  ["=ROUNDDOWN(-1250,-2)"],
  ["=LARGE(A1:A6,9)"],
  ["=MODE.SNGL(1,2,3)"],
  ["=MEDIAN(A1:A6,#REF!)"],
];
statisticsBook.recalculate();
assert.deepEqual(statisticsSheet.getRange("D1:D15").values.flat(), [20, 20, 20, 20, 2, 2, 1300, -1300, 12.35, -1300, 12.34, -1200, "#NUM!", "#N/A", "#REF!"]);
statisticsSheet.getRange("E1:E2").formulas = [["=ROUND(1.005,2)"], ["=MODE.SNGL(3,3,2,2)"]];
statisticsBook.recalculate();
assert.deepEqual(statisticsSheet.getRange("E1:E2").values.flat(), [1.01, 2]);
assert.match(statisticsBook.help("fx.MEDIAN").ndjson, /middle numeric value/);
assert.match(statisticsBook.help("fx.MODE.SNGL").ndjson, /most frequently/);
assert.match(statisticsBook.help("fx.RANK.EQ").ndjson, /descending by default/);
assert.match(statisticsBook.help("fx.ROUNDUP").ndjson, /away from zero/);
assert.match(statisticsBook.inspect({ kind: "formula", target: "Statistics!D1", maxChars: 4000 }).ndjson, /MEDIAN/);

const spillBook = Workbook.create();
const spillSheet = spillBook.worksheets.add("Spill");
spillSheet.getRange("A1").formulas = [["=SEQUENCE(2,3,10,2)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("A1:C2").values, [[10, 12, 14], [16, 18, 20]]);
const spillInspect = spillBook.inspect({ kind: "formula", target: "Spill!A1", maxChars: 8000 }).ndjson;
assert.match(spillInspect, /"spillRange":"A1:C2"/);
assert.match(spillInspect, /"spillValues":\[\[10,12,14\],\[16,18,20\]\]/);
const spillLayout = spillBook.layoutJson({ sheetName: "Spill", target: "Spill!B2" });
assert.equal(spillLayout.sheets[0].cells[0].spillParent, "Spill!A1");
spillSheet.getRange("E1").formulas = [["=TRANSPOSE(A1:C2)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("E1:F3").values, [[10, 16], [12, 18], [14, 20]]);
spillSheet.getRange("H1:J4").values = [["Item", "Region", "Score"], ["Alpha", "East", 10], ["Beta", "West", 15], ["Gamma", "East", 20]];
spillSheet.getRange("L1").formulas = [["=FILTER(H2:J4,I2:I4=\"East\")"]];
spillSheet.getRange("P1").formulas = [["=UNIQUE(I2:I4)"]];
spillSheet.getRange("R1").formulas = [["=SORT(H2:J4,3,-1)"]];
spillSheet.getRange("V1").formulas = [["=TAKE(H2:J4,2,-2)"]];
spillSheet.getRange("Y1").formulas = [["=DROP(H2:J4,1,1)"]];
spillSheet.getRange("AB1").formulas = [["=CHOOSECOLS(H2:J4,3,1)"]];
spillSheet.getRange("AE1").formulas = [["=CHOOSEROWS(H2:J4,3,1)"]];
spillSheet.getRange("AI1").formulas = [["=TAKE(H2:J4,,2)"]];
spillSheet.getRange("AL1").formulas = [["=DROP(H2:J4,,1)"]];
spillSheet.getRange("AO1:AP1").formulas = [["=TAKE(H2:J4,0)", "=CHOOSECOLS(H2:J4,0)"]];
spillSheet.getRange("AR1").formulas = [["=TOCOL(H2:I3)"]];
spillSheet.getRange("AT1").formulas = [["=TOROW(H2:I3,0,TRUE)"]];
spillSheet.getRange("AY1").formulas = [["=WRAPROWS(H2:H4,2)"]];
spillSheet.getRange("BB1").formulas = [["=WRAPCOLS(H2:H4,2,\"pad\")"]];
spillSheet.getRange("BE1:BF1").formulas = [["=WRAPROWS(H2:I3,2)", "=WRAPCOLS(H2:H4,0)"]];
spillSheet.getRange("H6:I7").values = [[1, null], ["#N/A", 4]];
spillSheet.getRange("BH1").formulas = [["=TOCOL(H6:I7,3)"]];
spillSheet.getRange("BJ1").formulas = [["=HSTACK(H2:H4,I2:I3)"]];
spillSheet.getRange("BM1").formulas = [["=VSTACK(H2:I3,H4:H4)"]];
spillSheet.getRange("BQ1").formulas = [["=EXPAND(H2:I3,3,3,\"pad\")"]];
spillSheet.getRange("BU1").formulas = [["=EXPAND(H2:I3,1,1)"]];
spillSheet.getRange("BW1").formulas = [["=EXPAND(H2:I3,,3)"]];
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("L1:N2").values, [["Alpha", "East", 10], ["Gamma", "East", 20]]);
assert.deepEqual(spillSheet.getRange("P1:P2").values, [["East"], ["West"]]);
assert.deepEqual(spillSheet.getRange("R1:T3").values, [["Gamma", "East", 20], ["Beta", "West", 15], ["Alpha", "East", 10]]);
assert.deepEqual(spillSheet.getRange("V1:W2").values, [["East", 10], ["West", 15]]);
assert.deepEqual(spillSheet.getRange("Y1:Z2").values, [["West", 15], ["East", 20]]);
assert.deepEqual(spillSheet.getRange("AB1:AC3").values, [[10, "Alpha"], [15, "Beta"], [20, "Gamma"]]);
assert.deepEqual(spillSheet.getRange("AE1:AG2").values, [["Gamma", "East", 20], ["Alpha", "East", 10]]);
assert.deepEqual(spillSheet.getRange("AI1:AJ3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "East"]]);
assert.deepEqual(spillSheet.getRange("AL1:AM3").values, [["East", 10], ["West", 15], ["East", 20]]);
assert.deepEqual(spillSheet.getRange("AO1:AP1").values, [["#CALC!", "#VALUE!"]]);
assert.deepEqual(spillSheet.getRange("AR1:AR4").values, [["Alpha"], ["East"], ["Beta"], ["West"]]);
assert.deepEqual(spillSheet.getRange("AT1:AW1").values, [["Alpha", "Beta", "East", "West"]]);
assert.deepEqual(spillSheet.getRange("AY1:AZ2").values, [["Alpha", "Beta"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BB1:BC2").values, [["Alpha", "Gamma"], ["Beta", "pad"]]);
assert.deepEqual(spillSheet.getRange("BE1:BF1").values, [["#VALUE!", "#NUM!"]]);
assert.deepEqual(spillSheet.getRange("BH1:BH2").values, [[1], [4]]);
assert.deepEqual(spillSheet.getRange("BJ1:BK3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BM1:BN3").values, [["Alpha", "East"], ["Beta", "West"], ["Gamma", "#N/A"]]);
assert.deepEqual(spillSheet.getRange("BQ1:BS3").values, [["Alpha", "East", "pad"], ["Beta", "West", "pad"], ["pad", "pad", "pad"]]);
assert.equal(spillSheet.getRange("BU1").values[0][0], "#VALUE!");
assert.deepEqual(spillSheet.getRange("BW1:BY2").values, [["Alpha", "East", "#N/A"], ["Beta", "West", "#N/A"]]);
assert.match(spillBook.inspect({ kind: "formula", target: "Spill!L1", maxChars: 8000 }).ndjson, /"spillRange":"L1:N2"/);
assert.match(spillBook.help("fx.SEQUENCE").ndjson, /dynamic array/);
assert.match(spillBook.help("fx.TRANSPOSE").ndjson, /spillRange/);
assert.match(spillBook.help("fx.FILTER").ndjson, /include array/);
assert.match(spillBook.help("fx.UNIQUE").ndjson, /unique rows/);
assert.match(spillBook.help("fx.SORT").ndjson, /column index/);
assert.match(spillBook.help("fx.TAKE").ndjson, /start or end/);
assert.match(spillBook.help("fx.DROP").ndjson, /spill the remainder/);
assert.match(spillBook.help("fx.CHOOSECOLS").ndjson, /Select and reorder/);
assert.match(spillBook.help("fx.CHOOSEROWS").ndjson, /Select and reorder/);
assert.match(spillBook.help("fx.TOCOL").ndjson, /one spilled column/);
assert.match(spillBook.help("fx.TOROW").ndjson, /one spilled row/);
assert.match(spillBook.help("fx.WRAPROWS").ndjson, /requested width/);
assert.match(spillBook.help("fx.WRAPCOLS").ndjson, /requested height/);
assert.match(spillBook.help("fx.HSTACK").ndjson, /horizontally/);
assert.match(spillBook.help("fx.VSTACK").ndjson, /vertically/);
assert.match(spillBook.help("fx.EXPAND").ndjson, /requested row/);
const spillXlsx = await SpreadsheetFile.exportXlsx(spillBook);
const spillZip = await JSZip.loadAsync(new Uint8Array(await spillXlsx.arrayBuffer()));
const spillWorksheetXml = await spillZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(spillWorksheetXml, /<f t="array" ref="A1:C2">SEQUENCE\(2,3,10,2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="E1:F3">TRANSPOSE\(A1:C2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="L1:N2">FILTER\(H2:J4,I2:I4="East"\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="V1:W2">TAKE\(H2:J4,2,-2\)<\/f>/);
const blockedSpillBook = Workbook.create();
const blockedSheet = blockedSpillBook.worksheets.add("Blocked");
blockedSheet.getRange("C1").values = [["blocked"]];
blockedSheet.getRange("A1").formulas = [["=SEQUENCE(1,3)"]];
blockedSpillBook.recalculate();
assert.equal(blockedSheet.getRange("A1").values[0][0], "#SPILL!");
assert.match(blockedSpillBook.verify().ndjson, /formulaError/);

const cycleBook = Workbook.create();
const cycleSheet = cycleBook.worksheets.add("Cycle");
cycleSheet.getRange("A1").formulas = [["=B1+1"]];
cycleSheet.getRange("B1").formulas = [["=A1+1"]];
cycleBook.recalculate();
assert.equal(cycleSheet.getRange("A1").values[0][0], "#CYCLE!");
const cycleGraph = cycleBook.formulaGraph();
assert.equal(cycleGraph.cycles.length, 1);
assert.match(cycleBook.inspect({ kind: "formulaGraph,formulaCycle" }).ndjson, /formulaCycle/);
assert.match(cycleBook.verify().ndjson, /formulaCycle/);

const missingSheetBook = Workbook.create();
missingSheetBook.worksheets.add("Main").getRange("A1").formulas = [["=Missing!A1+1"]];
assert.match(missingSheetBook.verify().ndjson, /missingFormulaSheet/);

const preview = await workbook.render({ sheetName: "Sheet1", range: "A1:C3" });
assert.equal(preview.type, "image/svg+xml");
const previewSvg = await preview.text();
assert.match(previewSvg, /<svg/);
assert.match(previewSvg, /TasksTable/);
assert.match(previewSvg, /Revenue Trend/);
assert.match(previewSvg, /Logo placeholder/);
assert.match(previewSvg, /RevenuePivot/);
assert.match(previewSvg, /Revenue sum/);
assert.match(previewSvg, /polyline/);
const layoutBlob = await workbook.render({ format: "layout", sheetName: "Sheet1", range: "A1:C3" });
assert.equal(layoutBlob.type, "application/vnd.open-office-artifact.layout+json");
const layout = JSON.parse(await layoutBlob.text());
assert.equal(layout.kind, "workbookLayout");
assert.equal(layout.sheets[0].name, "Sheet1");
assert.equal(layout.sheets[0].bounds.address, "A1:C3");
assert.ok(layout.sheets[0].cells.some((cell) => cell.address === "C2" && cell.formula === "=A2+B2"));
const c3LayoutCell = layout.sheets[0].cells.find((cell) => cell.address === "C3");
assert.equal(c3LayoutCell.computedStyle.fill, "green");
assert.deepEqual(c3LayoutCell.conditionalFormats.map((item) => item.id), [cf.id]);
const a2LayoutCell = layout.sheets[0].cells.find((cell) => cell.address === "A2");
assert.equal(a2LayoutCell.computedStyle.fill, "sky-100");
assert.deepEqual(a2LayoutCell.conditionalFormats.map((item) => item.id), [customCf.id]);
assert.match(previewSvg, /fill="#22c55e"/);
const scalePreviewSvg = await (await workbook.render({ sheetName: "Sheet1", range: "G2:G4" })).text();
assert.match(scalePreviewSvg, /fill="#22c55e"/);
const scaleLayout = workbook.layoutJson({ sheetName: "Sheet1", range: "G2:G4", target: "Sheet1!G4" });
assert.equal(scaleLayout.sheets[0].cells[0].computedStyle.fill, "#22c55e");
assert.ok(layout.sheets[0].tables.some((table) => table.name === "TasksTable"));
assert.ok(layout.sheets[0].pivots.some((pivot) => pivot.name === "RevenuePivot" && pivot.values.some((row) => row.includes("Revenue sum"))));
assert.ok(layout.sheets[0].charts.some((chart) => chart.title === "Revenue Trend"));
assert.ok(layout.sheets[0].images.some((item) => item.alt === "Logo placeholder"));
assert.ok(layout.sheets[0].sparklines.some((item) => item.targetRange === "H2:H2"));
const cellLayoutBlob = await workbook.render({ format: "layout", sheetName: "Sheet1", range: "A1:C3", target: "Sheet1!C2" });
assert.equal(cellLayoutBlob.metadata.target, "Sheet1!C2");
const cellLayout = JSON.parse(await cellLayoutBlob.text());
assert.deepEqual(cellLayout.sheets[0].cells.map((cell) => cell.address), ["C2"]);
assert.equal(cellLayout.sheets[0].tables.length, 0);
const imageLayout = workbook.layoutJson({ sheetName: "Sheet1", range: "A1:C3", target: image.id });
assert.deepEqual(imageLayout.sheets[0].images.map((item) => item.id), [image.id]);
assert.equal(imageLayout.sheets[0].cells.length, 0);
const chartSearchLayout = workbook.layoutJson({ sheetName: "Sheet1", search: "Revenue Trend" });
assert.deepEqual(chartSearchLayout.sheets[0].charts.map((item) => item.id), [chartFromRange.id]);
const workbookContextLayout = workbook.layoutJson({ sheetName: "Sheet1", target: image.id, before: 1 });
assert.deepEqual(workbookContextLayout.sheets[0].charts.map((item) => item.id), [chartFromConfig.id]);
assert.deepEqual(workbookContextLayout.sheets[0].images.map((item) => item.id), [image.id]);
const sheetTargetLayout = workbook.layoutJson({ target: "Sheet1" });
assert.equal(sheetTargetLayout.sheets.length, 1);
assert.ok(sheetTargetLayout.sheets[0].cells.length > 0);
assert.match(workbook.help("workbook.layoutJson").ndjson, /target\/search context slicing/);

const delimitedBook = Workbook.create();
const delimitedSheet = delimitedBook.worksheets.add("Data");
delimitedSheet.getRange("A1:D3").values = [
  ["Region", "Amount", "Note", "Detail"],
  ["North, East", 42, 'He said "ok"', "line one\nline two"],
  ["=literal", null, "plain", "tab\there"],
];
delimitedSheet.getRange("B3").formulas = [["=B2*2"]];
delimitedBook.recalculate();
const csv = await SpreadsheetFile.exportCsv(delimitedBook, { sheetName: "Data" });
assert.equal(csv.type, "text/csv");
assert.equal(csv.metadata.rows, 3);
assert.equal(csv.metadata.columns, 4);
assert.match(await csv.text(), /"North, East",42,"He said ""ok""","line one\nline two"/);
assert.doesNotMatch(await csv.text(), /=B2\*2/);
const csvInspect = await SpreadsheetFile.inspectDelimited(csv);
assert.equal(csvInspect.summary.rows, 3);
assert.equal(csvInspect.summary.columns, 4);
assert.equal(csvInspect.summary.quotedCells, 3);
assert.equal(csvInspect.summary.formulaLikeCells, 1);
const csvImported = await SpreadsheetFile.importCsv(csv, { sheetName: "Imported", coerceTypes: true });
assert.deepEqual(csvImported.worksheets.getItem("Imported").getRange("A2:D3").values, [
  ["North, East", 42, 'He said "ok"', "line one\nline two"],
  ["=literal", 84, "plain", "tab\there"],
]);
const formulaCsv = await SpreadsheetFile.exportCsv(delimitedBook, { sheetName: "Data", range: "A2:B3", formulas: true, includeBom: true, lineEnding: "\n" });
assert.equal(formulaCsv.bytes[0], 0xef);
assert.match(await formulaCsv.text(), /=B2\*2/);
const tsv = await SpreadsheetFile.exportTsv(delimitedBook, { sheetName: "Data", includeBom: true });
assert.equal(tsv.type, "text/tab-separated-values");
assert.equal((await SpreadsheetFile.inspectDelimited(tsv)).summary.hasBom, true);
assert.equal((await SpreadsheetFile.inspectDelimited(tsv)).summary.columns, 4);
const tsvImported = await SpreadsheetFile.importTsv(tsv, { coerceTypes: true });
assert.equal(tsvImported.worksheets.getItemAt(0).getRange("D3").values[0][0], "tab\there");
await assert.rejects(() => SpreadsheetFile.importCsv('a,"unterminated'), /unterminated quoted field/);
await assert.rejects(() => SpreadsheetFile.importCsv('a,"quoted"junk'), /unexpected content after a closing quote/);
await assert.rejects(() => SpreadsheetFile.inspectDelimited(csv, { maxBytes: 4 }), /exceeds maxBytes/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxBytes: 4 }), /exceeds maxBytes/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxRows: 2 }), /exceeds maxRows/);
await assert.rejects(() => SpreadsheetFile.exportCsv(delimitedBook, { maxColumns: 3 }), /exceeds maxColumns/);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(xlsx.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
const xlsxInspect = await SpreadsheetFile.inspectXlsx(xlsx, { includeText: true, maxChars: 16000 });
assert.equal(xlsxInspect.records[0].kind, "xlsxPackage");
assert.equal(xlsxInspect.ok, true);
assert.deepEqual(xlsxInspect.issues, []);
assert.equal(xlsxInspect.records[0].sheets, 1);
assert.ok(xlsxInspect.records[0].uncompressedBytes > 0);
assert.ok(xlsxInspect.records[0].relationshipReferences > 0);
assert.equal(xlsxInspect.records[0].relationshipReferenceIssues, 0);
assert.ok(xlsxInspect.parts.some((part) => part.path === "xl/workbook.xml" && part.contentType.includes("spreadsheetml.sheet.main+xml")));
const xlsxReferenceZip = await JSZip.loadAsync(new Uint8Array(await xlsx.arrayBuffer()));
const xlsxSheetXml = await xlsxReferenceZip.file("xl/worksheets/sheet1.xml").async("text");
const singleQuotedContentTypesXml = (await xlsxReferenceZip.file("[Content_Types].xml").async("text")).replace(/="([^"]*)"/g, "='$1'");
const singleQuotedWorkbookRelsXml = (await xlsxReferenceZip.file("xl/_rels/workbook.xml.rels").async("text")).replace(/="([^"]*)"/g, "='$1'");
const singleQuotedXlsx = await SpreadsheetFile.patchXlsx(xlsx, [
  { path: "[Content_Types].xml", xml: singleQuotedContentTypesXml },
  { path: "xl/_rels/workbook.xml.rels", xml: singleQuotedWorkbookRelsXml },
]);
assert.equal((await SpreadsheetFile.inspectXlsx(singleQuotedXlsx)).ok, true);
const brokenXlsxReferenceXml = xlsxSheetXml.replace(/<\/worksheet>\s*$/, '<drawing xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdMissingSourceReference"/></worksheet>');
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: brokenXlsxReferenceXml }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const invalidReferenceXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: brokenXlsxReferenceXml }], { validateResult: false });
const invalidReferenceXlsxInspect = await SpreadsheetFile.inspectXlsx(invalidReferenceXlsx);
assert.equal(invalidReferenceXlsxInspect.ok, false);
assert.ok(invalidReferenceXlsxInspect.issues.some((issue) => issue.type === "relationshipReferenceIdNotFound" && issue.path === "xl/worksheets/sheet1.xml" && issue.referenceAttribute === "r:id"));
assert.ok(xlsxReferenceZip.file("xl/media/image1.png"));
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/media/image1.png", remove: true }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const patchedXlsx = await SpreadsheetFile.patchXlsx(xlsx, { "customXml/review.json": { json: { status: "ok" } } });
assert.equal(patchedXlsx.type, xlsx.type);
assert.equal(patchedXlsx.metadata.patchedParts, 1);
assert.equal(patchedXlsx.metadata.validated, true);
assert.equal(patchedXlsx.metadata.validationIssues, 0);
assert.match((await SpreadsheetFile.inspectXlsx(patchedXlsx, { includeText: true, maxChars: 16000 })).ndjson, /review\.json/);
const recipeImageBytes = Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=", "base64");
const recipeImageXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{
  path: "xl/media/review.png",
  bytes: recipeImageBytes,
  recipe: { kind: "image", source: "xl/worksheets/sheet1.xml", id: "rIdReviewImage" },
}]);
assert.equal(recipeImageXlsx.metadata.recipesApplied, 1);
const recipeImageInspect = await SpreadsheetFile.inspectXlsx(recipeImageXlsx);
assert.equal(recipeImageInspect.ok, true);
assert.ok(recipeImageInspect.parts.some((part) => part.path === "xl/media/review.png" && part.contentType === "image/png"));
const recipeImageZip = await JSZip.loadAsync(new Uint8Array(await recipeImageXlsx.arrayBuffer()));
assert.match(await recipeImageZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text"), /Id="rIdReviewImage"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image"[^>]*Target="\.\.\/media\/review\.png"/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/media/unsupported.png", bytes: recipeImageBytes, recipe: { kind: "image", source: "xl/worksheets/sheet1.xml", sourceReference: true } }]), /root element wsDr/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/collision.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", id: "rId1", sourceReference: { name: "Collision" } } }]), /relationship Id rId1.*already targets/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/duplicate-id.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Duplicate Id", sheetId: 1 } } }]), /sheetId 1 already exists/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/duplicate-name.xml", xml: '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Sheet1" } } }]), /name Sheet1 already exists/);
const recipeTableXml = '<?xml version="1.0" encoding="UTF-8"?><table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="99" name="ReviewTable" displayName="ReviewTable" ref="A1:B2" totalsRowShown="0"><autoFilter ref="A1:B2"/><tableColumns count="2"><tableColumn id="1" name="A"/><tableColumn id="2" name="B"/></tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';
const recipeTableXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/tables/review.xml", xml: recipeTableXml, recipe: { kind: "table", source: "xl/worksheets/sheet1.xml", id: "rIdReviewTable", sourceReference: true } }]);
assert.equal(recipeTableXlsx.metadata.sourceReferencesUpdated, 1);
const recipeTableZip = await JSZip.loadAsync(new Uint8Array(await recipeTableXlsx.arrayBuffer()));
assert.match(await recipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /<tableParts count="2">[\s\S]*<tablePart r:id="rIdReviewTable"\/>/);
const removedRecipeTableXlsx = await SpreadsheetFile.patchXlsx(recipeTableXlsx, [{ path: "xl/tables/review.xml", remove: true, recipe: { kind: "table", source: "xl/worksheets/sheet1.xml", id: "rIdReviewTable", sourceReference: true } }]);
assert.equal((await SpreadsheetFile.inspectXlsx(removedRecipeTableXlsx)).ok, true);
const removedRecipeTableZip = await JSZip.loadAsync(new Uint8Array(await removedRecipeTableXlsx.arrayBuffer()));
assert.match(await removedRecipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /<tableParts count="1">/);
assert.doesNotMatch(await removedRecipeTableZip.file("xl/worksheets/sheet1.xml").async("text"), /rIdReviewTable/);
const recipeWorksheetXml = '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>';
const recipeWorksheetXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheetReview.xml", xml: recipeWorksheetXml, recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Review Data" } } }]);
assert.equal(recipeWorksheetXlsx.metadata.sourceReferencesUpdated, 1);
const recipeWorksheetZip = await JSZip.loadAsync(new Uint8Array(await recipeWorksheetXlsx.arrayBuffer()));
const recipeWorksheetRels = await recipeWorksheetZip.file("xl/_rels/workbook.xml.rels").async("text");
const recipeWorksheetRelId = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="worksheets\/sheetReview\.xml"/.exec(recipeWorksheetRels)?.[1];
assert.ok(recipeWorksheetRelId);
assert.match(await recipeWorksheetZip.file("xl/workbook.xml").async("text"), new RegExp(`<sheet name="Review Data" sheetId="2" r:id="${recipeWorksheetRelId}"/>`));
assert.ok((await SpreadsheetFile.importXlsx(recipeWorksheetXlsx)).worksheets.getItem("Review Data"));
const removedRecipeWorksheetXlsx = await SpreadsheetFile.patchXlsx(recipeWorksheetXlsx, [{ path: "xl/worksheets/sheetReview.xml", remove: true, recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Review Data" } } }]);
assert.equal((await SpreadsheetFile.inspectXlsx(removedRecipeWorksheetXlsx)).ok, true);
assert.equal((await SpreadsheetFile.importXlsx(removedRecipeWorksheetXlsx)).worksheets.getItem("Review Data"), undefined);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await removedRecipeWorksheetXlsx.arrayBuffer()))).file("xl/workbook.xml").async("text"), /Review Data/);
const replacementXlsx = await SpreadsheetFile.patchXlsx(xlsx, { "customXml/open-office-artifact.json": { json: { replaced: true } } }, { maxParts: xlsxInspect.parts.length });
assert.equal(replacementXlsx.metadata.patchedParts, 1);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/new-part.json", json: {} }], { maxParts: xlsxInspect.parts.length }), /would create .* maxParts/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "../evil.xml", text: "bad" }]), /Unsafe XLSX part path/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/unknown.xml", xml: "<x/>", recipe: "not-a-real-part" }]), /OOXML part recipe.*unsupported/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "customXml/large.txt", text: "12345" }], { maxPatchBytes: 4 }), /exceeds maxPatchBytes/);
await assert.rejects(() => SpreadsheetFile.inspectXlsx(xlsx, { maxParts: 1 }), /maxParts/);
await assert.rejects(() => SpreadsheetFile.inspectXlsx(xlsx, { maxPartBytes: 1 }), /maxPartBytes/);
const safelyRemovedStylesXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }]);
assert.equal(safelyRemovedStylesXlsx.metadata.contentTypesUpdated, 1);
assert.equal(safelyRemovedStylesXlsx.metadata.relationshipsUpdated, 1);
assert.equal((await SpreadsheetFile.inspectXlsx(safelyRemovedStylesXlsx)).ok, true);
await assert.rejects(() => SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }], { syncRelationships: false }), /invalid OOXML package.*relationshipTargetNotFound/);
const brokenRelationshipXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/styles.xml", remove: true }], { syncRelationships: false, validateResult: false });
assert.equal(brokenRelationshipXlsx.metadata.validated, false);
const brokenRelationshipInspect = await SpreadsheetFile.inspectXlsx(brokenRelationshipXlsx);
assert.equal(brokenRelationshipInspect.ok, false);
assert.ok(brokenRelationshipInspect.issues.some((issue) => issue.type === "relationshipTargetNotFound" && issue.target === "xl/styles.xml"));
const brokenRelationshipZip = await JSZip.loadAsync(new Uint8Array(await safelyRemovedStylesXlsx.arrayBuffer()));
assert.doesNotMatch(await brokenRelationshipZip.file("[Content_Types].xml").async("text"), /PartName="\/xl\/styles\.xml"/);
const xlsxBytes = new Uint8Array(await xlsx.arrayBuffer());
const zip = await JSZip.loadAsync(xlsxBytes);
const tablePartNames = Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name));
assert.equal(tablePartNames.length, 1);
const tableXml = await zip.file(tablePartNames[0]).async("text");
assert.match(tableXml, /displayName="TasksTable"/);
assert.match(tableXml, /ref="A1:D4"/);
assert.match(tableXml, /<tableColumns count="4">/);
const pivotPartNames = Object.keys(zip.files).filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name));
assert.equal(pivotPartNames.length, 2);
const pivotTableXml = await zip.file(pivotPartNames[0]).async("text");
assert.match(pivotTableXml, /<pivotTableDefinition/);
assert.match(pivotTableXml, /name="RevenuePivot"/);
assert.match(pivotTableXml, /cacheId="1"/);
assert.match(pivotTableXml, /<rowFields count="1"><field x="0"\/><\/rowFields>/);
assert.match(pivotTableXml, /<dataField name="Revenue sum" fld="1" subtotal="sum"\/>/);
const regionalPivotTableXml = await zip.file(pivotPartNames[1]).async("text");
assert.match(regionalPivotTableXml, /name="RegionalPivot"/);
assert.match(regionalPivotTableXml, /<colFields count="1"><field x="1"\/><\/colFields>/);
assert.match(regionalPivotTableXml, /<pivotField axis="axisCol" multipleItemSelectionAllowed="1" showAll="0"><items count="3"><item x="0"\/><item x="1" h="1"\/><item t="default"\/><\/items><\/pivotField>/);
assert.match(regionalPivotTableXml, /<dataField name="Profit total" fld="5" subtotal="sum"\/>/);
const pivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
assert.match(pivotCacheXml, /<pivotCacheDefinition/);
assert.match(pivotCacheXml, /r:id="rId1"/);
assert.match(pivotCacheXml, /<worksheetSource ref="F1:G4" sheet="Sheet1"\/>/);
assert.match(pivotCacheXml, /<cacheField name="Month"/);
assert.match(pivotCacheXml, /<cacheField name="Revenue"/);
const pivotCacheRecordsXml = await zip.file("xl/pivotCache/pivotCacheRecords1.xml").async("text");
assert.match(pivotCacheRecordsXml, /<pivotCacheRecords[^>]*count="3"/);
assert.match(pivotCacheRecordsXml, /<s v="Jan"\/><n v="100"\/>/);
assert.match(pivotCacheRecordsXml, /<s v="Mar"\/><n v="130"\/>/);
const pivotCacheDefinitionRelsXml = await zip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels").async("text");
assert.match(pivotCacheDefinitionRelsXml, /relationships\/pivotCacheRecords/);
assert.match(pivotCacheDefinitionRelsXml, /Target="pivotCacheRecords1\.xml"/);
const regionalPivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition2.xml").async("text");
assert.match(regionalPivotCacheXml, /refreshOnLoad="0" saveData="1" enableRefresh="0" invalid="1" missingItemsLimit="3" refreshedBy="QA Agent" refreshedDateIso="2026-07-12T00:00:00Z"/);
assert.match(regionalPivotCacheXml, /<cacheFields count="6">/);
assert.match(regionalPivotCacheXml, /<cacheField name="Profit" formula="\('Revenue'-'Cost'\)\*100%" databaseField="0" numFmtId="0"><sharedItems containsNumber="1" count="0"\/><\/cacheField>/);
const pivotTableDefinitionRelsXml = await zip.file("xl/pivotTables/_rels/pivotTable1.xml.rels").async("text");
assert.match(pivotTableDefinitionRelsXml, /relationships\/pivotCacheDefinition/);
assert.match(pivotTableDefinitionRelsXml, /Target="\.\.\/pivotCache\/pivotCacheDefinition1\.xml"/);
const worksheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(worksheetXml, /<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"\/><\/sheetView><\/sheetViews><sheetData>/);
assert.match(worksheetXml, /<conditionalFormatting sqref="C2:C3">/);
assert.match(worksheetXml, /<cfRule type="cellIs"[^>]*dxfId="0"/);
assert.match(worksheetXml, /<conditionalFormatting sqref="G2:G4"><cfRule type="colorScale"/);
assert.match(worksheetXml, /<color rgb="FF22C55E"\/>/);
assert.match(worksheetXml, /<cfvo type="percentile" val="50"\/>/);
assert.match(worksheetXml, /<dataValidations count="1">/);
assert.match(worksheetXml, /<dataValidation type="list"[^>]*sqref="D2:D3"/);
assert.match(worksheetXml, /<c r="A1" t="s" s="\d+"><v>\d+<\/v><\/c>/);
assert.match(worksheetXml, /<c r="B2" s="\d+"><v>3<\/v><\/c>/);
const sharedStringsXml = await zip.file("xl/sharedStrings.xml").async("text");
assert.match(sharedStringsXml, /<sst/);
assert.match(sharedStringsXml, /<t>Month<\/t>/);
assert.match(sharedStringsXml, /<t>Not Started<\/t>/);
const workbookXml = await zip.file("xl/workbook.xml").async("text");
assert.doesNotMatch(workbookXml, /<workbookPr\b[^>]*date1904=/);
assert.match(workbookXml, /<definedNames>/);
assert.match(workbookXml, /name="RevenueData"/);
assert.match(workbookXml, /Sheet1!G2:G4/);
assert.match(workbookXml, /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><pivotCache cacheId="2" r:id="rId\d+"\/><\/pivotCaches>/);
const stylesXml = await zip.file("xl/styles.xml").async("text");
assert.match(stylesXml, /<dxfs count="2">/);
assert.match(stylesXml, /<fgColor rgb="FF22C55E"/);
assert.match(stylesXml, /<fgColor rgb="FFE0F2FE"/);
assert.match(stylesXml, /<numFmt numFmtId="164" formatCode="#,##0"/);
assert.match(stylesXml, /<fgColor rgb="FF0F172A"/);
assert.match(stylesXml, /<fgColor rgb="FFE0F2FE"/);
assert.match(stylesXml, /<alignment horizontal="center" vertical="center" wrapText="1"\/>/);
assert.match(stylesXml, /<left style="thin"><color rgb="FF334155"\/><\/left>/);
assert.match(stylesXml, /applyAlignment="1"/);
assert.match(stylesXml, /applyBorder="1"/);
assert.match(stylesXml, /<b\/>/);
assert.match(worksheetXml, /<tableParts count="1">/);
assert.match(worksheetXml, /<tablePart r:id="rId1"\/>/);
assert.match(worksheetXml, /<x14:sparklineGroups/);
assert.match(worksheetXml, /<x14:sparklineGroup type="line"/);
assert.match(worksheetXml, /<xm:f>Sheet1!G2:G4<\/xm:f>/);
assert.match(worksheetXml, /<xm:sqref>H2:H2<\/xm:sqref>/);
assert.match(worksheetXml, /<x14:sparklineGroup type="column"/);
assert.match(worksheetXml, /<xm:sqref>H3<\/xm:sqref>/);
const thirdPartyViewXml = worksheetXml
  .replace('showGridLines="0"', 'showGridLines="1"')
  .replace('<pane xSplit="2" ySplit="1" topLeftCell="C2" activePane="bottomRight" state="frozen"/>', '<pane xSplit="1" ySplit="2" topLeftCell="B3" activePane="bottomRight" state="frozenSplit"/>');
const thirdPartyViewXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: thirdPartyViewXml }]);
const thirdPartyViewSheet = (await SpreadsheetFile.importXlsx(thirdPartyViewXlsx)).worksheets.getItem("Sheet1");
assert.equal(thirdPartyViewSheet.showGridLines, true);
assert.deepEqual(thirdPartyViewSheet.freezePanes.toJSON(), { rows: 2, columns: 1, frozen: true, topLeftCell: "B3", activePane: "bottomRight" });
const splitOnlyViewXlsx = await SpreadsheetFile.patchXlsx(xlsx, [{ path: "xl/worksheets/sheet1.xml", xml: thirdPartyViewXml.replace('state="frozenSplit"', 'state="split"') }]);
assert.equal((await SpreadsheetFile.importXlsx(splitOnlyViewXlsx)).worksheets.getItem("Sheet1").freezePanes.frozen, false);
const worksheetRelsXml = await zip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text");
assert.match(worksheetRelsXml, /Target="\.\.\/tables\/table1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/drawings\/drawing1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/threadedComments\/threadedComment1\.xml"/);
assert.match(worksheetRelsXml, /Target="\.\.\/pivotTables\/pivotTable1\.xml"/);
assert.match(worksheetRelsXml, /relationships\/pivotTable/);
assert.match(worksheetRelsXml, /relationships\/threadedComment/);
const drawingXml = await zip.file("xl/drawings/drawing1.xml").async("text");
assert.match(drawingXml, /<xdr:oneCellAnchor>/);
assert.match(drawingXml, /name="LogoImage"/);
assert.match(drawingXml, /descr="Logo placeholder"/);
const drawingRelsXml = await zip.file("xl/drawings/_rels/drawing1.xml.rels").async("text");
assert.match(drawingRelsXml, /Target="\.\.\/media\/image1\.png"/);
assert.match(drawingRelsXml, /Target="\.\.\/charts\/chart1\.xml"/);
const drawingChartXml = await zip.file("xl/drawings/drawing1.xml").async("text");
assert.match(drawingChartXml, /<xdr:graphicFrame>/);
assert.match(drawingChartXml, /<c:chart r:id="rId2"\/>/);
const chartXml = await zip.file("xl/charts/chart1.xml").async("text");
assert.match(chartXml, /<c:chartSpace/);
assert.match(chartXml, /Revenue Trend/);
assert.match(chartXml, /<c:lineChart>/);
assert.match(chartXml, /<c:v>130<\/c:v>/);
const threadedCommentsXml = await zip.file("xl/threadedComments/threadedComment1.xml").async("text");
assert.match(threadedCommentsXml, /<ThreadedComments/);
assert.match(threadedCommentsXml, /ref="C2"/);
assert.match(threadedCommentsXml, /Formula checks revenue sum/);
assert.match(threadedCommentsXml, /Reviewed by model/);
assert.match(threadedCommentsXml, /done="1"/);
const personsXml = await zip.file("xl/persons/person.xml").async("text");
assert.match(personsXml, /<personList/);
assert.match(personsXml, /displayName="Analyst"/);
const workbookRelsXml = await zip.file("xl/_rels/workbook.xml.rels").async("text");
assert.match(workbookRelsXml, /Target="persons\/person.xml"/);
assert.match(workbookRelsXml, /Target="pivotCache\/pivotCacheDefinition1\.xml"/);
assert.match(workbookRelsXml, /relationships\/pivotCacheDefinition/);
const mediaBytes = await zip.file("xl/media/image1.png").async("uint8array");
assert.ok(mediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
assert.match(contentTypesXml, /table1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.table\+xml"/);
assert.match(contentTypesXml, /chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/);
assert.match(contentTypesXml, /pivotTable1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotTable\+xml"/);
assert.match(contentTypesXml, /pivotCacheDefinition1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotCacheDefinition\+xml"/);
assert.match(contentTypesXml, /pivotCacheRecords1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.pivotCacheRecords\+xml"/);
assert.match(contentTypesXml, /threadedComment1\.xml" ContentType="application\/vnd\.ms-excel\.threadedcomments\+xml"/);
assert.match(contentTypesXml, /person\.xml" ContentType="application\/vnd\.ms-excel\.person\+xml"/);
assert.match(contentTypesXml, /sharedStrings\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sharedStrings\+xml"/);
const nativeOnlyZip = await JSZip.loadAsync(xlsxBytes);
nativeOnlyZip.remove("customXml/open-office-artifact.json");
const nativeOnlyWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await nativeOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const nativeOnlySheet = nativeOnlyWorkbook.worksheets.getItem("Sheet1");
assert.equal(nativeOnlySheet.images.items.length, 1);
assert.equal(nativeOnlySheet.charts.items.length, 2);
assert.equal(nativeOnlySheet.pivotTables.items.length, 2);
const nativeOnlyPivot = nativeOnlySheet.pivotTables.getItemOrNullObject("RevenuePivot");
assert.deepEqual(nativeOnlyPivot.computedValues(), [["Month", "Revenue sum"], ["Jan", 100], ["Feb", 120], ["Mar", 130]]);
assert.equal(nativeOnlyWorkbook.resolve("RevenuePivot"), nativeOnlyPivot);
assert.match(nativeOnlyWorkbook.inspect({ kind: "pivotTable", target: "RevenuePivot", maxChars: 4000 }).ndjson, /Revenue sum/);
const nativeOnlyRegionalPivot = nativeOnlySheet.pivotTables.getItemOrNullObject("RegionalPivot");
assert.deepEqual(nativeOnlyRegionalPivot.computedValues(), [["Region", "Q1 — Revenue total", "Q1 — Profit total"], ["East", 15, 6], ["West", 37, 13]]);
assert.deepEqual(nativeOnlyRegionalPivot.columnFields, ["Quarter"]);
assert.deepEqual(nativeOnlyRegionalPivot.calculatedFields, [{ name: "Profit", formula: "=('Revenue'-'Cost')*100%", numFmtId: 0, references: ["Revenue", "Cost"] }]);
assert.deepEqual(nativeOnlyRegionalPivot.filters, [{ field: "Quarter", exclude: ["Q2"] }]);
assert.deepEqual(nativeOnlyRegionalPivot.refreshPolicy, { refreshOnLoad: false, saveData: true, enableRefresh: false, invalid: true, missingItemsLimit: 3, refreshedBy: "QA Agent", refreshedDateIso: "2026-07-12T00:00:00Z" });
const nativeOnlyRegionalRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(nativeOnlyWorkbook));
assert.deepEqual(nativeOnlyRegionalRoundtrip.resolve("RegionalPivot").computedValues(), [["Region", "Q1 — Revenue total", "Q1 — Profit total"], ["East", 15, 6], ["West", 37, 13]]);
assert.deepEqual(nativeOnlyRegionalRoundtrip.resolve("RegionalPivot").filters, [{ field: "Quarter", exclude: ["Q2"] }]);
const unsupportedCalculatedZip = await JSZip.loadAsync(xlsxBytes);
unsupportedCalculatedZip.remove("customXml/open-office-artifact.json");
unsupportedCalculatedZip.file("xl/pivotCache/pivotCacheDefinition2.xml", regionalPivotCacheXml.replace(/formula="[^"]+"/, "formula=\"SUM('Revenue')\""));
const unsupportedCalculatedWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await unsupportedCalculatedZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const unsupportedCalculatedPivot = unsupportedCalculatedWorkbook.resolve("RegionalPivot");
assert.equal(unsupportedCalculatedPivot.calculatedFields[0].supported, false);
assert.equal(unsupportedCalculatedPivot.computedValues()[1][2], "#NAME?");
assert.match(unsupportedCalculatedWorkbook.verify().ndjson, /pivotCalculatedFieldUnsupported/);
const unsupportedCalculatedRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(unsupportedCalculatedWorkbook)).arrayBuffer()));
assert.match(await unsupportedCalculatedRoundtripZip.file("xl/pivotCache/pivotCacheDefinition2.xml").async("text"), /formula="SUM\('Revenue'\)"/);
const nativeOnlyImage = nativeOnlySheet.images.getItemOrNullObject("LogoImage");
assert.equal(nativeOnlyImage.alt, "Logo placeholder");
assert.match(nativeOnlyImage.dataUrl, /^data:image\/png;base64,/);
const nativeOnlyRevenueChart = nativeOnlySheet.charts.items.find((chart) => chart.title === "Revenue Trend");
assert.equal(nativeOnlyRevenueChart.type, "line");
assert.deepEqual(nativeOnlyRevenueChart.categories, ["Jan", "Feb", "Mar"]);
assert.deepEqual(nativeOnlyRevenueChart.series.items[0].values, [100, 120, 130]);
assert.ok(nativeOnlyRevenueChart.position.width > 100 && nativeOnlyRevenueChart.position.height > 80);
assert.equal(nativeOnlyWorkbook.resolve(nativeOnlyImage.id), nativeOnlyImage);
assert.match(nativeOnlyWorkbook.inspect({ kind: "drawing", target: nativeOnlyImage.id, maxChars: 4000 }).ndjson, /LogoImage/);
assert.equal(nativeOnlyWorkbook.worksheets.getItem("Sheet1").showGridLines, false);
assert.deepEqual(nativeOnlyWorkbook.worksheets.getItem("Sheet1").freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const nativeSparklineInspect = nativeOnlyWorkbook.inspect({ kind: "sparkline", maxChars: 12000 }).ndjson;
assert.match(nativeSparklineInspect, /"kind":"sparkline"/);
assert.match(nativeSparklineInspect, /"type":"line"/);
assert.match(nativeSparklineInspect, /"targetRange":"H2:H2"/);
assert.match(nativeSparklineInspect, /"sourceData":"G2:G4"/);
assert.match(nativeSparklineInspect, /"type":"column"/);
const nativeThreadInspect = nativeOnlyWorkbook.inspect({ kind: "thread", maxChars: 12000 }).ndjson;
assert.match(nativeThreadInspect, /"kind":"thread"/);
assert.match(nativeThreadInspect, /Formula checks revenue sum/);
assert.match(nativeThreadInspect, /Reviewed by model/);
assert.match(nativeThreadInspect, /"resolved":true/);
const nativeStyleInspect = nativeOnlyWorkbook.inspect({ kind: "style", range: "A1:C3", maxChars: 12000 }).ndjson;
assert.match(nativeStyleInspect, /"kind":"style"/);
assert.match(nativeStyleInspect, /"numberFormat":"#,##0"/);
assert.match(nativeStyleInspect, /"alignment":\{"horizontal":"center"/);
assert.match(nativeStyleInspect, /"border":\{"style":"thin"/);
const relocatedDrawingZip = await JSZip.loadAsync(xlsxBytes);
relocatedDrawingZip.remove("customXml/open-office-artifact.json");
const relocatedDrawingXml = (await relocatedDrawingZip.file("xl/drawings/drawing1.xml").async("text")).replace(
  /<xdr:oneCellAnchor>((?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?name="Chart 1"(?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?)<\/xdr:oneCellAnchor>/,
  (_anchor, body) => `<xdr:twoCellAnchor editAs="twoCell">${body.replace(/<xdr:ext[^>]*\/>/, '<xdr:to><xdr:col>13</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>10</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>')}</xdr:twoCellAnchor>`,
).replace(
  /<xdr:oneCellAnchor>((?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?name="ScoresChart"(?:(?!<\/xdr:oneCellAnchor>)[\s\S])*?)<\/xdr:oneCellAnchor>/,
  (_anchor, body) => `<xdr:absoluteAnchor>${body.replace(/<xdr:from>[\s\S]*?<\/xdr:from>/, '<xdr:pos x="1905000" y="2857500"/>')}</xdr:absoluteAnchor>`,
);
assert.match(relocatedDrawingXml, /<xdr:twoCellAnchor editAs="twoCell">/);
assert.match(relocatedDrawingXml, /<xdr:absoluteAnchor>/);
relocatedDrawingZip.file("xl/custom/visuals/workbook-drawing.xml", relocatedDrawingXml);
relocatedDrawingZip.remove("xl/drawings/drawing1.xml");
const relocatedDrawingRelationships = (await relocatedDrawingZip.file("xl/drawings/_rels/drawing1.xml.rels").async("text"))
  .replace("../media/image1.png", "../../assets/logo.png")
  .replace("../charts/chart1.xml", "chart-revenue.xml")
  .replace("../charts/chart2.xml", "chart-scores.xml");
relocatedDrawingZip.file("xl/custom/visuals/_rels/workbook-drawing.xml.rels", relocatedDrawingRelationships);
relocatedDrawingZip.remove("xl/drawings/_rels/drawing1.xml.rels");
relocatedDrawingZip.file("xl/assets/logo.png", await relocatedDrawingZip.file("xl/media/image1.png").async("uint8array"));
relocatedDrawingZip.remove("xl/media/image1.png");
const relocatedRevenueChartXml = (await relocatedDrawingZip.file("xl/charts/chart1.xml").async("text"))
  .replace(/<c:tx><c:v>([\s\S]*?)<\/c:v><\/c:tx>/, '<c:tx><c:strRef><c:f>Sheet1!$G$1</c:f><c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>$1</c:v></c:pt></c:strCache></c:strRef></c:tx>')
  .replace(/<c:cat><c:strLit>([\s\S]*?)<\/c:strLit><\/c:cat>/, '<c:cat><c:strRef><c:f>Sheet1!$F$2:$F$4</c:f><c:strCache>$1</c:strCache></c:strRef></c:cat>')
  .replace(/<c:val><c:numLit>([\s\S]*?)<\/c:numLit><\/c:val>/, '<c:val><c:numRef><c:f>Sheet1!$G$2:$G$4</c:f><c:numCache>$1</c:numCache></c:numRef></c:val>');
relocatedDrawingZip.file("xl/custom/visuals/chart-revenue.xml", relocatedRevenueChartXml);
relocatedDrawingZip.file("xl/custom/visuals/chart-scores.xml", await relocatedDrawingZip.file("xl/charts/chart2.xml").async("text"));
relocatedDrawingZip.remove("xl/charts/chart1.xml");
relocatedDrawingZip.remove("xl/charts/chart2.xml");
relocatedDrawingZip.file("xl/worksheets/_rels/sheet1.xml.rels", (await relocatedDrawingZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text")).replace("../drawings/drawing1.xml", "../custom/visuals/workbook-drawing.xml"));
relocatedDrawingZip.file("[Content_Types].xml", (await relocatedDrawingZip.file("[Content_Types].xml").async("text"))
  .replace("/xl/charts/chart1.xml", "/xl/custom/visuals/chart-revenue.xml")
  .replace("/xl/charts/chart2.xml", "/xl/custom/visuals/chart-scores.xml"));
const relocatedDrawingBlob = new FileBlob(await relocatedDrawingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type });
const relocatedDrawingInspect = await SpreadsheetFile.inspectXlsx(relocatedDrawingBlob);
assert.equal(relocatedDrawingInspect.ok, true);
assert.deepEqual(relocatedDrawingInspect.issues, []);
const relocatedDrawingBook = await SpreadsheetFile.importXlsx(relocatedDrawingBlob);
const relocatedDrawingSheet = relocatedDrawingBook.worksheets.getItem("Sheet1");
assert.equal(relocatedDrawingSheet.images.items.length, 1);
assert.equal(relocatedDrawingSheet.charts.items.length, 2);
assert.equal(relocatedDrawingSheet.images.items[0].alt, "Logo placeholder");
const relocatedRevenueChart = relocatedDrawingSheet.charts.items.find((chart) => chart.title === "Revenue Trend");
const relocatedScoreChart = relocatedDrawingSheet.charts.items.find((chart) => chart.title === "Scores");
assert.deepEqual(relocatedRevenueChart.series.items[0].values, [100, 120, 130]);
assert.equal(relocatedRevenueChart.series.items[0].formula, "Sheet1!$G$2:$G$4");
assert.equal(relocatedRevenueChart.series.items[0].categoryFormula, "Sheet1!$F$2:$F$4");
assert.ok(relocatedRevenueChart.position.width > 100 && relocatedRevenueChart.position.height > 80);
assert.deepEqual(relocatedScoreChart.position, { left: 240, top: 340, width: 240, height: 160 });
assert.match(relocatedDrawingSheet.toSvg(), /Revenue Trend/);
assert.match(relocatedDrawingSheet.toSvg(), /data:image\/png;base64/);
assert.ok(relocatedDrawingSheet.layoutJson().charts.some((chart) => chart.title === "Revenue Trend"));
assert.ok(relocatedDrawingBook.verify().issues.every((issue) => !["emptyChart", "emptyImage", "chartDataMismatch", "invalidImageDataUrl"].includes(issue.type)));
const relocatedDrawingRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(relocatedDrawingBook));
assert.equal(relocatedDrawingRoundtrip.worksheets.getItem("Sheet1").images.items.length, 1);
assert.equal(relocatedDrawingRoundtrip.worksheets.getItem("Sheet1").charts.items.length, 2);
const styleFidelityZip = await JSZip.loadAsync(xlsxBytes);
styleFidelityZip.remove("customXml/open-office-artifact.json");
const assignStyle = (xml, address, styleIndex) => xml.replace(new RegExp(`<c r="${address}"([^>]*)>`), (match, attrs) => `<c r="${address}"${attrs.replace(/\s+s="[^"]*"/, "")} s="${styleIndex}">`);
let styleFidelityXml = assignStyle(worksheetXml, "A2", 1);
styleFidelityXml = assignStyle(styleFidelityXml, "B2", 2);
styleFidelityXml = assignStyle(styleFidelityXml, "C2", 3);
styleFidelityZip.file("xl/worksheets/sheet1.xml", styleFidelityXml);
styleFidelityZip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><b val="false"/><i val="true"/><u val="double"/><strike val="true"/><sz val="12"/><color rgb="80123456"/><name val="Liberation Sans"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="7FFDE68A"/></patternFill></fill></fills><borders count="2"><border/><border diagonalUp="false"><left style="thin"><color rgb="40334155"/></left><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="4" fontId="1" fillId="2" borderId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1"><alignment horizontal="center" vertical="center" textRotation="30" wrapText="true" indent="1" shrinkToFit="true" readingOrder="2"/><protection locked="false" hidden="true"/></xf></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="10" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="right"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="1" applyNumberFormat="0" applyFont="0" applyFill="0" applyBorder="0" applyAlignment="0" applyProtection="0"/></cellXfs><dxfs count="0"/></styleSheet>`);
const styleFidelityBook = await SpreadsheetFile.importXlsx(new FileBlob(await styleFidelityZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const styleFidelitySheet = styleFidelityBook.worksheets.getItem("Sheet1");
assert.equal(styleFidelitySheet.getRange("A2").format.numberFormat, "mm-dd-yy");
assert.equal(styleFidelitySheet.getRange("B2").format.numberFormat, "0.00%");
assert.deepEqual(styleFidelitySheet.getRange("C2").format.font, { bold: false, italic: true, underline: "double", strike: true, color: "#123456", size: 12, name: "Liberation Sans" });
assert.equal(styleFidelitySheet.getRange("C2").format.fill, "#FDE68A");
assert.deepEqual(styleFidelitySheet.getRange("C2").format.alignment, { horizontal: "center", vertical: "center", wrapText: true, shrinkToFit: true, textRotation: 30, indent: 1, readingOrder: 2 });
assert.deepEqual(styleFidelitySheet.getRange("C2").format.protection, { locked: false, hidden: true });
const styleFidelityLayout = styleFidelitySheet.layoutJson({ range: "A2:C2" });
assert.deepEqual(styleFidelityLayout.cells.map((cell) => cell.displayValue), ["01-02-00", "300.00%", "5.00"]);
const styleFidelitySvg = styleFidelitySheet.toSvg();
assert.match(styleFidelitySvg, />300\.00%<\/text>/);
assert.match(styleFidelitySvg, /text-decoration="underline line-through"/);
assert.match(styleFidelitySvg, /transform="rotate\(-30 /);
const styleFidelityRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(styleFidelityBook));
assert.deepEqual(styleFidelityRoundtrip.worksheets.getItem("Sheet1").getRange("C2").format.alignment, styleFidelitySheet.getRange("C2").format.alignment);
assert.deepEqual(styleFidelityRoundtrip.worksheets.getItem("Sheet1").getRange("C2").format.protection, { locked: false, hidden: true });
const themeStyleZip = await JSZip.loadAsync(xlsxBytes);
themeStyleZip.remove("customXml/open-office-artifact.json");
themeStyleZip.file("xl/worksheets/sheet1.xml", assignStyle(worksheetXml, "A2", 1));
themeStyleZip.file("xl/_rels/workbook.xml.rels", workbookRelsXml.replace("</Relationships>", `<Relationship Id="rTheme99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="appearance/custom-theme.xml"/></Relationships>`));
themeStyleZip.file("xl/appearance/custom-theme.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Clean room theme"><a:themeElements><a:clrScheme name="Clean room"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="336699"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme></a:themeElements></a:theme>`);
themeStyleZip.file("xl/styles.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><fonts count="2"><font><sz val="11"/><name val="Aptos"/></font><font><sz val="11"/><color theme="4" tint="0.5"/><name val="Aptos"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor indexed="0"/></patternFill></fill></fills><borders count="2"><border/><border diagonalUp="1"><left style="thin"><color theme="4" tint="-0.25"/></left><right style="medium"><color indexed="0"/></right><top/><bottom style="double"><color rgb="FF112233"/></bottom><diagonal style="dashed"><color auto="1"/></diagonal></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1"/></cellXfs><colors><indexedColors><rgbColor rgb="FFABCDEF"/></indexedColors></colors></styleSheet>`);
const themeStyleBook = await SpreadsheetFile.importXlsx(new FileBlob(await themeStyleZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const themeStyleRange = themeStyleBook.worksheets.getItem("Sheet1").getRange("A2");
assert.equal(themeStyleRange.format.font.color, "#8CB3D9");
assert.equal(themeStyleRange.format.fill, "#ABCDEF");
assert.deepEqual(themeStyleRange.format.border, {
  left: { style: "thin", color: "#264D73" },
  right: { style: "medium", color: "#ABCDEF" },
  bottom: { style: "double", color: "#112233" },
  diagonal: { style: "dashed", color: "#000000" },
  diagonalUp: true,
});
assert.match(themeStyleBook.worksheets.getItem("Sheet1").toSvg(), /fill="#ABCDEF"/);
const themeStyleRoundtrip = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(themeStyleBook));
assert.equal(themeStyleRoundtrip.worksheets.getItem("Sheet1").getRange("A2").format.font.color, "#8CB3D9");
assert.deepEqual(themeStyleRoundtrip.worksheets.getItem("Sheet1").getRange("A2").format.border, themeStyleRange.format.border);
const nativeColorScaleInspect = nativeOnlyWorkbook.inspect({ kind: "conditionalFormat,computedStyle", target: "Sheet1!G4", maxChars: 12000 }).ndjson;
assert.match(nativeColorScaleInspect, /"ruleType":"colorScale"/);
assert.match(nativeColorScaleInspect, /"fill":"#22c55e"/);
const pivotRoundtripWorkbook = await SpreadsheetFile.importXlsx(xlsx);
assert.deepEqual(pivotRoundtripWorkbook.resolve("RevenuePivot").computedValues()[1], ["Jan", 100]);
assert.match(pivotRoundtripWorkbook.inspect({ kind: "pivotTable", target: "RevenuePivot", maxChars: 8000 }).ndjson, /Revenue sum/);
const sharedFormulaZip = await JSZip.loadAsync(xlsxBytes);
sharedFormulaZip.remove("customXml/open-office-artifact.json");
const sharedFormulaXml = worksheetXml
  .replace(/<c r="C2"([^>]*)>([\s\S]*?)<f>A2\+B2<\/f>([\s\S]*?)<\/c>/, `<c r="C2"$1>$2<f t="shared" si="0" ref="C2:C3">A2+B2</f>$3</c>`)
  .replace(/<c r="C3"([^>]*)>([\s\S]*?)<f>A3\+B3<\/f>([\s\S]*?)<\/c>/, `<c r="C3"$1>$2<f t="shared" si="0"></f>$3</c>`)
  .replace(/<c r="E3"([^>]*)>([\s\S]*?)<f>SUM\(RevenueData\)<\/f>([\s\S]*?)<\/c>/, `<c r="E3"$1>$2<f t="array" ref="E3:E4">SUM(G2:G4)</f>$3</c>`);
assert.match(sharedFormulaXml, /<f t="shared" si="0" ref="C2:C3">A2\+B2<\/f>/);
assert.match(sharedFormulaXml, /<f t="array" ref="E3:E4">SUM\(G2:G4\)<\/f>/);
sharedFormulaZip.file("xl/worksheets/sheet1.xml", sharedFormulaXml);
const sharedFormulaWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await sharedFormulaZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: xlsx.type }));
const sharedFormulaInspect = sharedFormulaWorkbook.inspect({ kind: "formula", target: "Sheet1!C3", maxChars: 8000 }).ndjson;
assert.match(sharedFormulaInspect, /"formula":"=A3\+B3"/);
assert.match(sharedFormulaInspect, /"formulaType":"shared"/);
assert.match(sharedFormulaInspect, /"sharedIndex":0/);
assert.match(sharedFormulaInspect, /"sharedRef":"C2:C3"/);
const arrayFormulaInspect = sharedFormulaWorkbook.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson;
assert.match(arrayFormulaInspect, /"formula":"=SUM\(G2:G4\)"/);
assert.match(arrayFormulaInspect, /"formulaType":"array"/);
assert.match(arrayFormulaInspect, /"arrayRef":"E3:E4"/);
assert.equal(sharedFormulaWorkbook.worksheets.getItem("Sheet1").getRange("C3").values[0][0], 12);
assert.equal(sharedFormulaWorkbook.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const sharedFormulaRoundtrip = await SpreadsheetFile.exportXlsx(sharedFormulaWorkbook);
const sharedFormulaRoundtripZip = await JSZip.loadAsync(new Uint8Array(await sharedFormulaRoundtrip.arrayBuffer()));
const sharedFormulaRoundtripXml = await sharedFormulaRoundtripZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(sharedFormulaRoundtripXml, /<f t="shared" si="0" ref="C2:C3">A2\+B2<\/f>/);
assert.match(sharedFormulaRoundtripXml, /<c r="C3"[^>]*>[\s\S]*<f t="shared" si="0"><\/f>/);
assert.match(sharedFormulaRoundtripXml, /<f t="array" ref="E3:E4">SUM\(G2:G4\)<\/f>/);
const sharedFormulaRoundtripImport = await SpreadsheetFile.importXlsx(sharedFormulaRoundtrip);
assert.match(sharedFormulaRoundtripImport.inspect({ kind: "formula", target: "Sheet1!C3", maxChars: 8000 }).ndjson, /"formulaType":"shared"/);
assert.match(sharedFormulaRoundtripImport.inspect({ kind: "formula", target: "Sheet1!E3", maxChars: 8000 }).ndjson, /"formulaType":"array"/);
const nativeRuleInspect = nativeOnlyWorkbook.inspect({ kind: "dataValidation,conditionalFormat", maxChars: 12000 }).ndjson;
assert.match(nativeRuleInspect, /"kind":"dataValidation"/);
assert.match(nativeRuleInspect, /"range":"D2:D3"/);
assert.match(nativeRuleInspect, /Not Started/);
assert.match(nativeRuleInspect, /"kind":"conditionalFormat"/);
assert.match(nativeRuleInspect, /"ruleType":"cellIs"/);
assert.match(nativeRuleInspect, /"ruleType":"expression"/);
const nativeComputedStyle = nativeOnlyWorkbook.inspect({ kind: "computedStyle", target: "Sheet1!C3", maxChars: 8000 }).ndjson;
assert.match(nativeComputedStyle, /"kind":"computedStyle"/);
assert.match(nativeComputedStyle, /"fill":"#22C55E"/);
assert.match(nativeOnlyWorkbook.inspect({ kind: "table", range: "A1:D4" }).ndjson, /"A"/);
assert.match(nativeOnlyWorkbook.inspect({ kind: "definedName", target: "RevenueData" }).ndjson, /Sheet1!G2:G4/);
assert.equal(nativeOnlyWorkbook.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.xlsx`);
await xlsx.save(out);
const loaded = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
assert.equal(loaded.worksheets.getItem("Sheet1").showGridLines, false);
assert.deepEqual(loaded.worksheets.getItem("Sheet1").freezePanes.toJSON(), { rows: 1, columns: 2, frozen: true, topLeftCell: "C2", activePane: "bottomRight" });
const roundtrip = loaded.inspect({ kind: "table,formula", range: "A1:C3" }).ndjson;
assert.match(roundtrip, /"values":\[\["A","B","Sum"\],\[2,3,5\],\[5,7,12\]\]/);
assert.match(roundtrip, /"formula":"=A2\+B2"/);
const roundtripMetadata = loaded.inspect({ kind: "dataValidation,conditionalFormat,thread,table,pivotTable,drawing", maxChars: 16000 }).ndjson;
assert.match(roundtripMetadata, /"kind":"dataValidation"/);
assert.match(roundtripMetadata, /"kind":"conditionalFormat"/);
assert.match(roundtripMetadata, /"kind":"thread"/);
assert.match(roundtripMetadata, /TasksTable/);
assert.match(roundtripMetadata, /RevenuePivot/);
assert.match(roundtripMetadata, /Revenue sum/);
assert.match(roundtripMetadata, /Revenue Trend/);
assert.match(roundtripMetadata, /ScoresChart/);
assert.match(roundtripMetadata, /LogoImage/);
assert.match(roundtripMetadata, /"kind":"sparkline"/);
assert.match(roundtripMetadata, /Reviewed by model/);
assert.match(loaded.inspect({ kind: "definedName", target: "RevenueData" }).ndjson, /Sheet1!G2:G4/);
assert.equal(loaded.worksheets.getItem("Sheet1").getRange("E3").values[0][0], 350);
const loadedThreadId = JSON.parse(roundtripMetadata.split("\n").find((line) => line.includes('"kind":"thread"'))).id;
assert.equal(loaded.resolve(loadedThreadId).resolved, true);
assert.deepEqual(loaded.trace("Sheet1!C3").tree.precedents.map((node) => node.address), ["A3", "B3"]);
const refreshOnlyBook = Workbook.create();
const refreshOnlySheet = refreshOnlyBook.worksheets.add("Source");
refreshOnlySheet.getRange("A1:B3").values = [["Team", "Score"], ["Red", 4], ["Blue", 6]];
refreshOnlySheet.pivotTables.add({
  name: "RefreshOnlyPivot",
  sourceRange: "A1:B3",
  targetRange: "D1:E3",
  rowFields: ["Team"],
  valueFields: [{ field: "Score", summarizeBy: "sum" }],
  refreshPolicy: { saveData: false, refreshOnLoad: true },
});
const refreshOnlyXlsx = await SpreadsheetFile.exportXlsx(refreshOnlyBook);
assert.equal((await SpreadsheetFile.inspectXlsx(refreshOnlyXlsx)).ok, true);
const refreshOnlyZip = await JSZip.loadAsync(new Uint8Array(await refreshOnlyXlsx.arrayBuffer()));
assert.equal(refreshOnlyZip.file("xl/pivotCache/pivotCacheRecords1.xml"), null);
assert.equal(refreshOnlyZip.file("xl/pivotCache/_rels/pivotCacheDefinition1.xml.rels"), null);
assert.doesNotMatch(await refreshOnlyZip.file("[Content_Types].xml").async("text"), /pivotCacheRecords/);
assert.doesNotMatch(await refreshOnlyZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /r:id=/);
assert.match(await refreshOnlyZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /refreshOnLoad="1" saveData="0"/);
refreshOnlyZip.remove("customXml/open-office-artifact.json");
const refreshOnlyNative = await SpreadsheetFile.importXlsx(new FileBlob(await refreshOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: refreshOnlyXlsx.type }));
const importedRefreshOnlyPivot = refreshOnlyNative.resolve("RefreshOnlyPivot");
assert.deepEqual(importedRefreshOnlyPivot.computedValues(), [["Team", "sum of Score"], ["Red", 4], ["Blue", 6]]);
assert.equal(importedRefreshOnlyPivot.refreshPolicy.saveData, false);
console.log("spreadsheet smoke ok");
