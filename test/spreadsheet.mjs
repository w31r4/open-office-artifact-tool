import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
sheet.getRange("A1:D1").format = { fill: "#0f172a", font: { bold: true, color: "#ffffff", name: "Aptos", size: 12 }, alignment: { horizontal: "center", vertical: "center", wrapText: true }, border: { style: "thin", color: "#334155" } };
sheet.getRange("B2:C3").setFormat({ numberFormat: "#,##0", fill: "sky-100" });
workbook.recalculate();

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
catalogSheet.getRange("F1:F20").formulas = [
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
];
catalogBook.recalculate();
assert.deepEqual(catalogSheet.getRange("F1:F20").values.flat(), ["ok", 1.23, 2, 30, 15, 20, "Al-100", "Alpha/Beta", 6, "MIXED", true, true, true, 4, 9, 6, "BX-200", 3, 15, 2]);
assert.match(catalogBook.help("fx.XLOOKUP").ndjson, /lookup/);
assert.match(catalogBook.help("fx.INDEX").ndjson, /1-based row/);
assert.match(catalogBook.help("fx.MATCH").ndjson, /1-based position/);
assert.match(catalogBook.help("fx.TEXTJOIN").ndjson, /delimiter/);
assert.match(catalogBook.inspect({ kind: "formula", maxChars: 20000 }).ndjson, /XLOOKUP/);

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
spillBook.recalculate();
assert.deepEqual(spillSheet.getRange("L1:N2").values, [["Alpha", "East", 10], ["Gamma", "East", 20]]);
assert.deepEqual(spillSheet.getRange("P1:P2").values, [["East"], ["West"]]);
assert.deepEqual(spillSheet.getRange("R1:T3").values, [["Gamma", "East", 20], ["Beta", "West", 15], ["Alpha", "East", 10]]);
assert.match(spillBook.inspect({ kind: "formula", target: "Spill!L1", maxChars: 8000 }).ndjson, /"spillRange":"L1:N2"/);
assert.match(spillBook.help("fx.SEQUENCE").ndjson, /dynamic array/);
assert.match(spillBook.help("fx.TRANSPOSE").ndjson, /spillRange/);
assert.match(spillBook.help("fx.FILTER").ndjson, /include array/);
assert.match(spillBook.help("fx.UNIQUE").ndjson, /unique rows/);
assert.match(spillBook.help("fx.SORT").ndjson, /column index/);
const spillXlsx = await SpreadsheetFile.exportXlsx(spillBook);
const spillZip = await JSZip.loadAsync(new Uint8Array(await spillXlsx.arrayBuffer()));
const spillWorksheetXml = await spillZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(spillWorksheetXml, /<f t="array" ref="A1:C2">SEQUENCE\(2,3,10,2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="E1:F3">TRANSPOSE\(A1:C2\)<\/f>/);
assert.match(spillWorksheetXml, /<f t="array" ref="L1:N2">FILTER\(H2:J4,I2:I4="East"\)<\/f>/);
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

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(xlsx.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
const xlsxBytes = new Uint8Array(await xlsx.arrayBuffer());
const zip = await JSZip.loadAsync(xlsxBytes);
const tablePartNames = Object.keys(zip.files).filter((name) => /^xl\/tables\/table\d+\.xml$/.test(name));
assert.equal(tablePartNames.length, 1);
const tableXml = await zip.file(tablePartNames[0]).async("text");
assert.match(tableXml, /displayName="TasksTable"/);
assert.match(tableXml, /ref="A1:D4"/);
assert.match(tableXml, /<tableColumns count="4">/);
const pivotPartNames = Object.keys(zip.files).filter((name) => /^xl\/pivotTables\/pivotTable\d+\.xml$/.test(name));
assert.equal(pivotPartNames.length, 1);
const pivotTableXml = await zip.file(pivotPartNames[0]).async("text");
assert.match(pivotTableXml, /<pivotTableDefinition/);
assert.match(pivotTableXml, /name="RevenuePivot"/);
assert.match(pivotTableXml, /cacheId="1"/);
assert.match(pivotTableXml, /<rowFields count="1"><field x="0"\/><\/rowFields>/);
assert.match(pivotTableXml, /<dataField name="Revenue sum" fld="1" subtotal="sum"\/>/);
const pivotCacheXml = await zip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
assert.match(pivotCacheXml, /<pivotCacheDefinition/);
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
const worksheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(worksheetXml, /<conditionalFormatting sqref="C2:C3">/);
assert.match(worksheetXml, /<cfRule type="cellIs"[^>]*dxfId="0"/);
assert.match(worksheetXml, /<dataValidations count="1">/);
assert.match(worksheetXml, /<dataValidation type="list"[^>]*sqref="D2:D3"/);
assert.match(worksheetXml, /<c r="A1" t="s" s="\d+"><v>\d+<\/v><\/c>/);
assert.match(worksheetXml, /<c r="B2" s="\d+"><v>3<\/v><\/c>/);
const sharedStringsXml = await zip.file("xl/sharedStrings.xml").async("text");
assert.match(sharedStringsXml, /<sst/);
assert.match(sharedStringsXml, /<t>Month<\/t>/);
assert.match(sharedStringsXml, /<t>Not Started<\/t>/);
const workbookXml = await zip.file("xl/workbook.xml").async("text");
assert.match(workbookXml, /<definedNames>/);
assert.match(workbookXml, /name="RevenueData"/);
assert.match(workbookXml, /Sheet1!G2:G4/);
assert.match(workbookXml, /<pivotCaches><pivotCache cacheId="1" r:id="rId\d+"\/><\/pivotCaches>/);
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
console.log("spreadsheet smoke ok");
