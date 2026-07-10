import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
workbook.recalculate();

const inspect = workbook.inspect({ kind: "workbook,sheet,table,formula", range: "A1:C3", include: "values,formulas" });
assert.match(inspect.ndjson, /"value":5|"values":/);
assert.match(inspect.ndjson, /"formula":"=A2\+B2"/);
assert.match(workbook.help("fx.PMT").ndjson, /fx.PMT/);
assert.match(workbook.help("range.dataValidation").ndjson, /validation rule/);
assert.match(workbook.help("range.conditionalFormats.add").ndjson, /conditional formatting/);
assert.match(workbook.help("workbook.comments.addThread").ndjson, /threaded comments/);
const statusRange = sheet.getRange("D2:D3");
statusRange.values = [["Not Started"], ["In Progress"]];
statusRange.dataValidation = { rule: { type: "list", values: ["Not Started", "In Progress", "Done"] } };
const cf = sheet.getRange("C2:C3").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 10, format: { fill: "green" } });
const customCf = sheet.getRange("A2:B3").conditionalFormats.addCustom("=A2>B2", { fill: "sky-100" });
workbook.comments.setSelf({ displayName: "Analyst" });
const thread = workbook.comments.addThread({ cell: sheet.getRange("C2") }, "Formula checks revenue sum.");
thread.addReply("Reviewed by model.").resolve();

const tasksTable = sheet.tables.add("A1:D3", true, "TasksTable");
tasksTable.style = "TableStyleMedium4";
tasksTable.showBandedColumns = true;
tasksTable.rows.add(null, [["8", "13", "21", "Done"]]);
assert.deepEqual(tasksTable.getDataRows()[0].slice(0, 3), [2, 3, 5]);
assert.equal(tasksTable.getHeaderRowRange().values[0][0], "A");

const chartSource = sheet.getRange("F1:G4");
chartSource.values = [["Month", "Revenue"], ["Jan", 100], ["Feb", 120], ["Mar", 130]];
const chartFromRange = sheet.charts.add("line", chartSource);
chartFromRange.title = "Revenue Trend";
chartFromRange.hasLegend = false;
chartFromRange.setPosition("I1", "M10");
const chartFromConfig = sheet.charts.add("bar", { name: "ScoresChart", title: "Scores", categories: ["A", "B"], series: [{ name: "Score", values: [9, 7] }], position: { left: 40, top: 220, width: 240, height: 160 } });
const image = sheet.images.add({
  name: "LogoImage",
  dataUrl: "data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHdpZHRoPSIxMCIgaGVpZ2h0PSIxMCI+PC9zdmc+",
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

const metadataInspect = workbook.inspect({ kind: "dataValidation,conditionalFormat,thread,table,drawing", maxChars: 12000 }).ndjson;
assert.match(metadataInspect, /"kind":"dataValidation"/);
assert.match(metadataInspect, /"type":"list"/);
assert.match(metadataInspect, /"kind":"conditionalFormat"/);
assert.match(metadataInspect, /"ruleType":"cellIs"/);
assert.match(metadataInspect, /"ruleType":"expression"/);
assert.match(metadataInspect, /"kind":"thread"/);
assert.match(metadataInspect, /Formula checks revenue sum/);
assert.match(metadataInspect, /TasksTable/);
assert.match(metadataInspect, /Revenue Trend/);
assert.match(metadataInspect, /ScoresChart/);
assert.match(metadataInspect, /LogoImage/);
assert.match(metadataInspect, /"kind":"sparkline"/);
assert.equal(workbook.resolve(thread.id).resolved, true);
assert.equal(workbook.resolve(cf.id).operator, "greaterThan");
assert.equal(workbook.resolve(customCf.id).formula, "=A2>B2");
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
assert.match(workbook.help("workbook.trace").ndjson, /precedent tree/);

const preview = await workbook.render({ sheetName: "Sheet1", range: "A1:C3" });
assert.equal(preview.type, "image/svg+xml");
const previewSvg = await preview.text();
assert.match(previewSvg, /<svg/);
assert.match(previewSvg, /TasksTable/);
assert.match(previewSvg, /Revenue Trend/);
assert.match(previewSvg, /Logo placeholder/);
assert.match(previewSvg, /polyline/);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(xlsx.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.xlsx`);
await xlsx.save(out);
const loaded = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
const roundtrip = loaded.inspect({ kind: "table,formula", range: "A1:C3" }).ndjson;
assert.match(roundtrip, /"values":\[\["A","B","Sum"\],\[2,3,5\],\[5,7,12\]\]/);
assert.match(roundtrip, /"formula":"=A2\+B2"/);
const roundtripMetadata = loaded.inspect({ kind: "dataValidation,conditionalFormat,thread,table,drawing", maxChars: 12000 }).ndjson;
assert.match(roundtripMetadata, /"kind":"dataValidation"/);
assert.match(roundtripMetadata, /"kind":"conditionalFormat"/);
assert.match(roundtripMetadata, /"kind":"thread"/);
assert.match(roundtripMetadata, /TasksTable/);
assert.match(roundtripMetadata, /Revenue Trend/);
assert.match(roundtripMetadata, /ScoresChart/);
assert.match(roundtripMetadata, /LogoImage/);
assert.match(roundtripMetadata, /"kind":"sparkline"/);
assert.match(roundtripMetadata, /Reviewed by model/);
const loadedThreadId = JSON.parse(roundtripMetadata.split("\n").find((line) => line.includes('"kind":"thread"'))).id;
assert.equal(loaded.resolve(loadedThreadId).resolved, true);
assert.deepEqual(loaded.trace("Sheet1!C3").tree.precedents.map((node) => node.address), ["A3", "B3"]);
console.log("spreadsheet smoke ok");
