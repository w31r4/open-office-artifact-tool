import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
sheet.getRange("A1:D1").format = { fill: "#0f172a", font: { bold: true, color: "#ffffff", name: "Aptos", size: 12 } };
sheet.getRange("B2:C3").setFormat({ numberFormat: "#,##0", fill: "sky-100" });
workbook.recalculate();

const inspect = workbook.inspect({ kind: "workbook,sheet,table,formula", range: "A1:C3", include: "values,formulas" });
assert.match(inspect.ndjson, /"value":5|"values":/);
assert.match(inspect.ndjson, /"formula":"=A2\+B2"/);
const styleInspect = workbook.inspect({ kind: "style", range: "A1:D3", maxChars: 8000 }).ndjson;
assert.match(styleInspect, /"kind":"style"/);
assert.match(styleInspect, /"numberFormat":"#,##0"/);
assert.match(styleInspect, /"fill":"sky-100"/);
assert.match(workbook.help("range.format").ndjson, /cell style/);
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
assert.match(workbook.help("workbook.formulaGraph").ndjson, /dependency graph/);

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
catalogSheet.getRange("F1:F16").formulas = [
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
];
catalogBook.recalculate();
assert.deepEqual(catalogSheet.getRange("F1:F16").values.flat(), ["ok", 1.23, 2, 30, 15, 20, "Al-100", "Alpha/Beta", 6, "MIXED", true, true, true, 4, 9, 6]);
assert.match(catalogBook.help("fx.XLOOKUP").ndjson, /lookup/);
assert.match(catalogBook.help("fx.TEXTJOIN").ndjson, /delimiter/);
assert.match(catalogBook.inspect({ kind: "formula", maxChars: 20000 }).ndjson, /XLOOKUP/);

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
assert.match(previewSvg, /polyline/);
const layoutBlob = await workbook.render({ format: "layout", sheetName: "Sheet1", range: "A1:C3" });
assert.equal(layoutBlob.type, "application/vnd.open-office-artifact.layout+json");
const layout = JSON.parse(await layoutBlob.text());
assert.equal(layout.kind, "workbookLayout");
assert.equal(layout.sheets[0].name, "Sheet1");
assert.equal(layout.sheets[0].bounds.address, "A1:C3");
assert.ok(layout.sheets[0].cells.some((cell) => cell.address === "C2" && cell.formula === "=A2+B2"));
assert.ok(layout.sheets[0].tables.some((table) => table.name === "TasksTable"));
assert.ok(layout.sheets[0].charts.some((chart) => chart.title === "Revenue Trend"));
assert.ok(layout.sheets[0].images.some((item) => item.alt === "Logo placeholder"));
assert.ok(layout.sheets[0].sparklines.some((item) => item.targetRange === "H2:H2"));
assert.match(workbook.help("workbook.layoutJson").ndjson, /layout JSON/);

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
const worksheetXml = await zip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(worksheetXml, /<c r="A1" t="s" s="\d+"><v>\d+<\/v><\/c>/);
assert.match(worksheetXml, /<c r="B2" s="\d+"><v>3<\/v><\/c>/);
const sharedStringsXml = await zip.file("xl/sharedStrings.xml").async("text");
assert.match(sharedStringsXml, /<sst/);
assert.match(sharedStringsXml, /<t>Month<\/t>/);
assert.match(sharedStringsXml, /<t>Not Started<\/t>/);
const stylesXml = await zip.file("xl/styles.xml").async("text");
assert.match(stylesXml, /<numFmt numFmtId="164" formatCode="#,##0"/);
assert.match(stylesXml, /<fgColor rgb="FF0F172A"/);
assert.match(stylesXml, /<fgColor rgb="FFE0F2FE"/);
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
const mediaBytes = await zip.file("xl/media/image1.png").async("uint8array");
assert.ok(mediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
assert.match(contentTypesXml, /table1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.table\+xml"/);
assert.match(contentTypesXml, /chart1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.drawingml\.chart\+xml"/);
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
assert.match(nativeOnlyWorkbook.inspect({ kind: "table", range: "A1:D4" }).ndjson, /"A"/);
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
