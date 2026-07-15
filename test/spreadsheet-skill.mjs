import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";
import { importXlsxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { nativeSpreadsheetRenderStatus, runSpreadsheetFixture, verifyWorkbookFile } from "../skills/spreadsheets/scripts/workflow.mjs";

const repoRoot = path.resolve(new URL("..", import.meta.url).pathname);
const fixturePath = path.join(repoRoot, "skills", "spreadsheets", "fixtures", "formula-summary.json");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-spreadsheet-skill-"));
const baselineDir = path.join(outputDir, "baselines");

try {
  const nativeAvailable = nativeSpreadsheetRenderStatus().available;
  const result = await runSpreadsheetFixture(fixturePath, {
    outputDir,
    renderFormat: "png",
    allSheets: true,
    nativeRender: nativeAvailable ? "required" : "off",
  });
  assert.equal(result.fixture.name, "formula-summary");
  assert.equal(result.roundtripCodec, "open-chestnut");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.renderFormat, "png");
  assert.equal(result.qa.summary.sheetRenders.length, 3);
  if (nativeAvailable) {
    assert.equal(result.qa.summary.nativeRender.status, "passed");
    assert.equal(result.qa.summary.nativeRender.pageCountMatches, true);
    assert.ok(result.qa.summary.nativeRender.pageCount >= 3);
  }
  for (const filePath of Object.values(result.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty skill output ${filePath}`);
  }
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(result.workbookPath));
  assert.equal(workbook.dateSystem, "1904");
  assert.equal(workbook.theme.name, "Agent Spreadsheet Theme");
  assert.equal(workbook.theme.colors.accent1, "#0F766E");
  assert.equal(workbook.worksheets.getItem("Summary").showGridLines, false);
  assert.deepEqual(workbook.worksheets.getItem("Summary").freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
  assert.ok(Math.abs(workbook.worksheets.getItem("Summary").getRange("B1:C4").format.columnWidthPx - 80) <= 1);
  assert.ok(workbook.worksheets.getItem("Summary").getRange("F1:F13").format.columnWidthPx > workbook.worksheets.getItem("Summary").getRange("G1:G13").format.columnWidthPx);
  assert.ok(Math.abs(workbook.worksheets.getItem("Summary").getRange("A1:D1").format.rowHeightPx - 28) < 0.01);
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("A1").format.border, {
    left: { style: "thin", color: "#334155" },
    right: { style: "medium", color: "#475569" },
    top: { style: "thin", color: "#64748B" },
    bottom: { style: "double", color: "#38BDF8" },
  });
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("B2:D4").values, [
    [100, 60, 0.4],
    [120, 70, 0.4166666666666667],
    [150, 90, 0.4],
  ]);
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("B2:D4").formulas, [
    ["=Inputs!B2", "=Inputs!C2", "=(B2-C2)/B2"],
    ["=Inputs!B3", "=Inputs!C3", "=(B3-C3)/B3"],
    ["=Inputs!B4", "=Inputs!C4", "=(B4-C4)/B4"],
  ]);
  assert.deepEqual(workbook.worksheets.getItem("DateTimeText").getRange("B2:B9").values, [[44650], [44650], [0.78125], [18], [45], [18], [1234.5], [0.125]]);
  assert.match(workbook.help("fx.DATEVALUE").ndjson, /ambiguous locale-numeric dates/);
  assert.match(workbook.help("fx.TIMEVALUE").ndjson, /fraction of one day/);
  assert.match(workbook.help("fx.VALUE").ndjson, /scientific notation/);
  assert.deepEqual(workbook.worksheets.getItem("Summary").mergedRanges, ["A15:G15"]);
  assert.equal(workbook.worksheets.getItem("Summary").getRange("A15").format.fill, "#0F766E");
  assert.equal(workbook.worksheets.getItem("Summary").getRange("A15").format.font.color, "#FFFFFF");
  const patternedStyle = workbook.worksheets.getItem("Summary").getRange("A14").format;
  assert.equal(patternedStyle.fill.patternType, "darkGrid");
  assert.equal(patternedStyle.fill.foreground.theme, 4);
  assert.equal(patternedStyle.fill.foreground.tint, 0.4);
  assert.equal(patternedStyle.fill.background, "#F8FAFC");
  assert.equal(patternedStyle.font.color.theme, 4);
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("B14").format.fill, patternedStyle.fill);
  assert.match(workbook.worksheets.getItem("Summary").toSvg(), /<pattern id="fill-13-0"/);
  const fixtureZip = await JSZip.loadAsync(await fs.readFile(result.workbookPath));
  assert.match(await fixtureZip.file("xl/styles.xml").async("text"), /patternType="darkGrid"><fgColor theme="4" tint="0.4"/);
  assert.match(await fixtureZip.file("xl/theme/theme1.xml").async("text"), /name="Agent Spreadsheet Theme"/);
  assert.match(await fixtureZip.file("xl/threadedComments/threadedComment1.xml").async("text"), /parentId="\{44444444-4444-4444-8444-444444444444\}"/);
  assert.match(await fixtureZip.file("xl/persons/person.xml").async("text"), /displayName="QA Reviewer"/);
  const nativeCommentZip = await JSZip.loadAsync(await fs.readFile(result.workbookPath));
  nativeCommentZip.remove("customXml/open-office-artifact.json");
  const nativeCommentWorkbook = await SpreadsheetFile.importXlsx(new FileBlob(await nativeCommentZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
  const nativeCommentThread = nativeCommentWorkbook.comments.threads.find((thread) => thread.target.address === "D2");
  assert.equal(nativeCommentThread?.resolved, true);
  assert.equal(nativeCommentThread?.comments[0].id, "{44444444-4444-4444-8444-444444444444}");
  assert.equal(nativeCommentThread?.comments[1].parentId, nativeCommentThread?.comments[0].id);
  assert.equal(nativeCommentThread?.comments[1].author, "QA Reviewer");
  const fixturePivotXml = await fixtureZip.file("xl/pivotTables/pivotTable1.xml").async("text");
  const fixturePivotCacheXml = await fixtureZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text");
  assert.match(fixturePivotXml, /x14:pivotFilter useWholeDay="0"/);
  assert.match(fixturePivotXml, /stringValue1="2026-03-31T17:00:00" stringValue2="2026-03-31T19:00:00"/);
  assert.match(fixturePivotCacheXml, /formula="IF\(AND\(ISNUMBER\('Revenue'\),'Revenue'&gt;='Cost',LEN\(TRIM\(&quot;  ok  &quot;\)\)=2,UPPER\(LEFT\(&quot;pass&quot;,1\)\)=&quot;P&quot;,LOWER\(RIGHT\(&quot;OK&quot;,1\)\)=&quot;k&quot;,MID\(&quot;margin&quot;,2,2\)=&quot;ar&quot;\),IFERROR\(ROUND\(SQRT\(POWER\('Revenue'-'Cost',2\)\)\/'Revenue',2\),0\),0\)"/);
  assert.match(fixturePivotCacheXml, /formula="IF\(AND\(DATE\(1904,1,1\)=0,YEAR\(DATE\(2026,3,31\)\)=2026,MONTH\(DATE\(2026,3,31\)\)=3,DAY\(DATE\(2026,3,31\)\)=31\),'Revenue'-'Cost',0\)"/);
  assert.match(fixturePivotCacheXml, /formula="IF\(AND\(DAY\(EDATE\(DATE\(2024,1,31\),1\)\)=29,DAY\(EOMONTH\(DATE\(2024,2,10\),0\)\)=29,DAYS\(DATE\(2024,3,1\),DATE\(2024,2,28\)\)=2,WEEKDAY\(DATE\(2026,3,31\),2\)=2\),'Revenue'-'Cost',0\)"/);
  assert.match(fixturePivotCacheXml, /formula="IF\(AND\(HOUR\(TIME\(27,0,0\)\)=3,MINUTE\(TIME\(0,750,0\)\)=30,SECOND\(TIME\(0,0,2000\)\)=20,HOUR\(DATE\(2026,3,31\)\+TIME\(18,45,30\)\)=18\),'Revenue'-'Cost',0\)"/);
  assert.match(fixturePivotCacheXml, /formula="IF\(AND\(NETWORKDAYS\(DATE\(2026,3,30\),DATE\(2026,4,5\),DATE\(2026,4,1\)\)=4,DAYS\(WORKDAY\(DATE\(2026,3,30\),5,DATE\(2026,4,3\)\),DATE\(2026,3,30\)\)=8,NETWORKDAYS.INTL\(DATE\(2026,3,30\),DATE\(2026,4,5\),11\)=6,DAYS\(WORKDAY.INTL\(DATE\(2026,3,30\),5,&quot;0000011&quot;\),DATE\(2026,3,30\)\)=7\),'Revenue'-'Cost',0\)"/);
  const summaryDrawings = workbook.worksheets.getItem("Summary");
  assert.equal(summaryDrawings.charts.items.length, 1);
  assert.equal(summaryDrawings.charts.items[0].title, "Quarter performance");
  assert.deepEqual(summaryDrawings.charts.items[0].categories, ["Jan", "Feb", "Mar"]);
  assert.deepEqual(summaryDrawings.charts.items[0].series.items[0].values, [100, 120, 150]);
  assert.equal(summaryDrawings.images.items.length, 1);
  assert.equal(summaryDrawings.images.items[0].alt, "Green status marker");
  assert.match(summaryDrawings.images.items[0].dataUrl, /^data:image\/svg\+xml;base64,/);
  assert.equal(workbook.resolve(summaryDrawings.charts.items[0].id), summaryDrawings.charts.items[0]);
  assert.equal(workbook.resolve(summaryDrawings.images.items[0].id), summaryDrawings.images.items[0]);
  assert.equal(summaryDrawings.pivotTables.items.length, 1);
  const summaryPivot = summaryDrawings.pivotTables.getItemOrNullObject("RevenuePivot");
  assert.deepEqual(summaryPivot.computedValues(), [["Period Year", "Period Quarter", "Period Month", "Month", "Period End", "Revenue total", "Margin rate", "Date contract", "Date shift contract", "Time contract", "Business day contract"], ["2026", "Q1", "Mar", "Mar", "2026-03-31T18:00:00Z", 150, 0.4, 60, 60, 60, 60]]);
  assert.deepEqual(summaryPivot.groupFields, [
    { name: "Period Year", sourceField: "Period End", groupBy: "years" },
    { name: "Period Quarter", sourceField: "Period End", groupBy: "quarters", parent: "Period Year" },
    { name: "Period Month", sourceField: "Period End", groupBy: "months", parent: "Period Quarter" },
  ]);
  assert.deepEqual(summaryPivot.filters, [{ field: "Month", include: ["Jan", "Mar"] }, { field: "Period End", type: "dateBetween", value1: "2026-03-31T17:00:00", value2: "2026-03-31T19:00:00", useWholeDay: false }]);
  assert.deepEqual(summaryPivot.calculatedFields, [
    { name: "Margin Rate", formula: '=IF(AND(ISNUMBER(\'Revenue\'),\'Revenue\'>=\'Cost\',LEN(TRIM("  ok  "))=2,UPPER(LEFT("pass",1))="P",LOWER(RIGHT("OK",1))="k",MID("margin",2,2)="ar"),IFERROR(ROUND(SQRT(POWER(\'Revenue\'-\'Cost\',2))/\'Revenue\',2),0),0)', numFmtId: 0, references: ["Revenue", "Cost"] },
    { name: "Date Contract", formula: "=IF(AND(DATE(1904,1,1)=0,YEAR(DATE(2026,3,31))=2026,MONTH(DATE(2026,3,31))=3,DAY(DATE(2026,3,31))=31),'Revenue'-'Cost',0)", numFmtId: 0, references: ["Revenue", "Cost"] },
    { name: "Date Shift Contract", formula: "=IF(AND(DAY(EDATE(DATE(2024,1,31),1))=29,DAY(EOMONTH(DATE(2024,2,10),0))=29,DAYS(DATE(2024,3,1),DATE(2024,2,28))=2,WEEKDAY(DATE(2026,3,31),2)=2),'Revenue'-'Cost',0)", numFmtId: 0, references: ["Revenue", "Cost"] },
    { name: "Time Contract", formula: "=IF(AND(HOUR(TIME(27,0,0))=3,MINUTE(TIME(0,750,0))=30,SECOND(TIME(0,0,2000))=20,HOUR(DATE(2026,3,31)+TIME(18,45,30))=18),'Revenue'-'Cost',0)", numFmtId: 0, references: ["Revenue", "Cost"] },
    { name: "Business Day Contract", formula: '=IF(AND(NETWORKDAYS(DATE(2026,3,30),DATE(2026,4,5),DATE(2026,4,1))=4,DAYS(WORKDAY(DATE(2026,3,30),5,DATE(2026,4,3)),DATE(2026,3,30))=8,NETWORKDAYS.INTL(DATE(2026,3,30),DATE(2026,4,5),11)=6,DAYS(WORKDAY.INTL(DATE(2026,3,30),5,"0000011"),DATE(2026,3,30))=7),\'Revenue\'-\'Cost\',0)', numFmtId: 0, references: ["Revenue", "Cost"] },
  ]);
  assert.deepEqual(nativeCommentWorkbook.resolve("RevenuePivot").calculatedFields, summaryPivot.calculatedFields);
  const nativePivotSecondZip = await JSZip.loadAsync(new Uint8Array(await (await SpreadsheetFile.exportXlsx(nativeCommentWorkbook)).arrayBuffer()));
  assert.match(await nativePivotSecondZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /formula="IF\(AND\(ISNUMBER\('Revenue'\),'Revenue'&gt;='Cost',LEN\(TRIM\(&quot;  ok  &quot;\)\)=2,UPPER\(LEFT\(&quot;pass&quot;,1\)\)=&quot;P&quot;,LOWER\(RIGHT\(&quot;OK&quot;,1\)\)=&quot;k&quot;,MID\(&quot;margin&quot;,2,2\)=&quot;ar&quot;\),IFERROR\(ROUND\(SQRT\(POWER\('Revenue'-'Cost',2\)\)\/'Revenue',2\),0\),0\)"/);
  assert.match(await nativePivotSecondZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /formula="IF\(AND\(DATE\(1904,1,1\)=0,YEAR\(DATE\(2026,3,31\)\)=2026,MONTH\(DATE\(2026,3,31\)\)=3,DAY\(DATE\(2026,3,31\)\)=31\),'Revenue'-'Cost',0\)"/);
  assert.match(await nativePivotSecondZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /formula="IF\(AND\(DAY\(EDATE\(DATE\(2024,1,31\),1\)\)=29,DAY\(EOMONTH\(DATE\(2024,2,10\),0\)\)=29,DAYS\(DATE\(2024,3,1\),DATE\(2024,2,28\)\)=2,WEEKDAY\(DATE\(2026,3,31\),2\)=2\),'Revenue'-'Cost',0\)"/);
  assert.match(await nativePivotSecondZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /formula="IF\(AND\(HOUR\(TIME\(27,0,0\)\)=3,MINUTE\(TIME\(0,750,0\)\)=30,SECOND\(TIME\(0,0,2000\)\)=20,HOUR\(DATE\(2026,3,31\)\+TIME\(18,45,30\)\)=18\),'Revenue'-'Cost',0\)"/);
  assert.match(await nativePivotSecondZip.file("xl/pivotCache/pivotCacheDefinition1.xml").async("text"), /formula="IF\(AND\(NETWORKDAYS\(DATE\(2026,3,30\),DATE\(2026,4,5\),DATE\(2026,4,1\)\)=4,DAYS\(WORKDAY\(DATE\(2026,3,30\),5,DATE\(2026,4,3\)\),DATE\(2026,3,30\)\)=8,NETWORKDAYS.INTL\(DATE\(2026,3,30\),DATE\(2026,4,5\),11\)=6,DAYS\(WORKDAY.INTL\(DATE\(2026,3,30\),5,&quot;0000011&quot;\),DATE\(2026,3,30\)\)=7\),'Revenue'-'Cost',0\)"/);
  assert.equal(summaryPivot.refreshPolicy.refreshOnLoad, false);
  assert.equal(summaryPivot.refreshPolicy.refreshedBy, "Spreadsheet skill");
  assert.equal(workbook.resolve("RevenuePivot"), summaryPivot);
  assert.match(workbook.help("workbook.structuredReferences").ndjson, /space intersection/);
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("G10:G13").values, [[0], [43889], [5], [2]]);
  assert.equal(workbook.worksheets.getItem("Summary").getRange("G14").values[0][0], 1);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /SummaryTable/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /Inputs!B2/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"freezePanes":\{"rows":1,"columns":1/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"customColumns":9/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"kind":"mergedCell"[\s\S]*"range":"A15:G15"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"drawingType":"chart"[\s\S]*"title":"Quarter performance"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"drawingType":"image"[\s\S]*"alt":"Green status marker"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"kind":"pivotTable"[\s\S]*"name":"RevenuePivot"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"filters":\[\{"field":"Month","include":\["Jan","Mar"\]/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"field":"Period End","type":"dateBetween"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"useWholeDay":false/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"groupBy":"quarters"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"calculatedFields":\[\{"name":"Margin Rate"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"name":"Margin Rate"/);
  assert.match(result.qa.workbook.inspect({ kind: "thread", maxChars: 4000 }).ndjson, /Margin evidence is approved/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /xl\/workbook\.xml/);
  assert.equal(result.qa.packageInspect.records[0].sheets, 3);
  assert.equal(result.qa.packageInspect.records[0].semanticIssues, 0);
  assert.deepEqual([...(await fs.readFile(result.qa.summary.files.preview)).subarray(0, 4)], [0x89, 0x50, 0x4e, 0x47]);

  const wasmResult = await runSpreadsheetFixture(path.join(repoRoot, "skills", "spreadsheets", "fixtures", "open-chestnut-basic.json"), {
    outputDir: path.join(outputDir, "open-chestnut-basic"),
    nativeRender: nativeSpreadsheetRenderStatus().available ? "required" : "off",
  });
  assert.equal(wasmResult.codec, "open-chestnut");
  assert.equal(wasmResult.roundtripCodec, "open-chestnut");
  assert.equal(wasmResult.qa.summary.packageOk, true);
  assert.equal(wasmResult.qa.summary.verifyOk, true);
  assert.equal(wasmResult.qa.summary.visualQaOk, true);
  assert.equal(wasmResult.qa.summary.renderFormat, "png");
  const wasmWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(wasmResult.workbookPath));
  assert.equal(wasmWorkbook.dateSystem, "1904");
  assert.deepEqual(wasmWorkbook.calculation, {
    mode: "automaticExceptTables",
    calculateOnSave: false,
    fullCalculationOnLoad: true,
    forceFullCalculation: true,
    iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 },
    fullPrecision: false,
  });
  assert.equal(wasmWorkbook.theme.name, "OpenChestnut Fixture");
  assert.equal(wasmWorkbook.theme.colors.accent1, "#0F766E");
  assert.deepEqual(wasmWorkbook.worksheets.items.map((item) => item.visibility), ["visible", "hidden", "visible", "visible", "veryHidden"]);
  assert.equal(wasmWorkbook.worksheets.getActiveWorksheet().name, "Icon Rules");
  assert.deepEqual(wasmWorkbook.worksheets.getSelectedWorksheets().map((sheet) => sheet.name), ["Summary", "Icon Rules"]);
  assert.equal(wasmWorkbook.windows.count, 2);
  assert.equal(wasmWorkbook.windows.getItemAt(1).getActiveWorksheet().name, "Advanced Filters");
  assert.deepEqual(wasmWorkbook.windows.getItemAt(1).getSelectedWorksheets().map((sheet) => sheet.name), ["Summary", "Advanced Filters"]);
  assert.deepEqual(wasmWorkbook.definedNames.items.map((item) => ({ name: item.name, refersTo: item.refersTo, scope: item.scope, comment: item.comment, hidden: item.hidden })), [
    { name: "SummaryValues", refersTo: "Summary!$B$2:$B$3", scope: undefined, comment: "Fixture summary values", hidden: false },
    { name: "DetailScores", refersTo: "Details!$C$2:$C$3", scope: "Details", comment: undefined, hidden: true },
  ]);
  assert.match(wasmWorkbook.inspect({ kind: "definedName", target: "SummaryValues" }).ndjson, /"hidden":false/);
  assert.deepEqual(wasmWorkbook.worksheets.getItem("Summary").getRange("B2:D3").values, [[42.5, 85, 127.5], [85, 170, null]]);
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").getRange("B2").format.numberFormat, "0.000 \"units\"");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").getRange("B3").format.numberFormat, "0.00%");
  const wasmHeaderStyle = wasmWorkbook.worksheets.getItem("Summary").getRange("A1").format;
  assert.equal(wasmHeaderStyle.font.bold, true);
  assert.equal(wasmHeaderStyle.font.underline, "double");
  assert.equal(wasmHeaderStyle.fill.patternType, "darkGrid");
  assert.equal(wasmHeaderStyle.fill.foreground.theme, 4);
  assert.equal(wasmHeaderStyle.fill.foreground.tint, 0.4);
  assert.equal(wasmHeaderStyle.border.bottom.style, "double");
  assert.deepEqual(wasmHeaderStyle.protection, { locked: false, hidden: true });
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").getRange("C2").format.fill, "#DCFCE7");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").getRange("C2").format.font.bold, true);
  assert.deepEqual(wasmWorkbook.worksheets.getItem("Summary").mergedRanges, ["A4:B4"]);
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").images.items.length, 3);
  const wasmImage = wasmWorkbook.worksheets.getItem("Summary").images.items[0];
  assert.equal(wasmImage.name, "OpenChestnut mark");
  assert.equal(wasmImage.alt, "OpenChestnut worksheet image");
  assert.equal(wasmImage.dataUrl, "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC");
  assert.deepEqual(wasmImage.anchor, { from: { row: 4, col: 0, rowOffsetPx: 6, colOffsetPx: 4 }, extent: { widthPx: 96, heightPx: 64 } });
  const wasmTwoCellImage = wasmWorkbook.worksheets.getItem("Summary").images.items[1];
  assert.equal(wasmTwoCellImage.name, "OpenChestnut two-cell mark");
  assert.equal(wasmTwoCellImage.alt, "OpenChestnut move-only two-cell image");
  assert.deepEqual(wasmTwoCellImage.anchor, {
    type: "twoCell",
    from: { row: 4, col: 2, rowOffsetPx: 5, colOffsetPx: 3 },
    to: { row: 7, col: 5, rowOffsetPx: 9, colOffsetPx: 7 },
    editAs: "oneCell",
  });
  const wasmAbsoluteImage = wasmWorkbook.worksheets.getItem("Summary").images.items[2];
  assert.equal(wasmAbsoluteImage.name, "OpenChestnut absolute mark");
  assert.equal(wasmAbsoluteImage.alt, "OpenChestnut page-relative image");
  assert.deepEqual(wasmAbsoluteImage.anchor, {
    type: "absolute",
    position: { leftPx: 350, topPx: 100 },
    extent: { widthPx: 80, heightPx: 60 },
  });
  assert.match(await fs.readFile(wasmResult.qa.summary.files.inspect, "utf8"), /"drawingType":"image"[\s\S]*"alt":"OpenChestnut worksheet image"/);
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("C2").formulaType, "shared");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("C3").sharedIndex, 7);
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("C3").sharedRef, "C2:C3");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("D2").formulaType, "array");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("D2").arrayRef, "D2:D3");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("E2").formulaType, "dynamicArray");
  assert.equal(wasmWorkbook.worksheets.getItem("Summary").store.get("E2").dynamicArrayRef, "E2:F3");
  assert.deepEqual(wasmWorkbook.worksheets.getItem("Summary").getRange("E2:F3").values, [[101, 102], [103, 104]]);
  assert.deepEqual(wasmWorkbook.worksheets.getItem("Summary").sortState, {
    reference: "A1:D3",
    caseSensitive: true,
    sortMethod: "stroke",
    columnSort: true,
    conditions: [{ reference: "A3:D3", descending: true, customList: "Double,Revenue" }, { reference: "A1:D1", descending: false }],
  });
  const wasmTable = wasmWorkbook.worksheets.getItem("Details").tables.getItemOrNullObject("CodecTable");
  assert.equal(wasmTable.isNullObject, undefined);
  assert.equal(wasmTable.range, "A1:C4");
  assert.equal(wasmTable.style, "TableStyleMedium4");
  assert.equal(wasmTable.showTotals, true);
  assert.equal(wasmTable.showFirstColumn, true);
  assert.equal(wasmTable.showRowStripes, false);
  assert.equal(wasmTable.showBandedColumns, true);
  assert.deepEqual(wasmTable.columnNames, ["Codec", "Runtime", "Score"]);
  assert.equal(wasmTable.columnDefinitions[0].totalsRowLabel, "Total");
  assert.equal(wasmTable.columnDefinitions[2].calculatedColumnFormula, "=LEN([@Codec])");
  assert.equal(wasmTable.columnDefinitions[2].totalsRowFormula, "=SUBTOTAL(109,[Score])");
  assert.deepEqual(wasmTable.filters, [
    { columnIndex: 0, kind: "values", values: ["OpenChestnut"], includeBlank: false },
    { columnIndex: 2, kind: "custom", matchAll: true, criteria: [{ operator: "greaterThanOrEqual", value: "5" }, { operator: "lessThanOrEqual", value: "20" }] },
  ]);
  assert.deepEqual(wasmTable.sortState, {
    reference: "A2:C3",
    caseSensitive: true,
    sortMethod: "stroke",
    conditions: [{ reference: "C2:C3", descending: true, customList: "20,10" }, { reference: "A2:A3", descending: false }],
  });
  const wasmAdvancedFilterTable = wasmWorkbook.worksheets.getItem("Advanced Filters").tables.getItemOrNullObject("AdvancedFilterTable");
  assert.deepEqual(wasmAdvancedFilterTable.filters, [
    { columnIndex: 0, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 15 }] },
    { columnIndex: 1, kind: "dynamic", type: "today", value: 45853, maxValue: 45854 },
    { columnIndex: 2, kind: "top10", top: true, percent: true, value: 10, filterValue: 95 },
  ]);
  const wasmIconTable = wasmWorkbook.worksheets.getItem("Icon Rules").tables.getItemOrNullObject("IconRuleTable");
  assert.deepEqual(wasmIconTable.filters, [
    { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 0 },
    { columnIndex: 1, kind: "icon", iconSet: "3Flags" },
  ]);
  assert.deepEqual(wasmIconTable.sortState, {
    reference: "A2:B3",
    caseSensitive: false,
    conditions: [
      { reference: "B2:B3", descending: true, kind: "icon", iconSet: "5Rating", iconId: 4 },
      { reference: "A2:A3", descending: false, kind: "icon", iconSet: "3Symbols2" },
    ],
  });
  const wasmColorTable = wasmWorkbook.worksheets.getItem("Color Rules").tables.getItemOrNullObject("ColorRuleTable");
  assert.deepEqual(wasmColorTable.filters, [
    { columnIndex: 0, kind: "color", target: "cell", color: "#E11D48" },
    { columnIndex: 1, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
  ]);
  assert.deepEqual(wasmColorTable.sortState, {
    reference: "A2:B3",
    caseSensitive: false,
    conditions: [
      { reference: "B2:B3", descending: true, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
      { reference: "A2:A3", descending: false, kind: "color", target: "cell", color: "#E11D48" },
    ],
  });
  const wasmZip = await JSZip.loadAsync(await fs.readFile(wasmResult.workbookPath));
  const wasmThemeXml = await wasmZip.file("xl/theme/theme1.xml").async("text");
  assert.match(wasmThemeXml, /name="OpenChestnut Fixture"/);
  assert.match(wasmThemeXml, /<a:accent1><a:srgbClr val="0F766E"/);
  const wasmStylesXml = await wasmZip.file("xl/styles.xml").async("text");
  assert.match(wasmStylesXml, /0\.000 &quot;units&quot;/);
  assert.match(wasmStylesXml, /patternType="darkGrid"/);
  assert.match(wasmStylesXml, /theme="4" tint="0\.4"/);
  assert.match(wasmStylesXml, /style="double"/);
  const wasmWorksheetXml = await wasmZip.file("xl/worksheets/sheet1.xml").async("text");
  assert.match(wasmWorksheetXml, /<x:f t="shared" ref="C2:C3" si="7">B2\*2<\/x:f>/);
  assert.match(wasmWorksheetXml, /<x:f t="shared" si="7"\s*\/>/);
  assert.match(wasmWorksheetXml, /<x:f t="array" ref="D2:D3">SUM\(B2:B3\)<\/x:f>/);
  assert.match(wasmWorksheetXml, /<x:c\b(?=[^>]*\br="E2")(?=[^>]*\bcm="1")[^>]*><x:f t="array" ref="E2:F3">SEQUENCE\(2,2,101,1\)<\/x:f>/);
  const wasmDynamicMetadataPath = Object.keys(wasmZip.files).find((partPath) => /^xl\/metadata\d*\.xml$/.test(partPath));
  assert.ok(wasmDynamicMetadataPath, "runnable OpenChestnut fixture must retain its dynamic-array metadata part");
  assert.match(await wasmZip.file(wasmDynamicMetadataPath).async("text"), /dynamicArrayProperties\b[^>]*fDynamic="1"[^>]*fCollapsed="0"/);
  assert.match(wasmWorksheetXml, /<x:sortState\b[^>]*ref="A1:D3"/);
  assert.match(wasmWorksheetXml, /<x:sortState\b[^>]*columnSort="1"/);
  assert.match(wasmWorksheetXml, /<x:sortCondition\b[^>]*ref="A3:D3"[^>]*customList="Double,Revenue"/);
  const wasmTableXml = await wasmZip.file("xl/tables/table1.xml").async("text");
  assert.match(wasmTableXml, /displayName="CodecTable"/);
  assert.match(wasmTableXml, /ref="A1:C4"/);
  assert.match(wasmTableXml, /<x:tableColumn id="1" name="Codec"/);
  assert.match(wasmTableXml, /totalsRowLabel="Total"/);
  assert.match(wasmTableXml, /<x:calculatedColumnFormula>LEN\(\[@Codec\]\)<\/x:calculatedColumnFormula>/);
  assert.match(wasmTableXml, /<x:totalsRowFormula>SUBTOTAL\(109,\[Score\]\)<\/x:totalsRowFormula>/);
  assert.match(wasmTableXml, /<x:filterColumn colId="0"><x:filters><x:filter val="OpenChestnut"\s*\/><\/x:filters><\/x:filterColumn>/);
  assert.match(wasmTableXml, /<x:customFilters and="1"><x:customFilter operator="greaterThanOrEqual" val="5"\s*\/><x:customFilter operator="lessThanOrEqual" val="20"\s*\/><\/x:customFilters>/);
  assert.match(wasmTableXml, /<x:sortState ref="A2:C3" caseSensitive="1" sortMethod="stroke"><x:sortCondition ref="C2:C3" descending="1" customList="20,10"\s*\/><x:sortCondition ref="A2:A3"\s*\/><\/x:sortState>/);
  assert.match(wasmTableXml, /showFirstColumn="1"/);
  assert.match(wasmTableXml, /showRowStripes="0"/);
  assert.match(wasmTableXml, /showColumnStripes="1"/);
  const wasmAdvancedFilterXml = await wasmZip.file("xl/tables/table2.xml").async("text");
  assert.match(wasmAdvancedFilterXml, /<x:filters calendarType="gregorian"><x:dateGroupItem year="2026" dateTimeGrouping="day" month="7" day="15"\s*\/><\/x:filters>/);
  assert.match(wasmAdvancedFilterXml, /<x:dynamicFilter type="today" val="45853" maxVal="45854"\s*\/>/);
  assert.match(wasmAdvancedFilterXml, /<x:top10 top="1" percent="1" val="10" filterVal="95"\s*\/>/);
  const wasmIconRuleXml = await wasmZip.file("xl/tables/table3.xml").async("text");
  assert.match(wasmIconRuleXml, /<x:iconFilter iconSet="3Arrows" iconId="0"\s*\/>/);
  assert.match(wasmIconRuleXml, /<x:iconFilter iconSet="3Flags"\s*\/>/);
  assert.match(wasmIconRuleXml, /<x:sortCondition ref="B2:B3" descending="1" sortBy="icon" iconSet="5Rating" iconId="4"\s*\/>/);
  assert.match(wasmIconRuleXml, /<x:sortCondition ref="A2:A3" sortBy="icon" iconSet="3Symbols2"\s*\/>/);
  const wasmColorRuleXml = await wasmZip.file("xl/tables/table4.xml").async("text");
  assert.match(wasmColorRuleXml, /<x:colorFilter dxfId="0" cellColor="1"\s*\/>/);
  assert.match(wasmColorRuleXml, /<x:colorFilter dxfId="1" cellColor="0"\s*\/>/);
  assert.match(wasmColorRuleXml, /<x:sortCondition ref="B2:B3" descending="1" sortBy="fontColor" dxfId="1"\s*\/>/);
  assert.match(wasmColorRuleXml, /<x:sortCondition ref="A2:A3" sortBy="cellColor" dxfId="0"\s*\/>/);
  assert.match(await wasmZip.file("xl/worksheets/sheet2.xml").async("text"), /<x:tableParts count="1"><x:tablePart\b[^>]*r:id="rIdTable1"\s*\/><\/x:tableParts>/);
  assert.match(await wasmZip.file("xl/worksheets/_rels/sheet2.xml.rels").async("text"), /Type="[^"]+\/table" Target="(?:\/xl|\.\.)\/tables\/table1\.xml"/);
  assert.match(await wasmZip.file("xl/worksheets/_rels/sheet3.xml.rels").async("text"), /Type="[^"]+\/table" Target="(?:\/xl|\.\.)\/tables\/table2\.xml"/);
  assert.match(await wasmZip.file("xl/worksheets/_rels/sheet4.xml.rels").async("text"), /Type="[^"]+\/table" Target="(?:\/xl|\.\.)\/tables\/table3\.xml"/);
  assert.match(await wasmZip.file("xl/worksheets/_rels/sheet5.xml.rels").async("text"), /Type="[^"]+\/table" Target="(?:\/xl|\.\.)\/tables\/table4\.xml"/);
  if (nativeSpreadsheetRenderStatus().available) assert.equal(wasmResult.qa.summary.nativeRender.status, "passed");

  const queryResult = await runSpreadsheetFixture(path.join(repoRoot, "skills", "spreadsheets", "fixtures", "open-chestnut-query-table.json"), {
    outputDir: path.join(outputDir, "open-chestnut-query-table"),
  });
  assert.equal(queryResult.codec, "open-chestnut");
  assert.equal(queryResult.roundtripCodec, "open-chestnut");
  assert.equal(queryResult.sourceQueryTable.query.name, "Warehouse sales refreshed");
  assert.equal(queryResult.sourceQueryTable.query.disableRefresh, true);
  assert.equal(queryResult.sourceQueryTable.query.backgroundRefresh, false);
  assert.equal(queryResult.sourceQueryTable.query.refresh.preserveSortFilterLayout, false);
  assert.equal(queryResult.sourceQueryTable.query.refresh.headersInLastRefresh, false);
  assert.equal(queryResult.sourceQueryTable.query.refresh.minimumVersion, 1);
  assert.deepEqual(queryResult.sourceQueryTable.query.refresh.fields, [
    { id: 1, name: "Territory", dataBound: false, fillFormulas: true, clipped: false, tableColumnId: 1 },
    { id: 2, name: "Revenue", dataBound: true, clipped: true, tableColumnId: 2 },
  ]);
  assert.deepEqual(queryResult.sourceQueryTable.query.refresh.deletedFieldNames, ["Legacy Territory", "Legacy Revenue"]);
  assert.deepEqual(queryResult.sourceQueryTable.query.refresh.sortState, {
    reference: "A2:B3",
    caseSensitive: false,
    sortMethod: "pinYin",
    columnSort: true,
    conditions: [
      { reference: "A3:B3", descending: false, customList: "80,120" },
      { reference: "A2:B2", descending: false, kind: "icon", iconSet: "3Arrows", iconId: 1 },
    ],
  });
  assert.deepEqual(queryResult.sourceConnections, [{
    connectionId: 7,
    name: "Fixture warehouse curated",
    type: 5,
    refreshedVersion: 8,
    description: "Curated without executing the source",
    keepAlive: true,
    intervalMinutes: 45,
    background: false,
    refreshOnLoad: true,
    saveData: false,
  }]);
  assert.equal(queryResult.qa.summary.packageOk, true);
  assert.equal(queryResult.qa.summary.verifyOk, true);
  assert.equal(queryResult.qa.summary.visualQaOk, true);
  assert.equal(queryResult.qa.summary.nativeRender.status, "skipped");
  assert.match(queryResult.qa.inspect.ndjson, /"kind":"connection"/);
  const queryWorkbook = await importXlsxWithOpenChestnut(await FileBlob.load(queryResult.workbookPath));
  assert.deepEqual(queryWorkbook.connections, queryResult.sourceConnections);
  const runnableQueryTable = queryWorkbook.worksheets.getItem("External Data").tables.getItemOrNullObject("ExternalSales");
  assert.deepEqual(runnableQueryTable.queryTable, {
    name: "Warehouse sales refreshed",
    connectionId: 7,
    headers: true,
    rowNumbers: false,
    disableRefresh: true,
    backgroundRefresh: false,
    firstBackgroundRefresh: false,
    refreshOnLoad: false,
    growShrinkType: "insertClear",
    fillFormulas: false,
    removeDataOnSave: false,
    disableEdit: false,
    preserveFormatting: true,
    adjustColumnWidth: true,
    intermediate: false,
    autoFormatId: 3,
    applyFontFormats: true,
    refresh: {
      preserveSortFilterLayout: false,
      fieldIdWrapped: false,
      headersInLastRefresh: false,
      minimumVersion: 1,
      nextId: 3,
      unboundColumnsLeft: 0,
      unboundColumnsRight: 0,
      fields: [
        { id: 1, name: "Territory", dataBound: false, fillFormulas: true, clipped: false, tableColumnId: 1 },
        { id: 2, name: "Revenue", dataBound: true, clipped: true, tableColumnId: 2 },
      ],
      deletedFieldNames: ["Legacy Territory", "Legacy Revenue"],
      sortState: {
        reference: "A2:B3",
        caseSensitive: false,
        sortMethod: "pinYin",
        columnSort: true,
        conditions: [
          { reference: "A3:B3", descending: false, customList: "80,120" },
          { reference: "A2:B2", descending: false, kind: "icon", iconSet: "3Arrows", iconId: 1 },
        ],
      },
    },
  });
  const queryZip = await JSZip.loadAsync(await fs.readFile(queryResult.workbookPath));
  const runnableConnectionXml = await queryZip.file("xl/connections.xml").async("text");
  assert.match(runnableConnectionXml, /name="Fixture warehouse curated"/);
  assert.match(runnableConnectionXml, /description="Curated without executing the source"/);
  assert.match(runnableConnectionXml, /keepAlive="1"/);
  assert.match(runnableConnectionXml, /interval="45"/);
  assert.match(runnableConnectionXml, /background="0"/);
  assert.match(runnableConnectionXml, /refreshOnLoad="1"/);
  assert.match(runnableConnectionXml, /saveData="0"/);
  assert.match(runnableConnectionXml, /Provider=Fixture.Provider;Data Source=fixture.invalid/);
  assert.match(runnableConnectionXml, /savePassword="0"/);
  assert.match(runnableConnectionXml, /credentials="integrated"/);
  assert.match(runnableConnectionXml, /<fixture:connectionOpaque value="kept"/);
  const runnableQueryXml = await queryZip.file("xl/queryTables/queryTable1.xml").async("text");
  assert.match(runnableQueryXml, /disableRefresh="1"/);
  assert.match(runnableQueryXml, /backgroundRefresh="0"/);
  assert.match(runnableQueryXml, /preserveSortFilterLayout="0"/);
  assert.match(runnableQueryXml, /headersInLastRefresh="0"/);
  assert.match(runnableQueryXml, /minimumVersion="1"/);
  assert.match(runnableQueryXml, /id="1" name="Territory" dataBound="0" tableColumnId="1" fillFormulas="1" clipped="0"/);
  assert.match(runnableQueryXml, /id="2" name="Revenue" dataBound="1" tableColumnId="2" clipped="1"/);
  assert.match(runnableQueryXml, /<x:queryTableFields count="2">/);
  assert.match(runnableQueryXml, /<x:deletedField name="Legacy Territory"/);
  assert.match(runnableQueryXml, /<x:deletedField name="Legacy Revenue"/);
  assert.match(runnableQueryXml, /<x:sortState ref="A2:B3" sortMethod="pinYin" columnSort="1">/);
  assert.match(runnableQueryXml, /<x:sortCondition ref="A3:B3" customList="80,120"/);
  assert.match(runnableQueryXml, /<x:sortCondition ref="A2:B2" sortBy="icon" iconSet="3Arrows" iconId="1"/);
  assert.match(runnableQueryXml, /<fixture:fieldOpaque value="kept"/);
  assert.match(runnableQueryXml, /<fixture:sortOpaque value="kept"/);
  assert.match(runnableQueryXml, /<fixture:opaque value="kept"/);
  assert.match(await fs.readFile(queryResult.qa.summary.files.packageInspect, "utf8"), /xl\/queryTables\/queryTable1\.xml/);

  const secondQa = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "second-qa"),
    sheetName: "Inputs",
    range: "A1:C4",
    renderFormat: "svg",
    nativeRender: "off",
  });
  assert.equal(secondQa.summary.verifyOk, true);
  assert.equal(secondQa.summary.packageOk, true);
  assert.equal(secondQa.summary.sheetName, "Inputs");
  await assert.rejects(
    () => verifyWorkbookFile(result.workbookPath, {
      outputDir: path.join(outputDir, "missing-baseline"),
      sheetName: "Summary",
      range: "A1:D4",
      renderFormat: "png",
      baselineDir: path.join(outputDir, "not-initialized-baselines"),
      nativeRender: "off",
    }),
    /Visual baseline is missing.*writeBaseline=true/,
  );

  const nativeStatus = nativeSpreadsheetRenderStatus();
  if (nativeStatus.commands.soffice) {
    const libreOffice = createLibreOfficeRenderer();
    const ods = await libreOffice({ input: await FileBlob.load(result.workbookPath), inputType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", outputType: "application/vnd.oasis.opendocument.spreadsheet", format: "ods", artifactKind: "workbook" });
    const libreOfficeXlsx = await libreOffice({ input: ods, inputType: "application/vnd.oasis.opendocument.spreadsheet", outputType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", format: "xlsx", artifactKind: "workbook", libreOffice: { inputExtension: "ods" } });
    const libreOfficeWorkbook = await SpreadsheetFile.importXlsx(libreOfficeXlsx);
    const libreOfficeSummary = libreOfficeWorkbook.worksheets.getItem("Summary");
    const displayCells = libreOfficeSummary.layoutJson({ range: "B2:D4" }).cells;
    assert.equal(displayCells.find((cell) => cell.address === "B2").displayValue, "$100");
    assert.equal(displayCells.find((cell) => cell.address === "D2").displayValue, "40.0%");
    assert.equal(libreOfficeSummary.getRange("A15").format.fill, "#0F766E");
    assert.equal(libreOfficeSummary.getRange("A15").format.alignment.horizontal, "center");
    assert.equal(libreOfficeSummary.getRange("A1").format.border.bottom.style, "double");
    assert.equal(libreOfficeSummary.getRange("G14").values[0][0], 1);
    assert.deepEqual(libreOfficeWorkbook.worksheets.getItem("DateTimeText").getRange("B2:B9").values, [[44650], [44650], [0.78125], [18], [45], [18], [1234.5], [0.125]]);
    assert.ok(libreOfficeSummary.charts.items.length >= 1);
    assert.ok(libreOfficeSummary.images.items.length >= 1);
    assert.ok(libreOfficeSummary.pivotTables.items.some((pivot) => pivot.name === "RevenuePivot"));
  }
  const intersectionResult = await runSpreadsheetFixture(path.join(repoRoot, "skills", "spreadsheets", "fixtures", "structured-intersection.json"), {
    outputDir: path.join(outputDir, "structured-intersection"),
    nativeRender: "off",
  });
  const intersectionWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(intersectionResult.workbookPath));
  assert.deepEqual(intersectionWorkbook.worksheets.getItem("Intersection").getRange("D2:D3").values, [[370], ["#NULL!"]]);
  assert.deepEqual(intersectionWorkbook.trace("Intersection!D2").tree.precedents.map((node) => node.address), ["B2", "B3", "B4"]);
  assert.match(await fs.readFile(intersectionResult.qa.summary.files.inspect, "utf8"), /IntersectionTable\[\[Month\]:\[Revenue\]\] IntersectionTable\[\[Revenue\]:\[Cost\]\]/);
  const intersectionZip = await JSZip.loadAsync(await fs.readFile(intersectionResult.workbookPath));
  assert.match(await intersectionZip.file("xl/worksheets/sheet1.xml").async("text"), /IntersectionTable\[\[Month\]:\[Revenue\]\] IntersectionTable\[\[Revenue\]:\[Cost\]\]/);
  if (nativeStatus.commands.soffice) {
    const libreOffice = createLibreOfficeRenderer();
    const intersectionXlsx = await FileBlob.load(intersectionResult.workbookPath);
    const ods = await libreOffice({ input: intersectionXlsx, inputType: intersectionXlsx.type, outputType: "application/vnd.oasis.opendocument.spreadsheet", format: "ods", artifactKind: "workbook" });
    const rewrittenXlsx = await libreOffice({ input: ods, inputType: "application/vnd.oasis.opendocument.spreadsheet", outputType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", format: "xlsx", artifactKind: "workbook", libreOffice: { inputExtension: "ods" } });
    const rewrittenWorkbook = await SpreadsheetFile.importXlsx(rewrittenXlsx);
    assert.deepEqual(rewrittenWorkbook.worksheets.getItem("Intersection").getRange("D2:D3").values, [[370], ["#NULL!"]]);
    assert.match(rewrittenWorkbook.worksheets.getItem("Intersection").getRange("D2").formulas[0][0], /\$A\$2:\$B\$4 Intersection!\$B\$2:\$C\$4/);
    assert.deepEqual(rewrittenWorkbook.trace("Intersection!D2").tree.precedents.map((node) => node.address), ["B2", "B3", "B4"]);
  }
  const csvPath = path.join(outputDir, "summary.csv");
  await (await SpreadsheetFile.exportCsv(workbook, { sheetName: "Summary", range: "A1:G15" })).save(csvPath);
  const csvQa = await verifyWorkbookFile(csvPath, {
    outputDir: path.join(outputDir, "csv-qa"),
    sheetName: "Summary",
    range: "A1:G15",
    renderFormat: "svg",
    nativeRender: nativeStatus.available ? "required" : "off",
    coerceTypes: true,
  });
  assert.equal(csvQa.summary.inputFormat, "csv");
  assert.equal(csvQa.summary.inputType, "text/csv");
  assert.equal(csvQa.packageInspect.summary.kind, "delimitedFile");
  assert.equal(csvQa.packageInspect.summary.rows, 15);
  assert.equal(csvQa.summary.verifyOk, true);
  assert.equal(csvQa.summary.visualQaOk, true);
  assert.match(await fs.readFile(csvQa.summary.files.packageInspect, "utf8"), /delimitedRow/);
  if (nativeStatus.available) assert.equal(csvQa.summary.nativeRender.status, "passed");

  const tsvPath = path.join(outputDir, "summary.tsv");
  await (await SpreadsheetFile.exportTsv(workbook, { sheetName: "Summary", range: "A1:G15" })).save(tsvPath);
  const tsvQa = await verifyWorkbookFile(tsvPath, {
    outputDir: path.join(outputDir, "tsv-qa"),
    sheetName: "Summary",
    range: "A1:G15",
    renderFormat: "svg",
    nativeRender: nativeStatus.available ? "required" : "off",
    coerceTypes: true,
  });
  assert.equal(tsvQa.summary.inputFormat, "tsv");
  assert.equal(tsvQa.summary.inputType, "text/tab-separated-values");
  assert.equal(tsvQa.packageInspect.summary.columns, 7);
  assert.equal(tsvQa.summary.verifyOk, true);
  if (nativeStatus.available) assert.equal(tsvQa.summary.nativeRender.status, "passed");

  const baselineWrite = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    sheetName: "Summary",
    range: "A1:D4",
    renderFormat: "png",
    baselineDir,
    writeBaseline: true,
    allSheets: true,
    nativeRender: nativeStatus.available ? "required" : "off",
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.baselinePath)).size > 100);
  const baselineCompare = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    sheetName: "Summary",
    range: "A1:D4",
    renderFormat: "png",
    baselineDir,
    allSheets: true,
    nativeRender: nativeStatus.available ? "required" : "off",
  });
  assert.equal(baselineCompare.summary.baselineCompared, true);
  assert.equal(baselineCompare.summary.pixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);
  assert.equal(baselineCompare.summary.allSheets, true);
  assert.deepEqual(baselineCompare.summary.sheetRenders.map((item) => item.sheetName).sort(), ["DateTimeText", "Inputs", "Summary"]);
  assert.ok(baselineCompare.summary.sheetRenders.every((item) => item.baselineCompared && item.pixelDiff.changed === false && item.ok));
  for (const sheetRender of baselineCompare.summary.sheetRenders) {
    assert.ok((await fs.stat(sheetRender.preview)).size > 100);
    assert.ok((await fs.stat(sheetRender.layout)).size > 20);
  }
  if (nativeStatus.available) {
    assert.equal(baselineCompare.summary.nativeRender.status, "passed");
    assert.equal(baselineCompare.summary.nativeRender.ok, true);
    assert.equal(baselineCompare.summary.nativeRender.pageCountMatches, true);
    assert.ok(baselineCompare.summary.nativeRender.pageCount >= 1);
    assert.ok(baselineCompare.summary.nativeRender.pages.every((page) => page.baselineCompared && page.pixelDiff.changed === false && page.ok));
  }
  const baselineMetadata = await sharp(await fs.readFile(baselineCompare.summary.baselinePath)).metadata();
  await sharp({ create: { width: baselineMetadata.width, height: baselineMetadata.height, channels: 3, background: { r: 0, g: 0, b: 0 } } }).png().toFile(baselineCompare.summary.baselinePath);
  const changedBaseline = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "baseline-changed"),
    sheetName: "Summary",
    range: "A1:D4",
    renderFormat: "png",
    baselineDir,
    nativeRender: "off",
    failOnIssues: false,
  });
  assert.equal(changedBaseline.summary.visualQaOk, false);
  assert.ok((await fs.stat(changedBaseline.summary.files.diff)).size > 100);
  assert.equal(changedBaseline.summary.sheetRenders[0].diffPath, changedBaseline.summary.files.diff);

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "spreadsheets", "SKILL.md"), "utf8");
  assert.match(skillText, /open-office-artifact-tool/);
  assert.match(skillText, /verify-workbook\.mjs/);
  assert.match(skillText, /baseline-dir/);
  assert.match(skillText, /LibreOffice/);
  assert.match(skillText, /\.csv/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("spreadsheet skill smoke ok");
