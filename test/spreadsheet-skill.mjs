import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { nativeSpreadsheetRenderStatus, runSpreadsheetFixture, verifyWorkbookFile } from "../skills/spreadsheets/scripts/workflow.mjs";

const repoRoot = path.resolve(new URL("..", import.meta.url).pathname);
const fixturePath = path.join(repoRoot, "skills", "spreadsheets", "fixtures", "formula-summary.json");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-spreadsheet-skill-"));
const baselineDir = path.join(outputDir, "baselines");

try {
  const result = await runSpreadsheetFixture(fixturePath, { outputDir, nativeRender: "off" });
  assert.equal(result.fixture.name, "formula-summary");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.renderFormat, "svg");
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
  assert.ok(Math.abs(workbook.worksheets.getItem("Summary").getRange("B1:C4").format.columnWidthPx - 96) <= 1);
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
  const fixturePivotXml = await fixtureZip.file("xl/pivotTables/pivotTable1.xml").async("text");
  assert.match(fixturePivotXml, /x14:pivotFilter useWholeDay="0"/);
  assert.match(fixturePivotXml, /stringValue1="2026-03-31T17:00:00" stringValue2="2026-03-31T19:00:00"/);
  const summaryDrawings = workbook.worksheets.getItem("Summary");
  assert.equal(summaryDrawings.charts.items.length, 1);
  assert.equal(summaryDrawings.charts.items[0].title, "Quarter performance");
  assert.deepEqual(summaryDrawings.charts.items[0].categories, ["Jan", "Feb", "Mar"]);
  assert.deepEqual(summaryDrawings.charts.items[0].series.items[0].values, [100, 120, 150]);
  assert.equal(summaryDrawings.images.items.length, 1);
  assert.equal(summaryDrawings.images.items[0].alt, "Green status marker");
  assert.match(summaryDrawings.images.items[0].dataUrl, /^data:image\/png;base64,/);
  assert.equal(workbook.resolve(summaryDrawings.charts.items[0].id), summaryDrawings.charts.items[0]);
  assert.equal(workbook.resolve(summaryDrawings.images.items[0].id), summaryDrawings.images.items[0]);
  assert.equal(summaryDrawings.pivotTables.items.length, 1);
  const summaryPivot = summaryDrawings.pivotTables.getItemOrNullObject("RevenuePivot");
  assert.deepEqual(summaryPivot.computedValues(), [["Period Year", "Period Quarter", "Period Month", "Month", "Period End", "Revenue total", "Gross profit"], ["2026", "Q1", "Mar", "Mar", "2026-03-31T18:00:00Z", 150, 60]]);
  assert.deepEqual(summaryPivot.groupFields, [
    { name: "Period Year", sourceField: "Period End", groupBy: "years" },
    { name: "Period Quarter", sourceField: "Period End", groupBy: "quarters", parent: "Period Year" },
    { name: "Period Month", sourceField: "Period End", groupBy: "months", parent: "Period Quarter" },
  ]);
  assert.deepEqual(summaryPivot.filters, [{ field: "Month", include: ["Jan", "Mar"] }, { field: "Period End", type: "dateBetween", value1: "2026-03-31T17:00:00", value2: "2026-03-31T19:00:00", useWholeDay: false }]);
  assert.deepEqual(summaryPivot.calculatedFields, [{ name: "Gross Profit", formula: "='Revenue'-'Cost'", numFmtId: 0, references: ["Revenue", "Cost"] }]);
  assert.equal(summaryPivot.refreshPolicy.refreshOnLoad, false);
  assert.equal(summaryPivot.refreshPolicy.refreshedBy, "Spreadsheet skill");
  assert.equal(workbook.resolve("RevenuePivot"), summaryPivot);
  assert.match(workbook.help("workbook.structuredReferences").ndjson, /special-character headers/);
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("G10:G13").values, [[0], [43889], [5], [2]]);
  assert.equal(workbook.worksheets.getItem("Summary").getRange("G14").values[0][0], 1);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /SummaryTable/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /Inputs!B2/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"freezePanes":\{"rows":1,"columns":1/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"customColumns":7/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"kind":"mergedCell"[\s\S]*"range":"A15:G15"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"drawingType":"chart"[\s\S]*"title":"Quarter performance"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"drawingType":"image"[\s\S]*"alt":"Green status marker"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"kind":"pivotTable"[\s\S]*"name":"RevenuePivot"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"filters":\[\{"field":"Month","include":\["Jan","Mar"\]/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"field":"Period End","type":"dateBetween"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"useWholeDay":false/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"groupBy":"quarters"/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"calculatedFields":\[\{"name":"Gross Profit"/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /xl\/workbook\.xml/);
  assert.equal(result.qa.packageInspect.records[0].sheets, 2);
  assert.match(await fs.readFile(result.qa.summary.files.preview, "utf8"), /<svg/);

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
    assert.ok(libreOfficeSummary.charts.items.length >= 1);
    assert.ok(libreOfficeSummary.images.items.length >= 1);
    assert.ok(libreOfficeSummary.pivotTables.items.some((pivot) => pivot.name === "RevenuePivot"));
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
  assert.deepEqual(baselineCompare.summary.sheetRenders.map((item) => item.sheetName).sort(), ["Inputs", "Summary"]);
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
