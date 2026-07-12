import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";

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
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("G10:G13").values, [[0], [43889], [5], [2]]);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /SummaryTable/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /Inputs!B2/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"freezePanes":\{"rows":1,"columns":1/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"customColumns":6/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /"kind":"mergedCell"[\s\S]*"range":"A15:G15"/);
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
