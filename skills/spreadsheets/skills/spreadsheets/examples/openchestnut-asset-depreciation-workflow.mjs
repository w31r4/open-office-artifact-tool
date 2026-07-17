import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const MONEY_FORMAT = "$#,##0;[Red]($#,##0);-";
const COUNT_FORMAT = "#,##0;[Red](#,##0);-";
const INPUT_FILL = "#FFF2CC";
const HEADER_FILL = "#0F172A";
const SECTION_FILL = "#1E3A5F";

function assertClose(actual, expected, tolerance = 1e-8) {
  assert.equal(typeof actual, "number");
  assert.ok(Math.abs(actual - expected) <= tolerance, `Expected ${actual} to be within ${tolerance} of ${expected}.`);
}

function styleMergedBand(sheet, range, columns, text, format) {
  const target = sheet.getRange(range);
  target.values = [[text, ...Array.from({ length: columns - 1 }, () => null)]];
  target.merge();
  target.format = format;
}

function styleTitle(sheet, range, columns, text) {
  styleMergedBand(sheet, range, columns, text, {
    fill: HEADER_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 30,
  });
}

function styleSection(sheet, range, columns, text) {
  styleMergedBand(sheet, range, columns, text, {
    fill: SECTION_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 22,
  });
}

export function buildAssetDepreciationWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const inputs = workbook.worksheets.add("Inputs");
  const depreciation = workbook.worksheets.add("Depreciation");
  const checks = workbook.worksheets.add("Checks");

  for (const sheet of [inputs, depreciation, checks]) sheet.showGridLines = false;

  styleTitle(inputs, "A1:D1", 4, "Fixed Asset Depreciation Model");
  styleSection(inputs, "A3:D3", 4, "Visible Assumptions");
  inputs.getRange("A4:D11").values = [
    ["Assumption", "Value", "Units", "Purpose"],
    ["Asset cost", 100000, "$", "Illustrative acquisition cost"],
    ["Estimated salvage value", 10000, "$", "Residual book value floor"],
    ["Useful life", 5, "years", "Annual modeled depreciation periods"],
    ["First-year months", 12, "months", "DB first-period proration"],
    ["DDB factor", 2, "multiple", "Double-declining balance factor"],
    ["Straight-line annual expense", null, "$ / year", "SLN result from visible assumptions"],
    ["Depreciable basis", null, "$", "Cost less estimated salvage value"],
  ];
  inputs.getRange("A4:D4").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  inputs.getRange("B10").formulas = [["=SLN(B5,B6,B7)"]];
  inputs.getRange("B11").formulas = [["=B5-B6"]];
  inputs.getRange("B5:B9").format = { fill: INPUT_FILL, font: { bold: true, color: "#0000FF" } };
  inputs.getRange("B10:B11").format = { font: { color: "#000000" } };
  inputs.getRange("B5:B6").format.numberFormat = MONEY_FORMAT;
  inputs.getRange("B7:B9").format.numberFormat = COUNT_FORMAT;
  inputs.getRange("B10:B11").format.numberFormat = MONEY_FORMAT;
  inputs.getRange("A1:A11").format.columnWidthPx = 208;
  inputs.getRange("B1:B11").format.columnWidthPx = 148;
  inputs.getRange("C1:C11").format.columnWidthPx = 126;
  inputs.getRange("D1:D11").format.columnWidthPx = 284;
  inputs.getRange("F3:G6").values = [
    ["Legend", "Meaning"],
    ["Blue text", "Editable illustrative assumption"],
    ["Green text", "Formula linked to another worksheet"],
    ["Black text", "Formula calculated within a worksheet"],
  ];
  inputs.getRange("F3:G3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" } };
  inputs.getRange("F4").format = { font: { bold: true, color: "#0000FF" } };
  inputs.getRange("F5").format = { font: { bold: true, color: "#008000" } };
  inputs.getRange("F6").format = { font: { bold: true, color: "#000000" } };
  inputs.getRange("F3:F6").format.columnWidthPx = 114;
  inputs.getRange("G3:G6").format.columnWidthPx = 254;
  inputs.freezePanes.freezeRows(4);

  styleTitle(depreciation, "A1:G1", 7, "Annual Depreciation Schedule");
  styleSection(depreciation, "A3:G3", 7, "Five-Year Asset Profile");
  depreciation.getRange("A4:G4").values = [["Year", "Opening DDB book value", "SLN expense", "DB expense", "DDB expense", "Closing DDB book value", "Closing DB book value"]];
  depreciation.getRange("A4:G4").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  depreciation.getRange("A5:A9").values = [[1], [2], [3], [4], [5]];
  depreciation.getRange("B5").formulas = [["='Inputs'!$B$5"]];
  depreciation.getRange("B6").formulas = [["=F5"]];
  depreciation.getRange("B6:B9").fillDown();
  depreciation.getRange("C5").formulas = [["=SLN('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7)"]];
  depreciation.getRange("C5:C9").fillDown();
  depreciation.getRange("D5").formulas = [["=DB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$8)"]];
  depreciation.getRange("D5:D9").fillDown();
  depreciation.getRange("E5").formulas = [["=DDB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$9)"]];
  depreciation.getRange("E5:E9").fillDown();
  depreciation.getRange("F5").formulas = [["=B5-E5"]];
  depreciation.getRange("F5:F9").fillDown();
  depreciation.getRange("G5").formulas = [["='Inputs'!$B$5-D5"]];
  depreciation.getRange("G6").formulas = [["=G5-D6"]];
  depreciation.getRange("G6:G9").fillDown();
  depreciation.getRange("B5:E9").format = { font: { color: "#008000" } };
  depreciation.getRange("B6:B9").format = { font: { color: "#000000" } };
  depreciation.getRange("F5:G9").format = { font: { color: "#000000" } };
  depreciation.getRange("A5:A9").format.numberFormat = COUNT_FORMAT;
  depreciation.getRange("B5:G9").format.numberFormat = MONEY_FORMAT;
  depreciation.getRange("A1:A9").format.columnWidthPx = 64;
  depreciation.getRange("B1:G9").format.columnWidthPx = 152;
  depreciation.freezePanes.freezeRows(4);

  styleTitle(checks, "A1:F1", 6, "Model Checks");
  checks.getRange("A3:F3").values = [["Check", "Actual", "Expected / minimum", "Difference", "Status", "Notes"]];
  checks.getRange("A3:F3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  checks.getRange("A4:F9").values = [
    ["DDB final book value", null, null, null, null, "DDB caps the final expense at the visible salvage value."],
    ["DDB total depreciation", null, null, null, null, "DDB recovers the visible depreciable basis."],
    ["SLN total depreciation", null, null, null, null, "Straight-line expense times life equals the same basis."],
    ["DB final book value floor", null, null, null, null, "Fixed-balance DB does not fall below salvage in this model."],
    ["Modeled depreciation periods", null, null, null, null, "Schedule rows match the visible useful-life input."],
    ["Overall model status", null, 0, null, null, "Zero failed checks required."],
  ];
  checks.getRange("B4:B9").formulas = [
    ["='Depreciation'!$F$9"],
    ["=SUM('Depreciation'!$E$5:$E$9)"],
    ["=SUM('Depreciation'!$C$5:$C$9)"],
    ["='Depreciation'!$G$9"],
    ["=COUNTIF('Depreciation'!$A$5:$A$9,\">0\")"],
    ["=COUNTIF(E4:E8,\"CHECK\")"],
  ];
  checks.getRange("C4").formulas = [["='Inputs'!$B$6"]];
  checks.getRange("C5:C6").formulas = [["='Inputs'!$B$11"], ["='Inputs'!$B$11"]];
  checks.getRange("C7").formulas = [["='Inputs'!$B$6"]];
  checks.getRange("C8").formulas = [["='Inputs'!$B$7"]];
  checks.getRange("D4:D9").formulas = [
    ["=B4-C4"],
    ["=B5-C5"],
    ["=B6-C6"],
    ["=B7-C7"],
    ["=B8-C8"],
    ["=B9-C9"],
  ];
  checks.getRange("E4:E9").formulas = [
    ["=IF(ABS(D4)<0.01,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D5)<0.01,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D6)<0.01,\"OK\",\"CHECK\")"],
    ["=IF(B7>=C7,\"OK\",\"CHECK\")"],
    ["=IF(B8=C8,\"OK\",\"CHECK\")"],
    ["=IF(B9=C9,\"OK\",\"CHECK\")"],
  ];
  checks.getRange("B4:C9").format = { font: { color: "#008000" } };
  checks.getRange("D4:D9").format = { font: { color: "#000000" } };
  checks.getRange("B4:D8").format.numberFormat = MONEY_FORMAT;
  checks.getRange("B9:D9").format.numberFormat = COUNT_FORMAT;
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "OK",
    format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
  });
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "CHECK",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  checks.getRange("A1:A9").format.columnWidthPx = 210;
  checks.getRange("B1:D9").format.columnWidthPx = 146;
  checks.getRange("E1:E9").format.columnWidthPx = 94;
  checks.getRange("F1:F9").format.columnWidthPx = 318;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Depreciation");
  workbook.recalculate();
  return workbook;
}

export async function createAssetDepreciationWorkbook(outputPath) {
  const workbook = buildAssetDepreciationWorkbook();
  const depreciation = workbook.worksheets.getItem("Depreciation");
  const checks = workbook.worksheets.getItem("Checks");

  assert.equal(depreciation.getRange("C5").values[0][0], 18000);
  assert.equal(depreciation.getRange("D5").values[0][0], 36900);
  assert.equal(depreciation.getRange("E5").values[0][0], 40000);
  assertClose(depreciation.getRange("E9").values[0][0], 2960);
  assert.equal(depreciation.getRange("F9").values[0][0], 10000);
  assert.ok(depreciation.getRange("G9").values[0][0] >= 10000);
  assert.deepEqual(checks.getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const inspection = workbook.inspect({
    kind: "workbook,sheet,formula",
    sheetName: "Depreciation",
    range: "A1:G9",
    maxChars: 16_000,
  });
  assert.match(inspection.ndjson, /SLN/);
  assert.match(inspection.ndjson, /DB/);
  assert.match(inspection.ndjson, /DDB/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Depreciation", range: "A1:G9", autoCrop: "all", format: "svg" });
  assert.equal(previewSvg.type, "image/svg+xml");
  const previewText = await previewSvg.text();
  assert.match(previewText, /Annual Depreciation Schedule/);
  assert.match(previewText, /\$40,000/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Depreciation").getRange("D5").formulas[0][0], "=DB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$8)");
  assert.equal(imported.worksheets.getItem("Depreciation").getRange("E5").formulas[0][0], "=DDB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$9)");
  assert.equal(imported.worksheets.getItem("Depreciation").getRange("F9").values[0][0], 10000);
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assert.equal(roundTrip.worksheets.getItem("Depreciation").getRange("F9").values[0][0], 10000);
  assert.deepEqual(roundTrip.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-asset-depreciation-workflow.xlsx");
  const result = await createAssetDepreciationWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
