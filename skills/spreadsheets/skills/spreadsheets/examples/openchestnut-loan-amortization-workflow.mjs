import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

const MONEY_FORMAT = "$#,##0;[Red]($#,##0);-";
const RATE_FORMAT = "0.0%;[Red](0.0%);-";
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

export function buildLoanAmortizationWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const inputs = workbook.worksheets.add("Inputs");
  const amortization = workbook.worksheets.add("Amortization");
  const checks = workbook.worksheets.add("Checks");

  for (const sheet of [inputs, amortization, checks]) sheet.showGridLines = false;

  styleTitle(inputs, "A1:D1", 4, "Loan Amortization Model");
  styleSection(inputs, "A3:D3", 4, "Visible Assumptions");
  inputs.getRange("A4:D10").values = [
    ["Assumption", "Value", "Units", "Purpose"],
    ["Original principal", 100000, "$", "Illustrative opening borrowing balance"],
    ["Annual interest rate", 0.12, "%", "Nominal annual rate"],
    ["Payments per year", 12, "payments/year", "Monthly payment cadence"],
    ["Term", 1, "years", "Illustrative one-year term"],
    ["Payment timing", 0, "0 = end / 1 = begin", "End-of-period payment for this example"],
    ["Periodic interest rate", null, "%", "Annual rate divided by payment cadence"],
  ];
  inputs.getRange("A4:D4").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  inputs.getRange("A11:D11").values = [["Total modeled periods", null, "periods", "Term multiplied by payment cadence"]];
  inputs.getRange("B10").formulas = [["=B6/B7"]];
  inputs.getRange("B11").formulas = [["=B7*B8"]];
  inputs.getRange("B5:B9").format = { fill: INPUT_FILL, font: { bold: true, color: "#0000FF" } };
  inputs.getRange("B10:B11").format = { font: { color: "#000000" } };
  inputs.getRange("B5").format.numberFormat = MONEY_FORMAT;
  inputs.getRange("B6:B6").format.numberFormat = RATE_FORMAT;
  inputs.getRange("B7:B8").format.numberFormat = COUNT_FORMAT;
  inputs.getRange("B9").format.numberFormat = COUNT_FORMAT;
  inputs.getRange("B10").format.numberFormat = RATE_FORMAT;
  inputs.getRange("B11").format.numberFormat = COUNT_FORMAT;
  inputs.getRange("A1:A11").format.columnWidthPx = 198;
  inputs.getRange("B1:B11").format.columnWidthPx = 148;
  inputs.getRange("C1:C11").format.columnWidthPx = 146;
  inputs.getRange("D1:D11").format.columnWidthPx = 270;
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

  styleTitle(amortization, "A1:F1", 6, "Loan Amortization Schedule");
  styleSection(amortization, "A3:F3", 6, "Twelve-Period Example");
  amortization.getRange("A4:F4").values = [["Period", "Opening balance", "Payment", "Interest", "Principal", "Closing balance"]];
  amortization.getRange("A4:F4").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  amortization.getRange("A5:A16").values = Array.from({ length: 12 }, (_, index) => [index + 1]);
  amortization.getRange("B5").formulas = [["='Inputs'!$B$5"]];
  amortization.getRange("B6").formulas = [["=F5"]];
  amortization.getRange("B6:B16").fillDown();
  amortization.getRange("C5").formulas = [["=PMT('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)"]];
  amortization.getRange("C5:C16").fillDown();
  amortization.getRange("D5").formulas = [["=IPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)"]];
  amortization.getRange("D5:D16").fillDown();
  amortization.getRange("E5").formulas = [["=PPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)"]];
  amortization.getRange("E5:E16").fillDown();
  amortization.getRange("F5").formulas = [["=B5+E5"]];
  amortization.getRange("F5:F16").fillDown();
  amortization.getRange("B5:E16").format = { font: { color: "#008000" } };
  amortization.getRange("B6:B16").format = { font: { color: "#000000" } };
  amortization.getRange("F5:F16").format = { font: { color: "#000000" } };
  amortization.getRange("A5:A16").format.numberFormat = COUNT_FORMAT;
  amortization.getRange("B5:F16").format.numberFormat = MONEY_FORMAT;
  amortization.getRange("A1:A16").format.columnWidthPx = 78;
  amortization.getRange("B1:F16").format.columnWidthPx = 138;
  amortization.freezePanes.freezeRows(4);

  styleTitle(checks, "A1:F1", 6, "Model Checks");
  checks.getRange("A3:F3").values = [["Check", "Actual", "Expected / minimum", "Difference", "Status", "Notes"]];
  checks.getRange("A3:F3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  checks.getRange("A4:F9").values = [
    ["Payment identity", null, 0, null, null, "Total payment equals total interest plus total principal."],
    ["First-period interest", null, null, null, null, "Uses the selected payment-timing convention."],
    ["Final closing balance", null, 0, null, null, "A fully amortized balance should be zero."],
    ["Total principal repaid", null, null, null, null, "Total principal should equal the negative original principal."],
    ["Modeled period count", null, null, null, null, "Schedule rows must equal the visible total-period input."],
    ["Overall model status", null, 0, null, null, "Zero failed checks required."],
  ];
  checks.getRange("B4:B9").formulas = [
    ["=SUM('Amortization'!$C$5:$C$16)-SUM('Amortization'!$D$5:$D$16)-SUM('Amortization'!$E$5:$E$16)"],
    ["='Amortization'!$D$5"],
    ["='Amortization'!$F$16"],
    ["=SUM('Amortization'!$E$5:$E$16)"],
    ["=COUNTIF('Amortization'!$A$5:$A$16,\">0\")"],
    ["=COUNTIF(E4:E8,\"CHECK\")"],
  ];
  checks.getRange("C5").formulas = [["=-'Inputs'!$B$5*'Inputs'!$B$10*(1-'Inputs'!$B$9)"]];
  checks.getRange("C7").formulas = [["=-'Inputs'!$B$5"]];
  checks.getRange("C8").formulas = [["='Inputs'!$B$11"]];
  checks.getRange("D4:D9").formulas = [
    ["=ABS(B4-C4)"],
    ["=ABS(B5-C5)"],
    ["=ABS(B6-C6)"],
    ["=ABS(B7-C7)"],
    ["=ABS(B8-C8)"],
    ["=ABS(B9-C9)"],
  ];
  checks.getRange("E4:E9").formulas = [
    ["=IF(D4<0.01,\"OK\",\"CHECK\")"],
    ["=IF(D5<0.01,\"OK\",\"CHECK\")"],
    ["=IF(D6<0.01,\"OK\",\"CHECK\")"],
    ["=IF(D7<0.01,\"OK\",\"CHECK\")"],
    ["=IF(B8=C8,\"OK\",\"CHECK\")"],
    ["=IF(B9=C9,\"OK\",\"CHECK\")"],
  ];
  checks.getRange("B4:C9").format = { font: { color: "#008000" } };
  checks.getRange("D4:D9").format = { font: { color: "#000000" } };
  checks.getRange("B4:D7").format.numberFormat = MONEY_FORMAT;
  checks.getRange("B8:D9").format.numberFormat = COUNT_FORMAT;
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "OK",
    format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
  });
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "CHECK",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  checks.getRange("A1:A9").format.columnWidthPx = 198;
  checks.getRange("B1:D9").format.columnWidthPx = 136;
  checks.getRange("E1:E9").format.columnWidthPx = 94;
  checks.getRange("F1:F9").format.columnWidthPx = 292;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Amortization");
  workbook.recalculate();
  return workbook;
}

export async function createLoanAmortizationWorkbook(outputPath) {
  const workbook = buildLoanAmortizationWorkbook();
  const amortization = workbook.worksheets.getItem("Amortization");
  const checks = workbook.worksheets.getItem("Checks");

  assertClose(amortization.getRange("C5").values[0][0], -8884.878867834168);
  assertClose(amortization.getRange("D5").values[0][0], -1000);
  assertClose(amortization.getRange("E5").values[0][0], -7884.878867834168);
  assertClose(amortization.getRange("D6").values[0][0], -921.1512113216584);
  assertClose(amortization.getRange("F16").values[0][0], 0, 1e-7);
  assert.deepEqual(checks.getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const inspection = workbook.inspect({
    kind: "workbook,sheet,formula",
    sheetName: "Amortization",
    range: "A1:F16",
    maxChars: 16_000,
  });
  assert.match(inspection.ndjson, /IPMT/);
  assert.match(inspection.ndjson, /PPMT/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Amortization", range: "A1:F16", autoCrop: "all", format: "svg" });
  assert.equal(previewSvg.type, "image/svg+xml");
  const previewText = await previewSvg.text();
  assert.match(previewText, /Loan Amortization Schedule/);
  assert.match(previewText, /\(\$8,885\)/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Amortization").getRange("D5").formulas[0][0], "=IPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.equal(imported.worksheets.getItem("Amortization").getRange("E5").format.numberFormat, MONEY_FORMAT);
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assertClose(roundTrip.worksheets.getItem("Amortization").getRange("F16").values[0][0], 0, 1e-7);
  assert.deepEqual(roundTrip.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-loan-amortization-workflow.xlsx");
  const result = await createLoanAmortizationWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
