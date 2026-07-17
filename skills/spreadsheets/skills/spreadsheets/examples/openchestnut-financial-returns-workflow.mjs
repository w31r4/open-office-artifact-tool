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

function assertClose(actual, expected, tolerance = 1e-9) {
  assert.equal(typeof actual, "number");
  assert.ok(Math.abs(actual - expected) <= tolerance, `Expected ${actual} to be within ${tolerance} of ${expected}.`);
}

function styleTitle(sheet, range, text) {
  sheet.getRange(range).values = [[text, null, null, null]];
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: HEADER_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 30,
  };
}

function styleSection(sheet, range, text) {
  sheet.getRange(range).values = [[text, null, null, null]];
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: SECTION_FILL,
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 22,
  };
}

export function buildFinancialReturnsWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const inputs = workbook.worksheets.add("Inputs");
  const returns = workbook.worksheets.add("Returns");
  const checks = workbook.worksheets.add("Checks");

  for (const sheet of [inputs, returns, checks]) sheet.showGridLines = false;

  styleTitle(inputs, "A1:D1", "Investment Returns Model");
  styleSection(inputs, "A3:D3", "Assumptions");
  inputs.getRange("A4:D10").values = [
    ["Assumption", "Value", "Units", "Purpose"],
    ["Periodic discount rate", 0.1, "%", "Used for NPV of annual periods"],
    ["Annual discount rate", 0.1, "%", "Used for XNPV on actual dates"],
    ["IRR / XIRR guess", 0.1, "%", "Explicit root-search starting point"],
    ["Loan rate per period", 0.01, "%", "Used for the illustrative PMT output"],
    ["Loan term", 12, "periods", "Used for the illustrative PMT output"],
    ["Loan principal", 100000, "$", "Positive borrowing balance"],
  ];
  inputs.getRange("A4:D4").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  inputs.getRange("B5:B10").format = { fill: INPUT_FILL, font: { bold: true, color: "#0000FF" } };
  inputs.getRange("B5:B8").format.numberFormat = RATE_FORMAT;
  inputs.getRange("B9").format.numberFormat = COUNT_FORMAT;
  inputs.getRange("B10").format.numberFormat = MONEY_FORMAT;

  styleSection(inputs, "A12:D12", "Cash Flow Schedule");
  inputs.getRange("A13:D18").values = [
    ["Period", "Date", "Cash flow", "Description"],
    [0, new Date("2026-01-01T00:00:00.000Z"), -100000, "Initial investment"],
    [1, new Date("2027-01-01T00:00:00.000Z"), 30000, "Year 1 distribution"],
    [2, new Date("2028-01-01T00:00:00.000Z"), 35000, "Year 2 distribution"],
    [3, new Date("2029-01-01T00:00:00.000Z"), 40000, "Year 3 distribution"],
    [4, new Date("2030-01-01T00:00:00.000Z"), 45000, "Year 4 distribution"],
  ];
  inputs.getRange("A13:D13").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  inputs.getRange("B14:C18").format = { font: { color: "#0000FF" } };
  inputs.getRange("B14:B18").format.numberFormat = "yyyy-mm-dd";
  inputs.getRange("C14:C18").format.numberFormat = MONEY_FORMAT;
  inputs.getRange("A1:A18").format.columnWidthPx = 176;
  inputs.getRange("B1:B18").format.columnWidthPx = 126;
  inputs.getRange("C1:C18").format.columnWidthPx = 126;
  inputs.getRange("D1:D18").format.columnWidthPx = 230;
  inputs.getRange("F3:G6").values = [
    ["Legend", "Meaning"],
    ["Blue text", "Editable assumption or cash-flow input"],
    ["Green text", "Formula linked to another worksheet"],
    ["Status", "Formula-driven check result"],
  ];
  inputs.getRange("F3:G3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" } };
  inputs.getRange("F4").format = { font: { color: "#0000FF", bold: true } };
  inputs.getRange("F5").format = { font: { color: "#008000", bold: true } };
  inputs.getRange("F6").format = { font: { color: "#000000", bold: true } };
  inputs.getRange("F3:F6").format.columnWidthPx = 104;
  inputs.getRange("G3:G6").format.columnWidthPx = 238;
  inputs.freezePanes.freezeRows(4);

  styleTitle(returns, "A1:C1", "Investment Return Outputs");
  returns.getRange("A3:C3").values = [["Metric", "Value", "Convention / formula purpose"]];
  returns.getRange("A3:C3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  returns.getRange("A4:C9").values = [
    ["PV of future cash flows", null, "Discounts periods 1-4 only"],
    ["Net present value", null, "Initial outlay plus PV of future cash flows"],
    ["Periodic IRR", null, "Cash flows indexed by period"],
    ["Date-based NPV", null, "Actual days from the first date / 365"],
    ["Date-based IRR", null, "Annualized actual-date return"],
    ["Loan payment per period", null, "Illustrative payment shown as a cash outflow"],
  ];
  returns.getRange("B4:B9").formulas = [
    ["=NPV('Inputs'!$B$5,'Inputs'!$C$15:$C$18)"],
    ["='Inputs'!$C$14+B4"],
    ["=IRR('Inputs'!$C$14:$C$18,'Inputs'!$B$7)"],
    ["=XNPV('Inputs'!$B$6,'Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18)"],
    ["=XIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18,'Inputs'!$B$7)"],
    ["=PMT('Inputs'!$B$8,'Inputs'!$B$9,'Inputs'!$B$10)"],
  ];
  returns.getRange("B4:B9").format = { fill: "#F0FDF4", font: { bold: true, color: "#008000" } };
  returns.getRange("B4:B5").format.numberFormat = MONEY_FORMAT;
  returns.getRange("B6").format.numberFormat = RATE_FORMAT;
  returns.getRange("B7").format.numberFormat = MONEY_FORMAT;
  returns.getRange("B8").format.numberFormat = RATE_FORMAT;
  returns.getRange("B9").format.numberFormat = MONEY_FORMAT;
  returns.getRange("A1:A9").format.columnWidthPx = 210;
  returns.getRange("B1:B9").format.columnWidthPx = 144;
  returns.getRange("C1:C9").format.columnWidthPx = 276;
  returns.freezePanes.freezeRows(3);

  styleTitle(checks, "A1:F1", "Model Checks");
  checks.getRange("A3:F3").values = [["Check", "Actual", "Expected / minimum", "Difference", "Status", "Notes"]];
  checks.getRange("A3:F3").format = { fill: HEADER_FILL, font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  checks.getRange("A4:F9").values = [
    ["Positive cash flow count", null, 1, null, null, "Requires at least one positive cash flow."],
    ["Negative cash flow count", null, 1, null, null, "Requires at least one negative cash flow."],
    ["Periodic return is numeric", null, true, null, null, "Bounded IRR solver converged."],
    ["Date-based return is numeric", null, true, null, null, "Date alignment and XIRR converged."],
    ["NPV reconciliation", null, null, null, null, "Initial outlay plus future PV."],
    ["Overall model status", null, 0, null, null, "Zero failed checks required."],
  ];
  // Write formulas only to formula cells: assigning a matrix with null formulas
  // would intentionally clear the neighboring static expected-value cells.
  checks.getRange("B4:B9").formulas = [
    ["=COUNTIF('Inputs'!$C$14:$C$18,\">0\")"],
    ["=COUNTIF('Inputs'!$C$14:$C$18,\"<0\")"],
    ["=ISNUMBER('Returns'!$B$6)"],
    ["=ISNUMBER('Returns'!$B$8)"],
    ["='Returns'!$B$5"],
    ["=COUNTIF(E4:E8,\"CHECK\")"],
  ];
  checks.getRange("C8").formulas = [["='Inputs'!$C$14+'Returns'!$B$4"]];
  checks.getRange("D4:D5").formulas = [["=B4-C4"], ["=B5-C5"]];
  checks.getRange("D8:D9").formulas = [["=B8-C8"], ["=B9-C9"]];
  checks.getRange("E4:E9").formulas = [
    ["=IF(B4>=C4,\"OK\",\"CHECK\")"],
    ["=IF(B5>=C5,\"OK\",\"CHECK\")"],
    ["=IF(B6=C6,\"OK\",\"CHECK\")"],
    ["=IF(B7=C7,\"OK\",\"CHECK\")"],
    ["=IF(ABS(D8)<'Inputs'!$B$10,\"OK\",\"CHECK\")"],
    ["=IF(B9=C9,\"OK\",\"CHECK\")"],
  ];
  checks.getRange("B4:C9").format = { font: { color: "#008000" } };
  checks.getRange("D4:D9").format = { font: { color: "#000000" } };
  checks.getRange("B4:D5").format.numberFormat = COUNT_FORMAT;
  checks.getRange("B8:D8").format.numberFormat = MONEY_FORMAT;
  checks.getRange("B9:D9").format.numberFormat = COUNT_FORMAT;
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "OK",
    format: { fill: "#DCFCE7", font: { bold: true, color: "#166534" } },
  });
  checks.getRange("E4:E9").conditionalFormats.add("containsText", {
    text: "CHECK",
    format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } },
  });
  checks.getRange("A1:A9").format.columnWidthPx = 208;
  checks.getRange("B1:D9").format.columnWidthPx = 130;
  checks.getRange("E1:E9").format.columnWidthPx = 94;
  checks.getRange("F1:F9").format.columnWidthPx = 270;
  checks.freezePanes.freezeRows(3);

  workbook.worksheets.setActiveWorksheet("Returns");
  workbook.recalculate();
  return workbook;
}

export async function createFinancialReturnsWorkbook(outputPath) {
  const workbook = buildFinancialReturnsWorkbook();
  const returns = workbook.worksheets.getItem("Returns");
  const checks = workbook.worksheets.getItem("Checks");

  assertClose(returns.getRange("B4").values[0][0], 116986.5446349293);
  assertClose(returns.getRange("B5").values[0][0], 16986.544634929305);
  assertClose(returns.getRange("B6").values[0][0], 0.1709368633949911);
  assertClose(returns.getRange("B7").values[0][0], 16970.673463254516);
  assertClose(returns.getRange("B8").values[0][0], 0.17083686863616765);
  assertClose(returns.getRange("B9").values[0][0], -8884.878867834168);
  assert.deepEqual(checks.getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const inspection = workbook.inspect({
    kind: "workbook,sheet,formula,computedStyle",
    sheetName: "Returns",
    range: "A1:C9",
    maxChars: 12_000,
  });
  assert.match(inspection.ndjson, /XIRR/);
  assert.match(inspection.ndjson, /"formula":"=NPV/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const previewSvg = await workbook.render({ sheetName: "Returns", range: "A1:C9", autoCrop: "all", format: "svg" });
  assert.equal(previewSvg.type, "image/svg+xml");
  const previewText = await previewSvg.text();
  assert.match(previewText, /Investment Return Outputs/);
  assert.match(previewText, /\(\$8,885\)/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Returns").getRange("B8").formulas[0][0], "=XIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18,'Inputs'!$B$7)");
  assert.equal(imported.worksheets.getItem("Returns").getRange("B7").format.numberFormat, MONEY_FORMAT);
  assert.deepEqual(imported.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  roundTrip.recalculate();
  assertClose(roundTrip.worksheets.getItem("Returns").getRange("B8").values[0][0], 0.17083686863616765);
  assert.deepEqual(roundTrip.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, previewSvg };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-financial-returns-workflow.xlsx");
  const result = await createFinancialReturnsWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
