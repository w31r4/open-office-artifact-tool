import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildProtectedInputWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Inputs");
  sheet.getRange("A1:D5").values = [
    ["Scenario", "Input", "Unit", "Calculated value"],
    ["Volume", 1200, "units", null],
    ["Price", 18.5, "USD", null],
    ["Variable cost", 7.25, "USD", null],
    ["Contribution", null, "USD", null],
  ];
  sheet.getRange("D2:D4").formulas = [["=B2"], ["=B3"], ["=B4"]];
  sheet.getRange("B5").formulas = [["=B2*(B3-B4)"]];
  sheet.getRange("D5").formulas = [["=B5"]];
  sheet.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("B2:B4").format = {
    fill: "#FFF2CC",
    font: { color: "#0000FF" },
    protection: { locked: false, hidden: false },
  };
  sheet.getRange("B5:D5").format = {
    fill: "#DCFCE7",
    font: { bold: true },
    protection: { locked: true, hidden: true },
  };
  sheet.getRange("B2:D5").format.numberFormat = "#,##0.00";
  sheet.getRange("A1:D5").format.autofitColumns();
  sheet.tables.add("A1:D5", true, "ProtectedInputs");
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;
  sheet.protection = {
    allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
  };
  return workbook;
}

export async function createProtectedInputWorkbook(outputPath) {
  const workbook = buildProtectedInputWorkbook();
  const sourceSheet = workbook.worksheets.getItem("Inputs");
  assert.deepEqual(sourceSheet.getRange("B2").format.protection, { locked: false, hidden: false });
  assert.deepEqual(sourceSheet.getRange("B5").format.protection, { locked: true, hidden: true });
  const inspection = workbook.inspect({ kind: "sheet,worksheetProtection,formula,style", sheetName: "Inputs", range: "A1:D5", maxChars: 12_000 });
  assert.match(inspection.ndjson, /"kind":"worksheetProtection"/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: "Inputs", range: "A1:D5", format: "svg" });
  assert.match(await preview.text(), /Contribution/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem("Inputs");
  assert.deepEqual(importedSheet.protection, {
    enabled: true,
    allow: ["selectLockedCells", "selectUnlockedCells", "sort", "autoFilter"],
  });
  assert.deepEqual(importedSheet.getRange("B2").format.protection, { locked: false, hidden: false });
  assert.deepEqual(importedSheet.getRange("B5").format.protection, { locked: true, hidden: true });

  // Tighten the selection policy through the source-bound public primitive;
  // OpenChestnut retains the native sheet and style graph around this edit.
  importedSheet.protection = { allow: ["selectUnlockedCells", "sort", "autoFilter"] };
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  const finalSheet = roundTrip.worksheets.getItem("Inputs");
  assert.deepEqual(finalSheet.protection, {
    enabled: true,
    allow: ["selectUnlockedCells", "sort", "autoFilter"],
  });
  assert.deepEqual(finalSheet.getRange("B2").format.protection, { locked: false, hidden: false });
  assert.deepEqual(finalSheet.getRange("B5").format.protection, { locked: true, hidden: true });
  const finalVerification = roundTrip.verify({ visualQa: true });
  assert.equal(finalVerification.ok, true, finalVerification.ndjson);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return {
    workbook: roundTrip,
    file: final,
    inspection,
    verification: finalVerification,
    preview,
    audit: {
      schema: "open-office-artifact-tool.spreadsheet-worksheet-protection-workflow.v1",
      provider: { requested: "open-chestnut", actual: "open-chestnut", fallbackUsed: false },
      securityBoundary: "editing-restriction-not-encryption",
      validation: {
        secondImport: true,
        unlockedInputRange: "B2:B4",
        hiddenFormulaRange: "B5:D5",
        allowedOperations: [...finalSheet.protection.allow],
      },
    },
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-worksheet-protection-workflow.xlsx");
  const result = await createProtectedInputWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok, audit: result.audit }));
}
