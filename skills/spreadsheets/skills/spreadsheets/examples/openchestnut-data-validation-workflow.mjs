import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildDataValidationWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Intake");
  sheet.getRange("A1:D5").values = [
    ["Request", "Status", "Priority", "Budget"],
    ["Workspace refresh", "Planned", 2, 8000],
    ["Billing controls", "In progress", 1, 12000],
    ["Archive cleanup", "Done", 4, 2500],
    ["New onboarding", "Planned", 3, 6000],
  ];
  sheet.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A1:D5").format.autofitColumns();
  sheet.getRange("D2:D5").format.numberFormat = "$#,##0";
  sheet.freezePanes.freezeRows(1);
  sheet.showGridLines = false;

  sheet.getRange("B2:B20").dataValidation = {
    rule: {
      type: "list",
      values: ["Planned", "In progress", "Done"],
      allowBlank: false,
      showInputMessage: true,
      promptTitle: "Choose a status",
      prompt: "Use one approved workflow state.",
      showErrorMessage: true,
      errorTitle: "Invalid status",
      error: "Choose a value from the list.",
      errorStyle: "stop",
      showDropdown: true,
    },
  };
  sheet.getRange("C2:C20").dataValidation = {
    rule: {
      type: "whole",
      operator: "between",
      formula1: 1,
      formula2: 5,
      allowBlank: false,
      showInputMessage: true,
      promptTitle: "Priority",
      prompt: "Enter a whole number from 1 to 5.",
      showErrorMessage: true,
      errorTitle: "Priority out of range",
      error: "Priority must be between 1 and 5.",
      errorStyle: "warning",
    },
  };
  sheet.getRange("D2:D20").dataValidation = {
    rule: {
      type: "custom",
      formula1: "=D2>=0",
      allowBlank: true,
      showErrorMessage: true,
      errorTitle: "Invalid budget",
      error: "Budget cannot be negative.",
      errorStyle: "stop",
    },
  };
  return workbook;
}

export async function createDataValidationWorkbook(outputPath) {
  const workbook = buildDataValidationWorkbook();
  const sheet = workbook.worksheets.getItem("Intake");
  assert.deepEqual(sheet.dataValidations.items.map((item) => item.rule.type), ["list", "whole", "custom"]);
  const inspection = workbook.inspect({ kind: "sheet,dataValidation", sheetName: sheet.name, range: "A1:D20", maxChars: 12_000 });
  assert.equal((inspection.ndjson.match(/"kind":"dataValidation"/g) || []).length, 3);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: sheet.name, range: "A1:D8", format: "svg" });
  assert.match(await preview.text(), /Workspace refresh/);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem(sheet.name);
  const status = importedSheet.dataValidations.items.find((item) => item.rule.type === "list");
  assert.deepEqual(status.rule.values, ["Planned", "In progress", "Done"]);
  assert.equal(status.rule.showDropdown, true);
  assert.equal(status.rule.showInputMessage, true);
  assert.equal(status.rule.showErrorMessage, true);

  status.rule.prompt = "Pick the current workflow state.";
  status.rule.errorStyle = "information";
  status.rule.showDropdown = false;
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  const finalSheet = roundTrip.worksheets.getItem(sheet.name);
  const finalStatus = finalSheet.dataValidations.items.find((item) => item.rule.type === "list");
  assert.equal(finalStatus.rule.prompt, "Pick the current workflow state.");
  assert.equal(finalStatus.rule.errorStyle, "information");
  assert.equal(finalStatus.rule.showDropdown, false);
  assert.equal(finalSheet.dataValidations.items.find((item) => item.rule.type === "custom").rule.formula1, "=D2>=0");
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
      schema: "open-office-artifact-tool.spreadsheet-data-validation-workflow.v1",
      provider: { requested: "open-chestnut", actual: "open-chestnut", fallbackUsed: false },
      validation: {
        authoredRuleTypes: ["list", "whole", "custom"],
        secondImport: true,
        editedPrompt: finalStatus.rule.prompt,
        editedErrorStyle: finalStatus.rule.errorStyle,
        dropdownVisible: finalStatus.rule.showDropdown,
      },
    },
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-data-validation-workflow.xlsx");
  const result = await createDataValidationWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok, audit: result.audit }));
}
