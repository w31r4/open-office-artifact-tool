import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildDataTableWorkbook() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const sheet = workbook.worksheets.add("Scenario Analysis");

  sheet.getRange("A1:B4").values = [
    [null, "Result"],
    [80, null],
    [100, null],
    [120, null],
  ];
  sheet.getRange("A1").formulas = [["=D1*D2"]];
  sheet.getRange("D1:D2").values = [[100], [0.2]];
  sheet.getRange("C1:D2").format = { fill: "#DBEAFE" };
  sheet.getRange("C1:C2").values = [["Base"], ["Rate"]];
  sheet.dataTables.add("A1:B4", { rowInput: "D1" });

  sheet.getRange("F1:H4").values = [
    [null, 0.1, 0.2],
    [80, null, null],
    [100, null, null],
    [120, null, null],
  ];
  sheet.getRange("F1").formulas = [["=D1*D2"]];
  sheet.dataTables.add("F1:H4", { rowInput: "D1", columnInput: "D2" });

  sheet.getRange("A1:H1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A1:H4").format.autofitColumns();
  sheet.getRange("B2:B4").format.numberFormat = "0.00";
  sheet.getRange("G2:H4").format.numberFormat = "0.00";
  sheet.freezePanes.freezeRows(1);
  return workbook;
}

export async function createDataTableWorkbook(outputPath) {
  const workbook = buildDataTableWorkbook();
  const sheet = workbook.worksheets.getItem("Scenario Analysis");
  const definitions = sheet.dataTables.__getDefinitions();
  assert.deepEqual(definitions.map((item) => item.displayFormula), ["{=TABLE(D1)}", "{=TABLE(D1,D2)}"]);

  const inspection = workbook.inspect({ kind: "sheet,dataTable", sheetName: sheet.name, maxChars: 8_000 });
  assert.match(inspection.ndjson, /"kind":"dataTable"/);
  const verification = workbook.verify();
  assert.equal(verification.ok, true, verification.ndjson);

  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedSheet = imported.worksheets.getItem(sheet.name);
  assert.deepEqual(importedSheet.dataTables.__getDefinitions(), definitions);
  importedSheet.getRange("J1").values = [["Host-calculated What-If outputs"]];
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  assert.deepEqual(roundTrip.worksheets.getItem(sheet.name).dataTables.__getDefinitions(), definitions);

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-data-table-workflow.xlsx");
  const result = await createDataTableWorkbook(outputPath);
  console.log(JSON.stringify({ outputPath, bytes: result.file.bytes.length, verified: result.verification.ok }));
}
