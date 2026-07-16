import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { SpreadsheetFile, Workbook } from "open-office-artifact-tool";

export function buildWorkbook() {
  const workbook = Workbook.create();
  const assumptions = workbook.worksheets.add("Assumptions");
  assumptions.getRange("A1").write([
    ["Driver", "Value"],
    ["Monthly growth", 0.08],
    ["Gross margin", 0.62],
  ]);
  assumptions.getRange("A1:B1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
  assumptions.getRange("B2:B3").setNumberFormat("0.0%");
  assumptions.getRange("A1:B3").format.autofitColumns();

  const forecast = workbook.worksheets.add("Forecast");
  forecast.getRange("A1").writeValues(["Month", "Revenue", "Gross Profit", "Growth"]);
  forecast.getRange("A2:A5").writeValues([["Jan"], ["Feb"], ["Mar"], ["Apr"]]);
  forecast.getRange("B2").write([[100]]);
  forecast.getRange("B3:B5").formulasR1C1 = [
    ["=R[-1]C*(1+'Assumptions'!R2C2)"],
    ["=R[-1]C*(1+'Assumptions'!R2C2)"],
    ["=R[-1]C*(1+'Assumptions'!R2C2)"],
  ];
  forecast.getRange("C2:C5").formulasR1C1 = [
    ["=RC[-1]*'Assumptions'!R3C2"],
    ["=RC[-1]*'Assumptions'!R3C2"],
    ["=RC[-1]*'Assumptions'!R3C2"],
    ["=RC[-1]*'Assumptions'!R3C2"],
  ];
  forecast.getRange("D2").write({ values: [[null]] });
  forecast.getRange("D3:D5").formulasR1C1 = [["=RC[-2]/R[-1]C[-2]-1"], ["=RC[-2]/R[-1]C[-2]-1"], ["=RC[-2]/R[-1]C[-2]-1"]];

  const header = forecast.getRange("A1:D1");
  header.format = { fill: "#123B5D", font: { bold: true, color: "#FFFFFF" } };
  forecast.getRange("B2:C5").setNumberFormat("$#,##0.00");
  forecast.getRange("D2:D5").setNumberFormat("0.0%");
  forecast.getRange("A1:D5").format.autofitColumns();
  forecast.freezePanes.freezeRows(1);

  const chart = forecast.charts.add("line", forecast.getRange("A1:B5"));
  chart.name = "Revenue trend";
  chart.title = "Revenue trend";
  chart.hasLegend = false;
  chart.setPosition("F1", "M14");
  return workbook;
}

export async function createWorkbook(outputPath) {
  const workbook = buildWorkbook();
  const forecast = workbook.worksheets.getItem("Forecast");
  assert.equal(forecast.getRange("A1").getCurrentRegion().address, "A1:D5");
  assert.equal(forecast.getRange("B3").formulaInfos[0][0].kind, "stored");
  assert.equal(forecast.getRange("B3").formulasR1C1[0][0], "=R[-1]C*(1+'Assumptions'!R2C2)");
  assert.ok(forecast.getRange("D3:D5").values.every(([value]) => Math.abs(value - 0.08) < 1e-12));

  const inspection = workbook.inspect({ kind: "sheet,formula,chart,style", sheetName: "Forecast", range: "A1:M14", maxChars: 16_000 });
  assert.match(inspection.ndjson, /Revenue trend/);
  assert.match(inspection.ndjson, /Assumptions/);
  const verification = workbook.verify({ visualQa: true });
  assert.equal(verification.ok, true, verification.ndjson);
  const preview = await workbook.render({ sheetName: "Forecast", autoCrop: "all", format: "svg" });
  assert.match(await preview.text(), /Revenue trend/);

  const first = await SpreadsheetFile.exportXlsx(workbook);
  const imported = await SpreadsheetFile.importXlsx(first);
  const importedForecast = imported.worksheets.getItem("Forecast");
  assert.deepEqual(importedForecast.getRange("B3:B5").formulas, [
    ["=B2*(1+'Assumptions'!$B$2)"],
    ["=B3*(1+'Assumptions'!$B$2)"],
    ["=B4*(1+'Assumptions'!$B$2)"],
  ]);
  importedForecast.getRange("A1:D5").getColumn(3).setNumberFormat("0.00%");
  const final = await SpreadsheetFile.exportXlsx(imported);
  const roundTrip = await SpreadsheetFile.importXlsx(final);
  assert.equal(roundTrip.worksheets.getItem("Forecast").getRange("D3").format.numberFormat, "0.00%");

  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await final.save(outputPath);
  return { workbook: roundTrip, file: final, inspection, verification, preview };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const outputPath = path.resolve(process.argv[2] || "openchestnut-range-workflow.xlsx");
  const result = await createWorkbook(outputPath);
  console.log(JSON.stringify({
    outputPath,
    bytes: result.file.bytes.length,
    sheets: result.workbook.worksheets.items.length,
    verified: result.verification.ok,
  }));
}
