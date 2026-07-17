import assert from "node:assert/strict";
import JSZip from "jszip";

import { SpreadsheetFile, Workbook, WorksheetDataTableCollection } from "../src/index.mjs";
import { WorksheetDataTableCollection as LeafWorksheetDataTableCollection } from "../src/spreadsheet/index.mjs";

assert.strictEqual(WorksheetDataTableCollection, LeafWorksheetDataTableCollection);

function buildWorkbook() {
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("WhatIf");
  workbook.worksheets.add("Other");

  sheet.getRange("A1:B3").values = [[null, 10], [1, null], [2, null]];
  sheet.getRange("A1").formulas = [["=D1*2"]];
  sheet.getRange("D1").values = [[5]];
  assert.equal(sheet.dataTables.add("A1:B3", { rowInput: "$D$1" }), undefined);

  sheet.getRange("E1:F3").values = [[null, 15], [1, null], [2, null]];
  sheet.getRange("E1").formulas = [["=D1*3"]];
  sheet.dataTables.add(sheet.getRange("E1:F3"), { columnInput: "WhatIf!D1" });

  sheet.getRange("I1:K3").values = [[null, 10, 20], [1, null, null], [2, null, null]];
  sheet.getRange("I1").formulas = [["=M1*M2"]];
  sheet.getRange("M1:M2").values = [[5], [7]];
  sheet.dataTables.add("I1:K3", { rowInput: "M1", columnInput: "M2" });
  return { workbook, sheet };
}

function formulaTags(xml) {
  return [...xml.matchAll(/<x:f\b[^>]*t="dataTable"[^>]*>/g)].map((match) => match[0]);
}

const { workbook, sheet } = buildWorkbook();
assert.ok(sheet.dataTables instanceof WorksheetDataTableCollection);
assert.deepEqual(sheet.dataTables.__getDefinitions(), [
  {
    range: { startRow: 1, startCol: 1, endRow: 2, endCol: 1 },
    formulaRef: "B2:B3",
    anchor: { row: 1, col: 1, address: "B2" },
    rowInput: "D1",
    rowOriented: true,
    twoVariable: false,
    displayFormula: "{=TABLE(D1)}",
  },
  {
    range: { startRow: 1, startCol: 5, endRow: 2, endCol: 5 },
    formulaRef: "F2:F3",
    anchor: { row: 1, col: 5, address: "F2" },
    columnInput: "D1",
    rowOriented: false,
    twoVariable: false,
    displayFormula: "{=TABLE(D1)}",
  },
  {
    range: { startRow: 1, startCol: 9, endRow: 2, endCol: 10 },
    formulaRef: "J2:K3",
    anchor: { row: 1, col: 9, address: "J2" },
    rowInput: "M1",
    columnInput: "M2",
    rowOriented: false,
    twoVariable: true,
    displayFormula: "{=TABLE(M1,M2)}",
  },
]);
const defensive = sheet.dataTables.__getDefinitions();
defensive[0].rowInput = "A1";
defensive[0].range.startRow = 99;
assert.equal(sheet.dataTables.__getDefinitions()[0].rowInput, "D1");
assert.equal(sheet.dataTables.__getDefinitions()[0].range.startRow, 1);
assert.match(workbook.inspect({ kind: "dataTable", maxChars: 8_000 }).ndjson, /"displayFormula":"\{=TABLE\(M1,M2\)\}"/);

const firstXlsx = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
const firstZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
const firstXml = await firstZip.file("xl/worksheets/sheet1.xml").async("text");
const firstTags = formulaTags(firstXml);
assert.equal(firstTags.length, 3);
assert.match(firstTags[0], /ref="B2:B3".*dt2D="0".*dtr="1".*r1="D1"/);
assert.match(firstTags[1], /ref="F2:F3".*dt2D="0".*dtr="0".*r2="D1"/);
assert.match(firstTags[2], /ref="J2:K3".*dt2D="1".*r1="M1".*r2="M2"/);
assert.doesNotMatch(firstTags[2], /\bdtr=/);

const imported = await SpreadsheetFile.importXlsx(firstXlsx);
const importedSheet = imported.worksheets.getItem("WhatIf");
assert.deepEqual(importedSheet.dataTables.__getDefinitions(), sheet.dataTables.__getDefinitions());
assert.equal(importedSheet.getRange("B2").formulas[0][0], "");
importedSheet.getRange("N1").values = [["ordinary edit"]];
const secondXlsx = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
const second = await SpreadsheetFile.importXlsx(secondXlsx);
assert.equal(second.worksheets.getItem("WhatIf").getRange("N1").values[0][0], "ordinary edit");
assert.deepEqual(second.worksheets.getItem("WhatIf").dataTables.__getDefinitions(), sheet.dataTables.__getDefinitions());

const addition = await SpreadsheetFile.importXlsx(firstXlsx);
const additionSheet = addition.worksheets.getItem("WhatIf");
additionSheet.getRange("O1:P2").values = [[null, 1], [2, null]];
additionSheet.getRange("O1").formulas = [["=D1"]];
additionSheet.dataTables.add("O1:P2", { rowInput: "D1" });
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(addition, { recalculate: false }),
  (error) => error?.code === "invalid_spreadsheet_data_table_topology" && /cannot add data table/i.test(error.message),
);

const removal = await SpreadsheetFile.importXlsx(firstXlsx);
removal.worksheets.getItem("WhatIf").dataTables._definitions.shift();
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(removal, { recalculate: false }),
  (error) => error?.code === "invalid_spreadsheet_data_table_topology" && /cannot remove or reorder/i.test(error.message),
);

const advancedZip = await JSZip.loadAsync(new Uint8Array(await firstXlsx.arrayBuffer()));
advancedZip.file("xl/worksheets/sheet1.xml", firstXml.replace(/(<x:f\b[^>]*t="dataTable"[^>]*r1="D1")/, "$1 ca=\"1\""));
const advancedBytes = await advancedZip.generateAsync({ type: "uint8array" });
const advanced = await SpreadsheetFile.importXlsx(advancedBytes);
const advancedSheet = advanced.worksheets.getItem("WhatIf");
const preservedAdvanced = await SpreadsheetFile.exportXlsx(advanced, { recalculate: false });
const preservedZip = await JSZip.loadAsync(new Uint8Array(await preservedAdvanced.arrayBuffer()));
assert.match(await preservedZip.file("xl/worksheets/sheet1.xml").async("text"), /t="dataTable"[^>]*ca="1"/);
advancedSheet.dataTables._definitions[0].rowInput = "D2";
await assert.rejects(
  () => SpreadsheetFile.exportXlsx(advanced, { recalculate: false }),
  (error) => error?.code === "unsupported_spreadsheet_data_table_edit" && /source-bound and read-only/i.test(error.message),
);

const invalidWorkbook = Workbook.create();
const invalidSheet = invalidWorkbook.worksheets.add("Invalid");
invalidSheet.getRange("A1").formulas = [["=1"]];
assert.throws(() => invalidSheet.dataTables.add("A1", { rowInput: "D1" }), /at least one row and one column/i);
assert.throws(() => invalidSheet.dataTables.add("A1:B2", {}), /at least a rowInput or columnInput/i);
assert.throws(() => invalidSheet.dataTables.add("A1:B2", { rowInput: "Other!D1" }), /must be on sheet "Invalid"/i);
assert.throws(() => invalidSheet.dataTables.add("A1:B2", { rowInput: "XFE1" }), /outside the worksheet/i);
invalidSheet.dataTables.add("A1:B2", { rowInput: "D1" });
assert.throws(() => invalidSheet.dataTables.add("A1:C3", { columnInput: "D2" }), /overlaps B2/i);
assert.throws(() => invalidSheet.dataTables.add("C1:D2", { rowInput: "D1" }), /top-left formula cell/i);

console.log("spreadsheet data table tests passed");
