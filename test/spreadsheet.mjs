import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { FileBlob, SpreadsheetFile, Workbook } from "../src/index.mjs";

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Sheet1");
sheet.getRange("A1:C3").values = [["A", "B", "Sum"], [2, 3, null], [5, 7, null]];
sheet.getRange("C2:C3").formulas = [["=A2+B2"], ["=A3+B3"]];
workbook.recalculate();

const inspect = workbook.inspect({ kind: "workbook,sheet,table,formula", range: "A1:C3", include: "values,formulas" });
assert.match(inspect.ndjson, /"value":5|"values":/);
assert.match(inspect.ndjson, /"formula":"=A2\+B2"/);
assert.match(workbook.help("fx.PMT").ndjson, /fx.PMT/);
const trace = workbook.trace("Sheet1!C2");
assert.equal(trace.tree.address, "C2");
assert.equal(trace.tree.value, 5);
assert.deepEqual(trace.tree.precedents.map((node) => node.address), ["A2", "B2"]);
assert.match(trace.ndjson, /"precedents":\["Sheet1!A2","Sheet1!B2"\]/);
assert.match(workbook.help("workbook.trace").ndjson, /precedent tree/);

const preview = await workbook.render({ sheetName: "Sheet1", range: "A1:C3" });
assert.equal(preview.type, "image/svg+xml");
assert.match(await preview.text(), /<svg/);

const xlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal(xlsx.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.xlsx`);
await xlsx.save(out);
const loaded = await SpreadsheetFile.importXlsx(await FileBlob.load(out));
const roundtrip = loaded.inspect({ kind: "table,formula", range: "A1:C3" }).ndjson;
assert.match(roundtrip, /"values":\[\["A","B","Sum"\],\[2,3,5\],\[5,7,12\]\]/);
assert.match(roundtrip, /"formula":"=A2\+B2"/);
assert.deepEqual(loaded.trace("Sheet1!C3").tree.precedents.map((node) => node.address), ["A3", "B3"]);
console.log("spreadsheet smoke ok");
