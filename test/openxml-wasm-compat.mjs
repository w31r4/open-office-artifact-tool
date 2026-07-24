import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { Workbook } from "../src/index.mjs";
import * as openChestnut from "../src/codecs/open-chestnut.mjs";
import * as openChestnutWire from "../src/generated/open_office/artifact/v1/office_artifact_pb.js";
import * as openXmlWasm from "../src/codecs/openxml-wasm.mjs";
import * as openXmlWasmWire from "../src/codecs/openxml-wasm-wire.mjs";

const importProbe = spawnSync(process.execPath, [
  "--input-type=module",
  "--eval",
  `await import(${JSON.stringify(pathToFileURL(path.join(import.meta.dirname, "..", "src", "codecs", "openxml-wasm.mjs")).href)});`,
], { encoding: "utf8" });
assert.equal(importProbe.status, 0, `legacy import probe failed: ${importProbe.stderr}`);
assert.equal(importProbe.stdout, "", "legacy import must not write output");
assert.equal(importProbe.stderr, "", "legacy import must not emit a warning");

for (const name of Object.keys(openChestnut)) {
  assert.equal(openXmlWasm[name], openChestnut[name], `legacy module must strictly re-export ${name}`);
}
for (const name of Object.keys(openChestnutWire)) {
  assert.equal(openXmlWasmWire[name], openChestnutWire[name], `legacy wire module must strictly re-export ${name}`);
}

for (const [legacyName, canonicalName] of Object.entries({
  OPEN_XML_WASM_PROTOCOL_VERSION: "OPEN_CHESTNUT_PROTOCOL_VERSION",
  OpenXmlWasmCodecError: "OpenChestnutCodecError",
  exportDocxWithOpenXmlWasm: "exportDocxWithOpenChestnut",
  exportPptxWithOpenXmlWasm: "exportPptxWithOpenChestnut",
  exportXlsxWithOpenXmlWasm: "exportXlsxWithOpenChestnut",
  importDocxWithOpenXmlWasm: "importDocxWithOpenChestnut",
  importPptxWithOpenXmlWasm: "importPptxWithOpenChestnut",
  importXlsxWithOpenXmlWasm: "importXlsxWithOpenChestnut",
  invokeOpenXmlWasm: "invokeOpenChestnut",
  openXmlWasmStatus: "openChestnutStatus",
})) {
  assert.equal(openXmlWasm[legacyName], openChestnut[canonicalName], `${legacyName} must alias ${canonicalName}`);
}

const workbook = Workbook.create();
workbook.worksheets.add("Compatibility").getRange("A1").values = [["OpenChestnut"]];
const file = await openXmlWasm.exportXlsxWithOpenXmlWasm(workbook);
const imported = await openXmlWasm.importXlsxWithOpenXmlWasm(file);
assert.equal(file.metadata.codec, "open-chestnut");
assert.equal(imported.worksheets.getItem("Compatibility").getRange("A1").values[0][0], "OpenChestnut");

console.log("openxml-wasm compatibility smoke ok");
