import assert from "node:assert/strict";
import { Workbook } from "../src/index.mjs";
import {
  OPEN_CHESTNUT_PROTOCOL_VERSION,
  OPEN_XML_WASM_PROTOCOL_VERSION,
  OpenChestnutCodecError,
  OpenXmlWasmCodecError,
  exportXlsxWithOpenChestnut,
  exportXlsxWithOpenXmlWasm,
  importXlsxWithOpenChestnut,
  importXlsxWithOpenXmlWasm,
  invokeOpenChestnut,
  invokeOpenXmlWasm,
  openChestnutStatus,
  openXmlWasmStatus,
} from "../src/codecs/openxml-wasm.mjs";

assert.equal(OPEN_XML_WASM_PROTOCOL_VERSION, OPEN_CHESTNUT_PROTOCOL_VERSION);
assert.equal(OpenXmlWasmCodecError, OpenChestnutCodecError);
assert.equal(exportXlsxWithOpenXmlWasm, exportXlsxWithOpenChestnut);
assert.equal(importXlsxWithOpenXmlWasm, importXlsxWithOpenChestnut);
assert.equal(invokeOpenXmlWasm, invokeOpenChestnut);
assert.equal(openXmlWasmStatus, openChestnutStatus);

const workbook = Workbook.create();
workbook.worksheets.add("Compatibility").getRange("A1").values = [["OpenChestnut"]];
const file = await exportXlsxWithOpenXmlWasm(workbook);
const imported = await importXlsxWithOpenXmlWasm(file);
assert.equal(file.metadata.codec, "open-chestnut");
assert.equal(imported.worksheets.getItem("Compatibility").getRange("A1").values[0][0], "OpenChestnut");

console.log("openxml-wasm compatibility smoke ok");
