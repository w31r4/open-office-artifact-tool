/**
 * @deprecated Use open-office-artifact-tool/codecs/open-chestnut.
 * This module is a compatibility-only name bridge; all implementation and
 * runtime ownership lives in open-chestnut.mjs.
 */
export {
  OPEN_CHESTNUT_PROTOCOL_VERSION,
  OPEN_CHESTNUT_PROTOCOL_VERSION as OPEN_XML_WASM_PROTOCOL_VERSION,
  OpenChestnutCodecError,
  OpenChestnutCodecError as OpenXmlWasmCodecError,
  exportDocxWithOpenChestnut,
  exportDocxWithOpenChestnut as exportDocxWithOpenXmlWasm,
  exportPptxWithOpenChestnut,
  exportPptxWithOpenChestnut as exportPptxWithOpenXmlWasm,
  exportXlsxWithOpenChestnut,
  exportXlsxWithOpenChestnut as exportXlsxWithOpenXmlWasm,
  importDocxWithOpenChestnut,
  importDocxWithOpenChestnut as importDocxWithOpenXmlWasm,
  importPptxWithOpenChestnut,
  importPptxWithOpenChestnut as importPptxWithOpenXmlWasm,
  importXlsxWithOpenChestnut,
  importXlsxWithOpenChestnut as importXlsxWithOpenXmlWasm,
  invokeOpenChestnut,
  invokeOpenChestnut as invokeOpenXmlWasm,
  openChestnutStatus,
  openChestnutStatus as openXmlWasmStatus,
} from "./open-chestnut.mjs";
