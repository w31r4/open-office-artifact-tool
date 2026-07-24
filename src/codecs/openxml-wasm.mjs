/**
 * @deprecated Import from open-office-artifact-tool/codecs/open-chestnut.
 *
 * This is a name-only compatibility bridge for the 0.x line. It has no
 * implementation, runtime, warning, selector, or fallback of its own: every
 * canonical export is the exact OpenChestnut binding and every historical
 * OpenXmlWasm name below is an alias for that same binding.
 */
export * from "./open-chestnut.mjs";

export {
  OPEN_CHESTNUT_PROTOCOL_VERSION as OPEN_XML_WASM_PROTOCOL_VERSION,
  OpenChestnutCodecError as OpenXmlWasmCodecError,
  exportDocxWithOpenChestnut as exportDocxWithOpenXmlWasm,
  exportPptxWithOpenChestnut as exportPptxWithOpenXmlWasm,
  exportXlsxWithOpenChestnut as exportXlsxWithOpenXmlWasm,
  importDocxWithOpenChestnut as importDocxWithOpenXmlWasm,
  importPptxWithOpenChestnut as importPptxWithOpenXmlWasm,
  importXlsxWithOpenChestnut as importXlsxWithOpenXmlWasm,
  invokeOpenChestnut as invokeOpenXmlWasm,
  openChestnutStatus as openXmlWasmStatus,
} from "./open-chestnut.mjs";
