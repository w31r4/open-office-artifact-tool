export class OpenXmlWasmCodecError extends Error {
  constructor(message, diagnostics = [], options = {}) {
    super(message, options);
    this.name = "OpenXmlWasmCodecError";
    this.code = diagnostics[0]?.code || options.code || "openxml_wasm_codec_error";
    this.diagnostics = diagnostics;
  }
}
