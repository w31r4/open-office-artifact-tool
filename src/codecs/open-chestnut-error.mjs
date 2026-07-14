export class OpenChestnutCodecError extends Error {
  constructor(message, diagnostics = [], options = {}) {
    super(message, options);
    this.name = "OpenChestnutCodecError";
    this.code = diagnostics[0]?.code || options.code || "open_chestnut_codec_error";
    this.diagnostics = diagnostics;
  }
}
