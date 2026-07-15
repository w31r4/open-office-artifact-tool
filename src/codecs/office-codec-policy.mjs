export const OFFICE_CODEC_IDS = Object.freeze(["javascript", "open-chestnut"]);

const OFFICE_CODEC_ID_SET = new Set(OFFICE_CODEC_IDS);

export function resolveOfficeCodec(options, apiName) {
  if (options == null) return "javascript";
  if (typeof options !== "object" || Array.isArray(options)) throw new TypeError(`${apiName} options must be an object.`);
  const codec = options.codec ?? "javascript";
  if (typeof codec !== "string" || !OFFICE_CODEC_ID_SET.has(codec)) {
    const suffix = codec === "openxml-wasm"
      ? " The deprecated openxml-wasm name is available only through its compatibility subpath."
      : "";
    throw new TypeError(`${apiName} codec must be javascript or open-chestnut; received ${String(codec)}.${suffix}`);
  }
  return codec;
}

export function codecDelegateOptions(options) {
  if (options == null) return {};
  const { codec: _codec, ...delegateOptions } = options;
  return delegateOptions;
}

export async function loadOpenChestnutCodec() {
  return import("./open-chestnut.mjs");
}
