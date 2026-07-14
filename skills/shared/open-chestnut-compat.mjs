export function normalizeOpenChestnutCodecName(value) {
  const name = String(value || "").trim().toLowerCase();
  return name === "openxml-wasm" ? "open-chestnut" : name;
}

export function documentOpenChestnutEdits(fixture = {}) {
  return fixture.openChestnutEdits ?? fixture.openXmlWasmEdits ?? [];
}

export function presentationOpenChestnutConfig(fixture = {}) {
  return fixture.openChestnut ?? fixture.openXmlWasm;
}
