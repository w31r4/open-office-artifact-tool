// Deprecated compatibility entrypoint. Canonical build ownership lives in
// build-open-chestnut.mjs; retain the old output variable for one migration
// window without teaching the new build implementation about the old name.
if (process.env.OPENXML_WASM_OUTPUT && !process.env.OPEN_CHESTNUT_OUTPUT) {
  process.env.OPEN_CHESTNUT_OUTPUT = process.env.OPENXML_WASM_OUTPUT;
}
await import("./build-open-chestnut.mjs");
