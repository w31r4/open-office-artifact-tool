import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const inputsPath = path.join(root, "scripts", "pdf-provider-ocr-release-inputs.v1.json");
const pythonInputsPath = path.join(root, "scripts", "pdf-provider-python-release-inputs.v1.json");
const nativeBuilder = path.join(root, "scripts", "build-ocr-native-payload.mjs");
const workflowPath = path.join(root, ".github", "workflows", "pdf-ocr-capability-packs.yml");

const [inputBytes, pythonInputBytes, nativeSource, workflowSource] = await Promise.all([
  fs.readFile(inputsPath),
  fs.readFile(pythonInputsPath),
  fs.readFile(nativeBuilder, "utf8"),
  fs.readFile(workflowPath, "utf8"),
]);
const inputs = JSON.parse(inputBytes);

assert.equal(inputs.schema, "open-office-artifact-tool.pdf-provider-ocr-release-inputs.v1");
assert.equal(inputs.schemaVersion, 1);
assert.equal(inputs.ocrCore.packId, "ocr-core");
assert.equal(inputs.ocrCore.version, "17.8.1-oat.1");
assert.equal(
  inputs.ocrCore.pythonInputs.sha256,
  crypto.createHash("sha256").update(pythonInputBytes).digest("hex"),
  "the OCR release lock must pin the exact isolated-Python wheel lock it builds",
);
assert.deepEqual(inputs.ocrCore.nativeBuild["darwin-arm64"].formulae, ["tesseract", "ghostscript", "poppler"]);
assert.deepEqual(inputs.ocrCore.nativeBuild["linux-x64"].packages, ["tesseract-ocr", "tesseract-ocr-eng", "ghostscript", "poppler-utils", "poppler-data", "libgs9-common", "fonts-droid-fallback", "fonts-urw-base35", "patchelf"]);

for (const [language, expected] of Object.entries({ eng: "ocr-language-eng", chi_sim: "ocr-language-chi-sim" })) {
  const languageInput = inputs.languages[language];
  assert.equal(languageInput.packId, expected);
  assert.equal(languageInput.version, "4.1.0-oat.1");
  assert.equal(languageInput.license, "Apache-2.0");
  assert.match(languageInput.url, /^https:\/\//);
  assert.match(languageInput.sha256, /^[a-f0-9]{64}$/);
  assert.ok(Number.isSafeInteger(languageInput.downloadBytes) && languageInput.downloadBytes > 0);
}
assert.match(inputs.licenseMaterial.tessdataFastApache20.url, /^https:\/\//);
assert.match(inputs.licenseMaterial.tessdataFastApache20.sha256, /^[a-f0-9]{64}$/);
assert.ok(inputs.licenseMaterial.tessdataFastApache20.downloadBytes > 1000);

// The release builder must force every package-local executable through its
// own relocated libraries, remove all bundled language data, and leave only a
// separately authorized language-pack directory for the provider adapter.
for (const sourceFragment of [
  "removeTraineddata",
  "DYLD_FALLBACK_LIBRARY_PATH",
  "LD_LIBRARY_PATH",
  "GS_LIB",
  "codesign",
  "patchelf",
  "writeLaunchers",
  "native library basename collision",
  "MACHO_MAGICS",
  "isMachOFile",
  "listMacLibraryFiles",
  "path.basename(source)",
  "contains a dangling symlink",
  "contains a symlink directory cycle",
  "resource-root",
]) assert.match(nativeSource, new RegExp(sourceFragment.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.match(nativeSource, /if \(!await isMachOFile\(target\)\) return false;/);
assert.match(nativeSource, /@loader_path\/\$\{path\.basename\(target\)\}/);
assert.match(workflowSource, /fonts-droid-fallback/);
assert.match(workflowSource, /--resource-root/);
assert.match(workflowSource, /fonts-urw-base35/);
assert.match(workflowSource, /libgs9-common/);
assert.match(workflowSource, /poppler-data/);
assert.match(workflowSource, /resource_target/);
assert.match(workflowSource, /resource_actual/);
assert.match(workflowSource, /dpkg-query -S/);
assert.match(workflowSource, /for root_formula in tesseract ghostscript poppler; do brew deps/);
assert.match(workflowSource, /unapproved Ghostscript resource target/);
assert.doesNotMatch(workflowSource, /brew deps --include-optional/);
assert.doesNotMatch(workflowSource, /brew deps --union tesseract ghostscript poppler/);

const invalidPlatform = spawnSync(process.execPath, [nativeBuilder,
  "--platform", "win32-x64",
  "--payload", root,
  "--notices", path.join(root, "package.json"),
  "--tesseract", process.execPath,
  "--ghostscript", process.execPath,
  "--pdftotext", process.execPath,
  "--ghostscript-root", root,
  "--tessdata-root", root,
  "--library-root", root,
], { cwd: root, encoding: "utf8" });
assert.equal(invalidPlatform.status, 2);
assert.match(invalidPlatform.stderr, /platform must be one of darwin-arm64, linux-x64/);

console.log("OCR PDF capability-pack build smoke ok");
