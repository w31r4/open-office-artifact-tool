import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "openxml-wasm-reproducibility-"));
const first = path.join(temporary, "first");
const second = path.join(temporary, "second");

try {
  build(first);
  build(second);
  const firstSnapshot = snapshot(first);
  const secondSnapshot = snapshot(second);
  if (JSON.stringify(firstSnapshot) !== JSON.stringify(secondSnapshot)) {
    const firstByPath = new Map(firstSnapshot.map((item) => [item.path, item]));
    const secondByPath = new Map(secondSnapshot.map((item) => [item.path, item]));
    const changed = [...new Set([...firstByPath.keys(), ...secondByPath.keys()])]
      .filter((file) => firstByPath.get(file)?.sha256 !== secondByPath.get(file)?.sha256)
      .slice(0, 20);
    throw new Error(`OpenXML WASM builds differ on the same host: ${changed.join(", ") || "file inventory changed"}.`);
  }
  console.log(`OpenXML WASM reproducibility ok: ${firstSnapshot.length} files.`);
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

function build(output) {
  const result = spawnSync(process.execPath, ["scripts/build-openxml-wasm.mjs"], {
    cwd: repoRoot,
    env: { ...process.env, OPENXML_WASM_OUTPUT: output },
    stdio: "inherit",
  });
  if (result.status !== 0) process.exit(result.status || 1);
}

function snapshot(root, base = root) {
  return fs.readdirSync(root, { withFileTypes: true }).flatMap((entry) => {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) return snapshot(target, base);
    const bytes = fs.readFileSync(target);
    return [{
      path: path.relative(base, target).split(path.sep).join("/"),
      bytes: bytes.byteLength,
      sha256: createHash("sha256").update(bytes).digest("hex"),
    }];
  }).sort((left, right) => left.path.localeCompare(right.path));
}
