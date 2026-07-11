import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-examples-"));
const examples = [
  "create-docx-report.mjs",
  "create-xlsx-dashboard.mjs",
  "create-pptx-compose.mjs",
  "parse-render-pdf.mjs",
  "render-via-playwright.mjs",
  "render-via-native-office.mjs",
];

for (const example of examples) {
  const result = spawnSync(process.execPath, [path.join(repoRoot, "examples", example)], {
    cwd: repoRoot,
    env: { ...process.env, OUTPUT_DIR: outputDir },
    encoding: "utf8",
  });
  assert.equal(result.status, 0, `${example} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
}

const outputs = await fs.readdir(outputDir);
assert.ok(outputs.includes("docx-report.docx"));
assert.ok(outputs.includes("xlsx-dashboard.xlsx"));
assert.ok(outputs.includes("pptx-compose.pptx"));
assert.ok(outputs.includes("modeled-report.pdf"));
assert.ok(outputs.includes("modeled-report.svg"));
await fs.rm(outputDir, { recursive: true, force: true });
console.log("examples smoke ok");
