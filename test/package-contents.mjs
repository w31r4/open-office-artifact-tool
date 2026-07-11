import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const result = spawnSync("npm", ["pack", "--dry-run", "--json", "--ignore-scripts"], {
  cwd: repoRoot,
  encoding: "utf8",
});
assert.equal(result.status, 0, `npm pack manifest failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
const report = JSON.parse(result.stdout)[0];
const files = report.files.map((item) => item.path);

for (const required of [
  "src/index.mjs",
  "native/OfficeBridge/src/OfficeBridge.csproj",
  "native/OfficeBridge/tests/BridgeProtocolTests.cs",
  "skills/spreadsheets/SKILL.md",
  "skills/spreadsheets/scripts/verify-workbook.mjs",
  "skills/spreadsheets/fixtures/formula-summary.json",
]) {
  assert.ok(files.includes(required), `npm package is missing ${required}`);
}
assert.ok(files.every((file) => !file.includes("/bin/") && !file.includes("/obj/")), "npm package must exclude dotnet bin/obj build output");
assert.ok(files.every((file) => !file.startsWith("handoff/") && !file.startsWith("reference/")), "npm package must exclude handoff and reference reference material");
assert.ok(report.unpackedSize < 1_000_000, `npm package unpacked size unexpectedly large: ${report.unpackedSize}`);

console.log("package contents smoke ok");

