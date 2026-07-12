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
const maxUnpackedBytes = 1_340_000;

for (const required of [
  "THIRD_PARTY_NOTICES.md",
  "src/index.mjs",
  "src/ooxml/docx-source-references.mjs",
  "src/ooxml/docx-settings.mjs",
  "src/ooxml/source-reference-xml.mjs",
  "src/ooxml/source-references.mjs",
  "src/shared/colors.mjs",
  "src/spreadsheet/formula-criteria.mjs",
  "src/spreadsheet/ooxml-drawings.mjs",
  "src/spreadsheet/ooxml-pivots.mjs",
  "src/spreadsheet/ooxml-styles.mjs",
  "src/spreadsheet/structured-references.mjs",
  "native/OfficeBridge/src/OfficeBridge.csproj",
  "native/OfficeBridge/tests/BridgeProtocolTests.cs",
  "skills/spreadsheets/SKILL.md",
  "skills/spreadsheets/scripts/verify-workbook.mjs",
  "skills/spreadsheets/fixtures/formula-summary.json",
  "skills/documents/SKILL.md",
  "skills/documents/scripts/verify-document.mjs",
  "skills/documents/fixtures/business-brief.json",
  "skills/documents/fixtures/package-comments.json",
  "skills/documents/fixtures/package-numbering.json",
  "skills/documents/fixtures/package-settings.json",
  "skills/presentations/SKILL.md",
  "skills/presentations/scripts/verify-presentation.mjs",
  "skills/presentations/fixtures/agent-readiness.json",
  "skills/pdf/SKILL.md",
  "skills/pdf/scripts/verify-pdf.mjs",
  "skills/pdf/fixtures/qa-report.json",
]) {
  assert.ok(files.includes(required), `npm package is missing ${required}`);
}
assert.ok(files.every((file) => !file.includes("/bin/") && !file.includes("/obj/")), "npm package must exclude dotnet bin/obj build output");
assert.ok(files.every((file) => !file.startsWith("handoff/") && !file.startsWith("reference/")), "npm package must exclude handoff and reference reference material");
assert.ok(report.unpackedSize < maxUnpackedBytes, `npm package unpacked size unexpectedly large: ${report.unpackedSize} (limit ${maxUnpackedBytes})`);

console.log("package contents smoke ok");
