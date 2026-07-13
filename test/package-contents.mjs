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
const maxPackedBytes = 500_000;
const maxUnpackedBytes = 2_150_000;

for (const required of [
  "THIRD_PARTY_NOTICES.md",
  "src/ooxml/docx-comments.mjs",
  "src/ooxml/docx-bibliography.mjs",
  "src/presentation/chart-trendline-svg.mjs",
  "src/presentation/ooxml-chart-data.mjs",
  "src/presentation/ooxml-charts.mjs",
  "src/presentation/ooxml-picture-bullets.mjs",
  "src/presentation/ooxml-hyperlinks.mjs",
  "src/presentation/ooxml-custom-shows.mjs",
  "src/ooxml/docx-links.mjs",
  "src/ooxml/docx-numbering.mjs",
  "src/ooxml/docx-sections.mjs",
  "src/pdf/table-grid.mjs",
  "src/pdf/reading-order.mjs",
  "src/pdf/accessibility.mjs",
  "src/index.mjs",
  "src/ooxml/docx-source-references.mjs",
  "src/ooxml/docx-settings.mjs",
  "src/ooxml/pptx-package-semantics.mjs",
  "src/ooxml/pptx-source-references.mjs",
  "src/ooxml/source-reference-xml.mjs",
  "src/ooxml/source-references.mjs",
  "src/presentation/ooxml-theme.mjs",
  "src/presentation/group-shapes.mjs",
  "src/presentation/opaque-objects.mjs",
  "src/presentation/text-paragraphs.mjs",
  "src/presentation/ooxml-masters.mjs",
  "src/presentation/ooxml-modern-comments.mjs",
  "src/presentation/master-graph.mjs",
  "src/shared/colors.mjs",
  "src/spreadsheet/formula-criteria.mjs",
  "src/spreadsheet/formula-coercion.mjs",
  "src/spreadsheet/ooxml-drawings.mjs",
  "src/spreadsheet/ooxml-pivots.mjs",
  "src/spreadsheet/ooxml-styles.mjs",
  "src/spreadsheet/ooxml-threaded-comments.mjs",
  "src/spreadsheet/pivot-dates.mjs",
  "src/spreadsheet/pivot-filters.mjs",
  "src/spreadsheet/pivot-formulas.mjs",
  "src/spreadsheet/pivot-groups.mjs",
  "src/spreadsheet/pivots.mjs",
  "src/spreadsheet/structured-references.mjs",
  "native/OfficeBridge/src/OfficeBridge.csproj",
  "native/OfficeBridge/tests/BridgeProtocolTests.cs",
  "skills/spreadsheets/SKILL.md",
  "skills/shared/visual-baselines.mjs",
  "skills/spreadsheets/scripts/verify-workbook.mjs",
  "skills/spreadsheets/fixtures/formula-summary.json",
  "skills/spreadsheets/fixtures/structured-intersection.json",
  "skills/documents/SKILL.md",
  "skills/documents/scripts/verify-document.mjs",
  "skills/documents/fixtures/business-brief.json",
  "skills/documents/fixtures/package-comments.json",
  "skills/documents/fixtures/package-numbering.json",
  "skills/documents/fixtures/package-settings.json",
  "skills/presentations/SKILL.md",
  "skills/presentations/scripts/verify-presentation.mjs",
  "skills/presentations/fixtures/agent-readiness.json",
  "skills/presentations/fixtures/package-drawing.json",
  "skills/presentations/fixtures/package-notes-comments.json",
  "skills/presentations/fixtures/modern-comments.json",
  "skills/pdf/SKILL.md",
  "skills/pdf/scripts/verify-pdf.mjs",
  "skills/pdf/fixtures/qa-report.json",
]) {
  assert.ok(files.includes(required), `npm package is missing ${required}`);
}
assert.ok(files.every((file) => !file.includes("/bin/") && !file.includes("/obj/")), "npm package must exclude dotnet bin/obj build output");
assert.ok(files.every((file) => !file.startsWith("handoff/") && !file.startsWith("reference/")), "npm package must exclude handoff and reference material");
assert.ok(report.size < maxPackedBytes, `npm package archive unexpectedly large: ${report.size} (limit ${maxPackedBytes})`);
assert.ok(report.unpackedSize < maxUnpackedBytes, `npm package unpacked size unexpectedly large: ${report.unpackedSize} (limit ${maxUnpackedBytes})`);

console.log("package contents smoke ok");
