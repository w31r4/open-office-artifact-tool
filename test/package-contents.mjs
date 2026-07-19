import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const packageMetadata = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
assert.equal(packageMetadata.license, "AGPL-3.0-or-later");
assert.equal(packageMetadata.dependencies.mupdf, "1.28.0");
assert.equal(packageMetadata.exports["./pdf/mupdf"], "./src/pdf/mupdf.mjs");
assert.equal(packageMetadata.bin, undefined, "MuPDF must not require an installer command");
assert.equal(packageMetadata.scripts.postinstall, undefined, "MuPDF must not require npm lifecycle hooks");
const pdfFacadeSource = await fs.readFile(path.join(repoRoot, "src", "pdf", "index.mjs"), "utf8");
assert.match(pdfFacadeSource, /await import\("\.\/mupdf\.mjs"\)/, "MuPDF must load only when a PDF operation needs it");
assert.doesNotMatch(pdfFacadeSource, /from\s+["']mupdf["']/, "the root PDF facade must not initialize MuPDF eagerly");
const presentationCodecSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "open-chestnut-presentation.mjs"), "utf8");
assert.match(presentationCodecSource, /from "\.\.\/presentation\/index\.mjs";/, "the Presentation codec must depend on the Presentation leaf module");
assert.doesNotMatch(presentationCodecSource, /from "\.\.\/index\.mjs";/, "the Presentation codec must not create a back-edge to the root entry");
const spreadsheetCodecSource = await fs.readFile(path.join(repoRoot, "src", "codecs", "open-chestnut.mjs"), "utf8");
assert.match(spreadsheetCodecSource, /from "\.\.\/spreadsheet\/index\.mjs";/, "the Spreadsheet codec must depend on the Spreadsheet leaf module");
assert.doesNotMatch(spreadsheetCodecSource, /from "\.\.\/index\.mjs";/, "the Spreadsheet codec must not create a back-edge to the root entry");
const skillsNpmIgnore = await fs.readFile(path.join(repoRoot, "skills", ".npmignore"), "utf8");
assert.match(skillsNpmIgnore, /__pycache__/);
assert.match(skillsNpmIgnore, /\*\.pyc/);
const result = spawnSync("npm", ["pack", "--dry-run", "--json", "--ignore-scripts"], {
  cwd: repoRoot,
  encoding: "utf8",
});
assert.equal(result.status, 0, `npm pack manifest failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
const report = JSON.parse(result.stdout)[0];
const files = report.files.map((item) => item.path);
// npm's gzip output varies between the macOS and Linux npm builds used by local
// and hosted gates. The qpdf slice measured 9,750,068 bytes on hosted Linux, so
// keep less than 25 KB of explicit cross-platform headroom instead of setting
// the budget to one machine's exact compressed byte count.
const maxPackedBytes = 9_775_000;
// The bundled OpenChestnut runtime is an audited product payload, not an
// optional download. Keep its unpacked budget tight while allowing the
// audited PDF provider/docs growth plus the bounded DOCX/PPTX modern-comment and
// native XLSX PivotTable codecs and runnable workflows. The prior 24.05 MB
// ceiling was exceeded by 3,767 bytes after the PivotTable source-built runtime
// and Skill slice (24,053,767 bytes on the audit host), so this keeps less than
// 97 KB of explicit headroom.
// The repository-only MIT Default Template Library is excluded from the npm
// tarball. Its retained Office/PNG sources must never consume this consumer
// package budget.
const maxUnpackedBytes = 24_150_000;

for (const required of [
  "LICENSE",
  "README.md",
  "README.en.md",
  "THIRD_PARTY_NOTICES.md",
  "docs/api.md",
  "docs/reference-skills.md",
  "docs/template-library-provenance.md",
  "proto/open_office/artifact/v1/office_artifact.proto",
  "src/generated/open_office/artifact/v1/office_artifact_pb.js",
  "src/codecs/open-chestnut.mjs",
  "src/codecs/open-chestnut-error.mjs",
  "src/codecs/open-chestnut-assets.mjs",
  "src/codecs/open-chestnut-presentation.mjs",
  "src/codecs/open-chestnut-spreadsheet-pivots.mjs",
  "runtime/open-chestnut/main.mjs",
  "runtime/open-chestnut/manifest.json",
  "runtime/open-chestnut/sbom.cdx.json",
  "runtime/open-chestnut/DOTNET-LICENSE.TXT",
  "runtime/open-chestnut/DOTNET-THIRD-PARTY-NOTICES.TXT",
  "runtime/open-chestnut/_framework/dotnet.native.wasm",
  "runtime/open-chestnut/_framework/OpenChestnut.Codec.wasm",
  "runtime/open-chestnut/_framework/OpenChestnut.Runtime.wasm",
  "src/ooxml/docx-comments.mjs",
  "src/ooxml/docx-bibliography.mjs",
  "src/ooxml/package.mjs",
  "src/presentation/chart-trendline-svg.mjs",
  "src/presentation/index.mjs",
  "src/presentation/ooxml-chart-data.mjs",
  "src/presentation/ooxml-charts.mjs",
  "src/presentation/ooxml-hyperlinks.mjs",
  "src/presentation/ooxml-custom-shows.mjs",
  "src/ooxml/docx-links.mjs",
  "src/ooxml/docx-numbering.mjs",
  "src/ooxml/docx-sections.mjs",
  "src/pdf/table-grid.mjs",
  "src/pdf/reading-order.mjs",
  "src/pdf/accessibility.mjs",
  "src/pdf/index.mjs",
  "src/pdf/mupdf.mjs",
  "src/document/index.mjs",
  "src/help/index.mjs",
  "src/index.mjs",
  "src/ooxml/docx-source-references.mjs",
  "src/ooxml/docx-settings.mjs",
  "src/ooxml/pptx-package-semantics.mjs",
  "src/ooxml/pptx-source-references.mjs",
  "src/ooxml/source-reference-xml.mjs",
  "src/ooxml/source-references.mjs",
  "src/presentation/ooxml-theme.mjs",
  "src/presentation/group-shapes.mjs",
  "src/presentation/native-objects.mjs",
  "src/presentation/compose.mjs",
  "src/presentation/custom-geometry.mjs",
  "src/presentation/text-paragraphs.mjs",
  "src/presentation/ooxml-masters.mjs",
  "src/presentation/ooxml-modern-comments.mjs",
  "src/shared/colors.mjs",
  "src/shared/binary.mjs",
  "src/shared/file-blob.mjs",
  "src/shared/ids.mjs",
  "src/shared/images.mjs",
  "src/shared/inspection.mjs",
  "src/shared/png.mjs",
  "src/shared/render-output.mjs",
  "src/shared/text-range.mjs",
  "src/shared/xml.mjs",
  "src/spreadsheet/formula-criteria.mjs",
  "src/spreadsheet/index.mjs",
  "src/spreadsheet/data-tables.mjs",
  "src/codecs/open-chestnut-spreadsheet-data-tables.mjs",
  "src/spreadsheet/formula-coercion.mjs",
  "src/spreadsheet/chart-source-data.mjs",
  "src/spreadsheet/ooxml-styles.mjs",
  "src/spreadsheet/ooxml-threaded-comments.mjs",
  "src/spreadsheet/pivot-dates.mjs",
  "src/spreadsheet/pivot-filters.mjs",
  "src/spreadsheet/pivot-formulas.mjs",
  "src/spreadsheet/pivot-groups.mjs",
  "src/spreadsheet/pivots.mjs",
  "src/spreadsheet/range-addressing.mjs",
  "src/spreadsheet/range-operations.mjs",
  "src/spreadsheet/structured-references.mjs",
  "native/OfficeBridge/src/OfficeBridge.csproj",
  "skills/documents/.codex-plugin/plugin.json",
  "skills/documents/README.md",
  "skills/documents/assets/icon.png",
  "skills/documents/skills/documents/SKILL.md",
  "skills/documents/skills/documents/agents/openai.yaml",
  "skills/documents/skills/documents/LICENSE.txt",
  "skills/documents/skills/documents/artifact_tool/API_QUICK_START.md",
  "skills/documents/skills/documents/examples/openchestnut-end-to-end.mjs",
  "skills/documents/skills/documents/examples/openchestnut-classic-comment-edit-workflow.mjs",
  "skills/documents/skills/documents/examples/end_to_end_smoke_test.md",
  "skills/documents/skills/documents/render_docx.py",
  "skills/documents/skills/documents/scripts/docx_ooxml_patch.py",
  "skills/documents/skills/documents/tasks/create_edit.md",
  "skills/spreadsheets/.codex-plugin/plugin.json",
  "skills/spreadsheets/.app.json",
  "skills/spreadsheets/README.md",
  "skills/spreadsheets/skills/spreadsheets/SKILL.md",
  "skills/spreadsheets/skills/spreadsheets/agents/openai.yaml",
  "skills/spreadsheets/skills/spreadsheets/artifact_tool_docs/API_QUICK_START.md",
  "skills/spreadsheets/skills/spreadsheets/features/charts.md",
  "skills/spreadsheets/skills/spreadsheets/features/pivot-tables.md",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-range-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-sparkline-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-data-table-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-pivot-table-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-loan-amortization-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-asset-depreciation-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-scatter-chart-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-bubble-chart-workflow.mjs",
  "skills/spreadsheets/skills/spreadsheets/examples/openchestnut-growth-assumption-edit-workflow.mjs",
  "skills/spreadsheets/skills/excel-live-control/SKILL.md",
  "skills/spreadsheets/skills/excel-live-control/agents/openai.yaml",
  "skills/spreadsheets/skills/excel-live-control/assets/file-spreadsheet.png",
  "skills/presentations/.codex-plugin/plugin.json",
  "skills/presentations/README.md",
  "skills/presentations/skills/presentations/SKILL.md",
  "skills/presentations/skills/presentations/agents/openai.yaml",
  "skills/presentations/skills/presentations/artifact_tool/API_QUICK_START.md",
  "skills/presentations/skills/presentations/examples/openchestnut-title-notes-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/openchestnut-slide-name-edit-workflow.mjs",
  "skills/presentations/skills/presentations/examples/openchestnut-slide-duplicate-workflow.mjs",
  "skills/presentations/skills/presentations/artifact_tool/api/references/ole-workbooks.spec.md",
  "skills/presentations/skills/presentations/container_tools/artifact_tool_utils.mjs",
  "skills/presentations/skills/presentations/container_tools/slides_test.py",
  "skills/presentations/skills/presentations/builtin_templates_support/scripts/create-presentation.mjs",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/artifact-tool-compose/index.mjs",
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library/assets/previews/layout-library.png",
  "skills/template-creator/.codex-plugin/plugin.json",
  "skills/template-creator/manifest.json",
  "skills/template-creator/README.md",
  "skills/template-creator/assets/icon.svg",
  "skills/template-creator/skills/template-creator/SKILL.md",
  "skills/template-creator/skills/template-creator/agents/agent.yaml",
  "skills/template-creator/skills/template-creator/manifest.txt",
  "skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs",
  "skills/pdf/.codex-plugin/plugin.json",
  "skills/pdf/README.md",
  "skills/pdf/skills/pdf/SKILL.md",
  "skills/pdf/skills/pdf/agents/openai.yaml",
  "skills/pdf/skills/pdf/manifest.txt",
  "skills/pdf/skills/pdf/artifact_tool/API_QUICK_START.md",
  "skills/pdf/skills/pdf/examples/public-api-end-to-end.mjs",
  "skills/pdf/skills/pdf/examples/accessible-board-report.mjs",
  "skills/pdf/skills/pdf/examples/provider-workflows.md",
  "skills/pdf/skills/pdf/examples/reportlab-report-spec.json",
  "skills/pdf/skills/pdf/examples/pymupdf-edit-operations.json",
  "skills/pdf/skills/pdf/examples/pymupdf-redaction-operations.json",
  "skills/pdf/skills/pdf/references/PROVIDER_MATRIX.md",
  "skills/pdf/skills/pdf/references/SAVE_POLICIES.md",
  "skills/pdf/skills/pdf/references/SECURITY_CHECKLIST.md",
  "skills/pdf/skills/pdf/references/PRODUCT_BOUNDARIES.md",
  "skills/pdf/skills/pdf/references/AUDIT_SCHEMA.md",
  "skills/pdf/skills/pdf/references/pdf-audit-v1.schema.json",
  "skills/pdf/skills/pdf/scripts/pdf_provider.py",
  "skills/pdf/skills/pdf/scripts/qpdf_provider.py",
  "skills/pdf/skills/pdf/scripts/mupdf.mjs",
  "skills/pdf/skills/pdf/scripts/reportlab_create.py",
  "skills/pdf/skills/pdf/scripts/pdfplumber_extract.py",
  "skills/pdf/skills/pdf/scripts/pypdf_edit.py",
  "skills/pdf/skills/pdf/scripts/pymupdf_edit.py",
  "skills/pdf/skills/pdf/scripts/residue_scan.py",
  "skills/pdf/skills/pdf/scripts/pdf_audit.py",
  "skills/pdf/skills/pdf/scripts/python_runtime.py",
  "skills/pdf/skills/pdf/tasks/create.md",
  "skills/pdf/skills/pdf/tasks/read_review.md",
  "skills/pdf/skills/pdf/tasks/edit_existing.md",
  "skills/pdf/skills/pdf/tasks/forms_annotations.md",
  "skills/pdf/skills/pdf/tasks/sign_verify.md",
  "skills/pdf/skills/pdf/tasks/redact.md",
  "skills/pdf/skills/pdf/tasks/accessibility.md",
  "skills/pdf/skills/pdf/tasks/render_review.md",
  "skills/pdf/skills/pdf/tasks/provider_setup.md",
  "skills/pdf/skills/pdf/tasks/repair_linearize.md",
]) {
  assert.ok(files.includes(required), `npm package is missing ${required}`);
}
assert.ok(files.every((file) => !file.includes("/bin/") && !file.includes("/obj/")), "npm package must exclude dotnet bin/obj build output");
for (const removed of [
  "src/codecs/openxml-wasm.mjs",
  "src/codecs/office-codec-policy.mjs",
  "skills/shared/open-chestnut-compat.mjs",
  "src/spreadsheet/ooxml-drawings.mjs",
  "src/spreadsheet/ooxml-pivots.mjs",
  "src/presentation/master-graph.mjs",
  "src/presentation/opaque-objects.mjs",
  "src/presentation/ooxml-picture-bullets.mjs",
]) assert.ok(!files.includes(removed), `npm package must not contain removed legacy Office implementation ${removed}`);
assert.ok(!files.includes("skills/reference-sync.json"), "npm package must exclude the repository-only reference source snapshot");
assert.ok(files.every((file) => !file.includes("/tests/") && !file.startsWith("test/")), "npm package must exclude development-only test sources");
assert.ok(files.every((file) => !file.includes(".DS_Store") && !file.includes("__pycache__") && !file.endsWith(".pyc")), "npm package must exclude local metadata and Python bytecode");
assert.ok(files.every((file) => !file.startsWith("reference/")), "npm package must exclude reference material");
assert.ok(!files.includes("native/OfficeBridge/OfficeBridge.sln"), "npm package must not publish a solution whose test project is repository-only");
assert.ok(files.every((file) => !file.startsWith("skills/default-template-library/")), "npm package must exclude the repository-only retained default template library");
assert.ok(files.every((file) => !file.startsWith("native/OpenChestnut/") && !file.startsWith("scripts/")), "npm runtime package must not duplicate repository-only OpenChestnut source or build tooling");
assert.ok(files.every((file) => !file.startsWith("evals/") && file !== "docs/agent-evals.md"), "npm runtime package must exclude the evaluator-side PromptBench and its oracle documentation");
assert.ok(!files.includes("docs/coverage.md") && !files.includes("docs/release.md") && !files.includes("docs/reference-runtime-architecture.md") && !files.includes("native/OpenChestnut/README.md"), "npm runtime package must exclude repository-only coverage, release history, and subsystem implementation notes");
assert.ok(report.size < maxPackedBytes, `npm package archive unexpectedly large: ${report.size} (limit ${maxPackedBytes})`);
assert.ok(report.unpackedSize < maxUnpackedBytes, `npm package unpacked size unexpectedly large: ${report.unpackedSize} (limit ${maxUnpackedBytes})`);

console.log("package contents smoke ok");
