import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import {
  DocumentFile,
  FileBlob,
  PdfFile,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const skillsRoot = path.join(repoRoot, "skills");
const pluginNames = ["documents", "spreadsheets", "presentations", "pdf", "template-creator", "default-template-library"];
const defaultTemplateSkills = [
  "artifact-template-analytics-dashboard",
  "artifact-template-business-review",
  "artifact-template-design-report",
  "artifact-template-experiment-analysis",
  "artifact-template-financial-budget",
  "artifact-template-investment-committee-memo",
  "artifact-template-legal-memorandum",
  "artifact-template-market-trends-report",
  "artifact-template-minimal-letterhead",
  "artifact-template-operating-calendar",
  "artifact-template-operating-review",
  "artifact-template-project-kickoff",
  "artifact-template-project-tracker",
  "artifact-template-sales-pipeline",
  "artifact-template-simple-dark-mode",
  "artifact-template-simple-light-mode",
  "artifact-template-strategy-memorandum",
  "artifact-template-system-design",
  "artifact-template-team-alignment",
  "artifact-template-three-statement-forecast",
];
const expectedSkills = new Map([
  ["documents", ["documents"]],
  ["spreadsheets", ["excel-live-control", "spreadsheets"]],
  ["presentations", ["presentations"]],
  ["pdf", ["pdf"]],
  ["template-creator", ["template-creator"]],
  ["default-template-library", defaultTemplateSkills],
]);
const expectedDeclaredSkillNames = new Map([
  ["documents", "documents"],
  ["excel-live-control", "excel-live-control"],
  ["spreadsheets", "Spreadsheets"],
  ["presentations", "Presentations"],
  ["pdf", "pdf"],
  ["template-creator", "template-creator"],
]);
for (const skillName of defaultTemplateSkills) expectedDeclaredSkillNames.set(skillName, skillName);

async function exists(file) {
  return fs.access(file).then(() => true, () => false);
}

async function walk(root) {
  const files = [];
  for (const entry of await fs.readdir(root, { withFileTypes: true })) {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) files.push(...await walk(target));
    else if (entry.isFile()) files.push(target);
  }
  return files;
}

function yamlValue(source, key) {
  return source.match(new RegExp(`^\\s*${key}:\\s*["']?([^"'\\n]+)["']?\\s*$`, "m"))?.[1]?.trim();
}

for (const pluginName of pluginNames) {
  const pluginRoot = path.join(skillsRoot, pluginName);
  const manifestPath = path.join(pluginRoot, ".codex-plugin", "plugin.json");
  const manifest = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  assert.equal(manifest.name, pluginName);
  assert.equal(manifest.version, "0.2.0");
  assert.equal(manifest.license, pluginName === "default-template-library" ? "MIT" : "AGPL-3.0-or-later");
  assert.equal(manifest.skills, "./skills/");
  assert.match(manifest.repository, /open-office-artifact-tool/);
  assert.ok(await exists(path.join(pluginRoot, "README.md")));
  for (const iconKey of ["composerIcon", "logo"]) {
    assert.ok(await exists(path.resolve(pluginRoot, manifest.interface[iconKey])), `${pluginName} ${iconKey} must resolve inside the plugin`);
  }

  const skillNames = (await fs.readdir(path.join(pluginRoot, "skills"), { withFileTypes: true }))
    .filter((entry) => entry.isDirectory())
    .map((entry) => entry.name)
    .sort();
  assert.deepEqual(skillNames, expectedSkills.get(pluginName));
  for (const skillName of skillNames) {
    const skillRoot = path.join(pluginRoot, "skills", skillName);
    const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
    const frontmatter = skillText.match(/^---\n([\s\S]*?)\n---/);
    assert.ok(frontmatter, `${pluginName}/${skillName} is missing YAML frontmatter`);
    assert.equal(yamlValue(frontmatter[1], "name"), expectedDeclaredSkillNames.get(skillName));
    const retainedTemplateSkill = pluginName === "default-template-library";
    const agentFilename = skillName === "template-creator" || retainedTemplateSkill ? "agent.yaml" : "openai.yaml";
    const agentText = await fs.readFile(path.join(skillRoot, "agents", agentFilename), "utf8");
    for (const iconKey of retainedTemplateSkill ? ["icon_large"] : ["icon_small", "icon_large"]) {
      const icon = yamlValue(agentText, iconKey);
      assert.ok(icon, `${pluginName}/${skillName} is missing ${iconKey}`);
      assert.ok(await exists(path.resolve(skillRoot, icon)), `${pluginName}/${skillName} ${iconKey} does not resolve`);
    }
  }
}

const templateCreatorManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "template-creator", "manifest.json"), "utf8"));
assert.equal(templateCreatorManifest.schemaVersion, 1);
assert.deepEqual(templateCreatorManifest.skills, ["skills/template-creator"]);
const defaultTemplateManifest = JSON.parse(await fs.readFile(path.join(skillsRoot, "default-template-library", "manifest.json"), "utf8"));
assert.equal(defaultTemplateManifest.schemaVersion, 1);
assert.deepEqual(defaultTemplateManifest.skills, [
  "skills/artifact-template-design-report",
  "skills/artifact-template-experiment-analysis",
  "skills/artifact-template-investment-committee-memo",
  "skills/artifact-template-legal-memorandum",
  "skills/artifact-template-minimal-letterhead",
  "skills/artifact-template-strategy-memorandum",
  "skills/artifact-template-system-design",
  "skills/artifact-template-business-review",
  "skills/artifact-template-market-trends-report",
  "skills/artifact-template-operating-review",
  "skills/artifact-template-project-kickoff",
  "skills/artifact-template-simple-dark-mode",
  "skills/artifact-template-simple-light-mode",
  "skills/artifact-template-team-alignment",
  "skills/artifact-template-analytics-dashboard",
  "skills/artifact-template-financial-budget",
  "skills/artifact-template-operating-calendar",
  "skills/artifact-template-project-tracker",
  "skills/artifact-template-sales-pipeline",
  "skills/artifact-template-three-statement-forecast",
]);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "LICENSE.md")), true);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "integrity.json")), true);
assert.equal(await exists(path.join(skillsRoot, "default-template-library", "catalog.json")), false);

assert.equal(await exists(path.join(skillsRoot, "documents", "SKILL.md")), false);
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "scripts")), false);
assert.equal(await exists(path.join(skillsRoot, "presentations", "fixtures")), false);
assert.ok(await exists(path.join(repoRoot, "test", "skill-harness", "spreadsheets", "scripts", "workflow.mjs")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "artifact_tool_docs", "API_QUICK_START.md")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "features", "charts.md")));
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "features", "pivot-tables.md")));
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "API_QUICK_START.md")), false);
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "charts.md")), false);
const spreadsheetSkillText = await fs.readFile(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "SKILL.md"), "utf8");
assert.match(spreadsheetSkillText, /artifact_tool_docs\/API_QUICK_START\.md/);
assert.match(spreadsheetSkillText, /features\/charts\.md/);
assert.match(spreadsheetSkillText, /features\/pivot-tables\.md/);
assert.match(spreadsheetSkillText, /openchestnut-pivot-table-workflow\.mjs/);
assert.match(spreadsheetSkillText, /openchestnut-financial-returns-workflow\.mjs/);
assert.match(spreadsheetSkillText, /openchestnut-loan-amortization-workflow\.mjs/);
assert.match(spreadsheetSkillText, /openchestnut-asset-depreciation-workflow\.mjs/);
assert.match(spreadsheetSkillText, /openchestnut-growth-assumption-edit-workflow\.mjs/);
assert.ok(await exists(path.join(skillsRoot, "spreadsheets", "skills", "spreadsheets", "examples", "openchestnut-growth-assumption-edit-workflow.mjs")));

const presentationApiRoot = path.join(skillsRoot, "presentations", "skills", "presentations", "artifact_tool", "api");
const presentationApiDocs = await fs.readFile(path.join(presentationApiRoot, "API_DOCS.md"), "utf8");
const presentationSpec = await fs.readFile(path.join(presentationApiRoot, "references", "presentation.spec.md"), "utf8");
const presentationLayoutSpec = await fs.readFile(path.join(presentationApiRoot, "references", "layout.spec.md"), "utf8");
assert.match(presentationApiDocs, /presentation\.view/);
assert.match(presentationSpec, /showGridlines\(\).*showGuides\(\)/s);
assert.match(presentationSpec, /gridSpacingCxEmu.*gridSpacingCyEmu/s);
assert.match(presentationLayoutSpec, /read-only `slideGuides`/);
const presentationSkillText = await fs.readFile(path.join(skillsRoot, "presentations", "skills", "presentations", "SKILL.md"), "utf8");
assert.match(presentationSkillText, /openchestnut-title-notes-edit-workflow\.mjs/);
assert.ok(await exists(path.join(skillsRoot, "presentations", "skills", "presentations", "examples", "openchestnut-title-notes-edit-workflow.mjs")));
assert.match(presentationSkillText, /openchestnut-slide-name-edit-workflow\.mjs/);
assert.ok(await exists(path.join(skillsRoot, "presentations", "skills", "presentations", "examples", "openchestnut-slide-name-edit-workflow.mjs")));

const documentsSkillRoot = path.join(skillsRoot, "documents", "skills", "documents");
const documentsManifest = (await fs.readFile(path.join(documentsSkillRoot, "manifest.txt"), "utf8"))
  .split(/\r?\n/)
  .map((entry) => entry.trim())
  .filter(Boolean);
assert.equal(new Set(documentsManifest).size, documentsManifest.length, "Documents manifest must not contain duplicates");
for (const entry of documentsManifest) {
  assert.equal(path.isAbsolute(entry), false, `Documents manifest entry must be relative: ${entry}`);
  assert.ok(!entry.split("/").includes(".."), `Documents manifest entry must stay inside the Skill: ${entry}`);
  assert.ok(await exists(path.join(documentsSkillRoot, entry)), `Documents manifest entry is missing: ${entry}`);
}
assert.ok(documentsManifest.includes("artifact_tool/API_QUICK_START.md"));
assert.ok(documentsManifest.includes("examples/openchestnut-end-to-end.mjs"));
assert.ok(documentsManifest.includes("examples/openchestnut-classic-comment-edit-workflow.mjs"));
assert.ok(!documentsManifest.includes("examples/end_to_end_smoke_test.md"));

const pdfSkillRoot = path.join(skillsRoot, "pdf", "skills", "pdf");
const pdfSkillText = await fs.readFile(path.join(pdfSkillRoot, "SKILL.md"), "utf8");
assert.match(pdfSkillText, /open-office-artifact-tool/);
assert.match(pdfSkillText, /PdfArtifact/);
assert.match(pdfSkillText, /createPdfjsParser/);
assert.match(pdfSkillText, /Poppler/);
assert.match(pdfSkillText, /ReportLab/);
assert.match(pdfSkillText, /pdfplumber/);
assert.match(pdfSkillText, /pypdf/);
assert.match(pdfSkillText, /PyMuPDF/);
assert.match(pdfSkillText, /pyHanko/);
assert.match(pdfSkillText, /veraPDF/);
assert.match(pdfSkillText, /rewrite/);
assert.match(pdfSkillText, /incremental/);
assert.match(pdfSkillText, /sanitize/);
assert.match(pdfSkillText, /silent fallback/i);
assert.match(pdfSkillText, /original bytes/i);
assert.ok(await exists(path.join(pdfSkillRoot, "artifact_tool", "API_QUICK_START.md")));
assert.ok(await exists(path.join(pdfSkillRoot, "examples", "public-api-end-to-end.mjs")));
assert.ok(await exists(path.join(pdfSkillRoot, "examples", "accessible-board-report.mjs")));
for (const relativePath of [
  "manifest.txt",
  "references/PROVIDER_MATRIX.md",
  "references/SAVE_POLICIES.md",
  "references/SECURITY_CHECKLIST.md",
  "references/PRODUCT_BOUNDARIES.md",
  "scripts/pdf_provider.py",
  "scripts/mupdf.mjs",
  "scripts/reportlab_create.py",
  "scripts/pdfplumber_extract.py",
  "scripts/pypdf_edit.py",
  "scripts/pymupdf_edit.py",
  "scripts/python_runtime.py",
  "scripts/residue_scan.py",
  "tasks/create.md",
  "tasks/read_review.md",
  "tasks/edit_existing.md",
  "tasks/forms_annotations.md",
  "tasks/sign_verify.md",
  "tasks/redact.md",
  "tasks/accessibility.md",
  "tasks/render_review.md",
]) assert.ok(await exists(path.join(pdfSkillRoot, relativePath)), `PDF Skill is missing ${relativePath}`);

const spreadsheetApp = JSON.parse(await fs.readFile(path.join(skillsRoot, "spreadsheets", ".app.json"), "utf8"));
assert.equal(
  spreadsheetApp.apps.connected_documents.id,
  "connector_office-artifact-tool_codex_document_control",
  "Excel live control must retain the host connector contract",
);

for (const file of (await walk(skillsRoot)).filter((item) => /\.(?:md|mjs|js|json|ya?ml|py)$/i.test(item))) {
  const source = await fs.readFile(file, "utf8");
  assert.doesNotMatch(source, /from\s+["']office-artifact-tool["']/, `${path.relative(repoRoot, file)} still imports the private package`);
}

const officialValidator = path.join(os.homedir(), ".codex", "skills", ".system", "plugin-creator", "scripts", "validate_plugin.py");
if (await exists(officialValidator)) {
  for (const pluginName of pluginNames.filter((name) => name !== "default-template-library")) {
    const validation = spawnSync("python3", [officialValidator, path.join(skillsRoot, pluginName)], { encoding: "utf8" });
    assert.equal(validation.status, 0, `${pluginName} failed the official plugin validator\n${validation.stdout}\n${validation.stderr}`);
  }
}

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-reference-skills-"));
const previousPackageDir = process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR;
try {
  process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR = repoRoot;

  const { createDocument, DEFAULT_BRIEF } = await import(
    "../skills/documents/skills/documents/examples/openchestnut-end-to-end.mjs"
  );
  const docxPath = path.join(tempRoot, "openchestnut-decision-brief.docx");
  const authoredDocument = await createDocument(docxPath);
  assert.equal(authoredDocument.verification.ok, true);
  assert.match(authoredDocument.inspection.ndjson, /Launch readiness decision brief/);
  const documentRoundTrip = await DocumentFile.importDocx(await FileBlob.load(docxPath));
  assert.equal(documentRoundTrip.blocks.find((block) => block.kind === "table")?.getCell(1, 1).value, "Verified");
  assert.equal(documentRoundTrip.blocks.filter((block) => block.kind === "listItem").length, 3);
  assert.equal(documentRoundTrip.comments[0]?.text, "Recommendation wording verified for the release record.");
  assert.equal(documentRoundTrip.bookmarks[0]?.name, "DecisionSection");
  assert.deepEqual(documentRoundTrip.contentControls.map((control) => [control.tag, control.alias, control.text]), [
    ["OWNER", "Brief owner", DEFAULT_BRIEF.owner],
  ]);
  assert.deepEqual(documentRoundTrip.notes.map((note) => [note.kind, note.text]), [
    ["footnote", "The final gate includes native rendering, package validation, and semantic re-import."],
    ["endnote", "Evidence snapshot dated 2026-07-17; retained with the release record."],
  ]);
  assert.equal(documentRoundTrip.blocks.some(
    (block) => block.kind === "hyperlink" && block.anchor === "DecisionSection",
  ), true);
  assert.equal(documentRoundTrip.headers[0]?.text, "LAUNCH READINESS | DECISION BRIEF");
  assert.equal(documentRoundTrip.footers[0]?.fieldInstruction, "PAGE");
  assert.deepEqual(documentRoundTrip.blocks.filter((block) => block.kind === "change").map(
    (block) => [block.changeType, block.text, block.author],
  ), [
    ["insert", "Final application-compatibility review is required before rollout.", "Lead reviewer"],
    ["delete", "Immediate unrestricted rollout.", "Release reviewer"],
  ]);
  const documentPackage = await JSZip.loadAsync(await fs.readFile(docxPath));
  const documentXml = await documentPackage.file("word/document.xml").async("text");
  assert.match(documentXml, /<w:ins\b/);
  assert.match(documentXml, /<w:del\b/);
  assert.match(documentXml, /<w:delText\b/);
  assert.match(documentXml, /<w:bookmarkStart\b[^>]*w:name="DecisionSection"/);
  assert.match(documentXml, /<w:bookmarkEnd\b/);
  assert.match(documentXml, /<w:hyperlink\b[^>]*w:anchor="DecisionSection"/);
  assert.match(documentXml, /<w:footnoteReference\b[^>]*w:id="1"/);
  assert.match(documentXml, /<w:endnoteReference\b[^>]*w:id="1"/);
  assert.match(documentXml, /<w:sdt>/);
  assert.match(documentXml, /<w:tag w:val="OWNER"\s*\/>/);
  assert.match(documentXml, /<w:t>Artifact Platform<\/w:t>/);
  const footnotesXml = await documentPackage.file("word/footnotes.xml").async("text");
  const endnotesXml = await documentPackage.file("word/endnotes.xml").async("text");
  for (const id of ["-1", "0", "1"]) assert.match(footnotesXml, new RegExp(`<w:footnote\\b[^>]*w:id="${id}"`));
  for (const id of ["-1", "0", "1"]) assert.match(endnotesXml, new RegExp(`<w:endnote\\b[^>]*w:id="${id}"`));
  assert.match(footnotesXml, /semantic re-import/);
  assert.match(endnotesXml, /retained with the release record/);

  const { ensureArtifactToolWorkspace, importArtifactTool } = await import(
    "../skills/presentations/skills/presentations/container_tools/artifact_tool_utils.mjs"
  );
  const workspace = path.join(tempRoot, "presentation-workspace");
  const prepared = await ensureArtifactToolWorkspace(workspace);
  assert.equal(prepared.packageDir, repoRoot);
  assert.equal(
    await fs.realpath(path.join(workspace, "node_modules", "open-office-artifact-tool")),
    await fs.realpath(repoRoot),
  );
  const importedPackage = await importArtifactTool(workspace);
  assert.equal(importedPackage.PresentationFile, PresentationFile);

  const layoutRoot = path.join(
    skillsRoot,
    "presentations",
    "skills",
    "presentations",
    "assets",
    "builtin_templates",
    "grid-layout-library",
  );
  const { buildPresentation, exportPresentation } = await import(
    "../skills/presentations/skills/presentations/builtin_templates_support/scripts/create-presentation.mjs"
  );
  const authoredPresentation = await buildPresentation(layoutRoot);
  assert.equal(authoredPresentation.slides.items.length, 26);
  const pptxPath = path.join(tempRoot, "reference-grid-layout-library.pptx");
  await exportPresentation(layoutRoot, pptxPath);
  const presentation = await PresentationFile.importPptx(await FileBlob.load(pptxPath));
  assert.equal(presentation.slides.items.length, 26);
  const allSlides = presentation.slides.items;
  const allShapes = allSlides.flatMap((slide) => slide.shapes.items);
  assert.equal(allShapes.filter((shape) => shape.geometry === "custom").length, 11);
  assert.equal(allSlides.flatMap((slide) => slide.images.items).length, 1);
  assert.equal(allSlides.flatMap((slide) => slide.connectors.items).length, 2);
  const firstText = allShapes.find((shape) => String(shape.text?.value || "").includes("Your presentation"));
  assert.ok(firstText, "the reference template headline must survive OpenChestnut export/import");
  assert.equal(firstText.text.value, "Your presentation \nheadline goes here");
  assert.ok(Math.abs(firstText.position.left - 41.33) < 0.02);
  assert.ok(Math.abs(firstText.position.top - 182.55) < 0.02);
  assert.equal(firstText.text.paragraphs[0].runs[0].style.fontSize, 80);
  assert.equal(firstText.text.paragraphs[0].runs[0].style.fontFamily, "Helvetica Neue");
  const pptxZip = await JSZip.loadAsync(await fs.readFile(pptxPath));
  const slideXml = await Promise.all(
    Object.keys(pptxZip.files)
      .filter((name) => /^ppt\/slides\/slide\d+\.xml$/.test(name))
      .map((name) => pptxZip.file(name).async("text")),
  );
  assert.equal(slideXml.reduce((count, xml) => count + (xml.match(/<a:custGeom\b/g)?.length || 0), 0), 11);

  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add("Summary");
  sheet.getRange("A1:C4").values = [
    ["Month", "Revenue", "EBITDA"],
    ["Jan", 100, 10],
    ["Feb", 120, 18],
    ["Mar", 130, 22],
  ];
  sheet.getRange("D1").values = [["Margin"]];
  sheet.getRange("D2").formulas = [["=C2/B2"]];
  sheet.getRange("D2:D4").fillDown();
  sheet.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("D2:D4").format.numberFormat = "0.0%";
  sheet.getRange("F1:G1").values = [["Month", "Revenue"]];
  sheet.getRange("F2:G2").formulas = [["=A2", "=B2"]];
  sheet.getRange("F2:G4").fillDown();
  const chart = sheet.charts.add("line", sheet.getRange("F1:G4"));
  chart.title = "Revenue Trend";
  chart.hasLegend = false;
  chart.setPosition("I1", "P15");
  const preview = await workbook.render({ sheetName: "Summary", autoCrop: "all", format: "svg" });
  assert.equal(preview.type, "image/svg+xml");
  assert.match(await preview.text(), /Revenue Trend/);
  const xlsx = await SpreadsheetFile.exportXlsx(workbook);
  const workbookRoundTrip = await SpreadsheetFile.importXlsx(xlsx);
  assert.deepEqual(workbookRoundTrip.worksheets.getItem("Summary").getRange("D2:D4").formulas, [
    ["=C2/B2"],
    ["=C3/B3"],
    ["=C4/B4"],
  ]);
  const csvWorkbook = await Workbook.fromCSV("Name,Value\nOpenChestnut,1", { sheetName: "Data" });
  assert.deepEqual(csvWorkbook.worksheets.getItem("Data").getRange("A1:B2").values, [["Name", "Value"], ["OpenChestnut", "1"]]);

  const { createWorkbook: createReferenceWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-range-workflow.mjs"
  );
  const spreadsheetPath = path.join(tempRoot, "openchestnut-range-workflow.xlsx");
  const authoredWorkbook = await createReferenceWorkbook(spreadsheetPath);
  assert.equal(authoredWorkbook.verification.ok, true);
  assert.match(authoredWorkbook.inspection.ndjson, /Revenue trend/);
  const spreadsheetRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(spreadsheetPath));
  assert.equal(spreadsheetRoundTrip.worksheets.getItem("Forecast").getRange("D3").format.numberFormat, "0.00%");
  assert.equal(spreadsheetRoundTrip.worksheets.getItem("Forecast").getRange("B3").formulasR1C1[0][0], "=R[-1]C*(1+'Assumptions'!R2C2)");

  const { createSparklineWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-sparkline-workflow.mjs"
  );
  const sparklinePath = path.join(tempRoot, "openchestnut-sparkline-workflow.xlsx");
  const authoredSparklines = await createSparklineWorkbook(sparklinePath);
  assert.equal(authoredSparklines.verification.ok, true);
  assert.match(authoredSparklines.inspection.ndjson, /"kind":"sparkline"/);
  const sparklineRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(sparklinePath));
  assert.deepEqual(sparklineRoundTrip.worksheets.getItem("Operating Trends").sparklineGroups.items.map((group) => group.type), ["line", "column"]);
  assert.equal(sparklineRoundTrip.worksheets.getItem("Operating Trends").sparklineGroups.items[0].seriesColor, "#F97316");

  const { createDataTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-data-table-workflow.mjs"
  );
  const dataTablePath = path.join(tempRoot, "openchestnut-data-table-workflow.xlsx");
  const authoredDataTables = await createDataTableWorkbook(dataTablePath);
  assert.equal(authoredDataTables.verification.ok, true);
  assert.match(authoredDataTables.inspection.ndjson, /"kind":"dataTable"/);
  const dataTableRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(dataTablePath));
  assert.deepEqual(
    dataTableRoundTrip.worksheets.getItem("Scenario Analysis").dataTables.__getDefinitions().map((item) => item.displayFormula),
    ["{=TABLE(D1)}", "{=TABLE(D1,D2)}"],
  );

  const { createPivotTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-pivot-table-workflow.mjs"
  );
  const pivotTablePath = path.join(tempRoot, "openchestnut-pivot-table-workflow.xlsx");
  const authoredPivotTable = await createPivotTableWorkbook(pivotTablePath);
  assert.equal(authoredPivotTable.verification.ok, true);
  assert.match(authoredPivotTable.inspection.ndjson, /"kind":"pivotTable"/);
  const pivotTableRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(pivotTablePath));
  assert.equal(pivotTableRoundTrip.worksheets.getItem("Pivot Summary").pivotTables.items[0].name, "Revenue by region");
  assert.deepEqual(pivotTableRoundTrip.worksheets.getItem("Pivot Summary").pivotTables.items[0].computedValues().at(-1), ["Grand Total", 380, 240, 620]);

  const { createFinancialReturnsWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-financial-returns-workflow.mjs"
  );
  const financialReturnsPath = path.join(tempRoot, "openchestnut-financial-returns-workflow.xlsx");
  const authoredFinancialReturns = await createFinancialReturnsWorkbook(financialReturnsPath);
  assert.equal(authoredFinancialReturns.verification.ok, true);
  assert.match(authoredFinancialReturns.inspection.ndjson, /XIRR/);
  assert.match(authoredFinancialReturns.inspection.ndjson, /MIRR/);
  const financialReturnsRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(financialReturnsPath));
  financialReturnsRoundTrip.recalculate();
  assert.equal(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B8").formulas[0][0], "=XIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18,'Inputs'!$B$7)");
  assert.equal(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B9").formulas[0][0], "=MIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$5,'Inputs'!$B$6)");
  assert.ok(Math.abs(financialReturnsRoundTrip.worksheets.getItem("Returns").getRange("B9").values[0][0] - 0.14400168352963139) < 1e-9);
  assert.deepEqual(financialReturnsRoundTrip.worksheets.getItem("Checks").getRange("E4:E10").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const { createLoanAmortizationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-loan-amortization-workflow.mjs"
  );
  const loanAmortizationPath = path.join(tempRoot, "openchestnut-loan-amortization-workflow.xlsx");
  const authoredLoanAmortization = await createLoanAmortizationWorkbook(loanAmortizationPath);
  assert.equal(authoredLoanAmortization.verification.ok, true);
  assert.match(authoredLoanAmortization.inspection.ndjson, /PPMT/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /PV/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /FV/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /NPER/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /CUMIPMT/);
  assert.match(authoredLoanAmortization.checksInspection.ndjson, /CUMPRINC/);
  const loanAmortizationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(loanAmortizationPath));
  loanAmortizationRoundTrip.recalculate();
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Amortization").getRange("D5").formulas[0][0], "=IPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Amortization").getRange("F16").values[0][0]) < 1e-7);
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B9").formulas[0][0], "=RATE('Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9,'Inputs'!$B$10)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B9").values[0][0] - 0.01) < 1e-10);
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B10").formulas[0][0], "=PV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B11").formulas[0][0], "=FV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B12").formulas[0][0], "=NPER('Inputs'!$B$10,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B13").formulas[0][0], "=CUMIPMT('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.equal(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B14").formulas[0][0], "=CUMPRINC('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B10").values[0][0] - 100000) < 1e-7);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B11").values[0][0]) < 1e-7);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B12").values[0][0] - 12) < 1e-10);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B13").values[0][0] + 6618.54641401005) < 1e-8);
  assert.ok(Math.abs(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("B14").values[0][0] + 100000) < 1e-8);
  assert.deepEqual(loanAmortizationRoundTrip.worksheets.getItem("Checks").getRange("E4:E15").values, Array.from({ length: 12 }, () => ["OK"]));

  const { createAssetDepreciationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-asset-depreciation-workflow.mjs"
  );
  const assetDepreciationPath = path.join(tempRoot, "openchestnut-asset-depreciation-workflow.xlsx");
  const authoredAssetDepreciation = await createAssetDepreciationWorkbook(assetDepreciationPath);
  assert.equal(authoredAssetDepreciation.verification.ok, true);
  assert.match(authoredAssetDepreciation.inspection.ndjson, /SLN/);
  assert.match(authoredAssetDepreciation.inspection.ndjson, /DDB/);
  const assetDepreciationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(assetDepreciationPath));
  assetDepreciationRoundTrip.recalculate();
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("D5").formulas[0][0], "=DB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$8)");
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("E5").formulas[0][0], "=DDB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$9)");
  assert.equal(assetDepreciationRoundTrip.worksheets.getItem("Depreciation").getRange("F9").values[0][0], 10000);
  assert.deepEqual(assetDepreciationRoundTrip.worksheets.getItem("Checks").getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);

  const { createScatterWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-scatter-chart-workflow.mjs"
  );
  const scatterPath = path.join(tempRoot, "openchestnut-scatter-chart-workflow.xlsx");
  const authoredScatter = await createScatterWorkbook(scatterPath);
  assert.equal(authoredScatter.verification.ok, true);
  assert.match(authoredScatter.inspection.ndjson, /"chartType":"scatter"/);
  const scatterRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(scatterPath));
  const scatterChart = scatterRoundTrip.worksheets.getItem("Relationship Analysis").charts.items[0];
  assert.equal(scatterChart.type, "scatter");
  assert.deepEqual(scatterChart.series.items[0].xValues, [10, 20, 25, 34, 45]);

  const { createBubbleWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-bubble-chart-workflow.mjs"
  );
  const bubblePath = path.join(tempRoot, "openchestnut-bubble-chart-workflow.xlsx");
  const authoredBubble = await createBubbleWorkbook(bubblePath);
  assert.equal(authoredBubble.verification.ok, true);
  assert.match(authoredBubble.inspection.ndjson, /"chartType":"bubble"/);
  const bubbleRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(bubblePath));
  const bubbleChart = bubbleRoundTrip.worksheets.getItem("Opportunity Analysis").charts.items[0];
  assert.equal(bubbleChart.type, "bubble");
  assert.deepEqual(bubbleChart.series.items[0].bubbleSizes, [4, 10, 12, 18, 27]);

  const { createPdf } = await import(
    "../skills/pdf/skills/pdf/examples/public-api-end-to-end.mjs"
  );
  const pdfPath = path.join(tempRoot, "release-readiness-scorecard.pdf");
  const pdfRenderDir = path.join(tempRoot, "release-readiness-scorecard-pages");
  const authoredPdf = await createPdf(pdfPath, { renderDir: pdfRenderDir });
  assert.equal(authoredPdf.verification.ok, true);
  assert.equal(authoredPdf.fileInspection.summary.tagged, true);
  assert.equal(authoredPdf.fileInspection.summary.figures, 2);
  assert.equal(authoredPdf.renderedPages.length, authoredPdf.pdf.pages.length);
  assert.ok(authoredPdf.renderedPages.every((page) => page.bytes > 1_000));
  const pdfRoundTrip = await PdfFile.importPdf(await FileBlob.load(pdfPath));
  assert.equal(pdfRoundTrip.pages[0].tables[0].getCell(3, 2).value, "Verified");
  assert.match(pdfRoundTrip.extractText(), /Release readiness scorecard/);
  const { createPdfjsParser } = await import("open-office-artifact-tool/pdf/pdfjs");
  const parsedPdf = await PdfFile.importPdf(await FileBlob.load(pdfPath), {
    parser: createPdfjsParser(),
    preferParser: true,
    parserName: "pdfjs",
  });
  assert.match(parsedPdf.extractText(), /Release readiness scorecard/);
} finally {
  if (previousPackageDir === undefined) delete process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR;
  else process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR = previousPackageDir;
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("reference skill plugins smoke ok");
