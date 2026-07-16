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
const pluginNames = ["documents", "spreadsheets", "presentations", "pdf"];
const expectedSkills = new Map([
  ["documents", ["documents"]],
  ["spreadsheets", ["excel-live-control", "spreadsheets"]],
  ["presentations", ["presentations"]],
  ["pdf", ["pdf"]],
]);
const expectedDeclaredSkillNames = new Map([
  ["documents", "documents"],
  ["excel-live-control", "excel-live-control"],
  ["spreadsheets", "Spreadsheets"],
  ["presentations", "Presentations"],
  ["pdf", "pdf"],
]);

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
  assert.equal(manifest.license, "MIT");
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
    const agentText = await fs.readFile(path.join(skillRoot, "agents", "openai.yaml"), "utf8");
    for (const iconKey of ["icon_small", "icon_large"]) {
      const icon = yamlValue(agentText, iconKey);
      assert.ok(icon, `${pluginName}/${skillName} is missing ${iconKey}`);
      assert.ok(await exists(path.resolve(skillRoot, icon)), `${pluginName}/${skillName} ${iconKey} does not resolve`);
    }
  }
}

assert.equal(await exists(path.join(skillsRoot, "documents", "SKILL.md")), false);
assert.equal(await exists(path.join(skillsRoot, "spreadsheets", "scripts")), false);
assert.equal(await exists(path.join(skillsRoot, "presentations", "fixtures")), false);
assert.ok(await exists(path.join(repoRoot, "test", "skill-harness", "spreadsheets", "scripts", "workflow.mjs")));

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
  for (const pluginName of pluginNames) {
    const validation = spawnSync("python3", [officialValidator, path.join(skillsRoot, pluginName)], { encoding: "utf8" });
    assert.equal(validation.status, 0, `${pluginName} failed the official plugin validator\n${validation.stdout}\n${validation.stderr}`);
  }
}

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-reference-skills-"));
const previousPackageDir = process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR;
try {
  process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR = repoRoot;

  const { createDocument } = await import(
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
  assert.equal(documentRoundTrip.headers[0]?.text, "LAUNCH READINESS | DECISION BRIEF");
  assert.equal(documentRoundTrip.footers[0]?.fieldInstruction, "PAGE");

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
    "codex-grid-layout-library",
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
