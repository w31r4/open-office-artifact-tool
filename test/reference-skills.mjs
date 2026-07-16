import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import {
  FileBlob,
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
} finally {
  if (previousPackageDir === undefined) delete process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR;
  else process.env.OPEN_OFFICE_ARTIFACT_TOOL_PACKAGE_DIR = previousPackageDir;
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("reference skill plugins smoke ok");
