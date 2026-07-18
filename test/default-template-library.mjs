import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  DocumentFile,
  PresentationFile,
  SpreadsheetFile,
} from "../src/index.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const pluginRoot = path.join(repoRoot, "skills", "default-template-library");
const generatorPath = path.join(pluginRoot, "scripts", "generate-template.mjs");
const catalog = JSON.parse(await fs.readFile(path.join(pluginRoot, "catalog.json"), "utf8"));
const manifest = JSON.parse(await fs.readFile(path.join(pluginRoot, "manifest.json"), "utf8"));
const plugin = JSON.parse(await fs.readFile(path.join(pluginRoot, ".codex-plugin", "plugin.json"), "utf8"));

const EXPECTED_READY = [
  ["artifact-template-design-report", "document", ".docx", "Design Report"],
  ["artifact-template-strategy-memorandum", "document", ".docx", "Strategy Memorandum"],
  ["artifact-template-operating-review", "presentation", ".pptx", "Operating Review"],
  ["artifact-template-project-kickoff", "presentation", ".pptx", "Project Kickoff"],
  ["artifact-template-financial-budget", "workbook", ".xlsx", "Financial Budget"],
  ["artifact-template-project-tracker", "workbook", ".xlsx", "Project Tracker"],
];
const NATIVE_PDF_PAGE_CONTRACTS = new Map([
  ["artifact-template-design-report", { pages: 3, reason: "three deliberate decision-record pages" }],
  ["artifact-template-strategy-memorandum", { pages: 2, reason: "one decision frame and one action page" }],
  ["artifact-template-operating-review", { pages: 3, reason: "one native page per slide" }],
  ["artifact-template-project-kickoff", { pages: 3, reason: "one native page per slide" }],
  ["artifact-template-financial-budget", { pages: 3, reason: "one native page per worksheet" }],
  ["artifact-template-project-tracker", { pages: 3, reason: "one native page per worksheet" }],
]);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
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

function runGenerator(args, options = {}) {
  return spawnSync(process.execPath, [generatorPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    ...options,
  });
}

function commandAvailable(command) {
  const probe = spawnSync(command, ["--version"], { encoding: "utf8" });
  return !probe.error && probe.status === 0;
}

async function assertNativePdfPreview(sourcePath, outputDirectory) {
  const converted = spawnSync("soffice", ["--headless", "--convert-to", "pdf", "--outdir", outputDirectory, sourcePath], {
    encoding: "utf8",
  });
  assert.equal(converted.status, 0, `LibreOffice could not render ${sourcePath}\nSTDOUT:\n${converted.stdout}\nSTDERR:\n${converted.stderr}`);
  const stem = path.basename(sourcePath, path.extname(sourcePath));
  const pdfPath = path.join(outputDirectory, `${stem}.pdf`);
  const info = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8" });
  assert.equal(info.status, 0, `Poppler could not inspect ${pdfPath}\nSTDOUT:\n${info.stdout}\nSTDERR:\n${info.stderr}`);
  const pages = Number(info.stdout.match(/^Pages:\s*([1-9]\d*)/m)?.[1]);
  assert.ok(Number.isInteger(pages), `native rendering must produce at least one page for ${sourcePath}`);
  const previewPrefix = path.join(outputDirectory, `${stem}-preview`);
  const preview = spawnSync("pdftoppm", ["-png", "-f", "1", "-l", String(pages), pdfPath, previewPrefix], { encoding: "utf8" });
  assert.equal(preview.status, 0, `Poppler could not rasterize ${pdfPath}\nSTDOUT:\n${preview.stdout}\nSTDERR:\n${preview.stderr}`);
  for (let page = 1; page <= pages; page += 1) {
    const previewPath = `${previewPrefix}-${page}.png`;
    const stat = await fs.stat(previewPath);
    assert.ok(stat.size > 0, `native preview is empty: ${previewPath}`);
  }
  return pages;
}

function resultFor(generatedById, id) {
  const result = generatedById.get(id);
  assert.ok(result, `missing generated result for ${id}`);
  return result;
}

assert.equal(plugin.name, "default-template-library");
assert.equal(plugin.version, "0.2.0");
assert.equal(plugin.license, "AGPL-3.0-or-later");
assert.equal(plugin.skills, "./skills/");
assert.match(plugin.repository, /open-office-artifact-tool/);
assert.equal(manifest.schemaVersion, 1);
assert.equal(manifest.catalog, "catalog.json");
assert.deepEqual(manifest.skills, [
  "skills/office-template-catalog",
  ...EXPECTED_READY.map(([id]) => `skills/${id}`),
]);
assert.equal(catalog.schemaVersion, 1);
assert.equal(catalog.templates.length, 20);
assert.deepEqual(catalog.provenancePolicy, {
  assetMode: "project-authored-source-free",
  retainedReferenceFiles: false,
  retainedPreviewFiles: false,
  thirdPartyTemplateAssets: "excluded-until-explicit-redistribution-authorization",
});

const ready = catalog.templates.filter((template) => template.status === "ready");
const planned = catalog.templates.filter((template) => template.status === "planned");
assert.equal(ready.length, 6);
assert.equal(planned.length, 14);
assert.deepEqual(ready.map((template) => [
  template.id,
  template.artifactKind,
  template.implementation.extension,
  template.displayName,
]), EXPECTED_READY);
for (const template of catalog.templates) {
  assert.match(template.id, /^artifact-template-[a-z0-9-]+$/);
  assert.ok(["document", "presentation", "workbook"].includes(template.artifactKind));
  assert.ok(["ready", "planned"].includes(template.status));
  if (template.status === "ready") {
    assert.equal(template.implementation.assetMode, "project-authored-source-free");
    assert.equal(template.implementation.generator, "scripts/generate-template.mjs");
    assert.equal(template.implementation.skill, `skills/${template.id}`);
  } else {
    assert.equal(template.implementation, undefined);
  }
}

const skillDirectories = (await fs.readdir(path.join(pluginRoot, "skills"), { withFileTypes: true }))
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();
assert.deepEqual(skillDirectories, [
  ...EXPECTED_READY.map(([id]) => id),
  "office-template-catalog",
].sort());
for (const skillName of skillDirectories) {
  const skillRoot = path.join(pluginRoot, "skills", skillName);
  const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
  const frontmatter = skillText.match(/^---\n([\s\S]*?)\n---/);
  assert.ok(frontmatter, `${skillName} must have frontmatter`);
  assert.equal(yamlValue(frontmatter[1], "name"), skillName);
  const agent = await fs.readFile(path.join(skillRoot, "agents", "openai.yaml"), "utf8");
  for (const key of ["icon_small", "icon_large"]) {
    const icon = yamlValue(agent, key);
    assert.ok(icon, `${skillName} must declare ${key}`);
    await fs.access(path.resolve(skillRoot, icon));
  }
}

for (const readyTemplate of ready) {
  const skillRoot = path.join(pluginRoot, readyTemplate.implementation.skill);
  const sidecar = JSON.parse(await fs.readFile(path.join(skillRoot, "template.json"), "utf8"));
  assert.equal(sidecar.templateId, readyTemplate.id);
  assert.equal(sidecar.artifactKind, readyTemplate.artifactKind);
  assert.equal(sidecar.generation.mode, "project-authored-source-free");
  assert.equal(sidecar.retainedReference, null);
  assert.equal(sidecar.retainedPreview, null);
}

const pluginFiles = await walk(pluginRoot);
for (const file of pluginFiles) {
  const relative = path.relative(pluginRoot, file).replaceAll(path.sep, "/");
  assert.doesNotMatch(relative, /(^|\/)(?:reference\.(?:docx|pptx|xlsx)|preview\.png)$/i, `template library must not retain a reference or preview binary: ${relative}`);
  assert.doesNotMatch(relative, /\.(?:docx|pptx|xlsx|png)$/i, `template library must remain source-only: ${relative}`);
  if (/\.(?:md|mjs|json|ya?ml|svg)$/i.test(relative)) {
    const source = await fs.readFile(file, "utf8");
    assert.doesNotMatch(source, /openai-templates|\.codex\/plugins\/cache/i, `template library must not point at a proprietary cache: ${relative}`);
  }
}

const help = runGenerator(["--help"]);
assert.equal(help.status, 0, help.stderr);
for (const [id] of EXPECTED_READY) assert.match(help.stdout, new RegExp(id));
const unavailable = runGenerator([
  "--template-id", "artifact-template-experiment-analysis",
  "--output", path.join(os.tmpdir(), "experiment-analysis.docx"),
]);
assert.notEqual(unavailable.status, 0);
assert.match(unavailable.stderr, /Unknown or unavailable source-free template/);

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-artifact-tool-default-template-library-"));
try {
  const generatedById = new Map();
  for (const template of ready) {
    const output = path.join(temporary, `${template.id}${template.implementation.extension}`);
    const audit = path.join(temporary, `${template.id}.audit.json`);
    const run = runGenerator([
      "--template-id", template.id,
      "--output", output,
      "--audit", audit,
    ]);
    assert.equal(run.status, 0, `${template.id} generator failed\nSTDOUT:\n${run.stdout}\nSTDERR:\n${run.stderr}`);
    const response = JSON.parse(run.stdout);
    const bytes = await fs.readFile(output);
    const report = JSON.parse(await fs.readFile(audit, "utf8"));
    assert.equal(response.templateId, template.id);
    assert.equal(response.artifactKind, template.artifactKind);
    assert.equal(report.template.id, template.id);
    assert.equal(report.template.provenance, "project-authored-source-free");
    assert.equal(report.template.retainedReference, false);
    assert.equal(report.source, null);
    assert.equal(report.provider.actual, "open-chestnut");
    assert.equal(report.provider.silentFallback, false);
    assert.equal(report.savePolicy.strategy, "create-new");
    assert.equal(report.output.sha256, sha256(bytes));
    assert.equal(report.validation.verify, true);
    assert.equal(report.validation.secondImport, true);
    assert.equal(report.validation.modelRender.type, "image/svg+xml");
    assert.ok(report.validation.modelRender.bytes > 0);
    generatedById.set(template.id, { template, output, audit, bytes });
  }

  const editedOutputs = [];
  for (const [id, , extension, title] of EXPECTED_READY.filter(([, kind]) => kind === "document")) {
    const result = resultFor(generatedById, id);
    const document = await DocumentFile.importDocx(result.bytes);
    const titleBlock = document.blocks.find((block) => block.text === title);
    assert.ok(titleBlock, `${title} title block is required`);
    document.resolve(`${titleBlock.id}/text`).text = `${title} — reviewed`;
    const editedBlob = await DocumentFile.exportDocx(document);
    const output = path.join(temporary, `reviewed-${id}${extension}`);
    await fs.writeFile(output, editedBlob.bytes);
    const reimported = await DocumentFile.importDocx(editedBlob);
    assert.equal(reimported.blocks.some((block) => block.text === `${title} — reviewed`), true);
    assert.equal(reimported.verify({ visualQa: true }).ok, true);
    editedOutputs.push({ templateId: id, output });
  }

  for (const [id, , extension, title] of EXPECTED_READY.filter(([, kind]) => kind === "presentation")) {
    const result = resultFor(generatedById, id);
    const presentation = await PresentationFile.importPptx(result.bytes);
    const titleShape = presentation.slides.getItem(0).shapes.items.find((shape) => shape.text.value === title);
    assert.ok(titleShape, `${title} title shape is required`);
    titleShape.text.set(`${title} — reviewed`);
    const editedBlob = await PresentationFile.exportPptx(presentation);
    const output = path.join(temporary, `reviewed-${id}${extension}`);
    await fs.writeFile(output, editedBlob.bytes);
    const reimported = await PresentationFile.importPptx(editedBlob);
    assert.equal(reimported.slides.getItem(0).shapes.items.some((shape) => shape.text.value === `${title} — reviewed`), true);
    assert.equal(reimported.verify({ visualQa: true }).ok, true);
    editedOutputs.push({ templateId: id, output });
  }

  const budgetResult = resultFor(generatedById, "artifact-template-financial-budget");
  const budget = await SpreadsheetFile.importXlsx(budgetResult.bytes);
  budget.worksheets.getItem("Assumptions").getRange("B5").values = [[125000]];
  budget.recalculate();
  const editedBudgetBlob = await SpreadsheetFile.exportXlsx(budget, { recalculate: false });
  const editedBudgetOutput = path.join(temporary, "reviewed-financial-budget.xlsx");
  await fs.writeFile(editedBudgetOutput, editedBudgetBlob.bytes);
  const editedBudget = await SpreadsheetFile.importXlsx(editedBudgetBlob);
  editedBudget.recalculate();
  assert.equal(editedBudget.worksheets.getItem("Assumptions").getRange("B5").values[0][0], 125000);
  assert.deepEqual(editedBudget.worksheets.getItem("Budget Summary").getRange("D4:D7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);
  assert.equal(editedBudget.verify({ visualQa: true }).ok, true);
  editedOutputs.push({ templateId: "artifact-template-financial-budget", output: editedBudgetOutput });

  const trackerResult = resultFor(generatedById, "artifact-template-project-tracker");
  const tracker = await SpreadsheetFile.importXlsx(trackerResult.bytes);
  tracker.worksheets.getItem("Work Plan").getRange("D9").values = [["Done"]];
  tracker.recalculate();
  const editedTrackerBlob = await SpreadsheetFile.exportXlsx(tracker, { recalculate: false });
  const editedTrackerOutput = path.join(temporary, "reviewed-project-tracker.xlsx");
  await fs.writeFile(editedTrackerOutput, editedTrackerBlob.bytes);
  const editedTracker = await SpreadsheetFile.importXlsx(editedTrackerBlob);
  editedTracker.recalculate();
  assert.equal(editedTracker.worksheets.getItem("Work Plan").getRange("D9").values[0][0], "Done");
  assert.equal(editedTracker.worksheets.getItem("Project Summary").getRange("B4").values[0][0], 3);
  assert.deepEqual(editedTracker.worksheets.getItem("Project Summary").getRange("D4:D8").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  assert.equal(editedTracker.verify({ visualQa: true }).ok, true);
  editedOutputs.push({ templateId: "artifact-template-project-tracker", output: editedTrackerOutput });

  if (commandAvailable("soffice") && commandAvailable("pdfinfo") && commandAvailable("pdftoppm")) {
    const rendered = path.join(temporary, "native-render");
    await fs.mkdir(rendered);
    for (const { templateId, output } of editedOutputs) {
      const pages = await assertNativePdfPreview(output, rendered);
      const expected = NATIVE_PDF_PAGE_CONTRACTS.get(templateId);
      if (expected) assert.equal(pages, expected.pages, `${templateId} must render ${expected.pages} native PDF page(s): ${expected.reason}.`);
    }
  }

  const firstReady = resultFor(generatedById, EXPECTED_READY[0][0]);
  const repeated = runGenerator([
    "--template-id", firstReady.template.id,
    "--output", firstReady.output,
    "--audit", firstReady.audit,
  ]);
  assert.notEqual(repeated.status, 0);
  assert.match(repeated.stderr, /already exists/);
} finally {
  await fs.rm(temporary, { force: true, recursive: true });
}

console.log("default template library smoke ok");
