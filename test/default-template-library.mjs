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
  const result = spawnSync(process.execPath, [generatorPath, ...args], {
    cwd: repoRoot,
    encoding: "utf8",
    ...options,
  });
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
  "skills/artifact-template-strategy-memorandum",
  "skills/artifact-template-project-kickoff",
  "skills/artifact-template-financial-budget",
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
assert.equal(ready.length, 3);
assert.equal(planned.length, 17);
assert.deepEqual(ready.map((template) => [template.id, template.artifactKind, template.implementation.extension]), [
  ["artifact-template-strategy-memorandum", "document", ".docx"],
  ["artifact-template-project-kickoff", "presentation", ".pptx"],
  ["artifact-template-financial-budget", "workbook", ".xlsx"],
]);
for (const template of catalog.templates) {
  assert.match(template.id, /^artifact-template-[a-z0-9-]+$/);
  assert.ok(["document", "presentation", "workbook"].includes(template.artifactKind));
  assert.ok(["ready", "planned"].includes(template.status));
  if (template.status === "ready") {
    assert.equal(template.implementation.assetMode, "project-authored-source-free");
    assert.ok(template.implementation.generator);
  } else {
    assert.equal(template.implementation, undefined);
  }
}

const skillDirectories = (await fs.readdir(path.join(pluginRoot, "skills"), { withFileTypes: true }))
  .filter((entry) => entry.isDirectory())
  .map((entry) => entry.name)
  .sort();
assert.deepEqual(skillDirectories, [
  "artifact-template-financial-budget",
  "artifact-template-project-kickoff",
  "artifact-template-strategy-memorandum",
  "office-template-catalog",
]);
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
assert.match(help.stdout, /artifact-template-strategy-memorandum/);
const unavailable = runGenerator([
  "--template-id", "artifact-template-design-report",
  "--output", path.join(os.tmpdir(), "design-report.docx"),
]);
assert.notEqual(unavailable.status, 0);
assert.match(unavailable.stderr, /Unknown or unavailable source-free template/);

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-artifact-tool-default-template-library-"));
try {
  const generated = [];
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
    generated.push({ template, output, audit, bytes });
  }

  const [documentResult, presentationResult, workbookResult] = generated;
  const document = await DocumentFile.importDocx(documentResult.bytes);
  const documentTitle = document.blocks.find((block) => block.text === "Strategy Memorandum");
  assert.ok(documentTitle);
  document.resolve(`${documentTitle.id}/text`).text = "Strategy Memorandum — reviewed";
  const editedDocument = await DocumentFile.importDocx(await DocumentFile.exportDocx(document));
  assert.equal(editedDocument.blocks.some((block) => block.text === "Strategy Memorandum — reviewed"), true);
  assert.equal(editedDocument.verify({ visualQa: true }).ok, true);

  const presentation = await PresentationFile.importPptx(presentationResult.bytes);
  assert.deepEqual(presentation.slides.items.map((slide) => slide.name), ["Kickoff overview", "Scope and plan", "Owners and decisions"]);
  const kickoffTitle = presentation.slides.getItem(0).shapes.items.find((shape) => shape.name === "kickoff-title");
  assert.ok(kickoffTitle);
  kickoffTitle.text.set("Project Kickoff — reviewed");
  const editedPresentation = await PresentationFile.importPptx(await PresentationFile.exportPptx(presentation));
  assert.equal(editedPresentation.slides.getItem(0).shapes.items.some((shape) => shape.text.value === "Project Kickoff — reviewed"), true);
  assert.equal(editedPresentation.verify({ visualQa: true }).ok, true);

  const workbook = await SpreadsheetFile.importXlsx(workbookResult.bytes);
  workbook.worksheets.getItem("Assumptions").getRange("B5").values = [[125000]];
  workbook.recalculate();
  const editedWorkbook = await SpreadsheetFile.importXlsx(await SpreadsheetFile.exportXlsx(workbook, { recalculate: false }));
  editedWorkbook.recalculate();
  assert.equal(editedWorkbook.worksheets.getItem("Assumptions").getRange("B5").values[0][0], 125000);
  assert.deepEqual(editedWorkbook.worksheets.getItem("Budget Summary").getRange("D4:D7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);
  assert.equal(editedWorkbook.verify({ visualQa: true }).ok, true);

  const repeated = runGenerator([
    "--template-id", documentResult.template.id,
    "--output", documentResult.output,
    "--audit", documentResult.audit,
  ]);
  assert.notEqual(repeated.status, 0);
  assert.match(repeated.stderr, /already exists/);
} finally {
  await fs.rm(temporary, { force: true, recursive: true });
}

console.log("default template library smoke ok");
