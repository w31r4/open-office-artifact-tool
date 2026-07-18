import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  DocumentFile,
  PresentationFile,
  SpreadsheetFile,
} from "open-office-artifact-tool";

import { TEMPLATE_DEFINITIONS } from "../templates/definitions.mjs";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function assertVerified(report, label) {
  if (!report?.ok) throw new Error(`${label} verification failed: ${report?.ndjson || JSON.stringify(report?.issues || [])}`);
}

async function assertPreview(preview, title, label) {
  assert.equal(preview.type, "image/svg+xml");
  assert.equal((await preview.text()).includes(title), true, `${label} preview must retain its title.`);
}

async function exportDocumentTemplate(definition, document) {
  const title = definition.validation.title;
  assertVerified(document.verify({ visualQa: true }), `${title} model`);
  const first = await DocumentFile.exportDocx(document);
  const imported = await DocumentFile.importDocx(first);
  assert.equal(imported.blocks.some((block) => block.text === title), true, "DOCX import must retain the template title.");
  assertVerified(imported.verify({ visualQa: true }), `${title} first import`);
  const final = await DocumentFile.exportDocx(imported);
  const reimported = await DocumentFile.importDocx(final);
  assert.equal(reimported.blocks.some((block) => block.text === title), true, "DOCX second import must retain the template title.");
  const preview = await reimported.render({ format: "svg" });
  await assertPreview(preview, title, title);
  return { file: final, preview, validation: { verify: true, secondImport: true } };
}

async function exportPresentationTemplate(definition, presentation) {
  const { title, slideNames } = definition.validation;
  assert.ok(slideNames?.length, `${title} must declare expected slide names.`);
  assertVerified(presentation.verify({ visualQa: true }), `${title} model`);
  const first = await PresentationFile.exportPptx(presentation);
  const imported = await PresentationFile.importPptx(first);
  assert.deepEqual(imported.slides.items.map((slide) => slide.name), slideNames, "PPTX import must retain all template slides.");
  assert.equal(imported.slides.getItem(0).shapes.items.some((shape) => shape.text.value === title), true, "PPTX import must retain the template title.");
  assertVerified(imported.verify({ visualQa: true }), `${title} first import`);
  const final = await PresentationFile.exportPptx(imported);
  const reimported = await PresentationFile.importPptx(final);
  assert.deepEqual(reimported.slides.items.map((slide) => slide.name), slideNames, "PPTX second import must retain all template slides.");
  const preview = await reimported.export({ format: "montage" });
  await assertPreview(preview, title, title);
  return { file: final, preview, validation: { verify: true, secondImport: true, slides: slideNames.length } };
}

async function exportWorkbookTemplate(definition, workbook) {
  const { title, sheetNames, previewSheet, previewRange } = definition.validation;
  assert.ok(sheetNames?.length && previewSheet && previewRange, `${title} must declare workbook validation metadata.`);
  assertVerified(workbook.verify({ visualQa: true }), `${title} model`);
  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.deepEqual(imported.worksheets.items.map((sheet) => sheet.name), sheetNames, "XLSX import must retain all template sheets.");
  assertVerified(imported.verify({ visualQa: true }), `${title} first import`);
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const reimported = await SpreadsheetFile.importXlsx(final);
  reimported.recalculate();
  assert.deepEqual(reimported.worksheets.items.map((sheet) => sheet.name), sheetNames, "XLSX second import must retain all template sheets.");
  const preview = await reimported.render({ sheetName: previewSheet, range: previewRange, autoCrop: "all", format: "svg" });
  await assertPreview(preview, title, title);
  return { file: final, preview, validation: { verify: true, secondImport: true, sheets: sheetNames.length } };
}

async function exportSourceFreeTemplate(definition, artifact) {
  if (definition.artifactKind === "document") return exportDocumentTemplate(definition, artifact);
  if (definition.artifactKind === "presentation") return exportPresentationTemplate(definition, artifact);
  if (definition.artifactKind === "workbook") return exportWorkbookTemplate(definition, artifact);
  throw new Error(`Unsupported template artifact kind: ${definition.artifactKind}`);
}

async function assertNewFile(target, label) {
  try {
    await fs.access(target);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}. Choose a new output path instead of overwriting it.`);
}

function normalizedOutputPath(value, definition) {
  const outputPath = path.resolve(requiredText(value, "outputPath"));
  if (path.extname(outputPath).toLowerCase() !== definition.extension) {
    throw new Error(`Output path must end in ${definition.extension}: ${outputPath}`);
  }
  return outputPath;
}

export async function generateTemplate({ templateId, outputPath, auditPath } = {}) {
  const id = requiredText(templateId, "templateId");
  const definition = TEMPLATE_DEFINITIONS[id];
  if (!definition) throw new RangeError(`Unknown or unavailable source-free template: ${id}`);
  const finalPath = normalizedOutputPath(outputPath, definition);
  const finalAuditPath = path.resolve(auditPath ? requiredText(auditPath, "auditPath") : `${finalPath}.audit.json`);
  if (finalAuditPath === finalPath) throw new Error("auditPath must be distinct from outputPath.");
  await Promise.all([
    assertNewFile(finalPath, "Template output"),
    assertNewFile(finalAuditPath, "Template audit"),
  ]);
  await Promise.all([
    fs.mkdir(path.dirname(finalPath), { recursive: true }),
    fs.mkdir(path.dirname(finalAuditPath), { recursive: true }),
  ]);

  const artifact = definition.build();
  const generated = await exportSourceFreeTemplate(definition, artifact);
  const temporaryOutput = `${finalPath}.tmp-${process.pid}-${crypto.randomUUID()}`;
  const temporaryAudit = `${finalAuditPath}.tmp-${process.pid}-${crypto.randomUUID()}`;
  let outputPromoted = false;
  try {
    await generated.file.save(temporaryOutput);
    const bytes = await fs.readFile(temporaryOutput);
    const audit = {
      schema: "open-office-artifact-tool.template-library.v1",
      status: "succeeded",
      template: {
        id,
        artifactKind: definition.artifactKind,
        provenance: "project-authored-source-free",
        retainedReference: false,
        retainedPreview: false,
      },
      source: null,
      output: { path: finalPath, bytes: bytes.length, sha256: sha256(bytes) },
      provider: { actual: "open-chestnut", silentFallback: false },
      savePolicy: { strategy: "create-new" },
      validation: {
        ...generated.validation,
        modelRender: { type: generated.preview.type, bytes: generated.preview.bytes.length },
      },
    };
    await fs.writeFile(temporaryAudit, `${JSON.stringify(audit, null, 2)}\n`);
    await fs.rename(temporaryOutput, finalPath);
    outputPromoted = true;
    await fs.rename(temporaryAudit, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryOutput, { force: true }),
      fs.rm(temporaryAudit, { force: true }),
      outputPromoted ? fs.rm(finalPath, { force: true }) : Promise.resolve(),
    ]);
    throw error;
  }
}

function usage() {
  return [
    "Usage:",
    "  node generate-template.mjs --template-id <id> --output <path> [--audit <path>]",
    "",
    "Ready template IDs:",
    ...Object.entries(TEMPLATE_DEFINITIONS).map(([id, definition]) => `  ${id} (${definition.extension})`),
  ].join("\n");
}

function parseCli(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--help" || token === "-h") return { help: true };
    if (!token.startsWith("--")) throw new Error(`Unexpected argument: ${token}`);
    const key = token.slice(2);
    if (!new Set(["template-id", "output", "audit"]).has(key)) throw new Error(`Unknown option: ${token}`);
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) throw new Error(`${token} requires a value.`);
    result[key] = value;
    index += 1;
  }
  return {
    templateId: result["template-id"],
    outputPath: result.output,
    auditPath: result.audit,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  try {
    const request = parseCli(process.argv.slice(2));
    if (request.help) {
      console.log(usage());
    } else {
      const result = await generateTemplate(request);
      console.log(JSON.stringify({
        templateId: result.audit.template.id,
        artifactKind: result.audit.template.artifactKind,
        outputPath: result.outputPath,
        auditPath: result.auditPath,
        outputSha256: result.audit.output.sha256,
      }));
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
