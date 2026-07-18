import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

const libraryRoot = path.resolve(import.meta.dirname, "..");
const integrity = JSON.parse(await fs.readFile(path.join(libraryRoot, "integrity.json"), "utf8"));

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function referenceFor(templateId) {
  const record = integrity.assets.find((asset) => asset.templateId === templateId && asset.role === "reference");
  if (!record) throw new RangeError(`Unknown retained template: ${templateId}`);
  return record;
}

async function assertNewFile(target, label) {
  try {
    await fs.lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}. Choose a distinct output path instead of overwriting it.`);
}

function normalizedOutputPath(value, reference) {
  const target = path.resolve(requiredText(value, "outputPath"));
  if (path.extname(target).toLowerCase() !== path.extname(reference.path).toLowerCase()) {
    throw new Error(`Output path must end in ${path.extname(reference.path)}: ${target}`);
  }
  return target;
}

export async function materializeTemplate({ templateId, outputPath, auditPath } = {}) {
  const id = requiredText(templateId, "templateId");
  const reference = referenceFor(id);
  const output = normalizedOutputPath(outputPath, reference);
  const audit = path.resolve(auditPath ? requiredText(auditPath, "auditPath") : `${output}.audit.json`);
  if (audit === output) throw new Error("auditPath must be distinct from outputPath.");
  await Promise.all([assertNewFile(output, "Template output"), assertNewFile(audit, "Template audit")]);
  await Promise.all([fs.mkdir(path.dirname(output), { recursive: true }), fs.mkdir(path.dirname(audit), { recursive: true })]);

  const source = path.join(libraryRoot, reference.path);
  const sourceBytes = await fs.readFile(source);
  if (sourceBytes.length !== reference.bytes || sha256(sourceBytes) !== reference.sha256) {
    throw new Error(`Retained source integrity check failed for ${id}.`);
  }

  const temporaryOutput = `${output}.tmp-${process.pid}-${crypto.randomUUID()}`;
  const temporaryAudit = `${audit}.tmp-${process.pid}-${crypto.randomUUID()}`;
  let outputPromoted = false;
  try {
    await fs.copyFile(source, temporaryOutput, fs.constants.COPYFILE_EXCL);
    const outputBytes = await fs.readFile(temporaryOutput);
    if (!outputBytes.equals(sourceBytes)) throw new Error(`Materialized output differs from retained source for ${id}.`);
    const report = {
      schema: "open-office-artifact-tool.default-template-library.v1",
      status: "succeeded",
      operation: "materialize-retained-reference",
      template: { id, sourceCommit: integrity.source.commit, license: integrity.source.license },
      source: { path: reference.path, bytes: sourceBytes.length, sha256: reference.sha256 },
      output: { path: output, bytes: outputBytes.length, sha256: sha256(outputBytes) },
      savePolicy: { strategy: "create-new", overwrite: false },
      validation: { byteIdenticalToRetainedReference: true },
    };
    await fs.writeFile(temporaryAudit, `${JSON.stringify(report, null, 2)}\n`);
    await fs.rename(temporaryOutput, output);
    outputPromoted = true;
    await fs.rename(temporaryAudit, audit);
    return { outputPath: output, auditPath: audit, audit: report };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryOutput, { force: true }),
      fs.rm(temporaryAudit, { force: true }),
      outputPromoted ? fs.rm(output, { force: true }) : Promise.resolve(),
    ]);
    throw error;
  }
}

function usage() {
  return [
    "Usage:",
    "  node materialize-template.mjs --template-id <id> --output <path> [--audit <path>]",
    "",
    "Creates a distinct byte-identical working copy of one retained repository template.",
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
  return { templateId: result["template-id"], outputPath: result.output, auditPath: result.audit };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  try {
    const request = parseCli(process.argv.slice(2));
    if (request.help) console.log(usage());
    else {
      const result = await materializeTemplate(request);
      console.log(JSON.stringify({ templateId: result.audit.template.id, outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
