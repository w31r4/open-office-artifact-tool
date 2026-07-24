#!/usr/bin/env node

import { createHash } from "node:crypto";
import { createReadStream } from "node:fs";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";

const SCRIPT_DIRECTORY = path.dirname(fileURLToPath(import.meta.url));
const SKILL_DIRECTORY = path.resolve(SCRIPT_DIRECTORY, "..");
const SIDECAR_NAME = "artifact-template.json";
const TEMPLATE_NAME_PATTERN = /^artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*$/u;
const HASH_PATTERN = /^[a-f0-9]{64}$/u;
const MAX_SIDECAR_BYTES = 128 * 1024;
const MAX_SKILL_BYTES = 256 * 1024;
const DEFAULT_MAX_CANDIDATES = 5;
const MAX_CANDIDATES = 20;
const VALID_KINDS = new Set(["document", "presentation", "spreadsheet"]);
const REFERENCE_EXTENSIONS = new Map([
  ["document", ".docx"],
  ["presentation", ".pptx"],
  ["spreadsheet", ".xlsx"],
]);
const VALID_DENSITIES = new Set(["sparse", "medium", "dense", "mixed"]);
const VALID_COLOR_MODES = new Set(["light", "dark", "neutral", "mixed"]);
const VALID_COMMITMENTS = new Set(["neutral", "opinionated"]);
const VALID_EDIT_LEVELS = new Set(["copy-only", "bounded-edit", "composable"]);
const USAGE = [
  "Usage: query-templates.mjs --kind <document|spreadsheet|presentation>",
  "  [--tag <tag>]... [--id <artifact-template-id>]",
  "  [--root <absolute-template-root>]... [--max <1-20>]",
].join("\n");

export async function queryTemplates({
  kind,
  tags = [],
  id = null,
  roots = null,
  maxCandidates = DEFAULT_MAX_CANDIDATES,
} = {}) {
  assertKind(kind);
  assertTemplateId(id, "--id", true);
  const normalizedTags = normalizeQueryTags(tags);
  if (!Number.isSafeInteger(maxCandidates) || maxCandidates < 1 || maxCandidates > MAX_CANDIDATES) {
    throw new Error(`maxCandidates must be an integer from 1 to ${MAX_CANDIDATES}.`);
  }

  const rootEntries = await resolveRoots(roots);
  const candidates = [];
  const invalid = [];
  const seenTemplatePaths = new Set();
  const claimedTemplateIds = new Set();

  for (const rootEntry of rootEntries) {
    for (const entry of await fs.readdir(rootEntry.path, { withFileTypes: true })) {
      if (!entry.isDirectory() || !TEMPLATE_NAME_PATTERN.test(entry.name)) continue;
      if (id != null && entry.name !== id) continue;
      if (claimedTemplateIds.has(entry.name)) continue;
      claimedTemplateIds.add(entry.name);
      const templatePath = path.join(rootEntry.path, entry.name);
      try {
        const canonicalTemplatePath = await fs.realpath(templatePath);
        if (seenTemplatePaths.has(canonicalTemplatePath)) continue;
        seenTemplatePaths.add(canonicalTemplatePath);
        const candidate = await readTemplate({
          expectedId: entry.name,
          root: rootEntry,
          templatePath,
        });
        if (candidate.kind !== kind) continue;
        if (id != null && candidate.id !== id) continue;
        candidate.matchedTags = matchedTags(candidate, normalizedTags);
        candidates.push(candidate);
      } catch (error) {
        invalid.push({
          id: entry.name,
          root: rootEntry.path,
          error: error instanceof Error ? error.message : String(error),
        });
      }
    }
  }

  invalid.sort((left, right) => left.id.localeCompare(right.id) || left.root.localeCompare(right.root));
  if (id != null) {
    const invalidRequested = invalid.find((entry) => entry.id === id);
    if (invalidRequested != null) {
      throw new Error(`Requested template ${id} is invalid: ${invalidRequested.error}`);
    }
    if (candidates.length === 0) {
      throw new Error(`Requested template ${id} was not found for kind ${kind}.`);
    }
  }

  candidates.sort((left, right) =>
    right.matchedTags.length - left.matchedTags.length ||
    left.id.localeCompare(right.id) ||
    left.templateRoot.localeCompare(right.templateRoot)
  );

  return {
    schemaVersion: 1,
    kind,
    requestedId: id,
    queryTags: normalizedTags,
    searchedRoots: rootEntries.map((entry) => ({
      path: entry.path,
      source: entry.source,
    })),
    candidates: candidates.slice(0, id == null ? maxCandidates : 1),
    invalid,
    selectionMade: false,
  };
}

async function resolveRoots(explicitRoots) {
  if (explicitRoots != null && !Array.isArray(explicitRoots)) {
    throw new Error("Template roots must be an array of paths.");
  }
  if (explicitRoots != null && explicitRoots.length > 20) {
    throw new Error("At most 20 template roots may be queried.");
  }
  const requested = explicitRoots == null || explicitRoots.length === 0
    ? defaultRoots()
    : explicitRoots.map((root) => ({ path: root, source: "explicit" }));
  const resolved = [];
  const seen = new Set();

  for (const entry of requested) {
    if (typeof entry.path !== "string" || entry.path.trim().length === 0) {
      throw new Error("Template roots must be non-empty paths.");
    }
    const absolutePath = path.resolve(entry.path);
    const stat = await fs.lstat(absolutePath).catch((error) => {
      if (error?.code === "ENOENT" && entry.source !== "explicit") return null;
      throw error;
    });
    if (stat == null) continue;
    if (!stat.isDirectory() || stat.isSymbolicLink()) {
      throw new Error(`Template root must be a real directory: ${absolutePath}`);
    }
    const canonicalPath = await fs.realpath(absolutePath);
    if (seen.has(canonicalPath)) continue;
    seen.add(canonicalPath);
    resolved.push({ path: canonicalPath, source: entry.source });
  }
  return resolved;
}

function defaultRoots() {
  const configured = (process.env.OFFICEKIT_TEMPLATE_ROOTS ?? "")
    .split(path.delimiter)
    .filter(Boolean)
    .map((root) => ({ path: root, source: "configured" }));
  const officeArtifactHome = process.env.OFFICE_ARTIFACT_HOME == null
    ? path.join(os.homedir(), ".office-artifact-tool")
    : path.resolve(process.env.OFFICE_ARTIFACT_HOME);
  return [
    ...configured,
    { path: path.join(officeArtifactHome, "skills"), source: "local-user" },
    { path: path.resolve(SKILL_DIRECTORY, ".."), source: "flat-installed" },
    {
      path: path.resolve(SKILL_DIRECTORY, "../../../default-template-library/skills"),
      source: "repository-default",
    },
  ];
}

async function readTemplate({ expectedId, root, templatePath }) {
  const templateStat = await fs.lstat(templatePath);
  if (!templateStat.isDirectory() || templateStat.isSymbolicLink()) {
    throw new Error("template root must be a real directory");
  }
  const sidecarPath = path.join(templatePath, SIDECAR_NAME);
  const sidecarStat = await fs.lstat(sidecarPath);
  if (!sidecarStat.isFile() || sidecarStat.isSymbolicLink()) {
    throw new Error(`${SIDECAR_NAME} must be a regular file`);
  }
  if (sidecarStat.size > MAX_SIDECAR_BYTES) {
    throw new Error(`${SIDECAR_NAME} exceeds the ${MAX_SIDECAR_BYTES}-byte budget`);
  }

  let metadata;
  try {
    metadata = JSON.parse(await fs.readFile(sidecarPath, "utf8"));
  } catch (error) {
    throw new Error(`${SIDECAR_NAME} is not valid JSON: ${error.message}`);
  }
  validateMetadata(metadata, expectedId);

  const skillPath = await resolveTemplateSkill(templatePath);
  const referencePath = await resolveAsset(
    templatePath,
    metadata.reference,
    metadata.provenance.referenceSha256,
    "reference",
  );
  const previewPath = await resolveAsset(
    templatePath,
    metadata.preview,
    metadata.provenance.previewSha256,
    "preview",
  );

  return {
    id: metadata.id,
    displayName: metadata.displayName,
    kind: metadata.kind,
    useWhen: metadata.useWhen,
    avoidWhen: metadata.avoidWhen,
    audiences: metadata.audiences,
    contentShapes: metadata.contentShapes,
    visualTraits: metadata.visualTraits,
    visualCommitment: metadata.visualCommitment,
    editProfile: metadata.editProfile,
    provenance: {
      license: metadata.provenance.license,
      source: metadata.provenance.source,
      referenceSha256: metadata.provenance.referenceSha256,
      previewSha256: metadata.provenance.previewSha256,
    },
    catalogSource: root.source,
    templateRoot: await fs.realpath(templatePath),
    skillPath,
    referencePath,
    previewPath,
  };
}

function validateMetadata(value, expectedId) {
  if (value == null || typeof value !== "object" || Array.isArray(value)) {
    throw new Error("metadata must be an object");
  }
  assertObjectKeys(
    value,
    "metadata",
    [
      "schemaVersion",
      "id",
      "displayName",
      "kind",
      "reference",
      "preview",
      "useWhen",
      "avoidWhen",
      "audiences",
      "contentShapes",
      "visualTraits",
      "visualCommitment",
      "editProfile",
      "provenance",
    ],
  );
  if (value.schemaVersion !== 2) {
    throw new Error("schemaVersion must be 2");
  }
  assertTemplateId(value.id, "id");
  if (value.id !== expectedId) {
    throw new Error(`id must match directory name ${expectedId}`);
  }
  assertShortString(value.displayName, "displayName", 80);
  assertKind(value.kind);
  assertStringArray(value.useWhen, "useWhen", { min: 1, max: 20 });
  assertStringArray(value.avoidWhen, "avoidWhen", { min: 0, max: 20 });
  assertStringArray(value.audiences, "audiences", { min: 0, max: 20 });
  assertStringArray(value.contentShapes, "contentShapes", { min: 0, max: 20 });
  assertRelativeAssetPath(value.reference, "reference");
  assertRelativeAssetPath(value.preview, "preview");
  if (path.posix.extname(value.reference).toLowerCase() !== REFERENCE_EXTENSIONS.get(value.kind)) {
    throw new Error(`${value.kind} templates must use a ${REFERENCE_EXTENSIONS.get(value.kind)} reference`);
  }
  if (path.posix.extname(value.preview).toLowerCase() !== ".png") {
    throw new Error("preview must use a .png file");
  }

  if (value.visualTraits == null || typeof value.visualTraits !== "object" || Array.isArray(value.visualTraits)) {
    throw new Error("visualTraits must be an object");
  }
  assertObjectKeys(
    value.visualTraits,
    "visualTraits",
    ["tone", "density", "colorMode", "structure"],
  );
  assertStringArray(value.visualTraits.tone, "visualTraits.tone", { min: 0, max: 12 });
  assertEnum(value.visualTraits.density, "visualTraits.density", VALID_DENSITIES);
  assertEnum(value.visualTraits.colorMode, "visualTraits.colorMode", VALID_COLOR_MODES);
  assertStringArray(value.visualTraits.structure, "visualTraits.structure", { min: 0, max: 12 });
  assertEnum(value.visualCommitment, "visualCommitment", VALID_COMMITMENTS);

  if (value.editProfile == null || typeof value.editProfile !== "object" || Array.isArray(value.editProfile)) {
    throw new Error("editProfile must be an object");
  }
  assertObjectKeys(
    value.editProfile,
    "editProfile",
    ["level", "verifiedOperations"],
  );
  assertEnum(value.editProfile.level, "editProfile.level", VALID_EDIT_LEVELS);
  assertStringArray(value.editProfile.verifiedOperations, "editProfile.verifiedOperations", {
    min: 0,
    max: 24,
  });
  if (value.editProfile.level === "copy-only" && value.editProfile.verifiedOperations.length !== 0) {
    throw new Error("copy-only templates cannot declare verifiedOperations");
  }

  if (value.provenance == null || typeof value.provenance !== "object" || Array.isArray(value.provenance)) {
    throw new Error("provenance must be an object");
  }
  assertObjectKeys(
    value.provenance,
    "provenance",
    ["license", "source", "referenceSha256", "previewSha256"],
  );
  assertShortString(value.provenance.license, "provenance.license", 120);
  assertShortString(value.provenance.source, "provenance.source", 500);
  assertHash(value.provenance.referenceSha256, "provenance.referenceSha256");
  assertHash(value.provenance.previewSha256, "provenance.previewSha256");
}

async function resolveAsset(templatePath, relativePath, expectedHash, label) {
  const resolved = path.resolve(templatePath, relativePath);
  const canonicalTemplatePath = await fs.realpath(templatePath);
  const stat = await fs.lstat(resolved);
  if (!stat.isFile() || stat.isSymbolicLink()) {
    throw new Error(`${label} must be a regular nonsymlink file`);
  }
  const canonicalAssetPath = await fs.realpath(resolved);
  if (!isInside(canonicalTemplatePath, canonicalAssetPath)) {
    throw new Error(`${label} escapes the template directory`);
  }
  const actualHash = await sha256File(canonicalAssetPath);
  if (actualHash !== expectedHash) {
    throw new Error(`${label} SHA-256 mismatch`);
  }
  return canonicalAssetPath;
}

async function resolveTemplateSkill(templatePath) {
  const candidate = path.join(templatePath, "SKILL.md");
  const stat = await fs.lstat(candidate);
  if (!stat.isFile() || stat.isSymbolicLink()) {
    throw new Error("SKILL.md must be a regular nonsymlink file");
  }
  if (stat.size > MAX_SKILL_BYTES) {
    throw new Error(`SKILL.md exceeds the ${MAX_SKILL_BYTES}-byte budget`);
  }
  const [canonicalTemplatePath, canonicalSkillPath] = await Promise.all([
    fs.realpath(templatePath),
    fs.realpath(candidate),
  ]);
  if (!isInside(canonicalTemplatePath, canonicalSkillPath)) {
    throw new Error("SKILL.md escapes the template directory");
  }
  return canonicalSkillPath;
}

function matchedTags(candidate, tags) {
  if (tags.length === 0) return [];
  const searchable = [
    candidate.id,
    candidate.displayName,
    ...candidate.useWhen,
    ...candidate.audiences,
    ...candidate.contentShapes,
    ...candidate.visualTraits.tone,
    candidate.visualTraits.density,
    candidate.visualTraits.colorMode,
    ...candidate.visualTraits.structure,
  ].map(normalizeTag);
  return tags.filter((tag) => searchable.some((value) => value === tag || value.includes(tag)));
}

function normalizeQueryTags(tags) {
  if (!Array.isArray(tags)) throw new Error("tags must be an array.");
  if (tags.length > 20) throw new Error("At most 20 query tags may be used.");
  const normalized = tags.map((tag) => {
    if (typeof tag !== "string" || tag.trim().length === 0 || tag.length > 80) {
      throw new Error("Each query tag must be a non-empty string of at most 80 characters.");
    }
    return normalizeTag(tag);
  });
  return [...new Set(normalized)];
}

function normalizeTag(value) {
  return value
    .normalize("NFKC")
    .toLowerCase()
    .trim()
    .replace(/[_\s]+/gu, "-");
}

function assertTemplateId(value, label, optional = false) {
  if (optional && value == null) return;
  if (typeof value !== "string" || !TEMPLATE_NAME_PATTERN.test(value)) {
    throw new Error(`${label} must be an artifact-template-* identifier.`);
  }
}

function assertKind(value) {
  if (!VALID_KINDS.has(value)) {
    throw new Error("kind must be document, spreadsheet, or presentation.");
  }
}

function assertShortString(value, label, max) {
  if (
    typeof value !== "string" ||
    value.trim().length === 0 ||
    value !== value.trim() ||
    value.length > max ||
    /[\0\r\n]/u.test(value)
  ) {
    throw new Error(`${label} must be one trimmed line of at most ${max} characters`);
  }
}

function assertStringArray(value, label, { min, max }) {
  if (!Array.isArray(value) || value.length < min || value.length > max) {
    throw new Error(`${label} must contain ${min}-${max} strings`);
  }
  const seen = new Set();
  for (const entry of value) {
    assertShortString(entry, label, 120);
    const normalized = normalizeTag(entry);
    if (seen.has(normalized)) throw new Error(`${label} must not contain duplicates`);
    seen.add(normalized);
  }
}

function assertEnum(value, label, allowed) {
  if (!allowed.has(value)) {
    throw new Error(`${label} must be one of ${[...allowed].join(", ")}`);
  }
}

function assertObjectKeys(value, label, allowedKeys) {
  const allowed = new Set(allowedKeys);
  const extra = Object.keys(value).filter((key) => !allowed.has(key));
  if (extra.length > 0) {
    throw new Error(`${label} contains unsupported fields: ${extra.join(", ")}`);
  }
}

function assertHash(value, label) {
  if (typeof value !== "string" || !HASH_PATTERN.test(value)) {
    throw new Error(`${label} must be a lowercase SHA-256 value`);
  }
}

function assertRelativeAssetPath(value, label) {
  if (
    typeof value !== "string" ||
    value.length === 0 ||
    path.isAbsolute(value) ||
    value.includes("\\") ||
    value.split("/").some((segment) => segment === "" || segment === "." || segment === "..")
  ) {
    throw new Error(`${label} must be a safe relative path`);
  }
}

function isInside(root, candidate) {
  const relative = path.relative(root, candidate);
  return relative.length > 0 && relative !== ".." && !relative.startsWith(`..${path.sep}`) && !path.isAbsolute(relative);
}

async function sha256File(filePath) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(filePath)) hash.update(chunk);
  return hash.digest("hex");
}

function parseArguments(args) {
  const request = { tags: [], roots: [] };
  for (let index = 0; index < args.length; index += 1) {
    const flag = args[index];
    if (flag === "--help" || flag === "-h") {
      return { help: true };
    }
    const value = args[index + 1];
    if (value == null || value.startsWith("--")) throw new Error(USAGE);
    index += 1;
    if (flag === "--kind") request.kind = value;
    else if (flag === "--tag") request.tags.push(value);
    else if (flag === "--id") request.id = value;
    else if (flag === "--root") request.roots.push(value);
    else if (flag === "--max") request.maxCandidates = Number(value);
    else throw new Error(USAGE);
  }
  if (request.kind == null) throw new Error(USAGE);
  return request;
}

async function main() {
  const request = parseArguments(process.argv.slice(2));
  if (request.help) {
    process.stdout.write(`${USAGE}\n`);
    return;
  }
  if (request.roots.length === 0) request.roots = null;
  const result = await queryTemplates(request);
  process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);
}

const invokedPath = process.argv[1] == null
  ? null
  : await fs.realpath(path.resolve(process.argv[1]));
const modulePath = await fs.realpath(fileURLToPath(import.meta.url));
if (invokedPath === modulePath) {
  main().catch((error) => {
    process.stderr.write(`${error instanceof Error ? error.message : String(error)}\n`);
    process.exitCode = 2;
  });
}
