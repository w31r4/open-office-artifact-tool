#!/usr/bin/env node

import { randomUUID } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import process from "node:process";
import { fileURLToPath } from "node:url";
import { crc32 } from "node:zlib";

const TEMPLATE_PREFIX = "artifact-template-";
const WRITE_LOCK_NAME = ".artifact-template-write-lock";
const LOCK_OWNER_FILENAME = "owner-pid";
const BACKUP_NAME_PATTERN = /^(artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*)\.backup-[0-9a-f-]+$/u;
const MAX_SKILL_NAME_LENGTH = 64;
const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const USAGE =
  "Usage: create-template-skill.mjs --reference-path <path> --preview-path <path> --display-name <name> --description <description> [--mode update --skill-name <name>]";
const artifactKinds = new Map([
  [
    ".docx",
    {
      kind: "document",
      workflow: "Documents",
      preservation:
        "Preserve page setup, sections, styles, lists, tables, headers, footers, and recurring page elements.",
    },
  ],
  [
    ".pptx",
    {
      kind: "presentation",
      workflow: "Presentations",
      preservation:
        "Preserve source slides, layouts, masters, typography, geometry, images, charts, tables, and recurring slide chrome.",
    },
  ],
  [
    ".xlsx",
    {
      kind: "spreadsheet",
      workflow: "Spreadsheets",
      preservation:
        "Preserve sheet structure, formulas, names, number formats, dimensions, tables, charts, validation, conditional formatting, and frozen panes.",
    },
  ],
]);

async function createTemplateSkill(
  rawRequest,
  officeArtifactHome = getDefaultOfficeArtifactHome(),
) {
  const request = await validateRequest(rawRequest);
  const artifact = artifactKinds.get(
    path.extname(request.referencePath).toLowerCase(),
  );
  const skillsRoot = path.join(officeArtifactHome, "skills");
  const releaseLock = await acquireWriteLock(
    path.join(officeArtifactHome, WRITE_LOCK_NAME),
  );
  try {
    await recoverInterruptedTemplateReplacements(skillsRoot);
    const identity =
      request.mode === "update"
        ? await getUpdateIdentity(skillsRoot, request, artifact.kind)
        : await getCreateIdentity(skillsRoot, request.displayName);
    const skillPath = path.join(skillsRoot, identity.skillName);
    const stagedSkill = await stageTemplateSkill({
      artifact,
      description: request.description,
      displayName: identity.displayName,
      parentDirectory: skillsRoot,
      previewPath: request.previewPath,
      referencePath: request.referencePath,
      skillName: identity.skillName,
      sourceSkillPath: request.mode === "update" ? skillPath : null,
    });
    try {
      if (request.mode === "update") {
        await replaceTemplateSkill(stagedSkill, skillPath);
      } else {
        await fs.rename(stagedSkill, skillPath);
      }
    } catch (error) {
      await fs.rm(stagedSkill, { force: true, recursive: true });
      throw error;
    }
    return {
      displayName: identity.displayName,
      kind: artifact.kind,
      skillName: identity.skillName,
      skillPath,
    };
  } finally {
    await releaseLock();
  }
}

async function validateRequest(rawRequest) {
  const mode = rawRequest.mode ?? "create";
  if (mode !== "create" && mode !== "update") {
    throw new Error("--mode must be 'create' or 'update'.");
  }
  const displayName = getRequiredString(
    rawRequest,
    "displayName",
    "--display-name",
  );
  const description = getRequiredString(
    rawRequest,
    "description",
    "--description",
  );
  assertSingleLine(displayName, "--display-name", 64);
  assertSingleLine(description, "--description", 600);
  const referencePath = path.resolve(
    getRequiredString(rawRequest, "referencePath", "--reference-path"),
  );
  const previewPath = path.resolve(
    getRequiredString(rawRequest, "previewPath", "--preview-path"),
  );
  const extension = path.extname(referencePath).toLowerCase();
  if (!artifactKinds.has(extension)) {
    throw new Error("--reference-path must end in .docx, .pptx, or .xlsx.");
  }
  await Promise.all([
    assertRegularFile(referencePath, "--reference-path"),
    assertRegularFile(previewPath, "--preview-path"),
  ]);
  if (path.extname(previewPath).toLowerCase() !== ".png") {
    throw new Error("--preview-path must end in .png.");
  }
  if (!hasValidPngStructure(await fs.readFile(previewPath))) {
    throw new Error("--preview-path must contain a valid PNG.");
  }

  const skillName = getOptionalString(
    rawRequest,
    "skillName",
    "--skill-name",
  );
  if (mode === "update") {
    if (skillName == null) {
      throw new Error("--skill-name is required for an explicit update.");
    }
    assertSkillName(skillName);
  } else if (skillName != null) {
    throw new Error(
      "--skill-name is only valid when --mode is 'update'.",
    );
  }

  return {
    description,
    displayName,
    mode,
    previewPath,
    referencePath,
    skillName,
  };
}

function getDefaultOfficeArtifactHome() {
  const homeDir = path.resolve(
    process.env.HOME ?? process.env.USERPROFILE ?? os.homedir(),
  );
  return path.resolve(
    process.env.OFFICE_ARTIFACT_HOME ?? path.join(homeDir, ".office-artifact-tool"),
  );
}

async function recoverInterruptedTemplateReplacements(skillsRoot) {
  await fs.mkdir(skillsRoot, { recursive: true });
  const backupsBySkillName = new Map();
  const entries = await fs.readdir(skillsRoot, { withFileTypes: true });
  for (const entry of entries) {
    if (!entry.isDirectory()) continue;
    const match = BACKUP_NAME_PATTERN.exec(entry.name);
    if (match == null) continue;
    const skillName = match[1];
    const backups = backupsBySkillName.get(skillName) ?? [];
    backups.push(path.join(skillsRoot, entry.name));
    backupsBySkillName.set(skillName, backups);
  }

  for (const [skillName, backups] of backupsBySkillName) {
    const skillPath = path.join(skillsRoot, skillName);
    if (await pathExists(skillPath)) {
      for (const backupPath of backups) {
        await fs.rm(backupPath, { force: true, recursive: true });
      }
      continue;
    }
    if (backups.length !== 1) {
      throw new Error(
        `Cannot recover ${skillName}: found ${backups.length} interrupted-update backups.`,
      );
    }
    await fs.rename(backups[0], skillPath);
  }
}

async function getCreateIdentity(skillsRoot, displayName) {
  const slug = getSlug(displayName);
  for (let index = 1; ; index += 1) {
    const suffix = index === 1 ? "" : `-${index}`;
    const baseLength = MAX_SKILL_NAME_LENGTH - suffix.length;
    const skillName =
      `${TEMPLATE_PREFIX}${slug}`.slice(0, baseLength).replace(/-+$/u, "") +
      suffix;
    if (!(await pathExists(path.join(skillsRoot, skillName)))) {
      return {
        displayName: index === 1 ? displayName : `${displayName} ${index}`,
        skillName,
      };
    }
  }
}

async function getUpdateIdentity(skillsRoot, request, kind) {
  const skillPath = path.join(skillsRoot, request.skillName);
  const sidecarPath = path.join(skillPath, "artifact-template.json");
  const sidecar = JSON.parse(await fs.readFile(sidecarPath, "utf8"));
  if (sidecar.schemaVersion !== 1 || sidecar.kind !== kind) {
    throw new Error(
      `${request.skillName} is not a version 1 ${kind} artifact template.`,
    );
  }
  return { displayName: request.displayName, skillName: request.skillName };
}

async function stageTemplateSkill({
  artifact,
  description,
  displayName,
  parentDirectory,
  previewPath,
  referencePath,
  skillName,
  sourceSkillPath,
}) {
  await fs.mkdir(parentDirectory, { recursive: true });
  const stagedSkill = await fs.mkdtemp(
    path.join(parentDirectory, `.${skillName}-stage-`),
  );
  const referenceFilename = `reference${path.extname(referencePath).toLowerCase()}`;
  try {
    if (sourceSkillPath != null) {
      await fs.cp(sourceSkillPath, stagedSkill, { recursive: true });
    }
    await Promise.all([
      fs.mkdir(path.join(stagedSkill, "agents"), { recursive: true }),
      fs.mkdir(path.join(stagedSkill, "assets"), { recursive: true }),
    ]);
    await Promise.all([
      fs.writeFile(
        path.join(stagedSkill, "SKILL.md"),
        getTemplateSkillMarkdown({
          artifact,
          description,
          displayName,
          skillName,
        }),
      ),
      fs.writeFile(
        path.join(stagedSkill, "agents", "agent.yaml"),
        getTemplateAgentYaml({ artifact, displayName, skillName }),
      ),
      fs.writeFile(
        path.join(stagedSkill, "artifact-template.json"),
        `${JSON.stringify(
          {
            schemaVersion: 1,
            kind: artifact.kind,
            reference: `assets/${referenceFilename}`,
            preview: "assets/preview.png",
          },
          null,
          2,
        )}\n`,
      ),
      fs.copyFile(
        referencePath,
        path.join(stagedSkill, "assets", referenceFilename),
      ),
      fs.copyFile(previewPath, path.join(stagedSkill, "assets", "preview.png")),
    ]);
    return stagedSkill;
  } catch (error) {
    await fs.rm(stagedSkill, { force: true, recursive: true });
    throw error;
  }
}

function getTemplateSkillMarkdown({
  artifact,
  description,
  displayName,
  skillName,
}) {
  const triggerDescription = `Create a ${artifact.kind} using the ${displayName} template and its retained reference file. Use when the user selects this template, names ${displayName}, or explicitly invokes ${skillName}. ${description}`;
  return `---
name: ${skillName}
description: ${JSON.stringify(triggerDescription)}
---

# ${displayName}

Create a new ${artifact.kind} from this template. Keep the reference file unchanged.

## Workflow

1. Read \`artifact-template.json\` and resolve its paths relative to this skill directory.
2. Use the matching ${artifact.workflow} workflow with the retained reference file.
3. Treat the user's prompt and available sources as the content input. Do not invent facts merely to fill a template slot.
4. Clone or import the reference instead of replacing its visual system with generic defaults.
5. Render and verify the finished ${artifact.kind}, then return the final artifact.

## Fidelity

${artifact.preservation}

User instructions control requested content and explicit deviations. The retained reference controls layout and formatting where the user has not requested a change.
`;
}

function getTemplateAgentYaml({ artifact, displayName, skillName }) {
  const candidate = `Create ${artifact.kind}s with the ${displayName} template`;
  const shortDescription =
    candidate.length <= 64
      ? candidate
      : `Create a ${artifact.kind} from this saved template`;
  return `interface:
  display_name: ${JSON.stringify(displayName)}
  short_description: ${JSON.stringify(shortDescription)}
  icon_large: "./assets/preview.png"
  default_prompt: ${JSON.stringify(`Use the ${skillName} skill to create a new ${artifact.kind} with this template.`)}
`;
}

async function replaceTemplateSkill(stagedPath, finalPath) {
  const backupPath = `${finalPath}.backup-${randomUUID()}`;
  await fs.rename(finalPath, backupPath);
  try {
    await fs.rename(stagedPath, finalPath);
  } catch (error) {
    try {
      await fs.rename(backupPath, finalPath);
    } catch (rollbackError) {
      throw new AggregateError(
        [error, rollbackError],
        "Template update failed and rollback was incomplete.",
      );
    }
    throw error;
  }
  await fs.rm(backupPath, { force: true, recursive: true });
}

async function acquireWriteLock(lockPath) {
  await fs.mkdir(path.dirname(lockPath), { recursive: true });
  let canRecover = true;
  while (true) {
    const candidatePath = `${lockPath}.pending-${randomUUID()}`;
    let lockError;
    try {
      await fs.writeFile(candidatePath, `${process.pid}\n`, { flag: "wx" });
      await fs.link(candidatePath, lockPath);
      return () => fs.rm(lockPath, { force: true, recursive: true });
    } catch (error) {
      lockError = error;
    } finally {
      await fs.rm(candidatePath, { force: true });
    }

    if (lockError?.code !== "EEXIST") {
      throw lockError;
    }
    if (
      canRecover &&
      (await isWriteLockStale(lockPath)) &&
      (await recoverStaleWriteLock(lockPath))
    ) {
      canRecover = false;
      continue;
    }
    throw new Error(
      `Another artifact template write is already in progress at ${lockPath}.`,
    );
  }
}

async function recoverStaleWriteLock(lockPath) {
  const recoveryPath = `${lockPath}.stale-${randomUUID()}`;
  try {
    await fs.rename(lockPath, recoveryPath);
  } catch (error) {
    if (error?.code === "ENOENT") {
      return false;
    }
    throw error;
  }

  let removed = false;
  try {
    if (!(await isWriteLockStale(recoveryPath))) {
      await fs.rename(recoveryPath, lockPath);
      return false;
    }
    await fs.rm(recoveryPath, { force: true, recursive: true });
    removed = true;
    return true;
  } catch (error) {
    if (!removed) {
      try {
        await fs.rename(recoveryPath, lockPath);
      } catch (restoreError) {
        throw new AggregateError(
          [error, restoreError],
          "Template write-lock recovery failed and could not restore the lock.",
        );
      }
    }
    throw error;
  }
}

async function isWriteLockStale(lockPath) {
  const stat = await fs.lstat(lockPath).catch((error) => {
    if (error?.code === "ENOENT") {
      return null;
    }
    throw error;
  });
  if (stat == null) {
    return false;
  }
  const ownerPath = stat.isDirectory()
    ? path.join(lockPath, LOCK_OWNER_FILENAME)
    : lockPath;
  const owner = await fs.readFile(ownerPath, "utf8").catch((error) => {
    if (error?.code === "ENOENT" || error?.code === "ENOTDIR") {
      return null;
    }
    throw error;
  });
  const ownerPid = Number(owner);
  if (Number.isSafeInteger(ownerPid) && ownerPid > 0) {
    try {
      process.kill(ownerPid, 0);
      return false;
    } catch (error) {
      if (error?.code === "ESRCH") {
        return true;
      }
      if (error?.code === "EPERM") {
        return false;
      }
      throw error;
    }
  }
  return false;
}

function getSlug(displayName) {
  const slug = displayName
    .normalize("NFKD")
    .replace(/[\u0300-\u036f]/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]+/gu, "-")
    .replace(/^-+|-+$/gu, "");
  if (slug.length === 0) {
    throw new Error(
      "--display-name must contain at least one ASCII letter or number.",
    );
  }
  return slug;
}

function assertSkillName(skillName) {
  if (
    skillName.length > MAX_SKILL_NAME_LENGTH ||
    !/^artifact-template-[a-z0-9]+(?:-[a-z0-9]+)*$/u.test(skillName)
  ) {
    throw new Error(
      "--skill-name must be a valid artifact-template skill name.",
    );
  }
}

function assertSingleLine(value, label, maxLength) {
  if (/[<>]/u.test(value)) {
    throw new Error(`${label} must not contain angle brackets.`);
  }
  if (value.length > maxLength || /[\0\r\n]/u.test(value)) {
    throw new Error(
      `${label} must be one line of at most ${maxLength} characters.`,
    );
  }
}

function hasValidPngStructure(bytes) {
  if (
    bytes.length < PNG_SIGNATURE.length ||
    !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)
  ) {
    return false;
  }
  let offset = PNG_SIGNATURE.length;
  let firstChunk = true;
  while (offset + 12 <= bytes.length) {
    const dataLength = bytes.readUInt32BE(offset);
    const dataStart = offset + 8;
    const crcOffset = dataStart + dataLength;
    const nextOffset = crcOffset + 4;
    if (nextOffset > bytes.length) {
      return false;
    }
    const type = bytes.toString("ascii", offset + 4, dataStart);
    if (firstChunk && (type !== "IHDR" || dataLength !== 13)) {
      return false;
    }
    if (
      crc32(bytes.subarray(offset + 4, crcOffset)) !==
      bytes.readUInt32BE(crcOffset)
    ) {
      return false;
    }
    if (type === "IEND") {
      return dataLength === 0 && nextOffset === bytes.length;
    }
    firstChunk = false;
    offset = nextOffset;
  }
  return false;
}

async function assertRegularFile(filePath, label) {
  const stat = await fs.stat(filePath);
  if (!stat.isFile()) {
    throw new Error(`${label} must point to a file.`);
  }
}

function getRequiredString(value, key, label) {
  const entry = value[key];
  if (typeof entry !== "string" || entry.trim().length === 0) {
    throw new Error(`${label} must be a non-empty string.`);
  }
  return entry.trim();
}

function getOptionalString(value, key, label) {
  const entry = value[key];
  if (entry == null) {
    return null;
  }
  if (typeof entry !== "string" || entry.trim().length === 0) {
    throw new Error(`${label} must be a non-empty string when provided.`);
  }
  return entry.trim();
}

async function pathExists(filePath) {
  return fs
    .access(filePath)
    .then(() => true)
    .catch((error) => {
      if (error?.code === "ENOENT") {
        return false;
      }
      throw error;
    });
}

const requestFlagToKey = new Map([
  ["--mode", "mode"],
  ["--skill-name", "skillName"],
  ["--reference-path", "referencePath"],
  ["--preview-path", "previewPath"],
  ["--display-name", "displayName"],
  ["--description", "description"],
]);

function getRequestFromArguments(args) {
  if (args.length === 0 || args.length % 2 !== 0) {
    throw new Error(USAGE);
  }

  const request = {};
  for (let index = 0; index < args.length; index += 2) {
    const flag = args[index];
    const key = requestFlagToKey.get(flag);
    if (key == null || Object.hasOwn(request, key)) {
      throw new Error(USAGE);
    }
    request[key] = args[index + 1];
  }
  return request;
}

async function main() {
  const request = getRequestFromArguments(process.argv.slice(2));
  const result = await createTemplateSkill(request);
  process.stdout.write(`${JSON.stringify(result)}\n`);
}

if (
  process.argv[1] != null &&
  fileURLToPath(import.meta.url) === path.resolve(process.argv[1])
) {
  main().catch((error) => {
    process.stderr.write(
      `${error instanceof Error ? error.message : String(error)}\n`,
    );
    process.exitCode = 1;
  });
}
