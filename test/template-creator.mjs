import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import sharp from "sharp";

import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "open-office-artifact-tool";

const packageRoot = path.resolve(import.meta.dirname, "..");
const creatorPath = path.join(
  packageRoot,
  "skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs",
);

try {
  await fs.access(creatorPath);
} catch (error) {
  if (error?.code === "ENOENT") {
    console.log("template creator smoke skipped: repository-only skills are not packaged");
    process.exit(0);
  }
  throw error;
}

const tempRoot = await fs.mkdtemp(
  path.join(os.tmpdir(), "office-artifact-tool-template-creator-"),
);
const home = path.join(tempRoot, "neutral-home");
const fixturesDirectory = path.join(tempRoot, "fixtures");

function runCreator(args) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, [creatorPath, ...args], {
      env: { ...process.env, OFFICE_ARTIFACT_HOME: home },
      stdio: ["ignore", "pipe", "pipe"],
    });
    let stdout = "";
    let stderr = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => {
      stdout += chunk;
    });
    child.stderr.on("data", (chunk) => {
      stderr += chunk;
    });
    child.once("error", reject);
    child.once("close", (code) => {
      resolve({ code, stderr, stdout });
    });
  });
}

async function runSuccessfulCreator(args) {
  const result = await runCreator(args);
  if (result.code !== 0) {
    throw new Error(`Template creator failed (${result.code}): ${result.stderr}`);
  }
  try {
    return JSON.parse(result.stdout);
  } catch (error) {
    throw new Error(`Template creator did not return JSON: ${result.stdout}\n${error}`);
  }
}

async function assertBytesEqual(actualPath, expectedPath, label) {
  const [actual, expected] = await Promise.all([
    fs.readFile(actualPath),
    fs.readFile(expectedPath),
  ]);
  if (!actual.equals(expected)) {
    throw new Error(`${label} did not retain the exact source bytes.`);
  }
}

async function assertGeneratedTemplate(
  result,
  {
    kind,
    referencePath,
    visualCommitment = "opinionated",
    editLevel = "copy-only",
    provenanceSource = "local-user-reference",
  },
) {
  const skillsRoot = path.join(home, "skills");
  if (path.dirname(result.skillPath) !== skillsRoot) {
    throw new Error(`Template was not written below OFFICE_ARTIFACT_HOME: ${result.skillPath}`);
  }
  if (result.kind !== kind || !result.skillName.startsWith("artifact-template-")) {
    throw new Error(`Unexpected template result: ${JSON.stringify(result)}`);
  }

  const extension = path.extname(referencePath).toLowerCase();
  const referenceName = `reference${extension}`;
  const agentPath = path.join(result.skillPath, "agents/agent.yaml");
  const legacyAgentPath = path.join(result.skillPath, "agents/openai.yaml");
  const sidecarPath = path.join(result.skillPath, "artifact-template.json");
  const skillPath = path.join(result.skillPath, "SKILL.md");
  const previewPath = path.join(result.skillPath, "assets/preview.png");
  const retainedReferencePath = path.join(result.skillPath, "assets", referenceName);

  await Promise.all([
    fs.access(agentPath),
    fs.access(sidecarPath),
    fs.access(skillPath),
    fs.access(previewPath),
    fs.access(retainedReferencePath),
  ]);
  if (await fs.access(legacyAgentPath).then(() => true).catch(() => false)) {
    throw new Error("Generated template retained the legacy agent metadata filename.");
  }

  const [sidecar, skillText, previewBytes, retainedReferenceBytes] = await Promise.all([
    fs.readFile(sidecarPath, "utf8").then(JSON.parse),
    fs.readFile(skillPath, "utf8"),
    fs.readFile(previewPath),
    fs.readFile(retainedReferencePath),
  ]);
  if (
    result.schemaVersion !== 2 ||
    sidecar.schemaVersion !== 2 ||
    sidecar.id !== result.skillName ||
    sidecar.displayName !== result.displayName ||
    sidecar.kind !== kind ||
    sidecar.reference !== `assets/${referenceName}` ||
    sidecar.preview !== "assets/preview.png" ||
    !Array.isArray(sidecar.useWhen) ||
    sidecar.useWhen.length === 0 ||
    sidecar.visualCommitment !== visualCommitment ||
    sidecar.editProfile?.level !== editLevel ||
    (editLevel === "copy-only" && sidecar.editProfile?.verifiedOperations?.length !== 0) ||
    sidecar.provenance?.license !== "user-provided" ||
    sidecar.provenance?.source !== provenanceSource ||
    sidecar.provenance?.referenceSha256 !== sha256(retainedReferenceBytes) ||
    sidecar.provenance?.previewSha256 !== sha256(previewBytes)
  ) {
    throw new Error(`Generated sidecar is invalid: ${JSON.stringify(sidecar)}`);
  }
  if (/codex|openai|plugin:\/\//iu.test(skillText)) {
    throw new Error("Generated template skill contains a product-specific reference.");
  }

  await Promise.all([
    assertBytesEqual(retainedReferencePath, referencePath, `${kind} reference`),
    assertBytesEqual(previewPath, path.join(fixturesDirectory, "preview.png"), `${kind} preview`),
  ]);
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

async function assertNoTransactionalResidue() {
  const [skillEntries, homeEntries] = await Promise.all([
    fs.readdir(path.join(home, "skills")),
    fs.readdir(home),
  ]);
  const skillResidue = skillEntries.filter(
    (entry) => entry.includes("-stage-") || entry.includes(".backup-"),
  );
  const lockResidue = homeEntries.filter(
    (entry) =>
      entry.startsWith(".artifact-template-write-lock.pending-") ||
      entry.startsWith(".artifact-template-write-lock.stale-"),
  );
  const residue = [...skillResidue, ...lockResidue];
  if (residue.length > 0) {
    throw new Error(`Template creator left transactional residue: ${residue.join(", ")}`);
  }
}

async function writePresentationFixture(filePath, slideCount) {
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });
  for (let index = 0; index < slideCount; index += 1) presentation.slides.add();
  const file = await PresentationFile.exportPptx(presentation);
  await file.save(filePath);
}

async function writePngFixture(filePath) {
  await sharp({
    create: {
      width: 320,
      height: 180,
      channels: 4,
      background: { r: 15, g: 118, b: 110, alpha: 1 },
    },
  }).png().toFile(filePath);
}

async function writeDocumentFixture(filePath) {
  const document = DocumentModel.create();
  const file = await DocumentFile.exportDocx(document);
  await file.save(filePath);
}

async function writeSpreadsheetFixture(filePath) {
  const workbook = Workbook.create();
  const worksheet = workbook.worksheets.add("Fixture");
  worksheet.getRange("A1:B2").values = [
    ["Kind", "Value"],
    ["Fixture", 1],
  ];
  const file = await SpreadsheetFile.exportXlsx(workbook);
  await file.save(filePath);
}

try {
  await fs.mkdir(fixturesDirectory, { recursive: true });
  const pptxPath = path.join(fixturesDirectory, "reference.pptx");
  const updatedPptxPath = path.join(fixturesDirectory, "updated-reference.pptx");
  const docxPath = path.join(fixturesDirectory, "reference.docx");
  const xlsxPath = path.join(fixturesDirectory, "reference.xlsx");
  const previewPath = path.join(fixturesDirectory, "preview.png");

  await Promise.all([
    writePresentationFixture(pptxPath, 1),
    writePresentationFixture(updatedPptxPath, 2),
    writeSpreadsheetFixture(xlsxPath),
    writeDocumentFixture(docxPath),
    writePngFixture(previewPath),
    fs.mkdir(home, { recursive: true }),
  ]);
  await fs.writeFile(
    path.join(home, ".artifact-template-write-lock"),
    "999999999\n",
  );

  const pptxSelection = {
    useWhen: ["quarterly project review"],
    avoidWhen: ["legal memorandum"],
    audiences: ["executive"],
    contentShapes: ["status", "risks", "decisions"],
    visualTraits: {
      tone: ["formal"],
      density: "medium",
      colorMode: "light",
      structure: ["sectioned"],
    },
    visualCommitment: "neutral",
    editProfile: {
      level: "bounded-edit",
      verifiedOperations: ["recognized-placeholder-title-text-replace"],
    },
    provenance: {
      license: "user-provided",
      source: "local-test-reference",
    },
  };
  const unsupportedSelection = structuredClone(pptxSelection);
  unsupportedSelection.visualTraits.undocumentedTrait = true;
  const unsupportedSelectionResult = await runCreator([
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Unsupported selection fixture",
    "--description", "This fixture must fail before writing a template.",
    "--selection-json", JSON.stringify(unsupportedSelection),
  ]);
  assert.notEqual(unsupportedSelectionResult.code, 0);
  assert.match(
    unsupportedSelectionResult.stderr,
    /visualTraits contains unsupported fields: undocumentedTrait/,
  );

  const pptxTemplate = await runSuccessfulCreator([
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Presentation fixture",
    "--description", "Create presentations from the fixture layout.",
    "--selection-json", JSON.stringify(pptxSelection),
  ]);
  await assertGeneratedTemplate(pptxTemplate, {
    kind: "presentation",
    referencePath: pptxPath,
    visualCommitment: "neutral",
    editLevel: "bounded-edit",
    provenanceSource: "local-test-reference",
  });
  const createdPptxMetadata = JSON.parse(
    await fs.readFile(path.join(pptxTemplate.skillPath, "artifact-template.json"), "utf8"),
  );
  for (const key of [
    "useWhen",
    "avoidWhen",
    "audiences",
    "contentShapes",
    "visualTraits",
    "visualCommitment",
    "editProfile",
  ]) {
    assert.deepEqual(createdPptxMetadata[key], pptxSelection[key], `selection metadata ${key}`);
  }
  assert.equal(createdPptxMetadata.provenance.source, "local-test-reference");

  const docxTemplate = await runSuccessfulCreator([
    "--reference-path", docxPath,
    "--preview-path", previewPath,
    "--display-name", "Document fixture",
    "--description", "Create documents from the fixture layout.",
  ]);
  await assertGeneratedTemplate(docxTemplate, { kind: "document", referencePath: docxPath });

  const xlsxTemplate = await runSuccessfulCreator([
    "--reference-path", xlsxPath,
    "--preview-path", previewPath,
    "--display-name", "Spreadsheet fixture",
    "--description", "Create spreadsheets from the fixture layout.",
  ]);
  await assertGeneratedTemplate(xlsxTemplate, { kind: "spreadsheet", referencePath: xlsxPath });

  const longDescription =
    "Create a detailed planning workbook for a recurring operating review with assumptions, " +
    "owners, milestones, risks, decisions, supporting evidence, and a concise executive summary.";
  const longDescriptionTemplate = await runSuccessfulCreator([
    "--reference-path", xlsxPath,
    "--preview-path", previewPath,
    "--display-name", "Long description fixture",
    "--description", longDescription,
  ]);
  await assertGeneratedTemplate(longDescriptionTemplate, {
    kind: "spreadsheet",
    referencePath: xlsxPath,
  });
  const longDescriptionMetadata = JSON.parse(
    await fs.readFile(
      path.join(longDescriptionTemplate.skillPath, "artifact-template.json"),
      "utf8",
    ),
  );
  assert.ok(longDescriptionMetadata.useWhen[0].length <= 120);
  assert.equal(longDescriptionMetadata.useWhen[0], longDescription.slice(0, 120).trimEnd());
  assert.match(
    await fs.readFile(path.join(longDescriptionTemplate.skillPath, "SKILL.md"), "utf8"),
    new RegExp(longDescription.replace(/[.*+?^${}()|[\]\\]/gu, "\\$&")),
  );

  const kindChange = await runCreator([
    "--mode", "update",
    "--skill-name", docxTemplate.skillName,
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Document fixture",
    "--description", "Attempt to change the document fixture kind.",
  ]);
  if (kindChange.code === 0) {
    throw new Error("Template creator accepted an artifact-kind-changing update.");
  }
  await assertGeneratedTemplate(docxTemplate, { kind: "document", referencePath: docxPath });

  const interruptedBackupPath = `${pptxTemplate.skillPath}.backup-11111111-1111-4111-8111-111111111111`;
  await fs.rename(pptxTemplate.skillPath, interruptedBackupPath);
  const sentinelPath = path.join(interruptedBackupPath, "sentinel.txt");
  await fs.writeFile(sentinelPath, "retain me\n");
  const updatedPptxTemplate = await runSuccessfulCreator([
    "--mode", "update",
    "--skill-name", pptxTemplate.skillName,
    "--reference-path", updatedPptxPath,
    "--preview-path", previewPath,
    "--display-name", "Updated presentation fixture",
    "--description", "Create presentations from the updated fixture layout.",
  ]);
  if (updatedPptxTemplate.skillPath !== pptxTemplate.skillPath) {
    throw new Error("Template update changed the template path.");
  }
  await assertGeneratedTemplate(updatedPptxTemplate, {
    kind: "presentation",
    referencePath: updatedPptxPath,
    visualCommitment: "neutral",
    editLevel: "bounded-edit",
    provenanceSource: "local-test-reference",
  });
  const updatedPptxMetadata = JSON.parse(
    await fs.readFile(path.join(updatedPptxTemplate.skillPath, "artifact-template.json"), "utf8"),
  );
  for (const key of [
    "useWhen",
    "avoidWhen",
    "audiences",
    "contentShapes",
    "visualTraits",
    "visualCommitment",
    "editProfile",
  ]) {
    assert.deepEqual(updatedPptxMetadata[key], pptxSelection[key], `preserved selection metadata ${key}`);
  }
  const restoredSentinel = await fs.readFile(
    path.join(updatedPptxTemplate.skillPath, "sentinel.txt"),
    "utf8",
  );
  if (restoredSentinel !== "retain me\n") {
    throw new Error("Template update did not preserve additional template-owned files.");
  }
  await assertNoTransactionalResidue();

  const linkedTemplatePath = path.join(home, "skills", "artifact-template-linked");
  const outsidePath = path.join(tempRoot, "outside-assets");
  await fs.mkdir(linkedTemplatePath, { recursive: true });
  await fs.mkdir(outsidePath, { recursive: true });
  await fs.writeFile(path.join(linkedTemplatePath, "artifact-template.json"), JSON.stringify({ schemaVersion: 1, kind: "document" }));
  await fs.symlink(outsidePath, path.join(linkedTemplatePath, "assets"), "dir");
  const linkedUpdate = await runCreator([
    "--mode", "update",
    "--skill-name", "artifact-template-linked",
    "--reference-path", docxPath,
    "--preview-path", previewPath,
    "--display-name", "Linked template",
    "--description", "This update must fail before following a template-owned symlink.",
  ]);
  if (linkedUpdate.code === 0 || !/reject symbolic links/i.test(linkedUpdate.stderr)) {
    throw new Error(`Template creator did not fail closed on a template-owned symlink: ${linkedUpdate.stderr}`);
  }
  if ((await fs.readdir(outsidePath)).length !== 0) {
    throw new Error("Template creator wrote through a template-owned symbolic link.");
  }
  await fs.rm(linkedTemplatePath, { recursive: true, force: true });
  await assertNoTransactionalResidue();

  const oversizedReferencePath = path.join(fixturesDirectory, "oversized.docx");
  const oversizedHandle = await fs.open(oversizedReferencePath, "w");
  await oversizedHandle.truncate(512 * 1024 * 1024 + 1);
  await oversizedHandle.close();
  const oversized = await runCreator([
    "--reference-path", oversizedReferencePath,
    "--preview-path", previewPath,
    "--display-name", "Oversized reference",
    "--description", "This input must be rejected before it is copied.",
  ]);
  if (oversized.code === 0 || !/input budget/i.test(oversized.stderr)) {
    throw new Error(`Template creator did not enforce its input budget: ${oversized.stderr}`);
  }
  await assertNoTransactionalResidue();

  const activeLockPath = path.join(home, ".artifact-template-write-lock");
  await fs.writeFile(activeLockPath, `${process.pid}\n`);
  const activeLock = await runCreator([
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Blocked presentation fixture",
    "--description", "Attempt to create while another writer owns the lock.",
  ]);
  if (activeLock.code === 0) {
    throw new Error("Template creator wrote through an active write lock.");
  }
  await fs.readFile(activeLockPath, "utf8");
  await fs.rm(activeLockPath, { force: true, recursive: true });

  if (await fs.access(path.join(home, ".artifact-template-write-lock")).then(() => true).catch(() => false)) {
    throw new Error("Template creator did not release the write lock.");
  }

  console.log("template creator smoke ok");
} finally {
  await fs.rm(tempRoot, { force: true, recursive: true });
}
