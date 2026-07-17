import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

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

async function assertGeneratedTemplate(result, { kind, referencePath }) {
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

  const [sidecar, skillText] = await Promise.all([
    fs.readFile(sidecarPath, "utf8").then(JSON.parse),
    fs.readFile(skillPath, "utf8"),
  ]);
  if (
    sidecar.schemaVersion !== 1 ||
    sidecar.kind !== kind ||
    sidecar.reference !== `assets/${referenceName}` ||
    sidecar.preview !== "assets/preview.png"
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
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });
  const slide = presentation.slides.add();
  const png = await presentation.export({ slide, format: "png", scale: 1 });
  await fs.writeFile(filePath, new Uint8Array(await png.arrayBuffer()));
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

  const pptxTemplate = await runSuccessfulCreator([
    "--reference-path", pptxPath,
    "--preview-path", previewPath,
    "--display-name", "Presentation fixture",
    "--description", "Create presentations from the fixture layout.",
  ]);
  await assertGeneratedTemplate(pptxTemplate, { kind: "presentation", referencePath: pptxPath });

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
  });
  const restoredSentinel = await fs.readFile(
    path.join(updatedPptxTemplate.skillPath, "sentinel.txt"),
    "utf8",
  );
  if (restoredSentinel !== "retain me\n") {
    throw new Error("Template update did not preserve additional template-owned files.");
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
