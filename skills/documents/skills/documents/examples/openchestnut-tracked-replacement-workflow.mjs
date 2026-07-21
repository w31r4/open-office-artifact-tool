import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.length) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

async function renderModel(document) {
  const preview = await document.render({ format: "svg" });
  if (!/<svg\b/i.test(await preview.text())) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

function uniqueLiteralIndex(value, search) {
  const first = value.indexOf(search);
  if (first < 0 || value.indexOf(search, first + 1) >= 0) {
    throw new Error("expectedText must contain search exactly once before native mutation.");
  }
  return first;
}

export async function addDocumentTrackedReplacement({
  inputPath,
  outputPath,
  auditPath,
  targetBlockIndex,
  expectedText,
  search,
  replacement,
  author,
  date,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const expected = requiredText(expectedText, "expectedText");
  const oldText = requiredText(search, "search");
  const newText = requiredText(replacement, "replacement");
  const reviewer = requiredText(author, "author");
  const matchIndex = uniqueLiteralIndex(expected, oldText);
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original document remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from both DOCX paths.");

  const source = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(source);
  const sourceBlob = new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) });
  const sourceDocument = await DocumentFile.importDocx(sourceBlob);
  const candidates = sourceDocument.blocks
    .map((block, index) => ({ block, index }))
    .filter(({ block }) => block.kind === "paragraph" && block.text === expected);
  const selectedIndex = targetBlockIndex === undefined
    ? (candidates.length === 1 ? candidates[0].index : undefined)
    : Number(targetBlockIndex);
  if (!Number.isInteger(selectedIndex) || selectedIndex < 0 || selectedIndex >= sourceDocument.blocks.length) {
    throw new Error(`Tracked replacement requires one explicit target block; exact-text discovery found ${candidates.length}.`);
  }
  const selected = sourceDocument.blocks[selectedIndex];
  if (selected.kind !== "paragraph" || selected.text !== expected) {
    throw new Error("targetBlockIndex does not identify the expected exact paragraph snapshot.");
  }

  const tracked = await DocumentFile.addTrackedReplacement(sourceBlob, {
    expectedSourceSha256: sourceSha256,
    targetBlockIndex: selectedIndex,
    expectedText: expected,
    search: oldText,
    replacement: newText,
    author: reviewer,
    date,
  });
  const output = Buffer.from(await tracked.arrayBuffer());
  const operation = tracked.metadata?.trackedReplacement;
  if (!operation || operation.sourceSha256 !== sourceSha256 || operation.outputSha256 !== sha256(output)) {
    throw new Error("OpenChestnut tracked-replacement audit does not bind the exact source and output bytes.");
  }
  if (operation.targetBlockIndex !== selectedIndex || operation.changedParts.join("\n") !== "word/document.xml") {
    throw new Error("OpenChestnut tracked replacement changed an unexpected target or OPC part.");
  }
  if (operation.deletedTextSha256 !== sha256(Buffer.from(oldText, "utf8")) ||
      operation.insertedTextSha256 !== sha256(Buffer.from(newText, "utf8"))) {
    throw new Error("OpenChestnut tracked-replacement audit does not match the requested text hashes.");
  }

  const acceptedProjection = `${expected.slice(0, matchIndex)}${newText}${expected.slice(matchIndex + oldText.length)}`;
  const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
  const reimportedTarget = reimported.blocks[selectedIndex];
  if (reimportedTarget?.kind !== "paragraph" || reimportedTarget.text !== acceptedProjection || reimportedTarget.textEditable !== false) {
    throw new Error("Tracked DOCX re-import does not expose the expected accepted-view, source-bound paragraph projection.");
  }
  const verification = reimported.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
  const modelRender = await renderModel(reimported);
  if (sha256(await fs.readFile(sourcePath)) !== sourceSha256) throw new Error("Source DOCX changed during tracked replacement.");

  const audit = {
    schema: "open-office-artifact-tool.docx-audit.v1",
    status: "succeeded",
    source: { path: sourcePath, sha256: sourceSha256, bytes: source.length },
    output: { path: finalPath, sha256: operation.outputSha256, bytes: output.length },
    provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite", overwrite: false },
    operation: {
      type: "add-tracked-replacement",
      targetBlockIndex: operation.targetBlockIndex,
      targetBodyIndex: operation.targetBodyIndex,
      sourceElementSha256: operation.sourceElementSha256,
      outputElementSha256: operation.outputElementSha256,
      deletedTextSha256: operation.deletedTextSha256,
      insertedTextSha256: operation.insertedTextSha256,
      deletedTextChars: operation.deletedTextChars,
      insertedTextChars: operation.insertedTextChars,
      deletionNativeRevisionId: operation.deletionNativeRevisionId,
      insertionNativeRevisionId: operation.insertionNativeRevisionId,
      changedParts: operation.changedParts,
    },
    warnings: ["Native LibreOffice/Word render review remains required for release-sensitive redline display."],
    validation: {
      reimport: { ok: true, blockCount: reimported.blocks.length, acceptedProjectionSha256: sha256(Buffer.from(acceptedProjection, "utf8")) },
      verify: { ok: true },
      modelRender: { ok: true, ...modelRender },
      sourceImmutable: { ok: true },
    },
  };

  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  let outputPublished = false;
  try {
    await fs.writeFile(finalPath, output, { flag: "wx" });
    outputPublished = true;
    await fs.writeFile(finalAuditPath, JSON.stringify(audit, null, 2), { flag: "wx" });
  } catch (error) {
    if (outputPublished) await fs.rm(finalPath, { force: true });
    throw error;
  }
  return { outputPath: finalPath, auditPath: finalAuditPath, audit };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [inputPath, outputPath, auditPath, requestPath] = process.argv.slice(2);
  const request = JSON.parse(await fs.readFile(path.resolve(requiredText(requestPath, "requestPath")), "utf8"));
  const result = await addDocumentTrackedReplacement({ inputPath, outputPath, auditPath, ...request });
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    targetBlockIndex: result.audit.operation.targetBlockIndex,
  }));
}
