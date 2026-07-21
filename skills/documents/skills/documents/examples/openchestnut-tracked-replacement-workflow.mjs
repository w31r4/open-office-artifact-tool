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

function trackedReplacementCandidates(document, expectedText) {
  const candidates = [];
  for (let blockIndex = 0; blockIndex < document.blocks.length; blockIndex += 1) {
    const block = document.blocks[blockIndex];
    if (block.kind === "paragraph" && block.text === expectedText) {
      candidates.push({ kind: "paragraph", blockIndex });
      continue;
    }
    if (block.kind !== "table") continue;
    for (let row = 0; row < block.rows; row += 1) {
      for (let column = 0; column < block.columns; column += 1) {
        const cell = block.getCell(row, column);
        if (cell.value === expectedText && cell.verticalMerge !== "continue") {
          candidates.push({ kind: "tableCell", blockIndex, row, column });
        }
      }
    }
  }
  return candidates;
}

function normalizeTrackedReplacementTarget(target, targetBlockIndex, candidates) {
  if (target !== undefined && targetBlockIndex !== undefined) {
    throw new Error("Use target or the paragraph-only targetBlockIndex compatibility option, not both.");
  }
  const selected = target !== undefined
    ? target
    : targetBlockIndex !== undefined
      ? { kind: "paragraph", blockIndex: Number(targetBlockIndex) }
      : candidates.length === 1
        ? candidates[0]
        : undefined;
  if (!selected || typeof selected !== "object" ||
      !["paragraph", "tableCell"].includes(selected.kind) ||
      !Number.isInteger(selected.blockIndex) || selected.blockIndex < 0 || selected.blockIndex > 0xffff_ffff) {
    throw new Error(`Tracked replacement requires one explicit paragraph/table-cell target; exact-text discovery found ${candidates.length}.`);
  }
  if (selected.kind === "paragraph") return { kind: "paragraph", blockIndex: selected.blockIndex };
  if (!Number.isInteger(selected.row) || selected.row < 0 || selected.row > 0xffff_ffff ||
      !Number.isInteger(selected.column) || selected.column < 0 || selected.column > 0xffff_ffff) {
    throw new Error("A tableCell target requires unsigned 32-bit physical row and column indexes.");
  }
  return { kind: "tableCell", blockIndex: selected.blockIndex, row: selected.row, column: selected.column };
}

function targetSnapshot(document, target) {
  const block = document.blocks[target.blockIndex];
  if (target.kind === "paragraph") return block?.kind === "paragraph"
    ? { text: block.text, sourceBound: block.textEditable === false }
    : undefined;
  if (block?.kind !== "table" || target.row >= block.rows || target.column >= block.columns) return undefined;
  const cell = block.getCell(target.row, target.column);
  return { text: cell.value, sourceBound: cell.editable === false };
}

export async function addDocumentTrackedReplacement({
  inputPath,
  outputPath,
  auditPath,
  target,
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
  const candidates = trackedReplacementCandidates(sourceDocument, expected);
  const selectedTarget = normalizeTrackedReplacementTarget(target, targetBlockIndex, candidates);
  const selected = targetSnapshot(sourceDocument, selectedTarget);
  if (!selected || selected.text !== expected) throw new Error("target does not identify the expected exact paragraph/table-cell snapshot.");

  const tracked = await DocumentFile.addTrackedReplacement(sourceBlob, {
    expectedSourceSha256: sourceSha256,
    target: selectedTarget,
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
  if (JSON.stringify(operation.target) !== JSON.stringify(selectedTarget) ||
      operation.targetBlockIndex !== selectedTarget.blockIndex ||
      operation.changedParts.join("\n") !== "word/document.xml") {
    throw new Error("OpenChestnut tracked replacement changed an unexpected target or OPC part.");
  }
  if (operation.deletedTextSha256 !== sha256(Buffer.from(oldText, "utf8")) ||
      operation.insertedTextSha256 !== sha256(Buffer.from(newText, "utf8"))) {
    throw new Error("OpenChestnut tracked-replacement audit does not match the requested text hashes.");
  }

  const acceptedProjection = `${expected.slice(0, matchIndex)}${newText}${expected.slice(matchIndex + oldText.length)}`;
  const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
  const reimportedTarget = targetSnapshot(reimported, selectedTarget);
  if (!reimportedTarget || reimportedTarget.text !== acceptedProjection || !reimportedTarget.sourceBound) {
    throw new Error("Tracked DOCX re-import does not expose the expected accepted-view, source-bound paragraph/table-cell projection.");
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
      target: operation.target,
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
    target: result.audit.operation.target,
    targetBlockIndex: result.audit.operation.targetBlockIndex,
  }));
}
