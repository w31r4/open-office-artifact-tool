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
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

function blockProjection(block) {
  const common = { kind: block.kind, styleId: block.styleId || "" };
  if (block.kind === "table") return { ...common, values: block.values };
  if (block.kind === "image") return { ...common, alt: block.alt, widthPx: block.widthPx, heightPx: block.heightPx };
  if (block.kind === "section") return { ...common, breakType: block.breakType, orientation: block.orientation, pageSize: block.pageSize, margins: block.margins };
  if (block.kind === "hyperlink") return { ...common, text: block.text, url: block.url, anchor: block.anchor, tooltip: block.tooltip, history: block.history };
  if (block.kind === "field") return { ...common, instruction: block.instruction, display: block.display, complex: block.complex };
  return { ...common, text: String(block.text ?? block.display ?? "") };
}

function expectedBlocks(document, mode) {
  return document.blocks.map((block) => {
    if (block.kind !== "change") return blockProjection(block);
    const retainText = mode === "accept" ? block.changeType === "insert" : block.changeType === "delete";
    return { kind: "paragraph", styleId: block.styleId || "", text: retainText ? block.text : "" };
  });
}

async function renderModel(document) {
  const preview = await document.render({ format: "svg" });
  if (!/<svg\b/i.test(await preview.text())) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function finalizeDocumentRevisions({
  inputPath,
  outputPath,
  auditPath,
  mode,
  keepTracking = false,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const requestedMode = requiredText(mode, "mode").toLowerCase();
  if (!new Set(["accept", "reject"]).has(requestedMode)) throw new TypeError("mode must be accept or reject.");
  if (typeof keepTracking !== "boolean") throw new TypeError("keepTracking must be a boolean.");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original document remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from both DOCX paths.");

  const source = await fs.readFile(sourcePath);
  const sourceSha256 = sha256(source);
  const sourceDocument = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const revisions = sourceDocument.blocks
    .filter((block) => block.kind === "change")
    .map((block) => ({ id: block.id, type: block.changeType, author: block.author, date: block.date, textSha256: sha256(Buffer.from(block.text, "utf8")) }));
  if (revisions.length === 0) throw new Error("The imported document exposes no tracked revisions to finalize.");
  const expected = expectedBlocks(sourceDocument, requestedMode);

  const finalized = await DocumentFile.finalizeRevisions(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }), {
    mode: requestedMode,
    keepTracking,
    expectedSourceSha256: sourceSha256,
  });
  const output = Buffer.from(await finalized.arrayBuffer());
  const operation = finalized.metadata?.revisionFinalization;
  if (!operation || operation.sourceSha256 !== sourceSha256 || operation.outputSha256 !== sha256(output)) {
    throw new Error("OpenChestnut revision-finalization audit does not bind the exact source and output bytes.");
  }
  if (operation.insertionCount + operation.deletionCount !== revisions.length) {
    throw new Error("OpenChestnut finalized a different number of revisions than the inspected source model exposed.");
  }
  const allowedParts = new Set(["word/document.xml", "word/settings.xml"]);
  if (!operation.changedParts.includes("word/document.xml") || operation.changedParts.some((part) => !allowedParts.has(part))) {
    throw new Error(`Revision finalization changed an unexpected OPC part: ${operation.changedParts.join(", ")}.`);
  }

  const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
  if (reimported.blocks.some((block) => block.kind === "change")) throw new Error("Finalized DOCX still exposes tracked revisions after re-import.");
  if (JSON.stringify(reimported.blocks.map(blockProjection)) !== JSON.stringify(expected)) {
    throw new Error("Finalized DOCX semantics do not match the requested accept/reject projection.");
  }
  if (reimported.settings.trackRevisions !== operation.trackingAfter) {
    throw new Error("Finalized DOCX tracking setting does not match the provider audit.");
  }
  const verification = reimported.verify({ visualQa: true });
  if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
  const modelRender = await renderModel(reimported);
  if (sha256(await fs.readFile(sourcePath)) !== sourceSha256) throw new Error("Source DOCX changed during revision finalization.");

  const audit = {
    schema: "open-office-artifact-tool.docx-audit.v1",
    status: "succeeded",
    source: { path: sourcePath, sha256: sourceSha256, bytes: source.length },
    output: { path: finalPath, sha256: operation.outputSha256, bytes: output.length },
    provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite", overwrite: false },
    operation: {
      type: "finalize-revisions",
      mode: requestedMode,
      keepTracking,
      insertionCount: operation.insertionCount,
      deletionCount: operation.deletionCount,
      trackingBefore: operation.trackingBefore,
      trackingAfter: operation.trackingAfter,
      changedParts: operation.changedParts,
      revisions,
    },
    warnings: ["Native LibreOffice/Word render review remains required for release-sensitive layout."],
    validation: {
      reimport: { ok: true, remainingRevisions: 0, blockCount: reimported.blocks.length },
      semanticProjection: { ok: true },
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
  const [inputPath, outputPath, auditPath, mode, ...flags] = process.argv.slice(2);
  const result = await finalizeDocumentRevisions({
    inputPath,
    outputPath,
    auditPath,
    mode,
    keepTracking: flags.includes("--keep-tracking"),
  });
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    remainingRevisions: result.audit.validation.reimport.remainingRevisions,
  }));
}
