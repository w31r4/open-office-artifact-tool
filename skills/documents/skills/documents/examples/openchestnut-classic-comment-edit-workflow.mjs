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

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function classicCommentSnapshot(comment) {
  return {
    id: comment.id,
    targetId: comment.targetId,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    parentId: comment.parentId,
    resolved: comment.resolved,
    paraId: comment.paraId,
    durableId: comment.durableId,
    dateUtc: comment.dateUtc,
    person: comment.person,
    intelligentPlaceholder: comment.intelligentPlaceholder,
  };
}

function sameSnapshot(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function resolveEditableClassicComment(document, anchorText, expectedCommentText) {
  const targets = document.blocks.filter((block) => block.kind === "paragraph" && block.text.includes(anchorText));
  if (targets.length !== 1) throw new Error("Expected exactly one paragraph containing the requested comment anchor; found " + targets.length + ".");
  const target = targets[0];
  const matches = document.comments.filter((comment) => comment.targetId === target.id);
  if (matches.length !== 1) throw new Error("Expected exactly one comment attached to the requested paragraph; found " + matches.length + ".");
  const comment = matches[0];
  if (comment.parentId || comment.resolved || comment.paraId || comment.durableId || comment.dateUtc || comment.person || comment.intelligentPlaceholder) {
    throw new Error("The selected comment has reply, resolved, presence, or modern identity metadata and is source-bound outside the classic-comment text-edit workflow.");
  }
  if (comment.text !== expectedCommentText) {
    throw new Error("The selected classic comment does not contain the expected original text; refusing to edit an ambiguous review record.");
  }
  return { target, comment };
}

async function renderModel(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function editClassicComment({
  inputPath,
  outputPath,
  auditPath,
  anchorText,
  expectedCommentText,
  replacementText,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original document remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and DOCX output paths.");
  const requestedAnchor = requiredText(anchorText, "anchorText");
  const originalText = requiredText(expectedCommentText, "expectedCommentText");
  const replacement = requiredText(replacementText, "replacementText");

  const source = await fs.readFile(sourcePath);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const { target, comment } = resolveEditableClassicComment(document, requestedAnchor, originalText);
  const targetText = target.text;
  const snapshot = classicCommentSnapshot(comment);
  comment.text = replacement;

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  let output;
  let reimported;
  let verification;
  let modelRender;
  try {
    const exported = await DocumentFile.exportDocx(document);
    await exported.save(temporaryPath);
    output = await fs.readFile(temporaryPath);
    reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = resolveEditableClassicComment(reimported, requestedAnchor, replacement);
    if (reimported.comments.length !== document.comments.length || !sameSnapshot(snapshot, classicCommentSnapshot(roundTrip.comment))) {
      throw new Error("DOCX export changed the classic comment identity, anchor, author, initials, date, or topology.");
    }
    if (roundTrip.target.text !== targetText || roundTrip.comment.text !== replacement) {
      throw new Error("DOCX export did not preserve the target paragraph or requested classic comment text.");
    }
    verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Document verification failed: " + verification.ndjson);
    modelRender = await renderModel(reimported);

    const audit = {
      schema: "open-office-artifact-tool.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "classic-comment-text-edit",
        anchorSha256: sha256(Buffer.from(requestedAnchor, "utf8")),
        commentId: snapshot.id,
        targetId: snapshot.targetId,
      },
      warnings: [],
      validation: {
        reimport: { ok: true, commentCount: reimported.comments.length, commentId: roundTrip.comment.id, targetId: roundTrip.comment.targetId },
        verify: { ok: verification.ok },
        modelRender: { ok: true, ...modelRender },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryPath, { force: true }),
      fs.rm(temporaryAuditPath, { force: true }),
    ]);
    throw error;
  }
}

function parseCli(argv) {
  const [
    inputPath,
    outputPath,
    auditPath,
    anchorText = "Decision: proceed with controlled rollout.",
    expectedCommentText = "Please confirm the final retention wording.",
    replacementText = "Approved after legal review.",
  ] = argv;
  return { inputPath, outputPath, auditPath, anchorText, expectedCommentText, replacementText };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editClassicComment(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    commentId: result.audit.validation.reimport.commentId,
  }));
}
