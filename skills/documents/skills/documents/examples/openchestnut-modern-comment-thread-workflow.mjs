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

function identity(comment) {
  return {
    id: comment.id,
    targetId: comment.targetId,
    parentId: comment.parentId,
    paraId: comment.paraId,
    durableId: comment.durableId,
    author: comment.author,
    initials: comment.initials,
    date: comment.date,
    dateUtc: comment.dateUtc,
    person: comment.person,
    intelligentPlaceholder: comment.intelligentPlaceholder,
  };
}

function resolveThread(document, anchorText, expectedRootText, expectedReplyText) {
  const targets = document.blocks.filter((block) => block.kind === "paragraph" && block.text.includes(anchorText));
  if (targets.length !== 1) throw new Error(`Expected exactly one paragraph containing the requested anchor; found ${targets.length}.`);
  const comments = document.comments.filter((comment) => comment.targetId === targets[0].id);
  const roots = comments.filter((comment) => !comment.parentId);
  if (roots.length !== 1) throw new Error(`Expected exactly one root comment; found ${roots.length}.`);
  const replies = comments.filter((comment) => comment.parentId === roots[0].id);
  if (replies.length !== 1 || comments.length !== 2) throw new Error("Expected one bounded root plus one direct reply and no other comment topology.");
  if (!roots[0].paraId || !replies[0].paraId) throw new Error("The selected thread has no verified commentsExtended paragraph identities.");
  if (roots[0].text !== expectedRootText || replies[0].text !== expectedReplyText) {
    throw new Error("The selected modern thread does not match both expected source texts; refusing an ambiguous edit.");
  }
  return { target: targets[0], root: roots[0], reply: replies[0] };
}

async function renderModel(document) {
  const preview = await document.render({ format: "svg" });
  if (!/<svg\b/i.test(await preview.text())) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function editModernCommentThread({
  inputPath,
  outputPath,
  auditPath,
  anchorText,
  expectedRootText,
  replacementRootText,
  expectedReplyText,
  replacementReplyText,
  resolved = true,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the DOCX paths.");
  const requestedAnchor = requiredText(anchorText, "anchorText");
  const originalRoot = requiredText(expectedRootText, "expectedRootText");
  const nextRoot = requiredText(replacementRootText, "replacementRootText");
  const originalReply = requiredText(expectedReplyText, "expectedReplyText");
  const nextReply = requiredText(replacementReplyText, "replacementReplyText");

  const source = await fs.readFile(sourcePath);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const thread = resolveThread(document, requestedAnchor, originalRoot, originalReply);
  const sourceTargetText = thread.target.text;
  const rootIdentity = identity(thread.root);
  const replyIdentity = identity(thread.reply);
  thread.root.text = nextRoot;
  thread.reply.text = nextReply;
  if (resolved) thread.root.resolve();
  else thread.root.reopen();

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await DocumentFile.exportDocx(document);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = resolveThread(reimported, requestedAnchor, nextRoot, nextReply);
    if (roundTrip.target.text !== sourceTargetText || JSON.stringify(identity(roundTrip.root)) !== JSON.stringify(rootIdentity) ||
        JSON.stringify(identity(roundTrip.reply)) !== JSON.stringify(replyIdentity)) {
      throw new Error("DOCX export changed the modern thread identity, people metadata, timestamp, or target anchor.");
    }
    if (roundTrip.root.resolved !== Boolean(resolved)) throw new Error("Requested modern comment resolved state did not survive re-import.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const modelRender = await renderModel(reimported);
    const audit = {
      schema: "open-office-artifact-tool.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "modern-comment-thread-text-and-resolved-edit",
        anchorSha256: sha256(Buffer.from(requestedAnchor, "utf8")),
        rootId: rootIdentity.id,
        replyId: replyIdentity.id,
        resolved: Boolean(resolved),
      },
      warnings: [],
      validation: {
        reimport: { ok: true, commentCount: reimported.comments.length, rootId: roundTrip.root.id, replyId: roundTrip.reply.id },
        verify: { ok: verification.ok },
        modelRender: { ok: true, ...modelRender },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [inputPath, outputPath, auditPath, anchorText, expectedRootText, replacementRootText, expectedReplyText, replacementReplyText, state = "resolved"] = process.argv.slice(2);
  const result = await editModernCommentThread({
    inputPath, outputPath, auditPath, anchorText, expectedRootText, replacementRootText, expectedReplyText, replacementReplyText,
    resolved: state !== "open",
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
