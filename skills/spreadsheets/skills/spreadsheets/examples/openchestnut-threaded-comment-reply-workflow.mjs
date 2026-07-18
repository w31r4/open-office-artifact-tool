import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
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
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function normalizedAddress(value) {
  return requiredText(value, "cell").toUpperCase();
}

function resolveThread(workbook, sheetName, address) {
  const match = workbook.comments.threads.filter((thread) => thread.target.sheetName === sheetName && thread.target.address === address);
  if (match.length !== 1) throw new Error(`Expected exactly one threaded comment at ${sheetName}!${address}; found ${match.length}.`);
  const thread = match[0];
  if (!thread.comments.length) throw new Error(`Threaded comment at ${sheetName}!${address} has no root comment.`);
  if (thread.comments.some((comment, index) => index > 0 && comment.parentId && comment.parentId !== thread.comments[0].id)) {
    throw new Error(`Threaded comment at ${sheetName}!${address} has nested or branched replies and is source-bound.`);
  }
  return thread;
}

async function renderAllSheets(workbook) {
  const sheets = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for sheet ${sheet.name} did not produce SVG.`);
    sheets.push({ sheet: sheet.name, bytes: preview.bytes.length, renderer: "model-svg" });
  }
  return sheets;
}

export async function replyAndResolveThreadedComment({ inputPath, outputPath, auditPath, sheetName, cell, reply, author = "Spreadsheet Agent" }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original workbook remains immutable.");
  const targetSheet = requiredText(sheetName, "sheetName");
  const targetCell = normalizedAddress(cell);
  const replyText = requiredText(reply, "reply");
  const replyAuthor = requiredText(author, "author");
  const source = await fs.readFile(sourcePath);
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const thread = resolveThread(workbook, targetSheet, targetCell);
  const rootId = thread.comments[0].id;
  const originalComments = thread.comments.map((comment) => ({
    id: comment.id,
    parentId: comment.parentId || null,
    personId: comment.personId,
    date: comment.date,
    text: comment.text,
  }));
  thread.addReply(replyText, { author: replyAuthor });
  thread.resolve();
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await exported.save(temporaryPath);
  let output;
  let reimported;
  let verification;
  let renders;
  try {
    output = await fs.readFile(temporaryPath);
    reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const roundTrip = resolveThread(reimported, targetSheet, targetCell);
    if (roundTrip.comments.length !== originalComments.length + 1) throw new Error("Threaded comment export did not retain the expected root/direct-reply topology.");
    const added = roundTrip.comments.at(-1);
    if (added.text !== replyText || added.parentId !== rootId || !roundTrip.resolved || !roundTrip.comments.every((comment) => comment.done)) {
      throw new Error("Threaded comment export did not retain the requested direct reply and resolved state.");
    }
    for (const [index, original] of originalComments.entries()) {
      const actual = roundTrip.comments[index];
      if (!actual || actual.id !== original.id || (actual.parentId || null) !== original.parentId || actual.personId !== original.personId || actual.date !== original.date || actual.text !== original.text) {
        throw new Error("Threaded comment export changed an existing comment identity, parent, date, or text.");
      }
    }
    verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed: ${verification.ndjson}`);
    renders = await renderAllSheets(reimported);
    await fs.rename(temporaryPath, finalPath);
  } catch (error) {
    await fs.rm(temporaryPath, { force: true });
    throw error;
  }
  const finalBytes = output || await fs.readFile(finalPath);
  const audit = {
    schema: "open-office-artifact-tool.xlsx-audit.v1",
    status: "succeeded",
    source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
    output: { path: finalPath, sha256: sha256(finalBytes), bytes: finalBytes.length },
    provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
    savePolicy: { strategy: "rewrite" },
    operation: { type: "threaded-comment-direct-reply-resolve", target: { sheet: targetSheet, cell: targetCell }, replySha256: sha256(Buffer.from(replyText, "utf8")) },
    warnings: [],
    validation: {
      reimport: { ok: true, commentCount: reimported.comments.threads.find((thread) => thread.target.sheetName === targetSheet && thread.target.address === targetCell)?.comments.length },
      verify: { ok: verification.ok },
      modelRender: { ok: true, sheets: renders },
    },
  };
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  await fs.writeFile(finalAuditPath, JSON.stringify(audit, null, 2));
  return { outputPath: finalPath, auditPath: finalAuditPath, audit };
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, sheetName = "Forecast", cell = "F19", reply = "Approved after sensitivity review", author = "Spreadsheet Agent"] = argv;
  return { inputPath, outputPath, auditPath, sheetName, cell, reply, author };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await replyAndResolveThreadedComment(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, comments: result.audit.validation.reimport.commentCount }));
}
