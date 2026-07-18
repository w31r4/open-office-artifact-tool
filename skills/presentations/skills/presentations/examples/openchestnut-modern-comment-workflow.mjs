import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, Presentation, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredPath(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty path.");
  return path.resolve(value);
}

function oneModernThread(presentation) {
  const threads = presentation.slides.items.flatMap((slide) => slide.comments.items.map((thread) => ({ slide, thread })));
  if (threads.length !== 1 || threads[0].thread.nativeFormat !== "modern") {
    throw new Error(`Expected exactly one bounded modern comment thread; found ${threads.length}.`);
  }
  return threads[0];
}

async function modelRender(slide) {
  const preview = await slide.export({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Presentation model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function createAndEditModernCommentThread({ outputPath, auditPath }) {
  const finalPath = requiredPath(outputPath, "outputPath");
  const finalAuditPath = requiredPath(auditPath, "auditPath");
  if (finalPath === finalAuditPath) throw new Error("outputPath and auditPath must be distinct.");

  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
    commentFormat: "modern",
  });
  const slide = presentation.slides.add({ name: "Decision review" });
  const target = slide.shapes.add({
    id: "decision-evidence",
    name: "Decision evidence",
    geometry: "roundRect",
    position: { left: 96, top: 96, width: 620, height: 132 },
    fill: "#E0F2FE",
    text: "Customer evidence is ready",
  });
  const thread = slide.comments.addThread({
    textMatch: { element: target, query: "Customer evidence", occurrence: 0 },
  }, "Confirm the customer evidence before delivery.", {
    id: "{11111111-1111-4111-8111-111111111111}",
    nativeFormat: "modern",
    position: { x: 1_234_500, y: 2_345_600, unit: "emu" },
    comments: [{
      nativeId: "{11111111-1111-4111-8111-111111111111}",
      author: "Review Owner",
      person: {
        id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
        name: "Review Owner",
        initials: "RO",
        userId: "review.owner@example.test",
        providerId: "None",
      },
      text: "Confirm the customer evidence before delivery.",
      created: "2026-07-19T02:55:00Z",
      status: "active",
    }],
  });
  thread.addReply("Evidence is attached for review.", {
    nativeId: "{22222222-2222-4222-8222-222222222222}",
    author: "Evidence Owner",
    person: {
      id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
      name: "Evidence Owner",
      initials: "EO",
      userId: "evidence.owner@example.test",
      providerId: "None",
    },
    created: "2026-07-19T03:05:00Z",
    status: "active",
  });

  const sourceBlob = await PresentationFile.exportPptx(presentation);
  const source = sourceBlob.bytes;
  const imported = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: "modern-comment-source.pptx" }));
  const importedEntry = oneModernThread(imported);
  const sourceIdentity = {
    threadId: importedEntry.thread.id,
    targetId: importedEntry.thread.targetId,
    nativeCommentIds: importedEntry.thread.comments.map((comment) => comment.nativeId),
    authors: importedEntry.thread.comments.map((comment) => comment.author),
    created: importedEntry.thread.comments.map((comment) => comment.created),
    position: importedEntry.thread.position,
    anchor: importedEntry.thread.nativeAnchor,
  };
  importedEntry.thread.comments[0].text = "Customer evidence confirmed for delivery.";
  importedEntry.thread.comments[1].text = "Recorded in the decision log.";
  importedEntry.thread.resolve();

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const outputBlob = await PresentationFile.exportPptx(imported);
    await outputBlob.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const roundTrip = oneModernThread(reimported);
    const roundTripIdentity = {
      threadId: roundTrip.thread.id,
      targetId: roundTrip.thread.targetId,
      nativeCommentIds: roundTrip.thread.comments.map((comment) => comment.nativeId),
      authors: roundTrip.thread.comments.map((comment) => comment.author),
      created: roundTrip.thread.comments.map((comment) => comment.created),
      position: roundTrip.thread.position,
      anchor: roundTrip.thread.nativeAnchor,
    };
    if (JSON.stringify(roundTripIdentity) !== JSON.stringify(sourceIdentity)) {
      throw new Error("Modern comment export changed fixed thread identity, author/date metadata, anchor, position, or topology.");
    }
    if (roundTrip.thread.comments[0].text !== "Customer evidence confirmed for delivery." ||
        roundTrip.thread.comments[1].text !== "Recorded in the decision log." ||
        !roundTrip.thread.resolved) {
      throw new Error("Modern comment text/status edits did not survive second import.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    const packageInspection = await PresentationFile.inspectPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }), { includeText: true });
    if (!packageInspection.ok) throw new Error("PPTX package inspection failed: " + packageInspection.ndjson);
    const render = await modelRender(roundTrip.slide);
    const audit = {
      schema: "open-office-artifact-tool.pptx-modern-comments-audit.v1",
      status: "succeeded",
      source: { sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "fixed-topology-modern-comment-text-status-edit",
        threadId: roundTrip.thread.id,
        targetId: roundTrip.thread.targetId,
        replyCount: roundTrip.thread.comments.length - 1,
      },
      warnings: [],
      validation: {
        fixedIdentityPreserved: true,
        rootTextExact: true,
        replyTextExact: true,
        resolved: true,
        package: { ok: packageInspection.ok },
        verify: { ok: verification.ok },
        modelRender: { ok: true, ...render },
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
  const [outputPath, auditPath] = process.argv.slice(2);
  const result = await createAndEditModernCommentThread({ outputPath, auditPath });
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    threadId: result.audit.operation.threadId,
  }));
}
