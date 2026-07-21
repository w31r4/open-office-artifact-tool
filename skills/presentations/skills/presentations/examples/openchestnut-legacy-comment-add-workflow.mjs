import { constants as fsConstants } from "node:fs";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const EMU_PER_PIXEL = 9_525;
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value.trim();
}

function finiteCoordinate(value, label) {
  const number = Number(value);
  if (!Number.isFinite(number)) throw new TypeError(label + " must be a finite number.");
  return number;
}

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of tag.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function relationshipPartPath(partPath) {
  const slash = partPath.lastIndexOf("/");
  const directory = slash < 0 ? "" : partPath.slice(0, slash + 1);
  const name = slash < 0 ? partPath : partPath.slice(slash + 1);
  return directory + "_rels/" + name + ".rels";
}

function resolveRelationshipTarget(sourcePart, target) {
  const resolved = new URL(target, "https://openchestnut.invalid/" + sourcePart);
  if (resolved.origin !== "https://openchestnut.invalid") throw new Error("Unexpected external PPTX relationship target.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) throw new Error("Unsafe PPTX relationship target.");
  return partPath;
}

async function relationshipEntries(zip, sourcePart) {
  const xml = await zip.file(relationshipPartPath(sourcePart))?.async("text");
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*>/gi)].map((match) => {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type || !attributes.Target) throw new Error("Malformed relationship for " + sourcePart + ".");
    const external = attributes.TargetMode?.toLowerCase() === "external";
    return {
      id: attributes.Id,
      type: attributes.Type,
      external,
      targetPart: external ? null : resolveRelationshipTarget(sourcePart, attributes.Target),
    };
  });
}

function relationshipType(entry, suffix) {
  return entry.type.endsWith("/" + suffix);
}

function exactlyOne(entries, suffix, label) {
  const matches = entries.filter((entry) => relationshipType(entry, suffix));
  if (matches.length !== 1 || matches[0].external || !matches[0].targetPart) {
    throw new Error(label + " must own exactly one internal " + suffix + " relationship.");
  }
  return matches[0];
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  if (!presentationXml) throw new Error("PPTX is missing ppt/presentation.xml.");
  const relationships = await relationshipEntries(zip, "ppt/presentation.xml");
  const slideById = new Map(relationships
    .filter((entry) => relationshipType(entry, "slide") && !entry.external)
    .map((entry) => [entry.id, entry.targetPart]));
  const paths = [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)].map((match) => {
    const partPath = slideById.get(xmlAttributes(match[0])["r:id"]);
    if (!partPath || !zip.file(partPath)) throw new Error("Presentation slide list contains an unresolved SlidePart.");
    return partPath;
  });
  if (!paths.length || new Set(paths).size !== paths.length) throw new Error("Presentation slide list must contain distinct SlideParts.");
  return paths;
}

async function entryHashes(zip) {
  const output = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) output.set(name, sha256(await entry.async("uint8array")));
  }
  return output;
}

async function auditAddedLegacyCommentGraph(sourceBytes, outputBytes, targetIndex) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourceSlides = await orderedSlidePartPaths(sourceZip);
  const outputSlides = await orderedSlidePartPaths(outputZip);
  if (JSON.stringify(sourceSlides) !== JSON.stringify(outputSlides)) throw new Error("Adding a review comment changed slide package identity or order.");
  const targetPart = sourceSlides[targetIndex];
  if (!targetPart) throw new Error("Target slide index is outside the source package slide list.");

  const sourcePresentationRelationships = await relationshipEntries(sourceZip, "ppt/presentation.xml");
  if (sourcePresentationRelationships.some((entry) => relationshipType(entry, "commentAuthors"))) {
    throw new Error("Source presentation already has a legacy comment author catalog.");
  }
  for (const slidePart of sourceSlides) {
    if ((await relationshipEntries(sourceZip, slidePart)).some((entry) => relationshipType(entry, "comments"))) {
      throw new Error("Source presentation already has a legacy SlideCommentsPart.");
    }
  }

  const outputPresentationRelationships = await relationshipEntries(outputZip, "ppt/presentation.xml");
  const authorsRelationship = exactlyOne(outputPresentationRelationships, "commentAuthors", "PresentationPart");
  if (authorsRelationship.targetPart !== "ppt/commentAuthors.xml" || !outputZip.file(authorsRelationship.targetPart)) {
    throw new Error("Legacy author relationship does not resolve to ppt/commentAuthors.xml.");
  }
  if ((await relationshipEntries(outputZip, authorsRelationship.targetPart)).length !== 0) {
    throw new Error("Legacy CommentAuthorsPart must be a closed leaf.");
  }

  let commentsPart;
  for (let index = 0; index < outputSlides.length; index += 1) {
    const comments = (await relationshipEntries(outputZip, outputSlides[index])).filter((entry) => relationshipType(entry, "comments"));
    if (index === targetIndex) commentsPart = exactlyOne(comments, "comments", "Target SlidePart").targetPart;
    else if (comments.length) throw new Error("A non-target slide unexpectedly gained a comments relationship.");
  }
  if (!commentsPart || !/^ppt\/comments\/comment\d+\.xml$/.test(commentsPart) || !outputZip.file(commentsPart)) {
    throw new Error("Target comments relationship does not resolve to a numbered SlideCommentsPart.");
  }
  if ((await relationshipEntries(outputZip, commentsPart)).length !== 0) throw new Error("SlideCommentsPart must be a closed leaf.");

  const before = await entryHashes(sourceZip);
  const after = await entryHashes(outputZip);
  const added = [...after.keys()].filter((name) => !before.has(name)).sort();
  const expectedAdded = ["ppt/commentAuthors.xml", commentsPart].sort();
  if (JSON.stringify(added) !== JSON.stringify(expectedAdded)) {
    throw new Error("Legacy comment transaction added unexpected package parts: " + added.join(", "));
  }
  const changed = [...before.keys()].filter((name) => after.has(name) && before.get(name) !== after.get(name)).sort();
  const expectedChanged = [
    "[Content_Types].xml",
    "ppt/_rels/presentation.xml.rels",
    relationshipPartPath(targetPart),
  ].sort();
  if (JSON.stringify(changed) !== JSON.stringify(expectedChanged)) {
    throw new Error("Legacy comment transaction changed unexpected package parts: " + changed.join(", "));
  }
  return { ok: true, targetSlidePart: targetPart, authorsPart: authorsRelationship.targetPart, commentsPart, addedParts: added, changedParts: changed };
}

function visibleSlideSnapshot(slide) {
  const value = structuredClone(slide.toProto());
  delete value.comments;
  return JSON.stringify(value);
}

async function modelSvg(slide) {
  const rendered = await slide.export({ format: "svg" });
  if (!/<svg\b/i.test(await rendered.text())) throw new Error("Presentation model render did not produce SVG.");
  return { bytes: rendered.bytes.length, sha256: sha256(rendered.bytes) };
}

async function copyExclusive(source, destination) {
  await fs.copyFile(source, destination, fsConstants.COPYFILE_EXCL);
}

export async function addPptxLegacyReviewComment({ inputPath, outputPath, auditPath, slideName, text, author, created, position }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (new Set([sourcePath, finalPath, finalAuditPath]).size !== 3) throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  const targetSlideName = requiredText(slideName, "slideName");
  const commentText = requiredText(text, "text");
  const commentAuthor = requiredText(author, "author");
  const createdAt = requiredText(created, "created");
  if (Number.isNaN(Date.parse(createdAt))) throw new TypeError("created must be an ISO-8601 timestamp.");
  const commentPosition = {
    x: finiteCoordinate(position?.x, "position.x"),
    y: finiteCoordinate(position?.y, "position.y"),
    unit: position?.unit || "px",
  };
  if (!new Set(["px", "emu"]).has(commentPosition.unit)) throw new TypeError("position.unit must be px or emu.");

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const candidates = presentation.slides.items.filter((slide) => slide.name === targetSlideName);
  if (candidates.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(targetSlideName) + "; found " + candidates.length + ".");
  const target = candidates[0];
  const targetIndex = presentation.slides.items.indexOf(target);
  const capability = target.comments.capability;
  if (!capability.sourceBound || capability.format !== "legacy" || capability.partPresent || !capability.addable || target.comments.items.length) {
    throw new Error("Selected slide is not in an imported, comment-free presentation whose package graph explicitly permits legacy review-comment creation.");
  }
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const visibleSnapshot = visibleSlideSnapshot(target);
  const sourceRender = await modelSvg(target);
  const thread = target.comments.addThread(undefined, commentText, {
    author: commentAuthor,
    created: createdAt,
    position: commentPosition,
  });
  if (presentation.resolve(thread.id) !== thread) throw new Error("New review comment cannot be resolved by its model ID.");

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  let publishedOutput = false;
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageGraph = await auditAddedLegacyCommentGraph(source, output, targetIndex);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (JSON.stringify(reimported.slides.items.map((slide) => slide.name)) !== JSON.stringify(originalNames)) throw new Error("PPTX export changed slide count, order, or names.");
    const roundTrip = reimported.slides.items[targetIndex];
    if (visibleSlideSnapshot(roundTrip) !== visibleSnapshot) throw new Error("Adding a nonvisual review comment changed visible target-slide semantics.");
    const outputThread = roundTrip.comments.items[0];
    const expectedXEmu = Math.round(commentPosition.x * (commentPosition.unit === "px" ? EMU_PER_PIXEL : 1));
    const expectedYEmu = Math.round(commentPosition.y * (commentPosition.unit === "px" ? EMU_PER_PIXEL : 1));
    if (roundTrip.comments.items.length !== 1 || outputThread.comments.length !== 1 ||
        outputThread.comments[0].text !== commentText || outputThread.comments[0].author !== commentAuthor ||
        Date.parse(outputThread.comments[0].created) !== Date.parse(createdAt) ||
        outputThread.nativeAnchor?.positionXEmu !== expectedXEmu ||
        outputThread.nativeAnchor?.positionYEmu !== expectedYEmu) {
      throw new Error("Reimported legacy comment does not exactly match the requested review annotation.");
    }
    const outputCapability = roundTrip.comments.capability;
    if (!outputCapability.sourceBound || outputCapability.format !== "legacy" || !outputCapability.partPresent || outputCapability.addable) {
      throw new Error("Reimported comment capability does not reflect the newly created source-bound comments part.");
    }
    const outputRender = await modelSvg(roundTrip);
    if (outputRender.sha256 !== sourceRender.sha256) throw new Error("Adding a nonvisual review comment changed the target slide model SVG.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source presentation changed during the transaction.");

    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", overwrite: false },
      operation: { type: "source-bound-legacy-comment-add", slideId: target.id, slideName: target.name, slideIndex: targetIndex, threadId: thread.id },
      precondition: { capability },
      validation: {
        package: packageGraph,
        reimport: { ok: true, slideCount: reimported.slides.count, visibleSemanticsPreserved: true, commentExact: true, capability: outputCapability },
        modelRender: { ok: true, renderer: "model-svg", sourceSha256: sourceRender.sha256, outputSha256: outputRender.sha256, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await copyExclusive(temporaryPath, finalPath);
    publishedOutput = true;
    await copyExclusive(temporaryAuditPath, finalAuditPath);
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    if (publishedOutput) await fs.rm(finalPath, { force: true });
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, slideName = "Imported review target", text = "Confirm the imported evidence.", author = "Review Owner", created = "2026-07-20T03:04:05Z", x = "360", y = "240"] = argv;
  return { inputPath, outputPath, auditPath, slideName, text, author, created, position: { x: Number(x), y: Number(y), unit: "px" } };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await addPptxLegacyReviewComment(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256, commentsPart: result.audit.validation.package.commentsPart }));
}
