import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const require = createRequire(import.meta.url);
const TOPOLOGY_PARTS = new Set([
  "[Content_Types].xml",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
]);

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

function xmlAttributes(tag) {
  const attributes = Object.create(null);
  for (const match of tag.matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])([\s\S]*?)\2/g)) attributes[match[1]] = match[3];
  return attributes;
}

function unescapeXml(value) {
  return String(value).replace(/&(#x[0-9a-f]+|#\d+|amp|apos|gt|lt|quot);/gi, (_, entity) => {
    const lower = entity.toLowerCase();
    if (lower === "amp") return "&";
    if (lower === "apos") return "'";
    if (lower === "gt") return ">";
    if (lower === "lt") return "<";
    if (lower === "quot") return '"';
    const codePoint = lower.startsWith("#x") ? Number.parseInt(lower.slice(2), 16) : Number.parseInt(lower.slice(1), 10);
    return Number.isSafeInteger(codePoint) && codePoint >= 0 && codePoint <= 0x10ffff ? String.fromCodePoint(codePoint) : "&" + entity + ";";
  });
}

function slideNameFromXml(xml, partPath) {
  const commonData = xml.match(/<(?:[A-Za-z_][\w.-]*:)?cSld\b[^>]*>/i)?.[0];
  if (!commonData) throw new Error("Slide part " + partPath + " has no p:cSld element.");
  const name = xmlAttributes(commonData).name;
  if (typeof name !== "string" || !name) throw new Error("Slide part " + partPath + " has no explicit non-empty p:cSld/@name; this workflow refuses a fallback name.");
  return unescapeXml(name);
}

function resolveRelationshipTarget(target) {
  const resolved = new URL(target, "https://openchestnut.invalid/ppt/presentation.xml");
  if (resolved.origin !== "https://openchestnut.invalid") throw new Error("Unexpected PPTX relationship target origin.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) throw new Error("Unsafe PPTX slide relationship target: " + JSON.stringify(target));
  return partPath;
}

async function requiredZipBytes(zip, partPath) {
  const file = zip.file(partPath);
  if (!file) throw new Error("PPTX is missing required part " + partPath + ".");
  return file.async("uint8array");
}

function packagePartPaths(zip) {
  return Object.keys(zip.files).filter((partPath) => !zip.files[partPath].dir).sort();
}

async function orderedSlidePartPaths(zip) {
  const presentationXml = Buffer.from(await requiredZipBytes(zip, "ppt/presentation.xml")).toString("utf8");
  const relationshipsXml = Buffer.from(await requiredZipBytes(zip, "ppt/_rels/presentation.xml.rels")).toString("utf8");
  const relationships = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type?.endsWith("/slide")) continue;
    if (attributes.TargetMode?.toLowerCase() === "external" || !attributes.Target) {
      throw new Error("Presentation slide relationship " + JSON.stringify(attributes.Id) + " is not an internal SlidePart.");
    }
    relationships.set(attributes.Id, resolveRelationshipTarget(attributes.Target));
  }
  const paths = [];
  for (const match of presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)) {
    const relationshipId = xmlAttributes(match[0])["r:id"];
    const target = relationships.get(relationshipId);
    if (!target) throw new Error("Presentation slide list references an unresolved relationship " + JSON.stringify(relationshipId) + ".");
    if (!zip.file(target)) throw new Error("Presentation slide relationship points at missing part " + target + ".");
    paths.push(target);
  }
  if (!paths.length || new Set(paths).size !== paths.length) throw new Error("Presentation slide list must contain distinct, resolvable SlideParts.");
  return paths;
}

function relationshipPartPath(partPath) {
  return path.posix.join(path.posix.dirname(partPath), "_rels", path.posix.basename(partPath) + ".rels");
}

function canonicalCloneSnapshot(slide) {
  const proto = slide.toProto();
  const locationById = new Map();
  const collectIds = (value, location) => {
    if (Array.isArray(value)) {
      value.forEach((item, index) => collectIds(item, `${location}[${index}]`));
      return;
    }
    if (!value || typeof value !== "object") return;
    if (typeof value.id === "string") locationById.set(value.id, location);
    for (const [key, item] of Object.entries(value)) collectIds(item, `${location}.${key}`);
  };
  const normalize = (value) => {
    if (Array.isArray(value)) return value.map(normalize);
    if (!value || typeof value !== "object") return value;
    const result = {};
    for (const [key, item] of Object.entries(value)) {
      if (key === "id" || key === "layoutId") continue;
      if ((key === "startTargetId" || key === "endTargetId") && typeof item === "string") {
        result[key] = locationById.get(item) || "unresolved:" + item;
      } else {
        result[key] = normalize(item);
      }
    }
    return result;
  };
  collectIds(proto, "slide");
  return JSON.stringify(normalize(proto));
}

async function modelSvg(slide) {
  const svg = await slide.export({ format: "svg" });
  const text = await svg.text();
  if (!/<svg\b/i.test(text)) throw new Error("Presentation model render did not produce SVG.");
  // These attributes are inspect locators, not visual content. A duplicate
  // receives fresh model IDs, so compare the rendered scene without them.
  const visualSvg = text.replace(/\sdata-[\w-]*id="[^"]*"/gi, "");
  return {
    rawSha256: sha256(svg.bytes),
    visualSha256: sha256(Buffer.from(visualSvg)),
    bytes: svg.bytes.length,
  };
}

async function assertWorkflowScope(zip, sourcePart) {
  const relationshipPart = zip.file(relationshipPartPath(sourcePart));
  if (!relationshipPart) throw new Error("Source SlidePart " + sourcePart + " has no relationship part.");
  const relationshipsXml = await relationshipPart.async("text");
  const relationshipTypes = [...relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)]
    .map((match) => xmlAttributes(match[0]).Type || "");
  if (relationshipTypes.some((type) => type.endsWith("/notesSlide"))) {
    throw new Error("This duplicate workflow intentionally accepts no speaker-notes leaf; use an explicit broader closed-graph workflow for notes.");
  }
  if (relationshipTypes.some((type) => type.endsWith("/comments"))) {
    throw new Error("This duplicate workflow intentionally accepts no legacy-comments leaf; use an explicit broader closed-graph workflow for comments.");
  }
}

async function assertDuplicatePackageScope(sourceBytes, outputBytes, sourceIndex) {
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBytes), JSZip.loadAsync(outputBytes)]);
  const sourceParts = packagePartPaths(sourceZip);
  const outputParts = packagePartPaths(outputZip);
  const sourceSlideParts = await orderedSlidePartPaths(sourceZip);
  const outputSlideParts = await orderedSlidePartPaths(outputZip);
  if (!Number.isInteger(sourceIndex) || sourceIndex < 0 || sourceIndex >= sourceSlideParts.length) {
    throw new Error("Resolved source slide index is outside the source PPTX slide list.");
  }
  if (outputSlideParts.length !== sourceSlideParts.length + 1) throw new Error("PPTX duplicate did not create exactly one additional SlidePart.");
  const sourcePart = sourceSlideParts[sourceIndex];
  const clonePart = outputSlideParts[sourceIndex + 1];
  const expectedSlideOrder = [...sourceSlideParts];
  expectedSlideOrder.splice(sourceIndex + 1, 0, clonePart);
  if (JSON.stringify(outputSlideParts) !== JSON.stringify(expectedSlideOrder)) {
    throw new Error("PPTX duplicate changed source slide order or did not insert the clone immediately after its source.");
  }
  if (sourceZip.file(clonePart) || !/^ppt\/slides\/slide\d+\.xml$/i.test(clonePart)) {
    throw new Error("PPTX duplicate did not allocate a distinct canonical SlidePart path.");
  }
  const newParts = outputParts.filter((partPath) => !sourceZip.file(partPath));
  const expectedNewParts = [clonePart, relationshipPartPath(clonePart)].sort();
  if (JSON.stringify(newParts) !== JSON.stringify(expectedNewParts)) {
    throw new Error("PPTX duplicate created an unexpected package part outside the bare source-bound clone profile.");
  }
  for (const partPath of sourceParts) {
    if (TOPOLOGY_PARTS.has(partPath)) continue;
    const [sourcePartBytes, outputPartBytes] = await Promise.all([
      requiredZipBytes(sourceZip, partPath),
      requiredZipBytes(outputZip, partPath),
    ]);
    if (!Buffer.from(sourcePartBytes).equals(Buffer.from(outputPartBytes))) {
      throw new Error("PPTX duplicate changed retained source part " + partPath + ".");
    }
  }
  const [sourceSlideBytes, retainedSourceSlideBytes] = await Promise.all([
    requiredZipBytes(sourceZip, sourcePart),
    requiredZipBytes(outputZip, sourcePart),
  ]);
  if (!Buffer.from(sourceSlideBytes).equals(Buffer.from(retainedSourceSlideBytes))) {
    throw new Error("PPTX duplicate changed its retained source SlidePart.");
  }
  return {
    sourcePart,
    clonePart,
    sourcePartCount: sourceParts.length,
    outputPartCount: outputParts.length,
    newPartPaths: newParts,
    retainedSourcePartsByteIdentical: true,
  };
}

export async function duplicatePptxSlide({ inputPath, outputPath, auditPath, expectedName }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const sourceName = requiredText(expectedName, "expectedName");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const candidates = presentation.slides.items.filter((slide) => slide.name === sourceName);
  if (candidates.length !== 1) throw new Error("Expected exactly one imported source slide named " + JSON.stringify(sourceName) + "; found " + candidates.length + ".");
  const target = candidates[0];
  const sourceIndex = presentation.slides.items.indexOf(target);
  const sourceZip = await JSZip.loadAsync(source);
  const sourceSlideParts = await orderedSlidePartPaths(sourceZip);
  if (sourceSlideParts.length !== presentation.slides.count) throw new Error("PPTX package slide order does not match the imported presentation model.");
  const sourcePart = sourceSlideParts[sourceIndex];
  await assertWorkflowScope(sourceZip, sourcePart);
  if (slideNameFromXml(Buffer.from(await requiredZipBytes(sourceZip, sourcePart)).toString("utf8"), sourcePart) !== sourceName) {
    throw new Error("The selected model slide does not match its source SlidePart p:cSld/@name.");
  }
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const sourceSemantic = canonicalCloneSnapshot(target);
  const sourceRender = await modelSvg(target);
  const clone = target.duplicate();
  if (presentation.slides.items.indexOf(clone) !== sourceIndex + 1 || clone.name !== target.name) {
    throw new Error("slide.duplicate did not create an adjacent same-name pending clone.");
  }
  if (canonicalCloneSnapshot(clone) !== sourceSemantic) {
    throw new Error("Pending slide clone changed the canonical source semantics before export.");
  }

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertDuplicatePackageScope(source, output, sourceIndex);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const expectedNames = [...originalNames];
    expectedNames.splice(sourceIndex + 1, 0, sourceName);
    if (JSON.stringify(reimported.slides.items.map((slide) => slide.name)) !== JSON.stringify(expectedNames)) {
      throw new Error("PPTX export changed source names/order or did not retain the adjacent clone.");
    }
    const retained = reimported.slides.items[sourceIndex];
    const roundTripClone = reimported.slides.items[sourceIndex + 1];
    if (canonicalCloneSnapshot(retained) !== sourceSemantic || canonicalCloneSnapshot(roundTripClone) !== sourceSemantic) {
      throw new Error("PPTX duplicate did not preserve source and clone semantic structure after reimport.");
    }
    const cloneRender = await modelSvg(roundTripClone);
    if (cloneRender.visualSha256 !== sourceRender.visualSha256) throw new Error("PPTX duplicate changed the clone model SVG render.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-slide-duplicate",
        expectedName: sourceName,
        sourceIndex,
        cloneIndex: sourceIndex + 1,
        sourcePart: packageScope.sourcePart,
        clonePart: packageScope.clonePart,
        scope: "canonical-inline-leaves-without-notes-or-comments",
      },
      warnings: [],
      validation: {
        package: {
          ok: true,
          sourcePartCount: packageScope.sourcePartCount,
          outputPartCount: packageScope.outputPartCount,
          sourcePart: packageScope.sourcePart,
          clonePart: packageScope.clonePart,
          newPartPaths: packageScope.newPartPaths,
          retainedSourcePartsByteIdentical: packageScope.retainedSourcePartsByteIdentical,
          cloneInsertedAdjacent: true,
        },
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          sourceAndCloneSemanticsEqual: true,
          sourceAndCloneNames: [retained.name, roundTripClone.name],
        },
        modelRender: {
          ok: true,
          renderer: "model-svg",
          sourceRawSha256: sourceRender.rawSha256,
          cloneRawSha256: cloneRender.rawSha256,
          visualSha256: sourceRender.visualSha256,
          identityAttributesIgnored: true,
          visualEquivalent: true,
        },
        verify: { ok: verification.ok },
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
    expectedName = "Clone source",
  ] = argv;
  return { inputPath, outputPath, auditPath, expectedName };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await duplicatePptxSlide(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    sourcePart: result.audit.operation.sourcePart,
    clonePart: result.audit.operation.clonePart,
  }));
}
