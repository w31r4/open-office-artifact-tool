import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

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

async function orderedSlidePartPaths(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  const relationshipsXml = await zip.file("ppt/_rels/presentation.xml.rels")?.async("text");
  if (!presentationXml || !relationshipsXml) throw new Error("PPTX is missing presentation.xml or its relationship part.");
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

function withoutIds(value) {
  if (Array.isArray(value)) return value.map(withoutIds);
  if (!value || typeof value !== "object") return value;
  return Object.fromEntries(Object.entries(value)
    .filter(([key]) => key !== "id")
    .map(([key, item]) => [key, withoutIds(item)]));
}

function semanticSnapshot(slide) {
  return JSON.stringify(withoutIds(slide.toProto()));
}

async function modelSvg(slide) {
  const svg = await slide.export({ format: "svg" });
  const bytes = svg.bytes;
  if (!/<svg\b/i.test(await svg.text())) throw new Error("Presentation model render did not produce SVG.");
  return { bytes, sha256: sha256(bytes) };
}

async function assertPackageScope(sourceBytes, outputBytes, targetIndex, expectedName, replacementName) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourcePaths = Object.keys(sourceZip.files).sort();
  const outputPaths = Object.keys(outputZip.files).sort();
  if (JSON.stringify(outputPaths) !== JSON.stringify(sourcePaths)) throw new Error("PPTX rename changed the package part topology.");
  const slideParts = await orderedSlidePartPaths(sourceZip);
  if (!Number.isInteger(targetIndex) || targetIndex < 0 || targetIndex >= slideParts.length) throw new Error("Resolved target slide index is outside the source PPTX slide list.");
  const targetPart = slideParts[targetIndex];
  const sourceSlideXml = await sourceZip.file(targetPart).async("text");
  const outputSlideXml = await outputZip.file(targetPart).async("text");
  if (slideNameFromXml(sourceSlideXml, targetPart) !== expectedName) throw new Error("The selected source SlidePart no longer has the expected p:cSld/@name.");
  if (slideNameFromXml(outputSlideXml, targetPart) !== replacementName) throw new Error("Saved SlidePart does not have the requested p:cSld/@name.");
  for (const partPath of sourcePaths) {
    const sourceEntry = sourceZip.files[partPath];
    if (sourceEntry.dir || partPath === targetPart) continue;
    const sourcePart = await sourceZip.file(partPath).async("uint8array");
    const outputPart = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(sourcePart).equals(Buffer.from(outputPart))) {
      throw new Error("PPTX rename changed non-target part " + partPath + ".");
    }
  }
  return { targetPart, partCount: sourcePaths.length };
}

export async function editPptxSlideName({ inputPath, outputPath, auditPath, expectedName, replacementName }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const originalName = requiredText(expectedName, "expectedName");
  const nextName = requiredText(replacementName, "replacementName");
  if (originalName === nextName) throw new Error("replacementName must differ from expectedName.");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const targetCandidates = presentation.slides.items.filter((slide) => slide.name === originalName);
  if (targetCandidates.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(originalName) + "; found " + targetCandidates.length + ".");
  const target = targetCandidates[0];
  const targetIndex = presentation.slides.items.indexOf(target);
  const sourceZip = await JSZip.loadAsync(source);
  const sourceSlideParts = await orderedSlidePartPaths(sourceZip);
  if (sourceSlideParts.length !== presentation.slides.count) throw new Error("PPTX package slide order does not match the imported presentation model.");
  const sourcePart = sourceSlideParts[targetIndex];
  if (slideNameFromXml(await sourceZip.file(sourcePart).async("text"), sourcePart) !== originalName) {
    throw new Error("The selected model slide does not match its source SlidePart p:cSld/@name.");
  }
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const expectedNames = [...originalNames];
  expectedNames[targetIndex] = nextName;
  const sourceSemantics = semanticSnapshot(target);
  const sourceRender = await modelSvg(target);
  target.name = nextName;

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, targetIndex, originalName, nextName);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const reimportedNames = reimported.slides.items.map((slide) => slide.name);
    if (JSON.stringify(reimportedNames) !== JSON.stringify(expectedNames)) throw new Error("PPTX export changed slide count/order or an unexpected slide name.");
    const roundTrip = reimported.slides.items[targetIndex];
    if (semanticSnapshot(roundTrip) !== sourceSemantics) throw new Error("PPTX rename changed target slide semantics beyond its name.");
    const outputRender = await modelSvg(roundTrip);
    if (outputRender.sha256 !== sourceRender.sha256) throw new Error("PPTX rename changed the target slide's model SVG render.");
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
        type: "source-bound-slide-name-edit",
        sourcePart,
        sourceIndex: targetIndex,
        expectedName: originalName,
        replacementName: nextName,
        nativeAttribute: "p:cSld/@name",
      },
      warnings: [],
      validation: {
        package: {
          ok: true,
          partCount: packageScope.partCount,
          targetPart: packageScope.targetPart,
          targetNameVerified: true,
          targetPartMayBeCanonicalized: true,
          nonTargetPartsByteIdentical: true,
        },
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          expectedSlideNames: expectedNames,
          targetSemanticsPreserved: true,
        },
        modelRender: {
          ok: true,
          renderer: "model-svg",
          sourceSha256: sourceRender.sha256,
          outputSha256: outputRender.sha256,
          byteIdentical: true,
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
    expectedName = "Go-no-go decision",
    replacementName = "Go decision: controlled rollout",
  ] = argv;
  return { inputPath, outputPath, auditPath, expectedName, replacementName };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxSlideName(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    sourcePart: result.audit.operation.sourcePart,
  }));
}
