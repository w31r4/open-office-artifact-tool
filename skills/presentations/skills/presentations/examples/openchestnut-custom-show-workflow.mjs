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
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
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

function customShowsFromXml(xml) {
  const list = xml.match(/<(?:[A-Za-z_][\w.-]*:)?custShowLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?custShowLst>/i);
  if (!list) return [];
  const shows = [];
  for (const match of list[1].matchAll(/<((?:[A-Za-z_][\w.-]*:)?custShow)\b([^>]*)>([\s\S]*?)<\/\1>/gi)) {
    const attributes = xmlAttributes("<x " + match[2] + ">");
    const slideList = match[3].match(/<(?:[A-Za-z_][\w.-]*:)?sldLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sldLst>/i);
    if (!attributes.name || !/^\d+$/.test(attributes.id || "") || !slideList) throw new Error("PPTX custom-show XML is outside the canonical bounded profile.");
    const relationshipIds = [...slideList[1].matchAll(/<(?:[A-Za-z_][\w.-]*:)?sld\b[^>]*>/gi)].map((entry) => xmlAttributes(entry[0])["r:id"]);
    if (!relationshipIds.length || relationshipIds.some((id) => !id)) throw new Error("PPTX custom show has an empty or unresolved slide list.");
    shows.push({ name: unescapeXml(attributes.name), nativeId: Number(attributes.id), relationshipIds });
  }
  if (!shows.length) throw new Error("PPTX custom-show list has no canonical shows.");
  return shows;
}

function orderedSlideRelationshipIds(xml) {
  const list = xml.match(/<(?:[A-Za-z_][\w.-]*:)?sldIdLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sldIdLst>/i);
  if (!list) throw new Error("PPTX presentation has no slide ID list.");
  const ids = [...list[1].matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*>/gi)].map((entry) => xmlAttributes(entry[0])["r:id"]);
  if (!ids.length || ids.some((id) => !id) || new Set(ids).size !== ids.length) throw new Error("PPTX slide list has missing or duplicate relationship IDs.");
  return ids;
}

async function slideRenderHashes(presentation) {
  return Promise.all(presentation.slides.items.map(async (slide) => {
    const rendered = await slide.export({ format: "svg" });
    if (!/<svg\b/i.test(await rendered.text())) throw new Error("Presentation model render did not produce SVG.");
    return sha256(rendered.bytes);
  }));
}

async function assertPackageScope(sourceBytes, outputBytes, targetIndex, expected) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourcePaths = Object.keys(sourceZip.files).sort();
  const outputPaths = Object.keys(outputZip.files).sort();
  if (JSON.stringify(sourcePaths) !== JSON.stringify(outputPaths)) throw new Error("Custom-show edit changed package part topology.");
  for (const partPath of sourcePaths) {
    if (sourceZip.files[partPath].dir || partPath === "ppt/presentation.xml") continue;
    const before = await sourceZip.file(partPath).async("uint8array");
    const after = await outputZip.file(partPath).async("uint8array");
    if (!Buffer.from(before).equals(Buffer.from(after))) throw new Error("Custom-show edit changed non-target part " + partPath + ".");
  }
  const sourceXml = await sourceZip.file("ppt/presentation.xml")?.async("text");
  const outputXml = await outputZip.file("ppt/presentation.xml")?.async("text");
  if (!sourceXml || !outputXml) throw new Error("PPTX is missing ppt/presentation.xml.");
  const sourceShows = customShowsFromXml(sourceXml);
  const outputShows = customShowsFromXml(outputXml);
  if (sourceShows.length !== outputShows.length || targetIndex < 0 || targetIndex >= sourceShows.length) throw new Error("Custom-show edit changed show topology.");
  for (let index = 0; index < sourceShows.length; index += 1) {
    if (sourceShows[index].nativeId !== outputShows[index].nativeId) throw new Error("Custom-show edit changed native identity.");
    if (index !== targetIndex && JSON.stringify(sourceShows[index]) !== JSON.stringify(outputShows[index])) throw new Error("Custom-show edit changed a non-target show.");
  }
  if (sourceShows[targetIndex].name !== expected.sourceName || outputShows[targetIndex].name !== expected.outputName) throw new Error("Custom-show edit did not apply the requested exact name change.");
  if (JSON.stringify(outputShows[targetIndex].relationshipIds) !== JSON.stringify(expected.relationshipIds)) throw new Error("Custom-show edit did not apply the requested ordered slide membership.");
  return { partCount: sourcePaths.length, nativeId: outputShows[targetIndex].nativeId };
}

export async function editPptxCustomShow({ inputPath, outputPath, auditPath, expectedName, replacementName, orderedSlideNames }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const sourceName = requiredText(expectedName, "expectedName");
  const outputName = requiredText(replacementName, "replacementName");
  if (sourceName === outputName) throw new Error("replacementName must differ from expectedName.");
  if (!Array.isArray(orderedSlideNames) || !orderedSlideNames.length) throw new Error("orderedSlideNames must contain at least one exact slide name.");
  const requestedSlideNames = orderedSlideNames.map((name, index) => requiredText(name, `orderedSlideNames[${index}]`));

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const showCandidates = presentation.customShows.items.filter((show) => show.name === sourceName);
  if (showCandidates.length !== 1) throw new Error("Expected exactly one imported custom show named " + JSON.stringify(sourceName) + "; found " + showCandidates.length + ".");
  const target = showCandidates[0];
  const targetIndex = presentation.customShows.items.indexOf(target);
  const selectedSlides = requestedSlideNames.map((name) => {
    const matches = presentation.slides.items.filter((slide) => slide.name === name);
    if (matches.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(name) + "; found " + matches.length + ".");
    return matches[0];
  });
  const sourceZip = await JSZip.loadAsync(source);
  const sourcePresentationXml = await sourceZip.file("ppt/presentation.xml")?.async("text");
  if (!sourcePresentationXml) throw new Error("PPTX is missing ppt/presentation.xml.");
  const relationshipIdBySlideId = new Map(orderedSlideRelationshipIds(sourcePresentationXml).map((relationshipId, index) => [presentation.slides.items[index].id, relationshipId]));
  const requestedRelationshipIds = selectedSlides.map((slide) => relationshipIdBySlideId.get(slide.id));
  if (requestedRelationshipIds.some((relationshipId) => !relationshipId)) throw new Error("Requested custom-show slide does not have a source presentation relationship.");
  const sourceRenderHashes = await slideRenderHashes(presentation);
  const sourceShowCount = presentation.customShows.count;
  const sourceNativeId = target.nativeId;
  target.name = outputName;
  target.setSlides(selectedSlides);

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, targetIndex, {
      sourceName,
      outputName,
      relationshipIds: requestedRelationshipIds,
    });
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (reimported.customShows.count !== sourceShowCount) throw new Error("PPTX export changed custom-show count.");
    const roundTrip = reimported.customShows.getItem(outputName);
    if (!roundTrip || roundTrip.nativeId !== sourceNativeId || JSON.stringify(roundTrip.slides.map((slide) => slide.name)) !== JSON.stringify(requestedSlideNames)) {
      throw new Error("PPTX custom show did not survive second import with the requested semantics.");
    }
    const outputRenderHashes = await slideRenderHashes(reimported);
    if (JSON.stringify(outputRenderHashes) !== JSON.stringify(sourceRenderHashes)) throw new Error("Custom-show edit changed slide model rendering.");
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
        type: "source-bound-custom-show-edit",
        sourceName,
        replacementName: outputName,
        nativeId: sourceNativeId,
        orderedSlideNames: requestedSlideNames,
      },
      warnings: [],
      validation: {
        package: { ok: true, partCount: packageScope.partCount, onlyPresentationPartChanged: true, nativeIdPreserved: true },
        reimport: { ok: true, customShowCount: sourceShowCount, orderedSlideNames: requestedSlideNames },
        modelRender: { ok: true, sourceSha256: sourceRenderHashes, outputSha256: outputRenderHashes, slidesByteIdentical: true },
        verify: { ok: verification.ok },
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

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, expectedName, replacementName, slideNames] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    expectedName,
    replacementName,
    orderedSlideNames: String(slideNames || "").split(",").map((name) => name.trim()).filter(Boolean),
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxCustomShow(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, outputSha256: result.audit.output.sha256 }));
}
