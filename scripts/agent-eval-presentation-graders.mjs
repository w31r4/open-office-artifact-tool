import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  PPTX_CLOSED_LEAF_CLONE_FIXTURE,
  PPTX_RICH_NOTES_FIXTURE,
  PPTX_SLIDE_NAME_FIXTURE,
  PPTX_TITLE_NOTES_FIXTURE,
} from "./agent-eval-office-fixtures.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

export const pptxGradedCaseIds = new Set([
  "pptx-title-and-notes-edit",
  "pptx-source-bound-slide-name-edit",
  "pptx-closed-leaf-slide-clone",
]);

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const SHIPPED_RICH_NOTES_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/presentations|node_modules\/open-office-artifact-tool\/skills\/presentations\/skills\/presentations)\/examples\/openchestnut-rich-speaker-notes-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_SLIDE_NAME_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/presentations|node_modules\/open-office-artifact-tool\/skills\/presentations\/skills\/presentations)\/examples\/openchestnut-slide-name-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_SLIDE_DUPLICATE_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/presentations|node_modules\/open-office-artifact-tool\/skills\/presentations\/skills\/presentations)\/examples\/openchestnut-slide-duplicate-workflow\.mjs(?:$|[\s"'`])/i;
const CLONE_TOPOLOGY_PARTS = new Set([
  "[Content_Types].xml",
  "ppt/presentation.xml",
  "ppt/_rels/presentation.xml.rels",
]);
const CUSTOM_SHOW_ACTION = /^ppaction:\/\/customshow\?id=([0-9]+)(?:&return=(true|false))?$/i;

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function xmlAttributes(opening = "") {
  const attributes = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    attributes[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return attributes;
}

function semanticXmlAttributes(opening = "") {
  return xmlAttributes(String(opening).replace(/\sxmlns(?::[\w.-]+)?="[^"]*"/gi, ""));
}

function drawingTexts(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")));
}

function shapeTextByName(xml = "", name) {
  for (const shape of String(xml).matchAll(/<(?:[\w.-]+:)?sp\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?sp>/g)) {
    const properties = /<(?:[\w.-]+:)?cNvPr\b[^>]*>/.exec(shape[0])?.[0] || "";
    if (xmlAttributes(properties).name === name) return drawingTexts(shape[0]).join("\n");
  }
  return null;
}

function slideName(xml = "") {
  const opening = /<(?:[\w.-]+:)?cSld\b[^>]*>/.exec(String(xml))?.[0] || "";
  return xmlAttributes(opening).name || null;
}

function directBackground(xml = "") {
  const background = /<(?:[\w.-]+:)?bg\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?bg>/.exec(String(xml))?.[0] || "";
  const color = /<(?:[\w.-]+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(background)?.[1];
  return color ? "#" + color.toUpperCase() : null;
}

function numericPptxOrder(left, right) {
  return Number(/\d+/.exec(left)?.[0]) - Number(/\d+/.exec(right)?.[0]);
}

async function partHashes(zip, paths) {
  const hashes = {};
  for (const filePath of paths) hashes[filePath] = sha256(await zip.file(filePath).async("uint8array"));
  return hashes;
}

export async function inspectTitleNotesPptx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const slidePaths = paths.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const notesPaths = paths.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const slides = [];
  for (const slidePath of slidePaths) {
    const xml = await zip.file(slidePath).async("text");
    slides.push({
      path: slidePath,
      name: slideName(xml),
      texts: drawingTexts(xml),
      title: shapeTextByName(xml, PPTX_TITLE_NOTES_FIXTURE.titleShapeName),
      background: directBackground(xml),
    });
  }
  const notes = {};
  for (const notesPath of notesPaths) notes[notesPath] = drawingTexts(await zip.file(notesPath).async("text")).join("\n");
  const target = slides.find((slide) => slide.name === PPTX_TITLE_NOTES_FIXTURE.targetSlideName) || null;
  const untouched = slides.find((slide) => slide.name === PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName) || null;
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await partHashes(zip, paths),
    slidePaths,
    notesPaths,
    slides,
    target,
    untouched,
    targetNotesPath: "ppt/notesSlides/notesSlide1.xml",
    targetNotes: notes["ppt/notesSlides/notesSlide1.xml"] || null,
  };
}

function notesBodyShapes(xml = "") {
  const bodies = [];
  for (const shape of String(xml).matchAll(/<(?:[\w.-]+:)?sp\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?sp>/gi)) {
    const placeholder = /<(?:[\w.-]+:)?ph\b[^>]*>/.exec(shape[0])?.[0] || "";
    if (String(xmlAttributes(placeholder).type || "").toLowerCase() !== "body") continue;
    const body = /<(?:[\w.-]+:)?txBody\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?txBody>/i.exec(shape[0]);
    if (body) bodies.push(body[1]);
  }
  return {
    shapeCount: [...String(xml).matchAll(/<(?:[\w.-]+:)?sp\b[^>]*>/gi)].length,
    bodies,
  };
}

function notesOutsideBodySha256(xml = "") {
  const withoutBody = String(xml).replace(
    /(<(?:[\w.-]+:)?txBody\b[^>]*>)[\s\S]*?(<\/(?:[\w.-]+:)?txBody>)/i,
    "$1$2",
  );
  // Open XML SDK can move an equivalent xmlns declaration between the root
  // and a descendant while serializing. This evaluator owns a narrow,
  // namespace-fixed fixture, so discard declaration placement but retain every
  // non-body element, attribute, and text node for the mutation boundary.
  const namespacePlacementNormalized = withoutBody.replace(/\s+xmlns(?::[\w.-]+)?="[^"]*"/g, "");
  return sha256(Buffer.from(namespacePlacementNormalized));
}

function richRunStyle(runXml = "") {
  const properties = /<(?:[\w.-]+:)?rPr\b[^>]*>/.exec(String(runXml))?.[0] || "";
  const attributes = xmlAttributes(properties);
  const color = /<(?:[\w.-]+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/i.exec(String(runXml))?.[1] || null;
  const fontFamily = xmlAttributes(/<(?:[\w.-]+:)?latin\b[^>]*>/.exec(String(runXml))?.[0] || "").typeface || null;
  return {
    bold: attributes.b === "1",
    italic: attributes.i === "1",
    fontSize: /^[0-9]+$/.test(String(attributes.sz || "")) ? Number(attributes.sz) : null,
    color: color ? "#" + color.toUpperCase() : null,
    fontFamily,
  };
}

function inspectRichNotesBody(xml = "") {
  const bodyShapes = notesBodyShapes(xml);
  const body = bodyShapes.bodies[0] || null;
  if (!body) return { present: false, shapeCount: bodyShapes.shapeCount, bodyShapeCount: 0, paragraphs: [] };
  const paragraphs = [];
  for (const match of body.matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/gi)) {
    const paragraphXml = match[0];
    const bodyXml = match[1];
    const bulletCharacter = xmlAttributes(/<(?:[\w.-]+:)?buChar\b[^>]*>/.exec(paragraphXml)?.[0] || "").char || null;
    const autoNumberAttributes = xmlAttributes(/<(?:[\w.-]+:)?buAutoNum\b[^>]*>/.exec(paragraphXml)?.[0] || "");
    const autoNumber = autoNumberAttributes.type
      ? { type: autoNumberAttributes.type, startAt: autoNumberAttributes.startAt === undefined ? null : Number(autoNumberAttributes.startAt) }
      : null;
    const runs = [];
    for (const run of bodyXml.matchAll(/<(?:[\w.-]+:)?r\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?r>/gi)) {
      const text = /<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/i.exec(run[0])?.[1];
      if (text === undefined) continue;
      runs.push({ text: decodeXml(text.replace(/<[^>]+>/g, "")), style: richRunStyle(run[0]) });
    }
    paragraphs.push({ bulletCharacter, autoNumber, runs });
  }
  return {
    present: true,
    shapeCount: bodyShapes.shapeCount,
    bodyShapeCount: bodyShapes.bodies.length,
    paragraphs,
  };
}

/**
 * This intentionally decodes the evaluator's own narrow NotesSlide shape
 * rather than trusting the candidate model. It is limited to the canonical
 * body placeholder and ordinary DrawingML runs generated for this case.
 */
export async function inspectRichNotesPptx(filePath) {
  const fixture = PPTX_RICH_NOTES_FIXTURE;
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const slidePaths = paths.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const notesPaths = paths.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const slides = [];
  for (const slidePath of slidePaths) {
    const xml = await zip.file(slidePath).async("text");
    slides.push({
      path: slidePath,
      name: slideName(xml),
      texts: drawingTexts(xml),
      title: shapeTextByName(xml, fixture.titleShapeName),
      background: directBackground(xml),
    });
  }
  const targetNotesPath = "ppt/notesSlides/notesSlide1.xml";
  const targetNotesXml = await zip.file(targetNotesPath)?.async("text") || "";
  const richNotes = inspectRichNotesBody(targetNotesXml);
  richNotes.outsideBodySha256 = notesOutsideBodySha256(targetNotesXml);
  const target = slides.find((slide) => slide.name === fixture.targetSlideName) || null;
  const untouched = slides.find((slide) => slide.name === fixture.untouchedSlideName) || null;
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await partHashes(zip, paths),
    slidePaths,
    notesPaths,
    slides,
    target,
    untouched,
    targetNotesPath,
    targetNotes: richNotes.paragraphs.map((paragraph) => paragraph.runs.map((run) => run.text).join("")).join("\n"),
    richNotes,
  };
}

function relationshipPartPath(partPath) {
  return path.posix.join(path.posix.dirname(partPath), "_rels", path.posix.basename(partPath) + ".rels");
}

function resolveOpcRelationshipTarget(sourcePart, target) {
  if (typeof sourcePart !== "string" || sourcePart.startsWith("/") || sourcePart.includes("\\") || sourcePart.split("/").includes("..")) {
    throw new Error("Invalid OPC relationship source part: " + JSON.stringify(sourcePart));
  }
  if (typeof target !== "string" || !target || /[\\?#]/.test(target) || /%[0-9a-f]{2}/i.test(target)) {
    throw new Error("OPC relationship target is unsafe: " + JSON.stringify(target));
  }
  const resolved = target.startsWith("/")
    ? target.replace(/^\/+/, "")
    : path.posix.normalize(path.posix.join(sourcePart ? path.posix.dirname(sourcePart) : ".", target));
  if (!resolved || resolved === "." || resolved.startsWith("../") || resolved.split("/").includes("..")) {
    throw new Error("OPC relationship target escapes the package: " + JSON.stringify(target));
  }
  return resolved.replace(/^\.\//, "");
}

function resolvePptxRelationshipTarget(sourcePart, target) {
  if (typeof sourcePart !== "string" || !sourcePart.startsWith("ppt/") || !sourcePart.endsWith(".xml")) {
    throw new Error("Invalid PPTX relationship source part: " + JSON.stringify(sourcePart));
  }
  const resolved = resolveOpcRelationshipTarget(sourcePart, target);
  if (!resolved.startsWith("ppt/")) {
    throw new Error("PPTX relationship target escapes the package: " + JSON.stringify(target));
  }
  return resolved;
}

async function relationshipEntries(zip, sourcePart, { required = true } = {}) {
  const partPath = relationshipPartPath(sourcePart);
  const file = zip.file(partPath);
  if (!file) {
    if (!required) return [];
    throw new Error("PPTX relationship part is missing: " + partPath);
  }
  const xml = await file.async("text");
  const ids = new Set();
  const entries = [];
  for (const match of xml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type || ids.has(attributes.Id)) {
      throw new Error("PPTX relationship part has a missing or duplicate identity: " + partPath);
    }
    ids.add(attributes.Id);
    const external = String(attributes.TargetMode || "").toLowerCase() === "external";
    entries.push({
      id: attributes.Id,
      type: attributes.Type,
      external,
      targetPart: external ? null : resolvePptxRelationshipTarget(sourcePart, attributes.Target),
    });
  }
  return entries;
}

function relationshipWithSuffix(entries, suffix) {
  return entries.filter((entry) => String(entry.type).toLowerCase().endsWith("/" + suffix.toLowerCase()));
}

async function contentTypeForPart(zip, partPath) {
  const xml = await zip.file("[Content_Types].xml")?.async("text");
  if (!xml) throw new Error("PPTX has no [Content_Types].xml.");
  for (const match of xml.matchAll(/<(?:[\w.-]+:)?Override\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (String(attributes.PartName || "").replace(/^\/+/, "") === partPath) return attributes.ContentType || null;
  }
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  for (const match of xml.matchAll(/<(?:[\w.-]+:)?Default\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (String(attributes.Extension || "").toLowerCase() === extension) return attributes.ContentType || null;
  }
  return null;
}

function relationshipSourcePart(relationshipPart) {
  if (relationshipPart === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/i.exec(relationshipPart);
  if (!match) throw new Error("Invalid PPTX relationship-part path: " + relationshipPart);
  return `${match[1] || ""}${match[2]}`;
}

async function packageInboundRelationshipCount(zip, targetPart) {
  let count = 0;
  const relationshipParts = Object.keys(zip.files)
    .filter((partPath) => !zip.files[partPath].dir && (partPath === "_rels/.rels" || /(?:^|\/)_rels\/[^/]+\.rels$/i.test(partPath)));
  for (const relationshipPart of relationshipParts) {
    const sourcePart = relationshipSourcePart(relationshipPart);
    const xml = await zip.file(relationshipPart).async("text");
    for (const match of xml.matchAll(/<Relationship\b[^>]*>/gi)) {
      const attributes = xmlAttributes(match[0]);
      if (!String(attributes.Type || "").toLowerCase().endsWith("/package") || String(attributes.TargetMode || "").toLowerCase() === "external") continue;
      if (resolveOpcRelationshipTarget(sourcePart, attributes.Target) === targetPart) count++;
    }
  }
  return count;
}

function exactlyOneRelationship(entries, suffix, label) {
  const matches = relationshipWithSuffix(entries, suffix);
  if (matches.length !== 1 || matches[0].external || !matches[0].targetPart) {
    throw new Error(label + " must contain exactly one internal " + suffix + " relationship.");
  }
  return matches[0];
}

async function orderedSlideParts(zip) {
  const [presentationXml, entries] = await Promise.all([
    zip.file("ppt/presentation.xml")?.async("text"),
    relationshipEntries(zip, "ppt/presentation.xml"),
  ]);
  if (!presentationXml) throw new Error("PPTX has no ppt/presentation.xml.");
  const relationships = new Map(entries.filter((entry) => String(entry.type).endsWith("/slide"))
    .map((entry) => [entry.id, entry]));
  const parts = [];
  for (const match of presentationXml.matchAll(/<(?:[\w.-]+:)?sldId\b[^>]*>/g)) {
    // xmlAttributes normalizes namespace prefixes, so r:id is exposed as id.
    const relationship = relationships.get(xmlAttributes(match[0]).id);
    if (!relationship?.targetPart || !zip.file(relationship.targetPart)) {
      throw new Error("PPTX slide list contains an unresolved SlidePart relationship.");
    }
    parts.push(relationship.targetPart);
  }
  if (!parts.length || new Set(parts).size !== parts.length) throw new Error("PPTX slide list must contain distinct SlideParts.");
  return parts;
}

async function inspectCanonicalCustomShows(zip) {
  const [presentationXml, entries] = await Promise.all([
    zip.file("ppt/presentation.xml")?.async("text"),
    relationshipEntries(zip, "ppt/presentation.xml"),
  ]);
  if (!presentationXml) throw new Error("PPTX has no ppt/presentation.xml.");
  const slidePartByRelationshipId = new Map(
    relationshipWithSuffix(entries, "slide").map((entry) => [entry.id, entry.targetPart]),
  );
  const listPattern = /<(?:[\w.-]+:)?custShowLst\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?custShowLst>/gi;
  const lists = [...presentationXml.matchAll(listPattern)];
  const listOpeningCount = [...presentationXml.matchAll(/<(?:[\w.-]+:)?custShowLst\b/gi)].length;
  if (!listOpeningCount) return [];
  if (lists.length !== 1 || listOpeningCount !== 1 || Object.keys(semanticXmlAttributes(lists[0][1])).length) {
    throw new Error("PPTX closed-leaf fixture contains a non-canonical custom-show list.");
  }
  const list = lists[0];
  const showPattern = /<(?:[\w.-]+:)?custShow\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?custShow>/gi;
  const showMatches = [...list[2].matchAll(showPattern)];
  if (!showMatches.length || list[2].replace(showPattern, "").trim()) {
    throw new Error("PPTX closed-leaf fixture contains an unmodeled custom-show child.");
  }
  const shows = [];
  const names = new Set();
  const nativeIds = new Set();
  for (const match of showMatches) {
    const attributes = semanticXmlAttributes(match[1]);
    const nativeId = Number(attributes.id);
    const normalizedName = String(attributes.name || "").toLowerCase();
    const slideList = /^\s*<(?:[\w.-]+:)?sldLst\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?sldLst>\s*$/i.exec(match[2]);
    if (Object.keys(attributes).length !== 2 || !attributes.name || names.has(normalizedName)
        || !/^[0-9]+$/.test(attributes.id || "") || !Number.isInteger(nativeId)
        || nativeId < 0 || nativeId > 0xffffffff || nativeIds.has(nativeId)
        || !slideList || Object.keys(semanticXmlAttributes(slideList?.[1] || "")).length) {
      throw new Error("PPTX closed-leaf fixture contains a non-canonical custom show.");
    }
    const entryPattern = /<(?:[\w.-]+:)?sld\b([^>]*)\/\s*>/gi;
    const entryMatches = [...slideList[2].matchAll(entryPattern)];
    if (!entryMatches.length || slideList[2].replace(entryPattern, "").trim()) {
      throw new Error("PPTX custom show contains an unmodeled slide-list child.");
    }
    const slideParts = entryMatches.map((entry) => {
      const entryAttributes = semanticXmlAttributes(entry[1]);
      if (Object.keys(entryAttributes).length !== 1 || !entryAttributes.id) {
        throw new Error("PPTX custom-show member is not one relationship-only slide entry.");
      }
      const target = slidePartByRelationshipId.get(entryAttributes.id);
      if (!target || !zip.file(target)) throw new Error("PPTX custom show references an unresolved SlidePart.");
      return target;
    });
    names.add(normalizedName);
    nativeIds.add(nativeId);
    shows.push({ name: attributes.name, nativeId, slideParts });
  }
  return shows;
}

function elementTexts(xml, localName) {
  const expression = new RegExp(`<([\\w.-]+:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/([\\w.-]+:)?${localName}>`, "gi");
  return [...String(xml).matchAll(expression)].map((match) => decodeXml(match[2].replace(/<[^>]+>/g, "")));
}

async function inspectClosedLeafSlide(zip, partPath) {
  const [xml, entries] = await Promise.all([
    zip.file(partPath)?.async("text"),
    relationshipEntries(zip, partPath),
  ]);
  if (!xml) throw new Error("PPTX SlidePart is missing: " + partPath);
  const notesRelationships = relationshipWithSuffix(entries, "notesSlide");
  const commentRelationships = relationshipWithSuffix(entries, "comments");
  const chartRelationships = relationshipWithSuffix(entries, "chart");
  const packageRelationships = relationshipWithSuffix(entries, "package");
  const imageRelationships = relationshipWithSuffix(entries, "image");
  const customShowActions = [...xml.matchAll(/<(?:[\w.-]+:)?hlinkClick\b[^>]*>/gi)]
    .map((match) => xmlAttributes(match[0]))
    .filter((attributes) => String(attributes.action || "").toLowerCase().startsWith("ppaction://customshow"))
    .map((attributes) => {
      const parsed = CUSTOM_SHOW_ACTION.exec(attributes.action || "");
      return {
        action: attributes.action || "",
        relationshipId: attributes.id || "",
        nativeId: parsed ? Number(parsed[1]) : null,
        returnToSlide: parsed?.[2] ? parsed[2].toLowerCase() === "true" : null,
      };
    });
  if (notesRelationships.length > 1 || commentRelationships.length > 1) {
    throw new Error("PPTX closed-leaf fixture has ambiguous slide relationships.");
  }
  let notes = null;
  if (notesRelationships.length === 1) {
    const relationship = exactlyOneRelationship(entries, "notesSlide", "PPTX notes leaf");
    const [notesXml, notesEntries] = await Promise.all([
      zip.file(relationship.targetPart)?.async("text"),
      relationshipEntries(zip, relationship.targetPart),
    ]);
    if (!notesXml) throw new Error("PPTX NotesSlide part is missing: " + relationship.targetPart);
    const notesMaster = exactlyOneRelationship(notesEntries, "notesMaster", "PPTX notes leaf");
    const slide = exactlyOneRelationship(notesEntries, "slide", "PPTX notes leaf");
    notes = {
      part: relationship.targetPart,
      relationshipPart: relationshipPartPath(relationship.targetPart),
      relationship: { id: relationship.id, type: relationship.type },
      notesMaster: { id: notesMaster.id, type: notesMaster.type, targetPart: notesMaster.targetPart },
      slide: { id: slide.id, type: slide.type, targetPart: slide.targetPart },
      text: drawingTexts(notesXml).join("\n"),
      sha256: sha256(Buffer.from(notesXml)),
    };
  }
  let comments = null;
  if (commentRelationships.length === 1) {
    const relationship = exactlyOneRelationship(entries, "comments", "PPTX comments leaf");
    const commentsXml = await zip.file(relationship.targetPart)?.async("text");
    if (!commentsXml) throw new Error("PPTX SlideComments part is missing: " + relationship.targetPart);
    const childEntries = await relationshipEntries(zip, relationship.targetPart, { required: false });
    comments = {
      part: relationship.targetPart,
      relationship: { id: relationship.id, type: relationship.type },
      texts: elementTexts(commentsXml, "text"),
      childRelationshipCount: childEntries.length,
      sha256: sha256(Buffer.from(commentsXml)),
    };
  }
  const chartReferenceIds = [...xml.matchAll(/<(?:[\w.-]+:)?chart\b[^>]*>/gi)]
    .map((match) => xmlAttributes(match[0]).id)
    .filter(Boolean);
  if (new Set(chartReferenceIds).size !== chartReferenceIds.length
      || chartReferenceIds.length !== chartRelationships.length) {
    throw new Error("PPTX closed-chart fixture has duplicate, orphan, or unmodeled chart relationships.");
  }
  const chartRelationshipById = new Map(chartRelationships.map((relationship) => [relationship.id, relationship]));
  const charts = [];
  for (const relationshipId of chartReferenceIds) {
    const relationship = chartRelationshipById.get(relationshipId);
    if (!relationship || relationship.external || !relationship.targetPart
        || !/^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i.test(relationship.targetPart)) {
      throw new Error("PPTX chart leaf must use one internal numbered ChartPart relationship.");
    }
    const chartXml = await zip.file(relationship.targetPart)?.async("text");
    if (!chartXml) throw new Error("PPTX ChartPart is missing: " + relationship.targetPart);
    const childEntries = await relationshipEntries(zip, relationship.targetPart, { required: false });
    charts.push({
      part: relationship.targetPart,
      relationship: { id: relationship.id, type: relationship.type },
      texts: drawingTexts(chartXml),
      values: elementTexts(chartXml, "v"),
      childRelationshipCount: childEntries.length,
      bytes: Buffer.byteLength(chartXml),
      sha256: sha256(Buffer.from(chartXml)),
    });
  }
  const oleFrames = [...xml.matchAll(/<(?:[\w.-]+:)?graphicFrame\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?graphicFrame\s*>/gi)]
    .filter((match) => /<(?:[\w.-]+:)?oleObj\b/i.test(match[0]));
  const allOleObjects = [...xml.matchAll(/<(?:[\w.-]+:)?oleObj\b/gi)].length;
  if (oleFrames.length !== allOleObjects || oleFrames.length !== packageRelationships.length) {
    throw new Error("PPTX embedded-XLSX OLE fixture has an orphan, nested, or unmodeled package relationship.");
  }
  const packageRelationshipById = new Map(packageRelationships.map((relationship) => [relationship.id, relationship]));
  const imageRelationshipById = new Map(imageRelationships.map((relationship) => [relationship.id, relationship]));
  const usedPackageIds = new Set();
  const oleWorkbooks = [];
  for (const frameMatch of oleFrames) {
    const frame = frameMatch[0];
    const groupDepth = [...xml.slice(0, frameMatch.index).matchAll(/<(\/)?(?:[\w.-]+:)?grpSp\b[^>]*>/gi)]
      .reduce((depth, match) => depth + (match[1] ? -1 : /\/\s*>$/.test(match[0]) ? 0 : 1), 0);
    const oleTags = [...frame.matchAll(/<(?:[\w.-]+:)?oleObj\b[^>]*>/gi)];
    const embeds = [...frame.matchAll(/<(?:[\w.-]+:)?embed\b[^>]*\/?>/gi)];
    const links = [...frame.matchAll(/<(?:[\w.-]+:)?link\b/gi)];
    const pictures = [...frame.matchAll(/<(?:[\w.-]+:)?pic\b/gi)];
    const blips = [...frame.matchAll(/<(?:[\w.-]+:)?blip\b[^>]*>/gi)];
    const relationshipAttributes = [...frame.matchAll(/\br:(?:id|embed|link)="[^"]*"/gi)];
    if (groupDepth !== 0 || oleTags.length !== 1 || embeds.length !== 1 || links.length || pictures.length !== 1 || blips.length !== 1 || relationshipAttributes.length !== 2) {
      throw new Error("PPTX embedded-XLSX OLE fixture is outside the canonical top-level embed-plus-preview frame profile.");
    }
    const packageId = xmlAttributes(oleTags[0][0]).id;
    const previewId = xmlAttributes(blips[0][0]).embed;
    const packageRelationship = packageRelationshipById.get(packageId);
    const previewRelationship = imageRelationshipById.get(previewId);
    if (!packageId || !previewId || !usedPackageIds.add(packageId)
        || !packageRelationship || packageRelationship.external || !packageRelationship.targetPart
        || !previewRelationship || previewRelationship.external || !previewRelationship.targetPart
        || !packageRelationship.targetPart.toLowerCase().endsWith(".xlsx")
        || await contentTypeForPart(zip, packageRelationship.targetPart) !== "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        || !/^ppt\/media\/[^/]+\.[a-z0-9]+$/i.test(previewRelationship.targetPart)
        || !String(await contentTypeForPart(zip, previewRelationship.targetPart)).toLowerCase().startsWith("image/")) {
      throw new Error("PPTX embedded-XLSX OLE fixture has an invalid package or preview binding.");
    }
    const [workbookBytes, previewBytes, childRelationships, inboundCount] = await Promise.all([
      zip.file(packageRelationship.targetPart)?.async("uint8array"),
      zip.file(previewRelationship.targetPart)?.async("uint8array"),
      relationshipEntries(zip, packageRelationship.targetPart, { required: false }),
      packageInboundRelationshipCount(zip, packageRelationship.targetPart),
    ]);
    if (!workbookBytes?.length || !previewBytes?.length || childRelationships.length || inboundCount !== 1) {
      throw new Error("PPTX embedded-XLSX OLE package is missing, shared, or owns a child relationship graph.");
    }
    oleWorkbooks.push({
      part: packageRelationship.targetPart,
      relationship: { id: packageRelationship.id, type: packageRelationship.type },
      bytes: workbookBytes.length,
      sha256: sha256(workbookBytes),
      childRelationshipCount: childRelationships.length,
      inboundRelationshipCount: inboundCount,
      previewPart: previewRelationship.targetPart,
      previewRelationship: { id: previewRelationship.id, type: previewRelationship.type },
      previewSha256: sha256(previewBytes),
    });
  }
  if (usedPackageIds.size !== packageRelationshipById.size) {
    throw new Error("PPTX embedded-XLSX OLE fixture has an unused package relationship.");
  }
  return {
    part: partPath,
    name: slideName(xml),
    texts: drawingTexts(xml),
    background: directBackground(xml),
    sha256: sha256(Buffer.from(xml)),
    notes,
    comments,
    charts,
    oleWorkbooks,
    customShowActions,
  };
}

/**
 * Independent package-level evidence for the bounded notes/comments/chart
 * clone profile. It deliberately reads OPC relationships and bytes rather
 * than trusting the workflow audit or the presentation model.
 */
export async function inspectClosedLeafClonePptx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const [slideParts, presentationRelationships, customShows] = await Promise.all([
    orderedSlideParts(zip),
    relationshipEntries(zip, "ppt/presentation.xml"),
    inspectCanonicalCustomShows(zip),
  ]);
  const authorRelationships = relationshipWithSuffix(presentationRelationships, "commentAuthors");
  if (authorRelationships.length > 1) throw new Error("PPTX has more than one CommentAuthors relationship.");
  const authorRelationship = authorRelationships[0] || null;
  const authorPart = authorRelationship?.targetPart || null;
  if (authorRelationship && (authorRelationship.external || !authorPart || !zip.file(authorPart))) {
    throw new Error("PPTX CommentAuthors relationship is invalid.");
  }
  const slides = await Promise.all(slideParts.map((partPath) => inspectClosedLeafSlide(zip, partPath)));
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await partHashes(zip, paths),
    slides,
    customShows,
    commentAuthors: authorRelationship ? {
      part: authorPart,
      relationship: { id: authorRelationship.id, type: authorRelationship.type },
      sha256: sha256(await zip.file(authorPart).async("uint8array")),
    } : null,
  };
}

function auditProvider(audit) {
  const provider = audit?.provider;
  return String(typeof provider === "string" ? provider : provider?.actual || provider?.selected || provider?.name || "");
}

function auditVersion(audit) {
  const provider = audit?.provider;
  return String(provider?.version || audit?.providerVersion || "");
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  return provider.silentFallback === false || provider.fallbackUsed === false || audit?.silentFallback === false || audit?.fallbackUsed === false;
}

function auditStrategy(audit) {
  const policy = audit?.savePolicy || audit?.save_strategy;
  return String(typeof policy === "string" ? policy : policy?.strategy || audit?.strategy || "");
}

function auditOperation(audit) {
  const operation = audit?.operation;
  return String(typeof operation === "string" ? operation : operation?.type || operation?.name || "");
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[`${side}Sha256`] || "");
}

function sameArray(left, right) {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function sameValue(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function packageChanges(source, output) {
  const paths = [...new Set([...source.paths, ...output.paths])].sort();
  return paths.filter((filePath) => source.partHashes[filePath] !== output.partHashes[filePath]);
}

function visualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount === 2;
  const targetChanged = pageCountsMatch && source.pages?.[0]?.pixelSha256 !== output.pages?.[0]?.pixelSha256;
  const untouchedStable = pageCountsMatch
    && source.pages?.[1]?.width === output.pages?.[1]?.width
    && source.pages?.[1]?.height === output.pages?.[1]?.height
    && source.pages?.[1]?.pixelSha256 === output.pages?.[1]?.pixelSha256;
  return { available, rendered, pageCountsMatch, targetChanged, untouchedStable };
}

function stableVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount === 2;
  const pageStable = (index) => pageCountsMatch
    && source.pages?.[index]?.width === output.pages?.[index]?.width
    && source.pages?.[index]?.height === output.pages?.[index]?.height
    && source.pages?.[index]?.pixelSha256 === output.pages?.[index]?.pixelSha256;
  return {
    available,
    rendered,
    pageCountsMatch,
    targetStable: pageStable(0),
    untouchedStable: pageStable(1),
  };
}

function usedTypedRichNotesRoundTrip(commandText) {
  const directPublicApi = /PresentationFile\.importPptx/i.test(commandText)
    && /PresentationFile\.exportPptx/i.test(commandText);
  return directPublicApi || SHIPPED_RICH_NOTES_WORKFLOW.test(commandText);
}

function usedTypedSlideNameRoundTrip(commandText) {
  const directPublicApi = /PresentationFile\.importPptx/i.test(commandText)
    && /PresentationFile\.exportPptx/i.test(commandText);
  return directPublicApi || SHIPPED_SLIDE_NAME_WORKFLOW.test(commandText);
}

function usedTypedClosedLeafClone(commandText) {
  const directPublicApi = /PresentationFile\.importPptx/i.test(commandText)
    && /PresentationFile\.exportPptx/i.test(commandText);
  return directPublicApi || (SHIPPED_SLIDE_DUPLICATE_WORKFLOW.test(commandText)
    && /--allow-closed-leaves\b/i.test(commandText));
}

function cloneVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const expectedPageCounts = source?.pageCount === 2 && output?.pageCount === 3;
  const samePage = (left, right) => left?.width === right?.width
    && left?.height === right?.height
    && left?.pixelSha256 === right?.pixelSha256;
  return {
    available,
    rendered,
    expectedPageCounts,
    retainedSourceStable: samePage(source?.pages?.[0], output?.pages?.[0]),
    cloneMatchesSource: samePage(source?.pages?.[0], output?.pages?.[1]),
    appendixStable: samePage(source?.pages?.[1], output?.pages?.[2]),
  };
}

function closedChartSemanticsEqual(left = [], right = []) {
  return left.length === right.length && left.every((chart, index) => {
    const candidate = right[index];
    return chart?.relationship?.id === candidate?.relationship?.id
      && chart?.relationship?.type === candidate?.relationship?.type
      && chart?.sha256 === candidate?.sha256
      && chart?.childRelationshipCount === candidate?.childRelationshipCount
      && sameArray(chart?.texts || [], candidate?.texts || [])
      && sameArray(chart?.values || [], candidate?.values || []);
  });
}

function closedOleWorkbookSemanticsEqual(left = [], right = []) {
  return left.length === right.length && left.every((workbook, index) => {
    const candidate = right[index];
    return workbook?.relationship?.id === candidate?.relationship?.id
      && workbook?.relationship?.type === candidate?.relationship?.type
      && workbook?.sha256 === candidate?.sha256
      && workbook?.childRelationshipCount === 0
      && candidate?.childRelationshipCount === 0
      && workbook?.inboundRelationshipCount === 1
      && candidate?.inboundRelationshipCount === 1
      && workbook?.previewRelationship?.id === candidate?.previewRelationship?.id
      && workbook?.previewRelationship?.type === candidate?.previewRelationship?.type
      && workbook?.previewPart === candidate?.previewPart
      && workbook?.previewSha256 === candidate?.previewSha256;
  });
}

function closedLeafSemanticsEqual(left, right) {
  return left?.name === right?.name
    && sameArray(left?.texts || [], right?.texts || [])
    && left?.background === right?.background
    && left?.notes?.text === right?.notes?.text
    && sameArray(left?.comments?.texts || [], right?.comments?.texts || [])
    && closedChartSemanticsEqual(left?.charts, right?.charts)
    && closedOleWorkbookSemanticsEqual(left?.oleWorkbooks, right?.oleWorkbooks)
    && JSON.stringify(left?.customShowActions || []) === JSON.stringify(right?.customShowActions || []);
}

/**
 * Grade the deliberately small closed-leaf clone profile. The only acceptable
 * graph change is an adjacent cloned SlidePart plus a distinct byte-identical
 * ChartPart, cloned NotesSlide/relationship part, and cloned SlideComments XML;
 * both global catalog parts remain shared. This is not a general relationship-
 * graph clone oracle.
 */
export function gradePptxClosedLeafCloneEvidence({ evidence, audit, commands }) {
  const fixture = PPTX_CLOSED_LEAF_CLONE_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const visual = cloneVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const [sourceSlide, sourceAppendix] = source.slides;
  const [retainedSource, clone, outputAppendix] = output.slides;
  const newPaths = output.paths.filter((entry) => !source.paths.includes(entry)).sort();
  const changedSourcePaths = source.paths.filter((entry) => !CLONE_TOPOLOGY_PARTS.has(entry)
    && source.partHashes[entry] !== output.partHashes[entry]);
  const expectedNewPaths = [
    clone?.part,
    clone ? relationshipPartPath(clone.part) : null,
    ...(clone?.charts || []).map((chart) => chart.part),
    ...(clone?.oleWorkbooks || []).map((workbook) => workbook.part),
    clone?.notes?.part,
    clone?.notes?.relationshipPart,
    clone?.comments?.part,
  ].filter(Boolean).sort();
  const operation = audit?.operation && typeof audit.operation === "object" ? audit.operation : {};
  const packageValidation = audit?.validation?.package || {};
  const closedLeaves = operation.closedLeaves || {};
  const sourceLeaves = sourceSlide?.notes && sourceSlide?.comments;
  const cloneLeaves = clone?.notes && clone?.comments;
  const sourceChart = sourceSlide?.charts?.[0] || null;
  const cloneChart = clone?.charts?.[0] || null;
  const sourceOleWorkbook = sourceSlide?.oleWorkbooks?.[0] || null;
  const cloneOleWorkbook = clone?.oleWorkbooks?.[0] || null;
  const sourceCustomShow = source.customShows?.[0] || null;
  const outputCustomShow = output.customShows?.[0] || null;
  const sourceCustomShowAction = sourceSlide?.customShowActions?.[0] || null;
  const cloneCustomShowAction = clone?.customShowActions?.[0] || null;
  const expectedChartValues = [
    fixture.chartSeriesName,
    ...fixture.chartCategories,
    ...fixture.chartValues.map(String),
  ];
  const canonicalFixture = source.slides.length === 2
    && sourceSlide?.name === fixture.sourceSlideName
    && sourceSlide?.texts.includes(fixture.sourceTitle)
    && sourceSlide?.texts.includes(fixture.sourceSupportingText)
    && sourceSlide?.texts.includes(fixture.customShowText)
    && sourceSlide?.background === fixture.sourceBackground
    && sourceSlide?.notes?.text === fixture.sourceNotes
    && sameArray(sourceSlide?.comments?.texts || [], [fixture.sourceComment])
    && sourceSlide?.comments?.childRelationshipCount === 0
    && sourceSlide?.charts?.length === 1
    && sourceChart?.texts?.includes(fixture.chartTitle)
    && sameArray(sourceChart?.values || [], expectedChartValues)
    && sourceChart?.childRelationshipCount === 0
    && sourceSlide?.oleWorkbooks?.length === 1
    && sourceOleWorkbook?.part === fixture.oleWorkbookPart
    && sourceOleWorkbook?.relationship?.id === fixture.oleWorkbookRelationshipId
    && sourceOleWorkbook?.childRelationshipCount === 0
    && sourceOleWorkbook?.inboundRelationshipCount === 1
    && sourceOleWorkbook?.previewPart === fixture.olePreviewPart
    && sourceOleWorkbook?.previewRelationship?.id === fixture.olePreviewRelationshipId
    && sourceAppendix?.name === fixture.appendixSlideName
    && sourceAppendix?.texts.includes(fixture.appendixText)
    && sourceAppendix?.background === fixture.appendixBackground
    && sourceAppendix?.charts?.length === 0
    && source.customShows?.length === 1
    && sourceCustomShow?.name === fixture.customShowName
    && sourceCustomShow?.nativeId === fixture.customShowNativeId
    && sameArray(sourceCustomShow?.slideParts || [], [sourceSlide?.part, sourceAppendix?.part])
    && sourceSlide?.customShowActions?.length === 1
    && sourceCustomShowAction?.relationshipId === ""
    && sourceCustomShowAction?.nativeId === fixture.customShowNativeId
    && sourceCustomShowAction?.returnToSlide === true
    && Boolean(source.commentAuthors?.part);
  const cloneSemantics = output.slides.length === 3
    && retainedSource?.part === sourceSlide?.part
    && clone?.name === fixture.sourceSlideName
    && outputAppendix?.part === sourceAppendix?.part
    && outputAppendix?.name === fixture.appendixSlideName
    && closedLeafSemanticsEqual(sourceSlide, retainedSource)
    && closedLeafSemanticsEqual(sourceSlide, clone)
    && closedLeafSemanticsEqual(sourceAppendix, outputAppendix);
  const closedChartGraph = Boolean(sourceChart && cloneChart)
    && sourceSlide.charts.length === 1
    && clone.charts.length === 1
    && sourceChart.part !== cloneChart.part
    && source.paths.includes(sourceChart.part)
    && !source.paths.includes(cloneChart.part)
    && sourceChart.relationship.id === cloneChart.relationship.id
    && sourceChart.relationship.type === cloneChart.relationship.type
    && sourceChart.sha256 === cloneChart.sha256
    && sourceChart.childRelationshipCount === 0
    && cloneChart.childRelationshipCount === 0;
  const closedOleWorkbookGraph = Boolean(sourceOleWorkbook && cloneOleWorkbook)
    && sourceSlide.oleWorkbooks.length === 1
    && clone.oleWorkbooks.length === 1
    && sourceOleWorkbook.part !== cloneOleWorkbook.part
    && source.paths.includes(sourceOleWorkbook.part)
    && !source.paths.includes(cloneOleWorkbook.part)
    && sourceOleWorkbook.relationship.id === cloneOleWorkbook.relationship.id
    && sourceOleWorkbook.relationship.type === cloneOleWorkbook.relationship.type
    && sourceOleWorkbook.sha256 === cloneOleWorkbook.sha256
    && sourceOleWorkbook.childRelationshipCount === 0
    && cloneOleWorkbook.childRelationshipCount === 0
    && sourceOleWorkbook.inboundRelationshipCount === 1
    && cloneOleWorkbook.inboundRelationshipCount === 1
    && sourceOleWorkbook.previewRelationship.id === cloneOleWorkbook.previewRelationship.id
    && sourceOleWorkbook.previewRelationship.type === cloneOleWorkbook.previewRelationship.type
    && sourceOleWorkbook.previewPart === cloneOleWorkbook.previewPart
    && sourceOleWorkbook.previewSha256 === cloneOleWorkbook.previewSha256;
  const closedLeafGraph = Boolean(sourceLeaves && cloneLeaves)
    && sourceSlide.notes.part !== clone.notes.part
    && sourceSlide.notes.relationship.id === clone.notes.relationship.id
    && sourceSlide.notes.relationship.type === clone.notes.relationship.type
    && sourceSlide.notes.notesMaster.id === clone.notes.notesMaster.id
    && sourceSlide.notes.notesMaster.type === clone.notes.notesMaster.type
    && sourceSlide.notes.notesMaster.targetPart === clone.notes.notesMaster.targetPart
    && sourceSlide.notes.slide.id === clone.notes.slide.id
    && sourceSlide.notes.slide.type === clone.notes.slide.type
    && sourceSlide.notes.slide.targetPart === sourceSlide.part
    && clone.notes.slide.targetPart === clone.part
    && sourceSlide.notes.sha256 === clone.notes.sha256
    && sourceSlide.comments.part !== clone.comments.part
    && sourceSlide.comments.relationship.id === clone.comments.relationship.id
    && sourceSlide.comments.relationship.type === clone.comments.relationship.type
    && sourceSlide.comments.sha256 === clone.comments.sha256
    && clone.comments.childRelationshipCount === 0
    && source.commentAuthors?.part === output.commentAuthors?.part
    && source.commentAuthors?.sha256 === output.commentAuthors?.sha256;
  const customShowGraph = Boolean(sourceCustomShow && outputCustomShow && sourceCustomShowAction && cloneCustomShowAction)
    && output.customShows.length === 1
    && JSON.stringify(sourceCustomShow) === JSON.stringify(outputCustomShow)
    && sameArray(outputCustomShow.slideParts, [retainedSource?.part, outputAppendix?.part])
    && !outputCustomShow.slideParts.includes(clone?.part)
    && sourceCustomShowAction.action === cloneCustomShowAction.action
    && sourceCustomShowAction.relationshipId === ""
    && cloneCustomShowAction.relationshipId === ""
    && sourceCustomShowAction.nativeId === fixture.customShowNativeId
    && cloneCustomShowAction.nativeId === fixture.customShowNativeId
    && sourceCustomShowAction.returnToSlide === true
    && cloneCustomShowAction.returnToSlide === true;
  const packageScope = sameArray(newPaths, expectedNewPaths)
    && changedSourcePaths.length === 0
    && customShowGraph
    && sourceSlide?.sha256 === retainedSource?.sha256
    && sourceAppendix?.sha256 === outputAppendix?.sha256
    && source.paths.every((entry) => !CLONE_TOPOLOGY_PARTS.has(entry)
      ? source.partHashes[entry] === output.partHashes[entry]
      : true);
  const commandText = commands.join("\n");
  return [
    check("pptx-clone-machine:canonical-closed-leaf-fixture", "machine", canonicalFixture, {
      sourceSlide,
      sourceAppendix,
      commentAuthors: source.commentAuthors,
    }),
    check("pptx-clone-machine:adjacent-semantic-clone", "machine", cloneSemantics, {
      sourceSlides: source.slides,
      outputSlides: output.slides,
    }),
    check("pptx-clone-machine:closed-leaves-copied-and-catalogs-shared", "machine", closedLeafGraph, {
      sourceNotes: sourceSlide?.notes || null,
      cloneNotes: clone?.notes || null,
      sourceComments: sourceSlide?.comments || null,
      cloneComments: clone?.comments || null,
      sourceCommentAuthors: source.commentAuthors || null,
      outputCommentAuthors: output.commentAuthors || null,
    }),
    check("pptx-clone-machine:chart-part-copied-to-independent-leaf", "machine", closedChartGraph, {
      sourceChart,
      cloneChart,
    }),
    check("pptx-clone-machine:ole-workbook-copied-to-independent-package", "machine", closedOleWorkbookGraph, {
      sourceOleWorkbook,
      cloneOleWorkbook,
    }),
    check("pptx-clone-machine:custom-show-action-retained-without-membership-drift", "machine", customShowGraph, {
      sourceCustomShow,
      outputCustomShow,
      sourceCustomShowAction,
      cloneCustomShowAction,
    }),
    check("pptx-clone-machine:only-approved-clone-parts-added", "machine", sameArray(newPaths, expectedNewPaths), {
      newPaths,
      expectedNewPaths,
    }),
    check("pptx-clone-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("pptx-clone-visual:native-render", "visual", visual.available && visual.rendered && visual.expectedPageCounts, {
      visual: evidence.visual,
    }),
    check("pptx-clone-visual:source-clone-and-appendix-pixels-stable", "visual", visual.retainedSourceStable
      && visual.cloneMatchesSource && visual.appendixStable, { visual }),
    gate("pptx-clone-security:source-parts-byte-preserved-and-graph-bounded", "security", packageScope && closedLeafGraph && closedChartGraph && closedOleWorkbookGraph && customShowGraph, {
      changedSourcePaths,
      newPaths,
      expectedNewPaths,
      sourceParts: source.paths,
      outputParts: output.paths,
    }),
    gate("pptx-clone-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("pptx-clone-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("pptx-clone-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("pptx-clone-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("pptx-clone-trace:closed-leaf-operation-contract", "trace", auditOperation(audit) === "source-bound-slide-duplicate"
      && operation.expectedName === fixture.sourceSlideName
      && operation.sourcePart === sourceSlide?.part
      && operation.clonePart === clone?.part
      && operation.allowClosedLeaves === true
      && closedLeaves.speakerNotes === true
      && closedLeaves.legacyComments === true
      && operation.chartParts?.count === 1
      && operation.oleWorkbookParts?.count === 1
      && operation.runHyperlinks?.customShowCount === 1
      && operation.runHyperlinks?.customShowActions?.[0]?.nativeId === fixture.customShowNativeId
      && operation.customShows?.count === 1
      && operation.customShows?.shows?.[0]?.nativeId === fixture.customShowNativeId
      && sameArray(operation.chartParts?.sourceParts || [], [sourceChart?.part])
      && sameArray(operation.chartParts?.relationshipIds || [], [sourceChart?.relationship?.id])
      && sameArray(operation.oleWorkbookParts?.sourceParts || [], [sourceOleWorkbook?.part])
      && sameArray(operation.oleWorkbookParts?.relationshipIds || [], [sourceOleWorkbook?.relationship?.id])
      && sameArray(operation.oleWorkbookParts?.previewParts || [], [sourceOleWorkbook?.previewPart]), {
      operation: audit?.operation || null,
    }),
    check("pptx-clone-trace:package-validation-contract", "trace", packageValidation.ok === true
      && packageValidation.retainedSourcePartsByteIdentical === true
      && packageValidation.cloneInsertedAdjacent === true
      && packageValidation.chartParts?.count === 1
      && packageValidation.chartParts?.independentParts === true
      && packageValidation.chartParts?.allPayloadsByteIdentical === true
      && packageValidation.chartParts?.parts?.[0]?.relationshipId === sourceChart?.relationship?.id
      && packageValidation.chartParts?.parts?.[0]?.sourcePart === sourceChart?.part
      && packageValidation.chartParts?.parts?.[0]?.clonePart === cloneChart?.part
      && packageValidation.chartParts?.parts?.[0]?.sourceSha256 === sourceChart?.sha256
      && packageValidation.oleWorkbookParts?.count === 1
      && packageValidation.oleWorkbookParts?.independentParts === true
      && packageValidation.oleWorkbookParts?.allPayloadsByteIdentical === true
      && packageValidation.oleWorkbookParts?.previewPartsShared === true
      && packageValidation.oleWorkbookParts?.parts?.[0]?.relationshipId === sourceOleWorkbook?.relationship?.id
      && packageValidation.oleWorkbookParts?.parts?.[0]?.sourcePart === sourceOleWorkbook?.part
      && packageValidation.oleWorkbookParts?.parts?.[0]?.clonePart === cloneOleWorkbook?.part
      && packageValidation.oleWorkbookParts?.parts?.[0]?.sourceSha256 === sourceOleWorkbook?.sha256
      && packageValidation.oleWorkbookParts?.parts?.[0]?.previewPart === sourceOleWorkbook?.previewPart
      && packageValidation.oleWorkbookParts?.parts?.[0]?.previewShared === true
      && packageValidation.runHyperlinks?.customShowCount === 1
      && packageValidation.customShows?.count === 1
      && packageValidation.customShows?.exactSourceMembershipRetained === true
      && packageValidation.closedLeaves?.speakerNotes?.notesXmlByteIdentical === true
      && packageValidation.closedLeaves?.legacyComments?.commentsXmlByteIdentical === true, {
      packageValidation,
    }),
    check("pptx-clone-trace:typed-roundtrip", "trace", usedTypedClosedLeafClone(commandText), {
      expected: "public PresentationFile importPptx/exportPptx calls or the integrity-protected published duplicate workflow with --allow-closed-leaves",
    }),
    check("pptx-clone-trace:second-import", "trace", audit?.validation?.reimport?.ok === true
      && audit?.validation?.reimport?.sourceAndCloneSemanticsEqual === true
      && audit?.validation?.reimport?.sourceAndCloneClosedLeavesEqual === true
      && audit?.validation?.reimport?.sourceAndCloneOleWorkbookBindingsIndependent === true
      && audit?.validation?.reimport?.customShowMembershipRetained === true, {
      validation: audit?.validation?.reimport || null,
    }),
  ];
}

export function gradePptxRichNotesEvidence({ evidence, audit, commands }) {
  const fixture = PPTX_RICH_NOTES_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const visual = visualEvidence(evidence.visual?.source, evidence.visual?.output);
  const changedPaths = packageChanges(source, output);
  const expectedChangedPaths = [source.target?.path, source.targetNotesPath].filter(Boolean).sort();
  const sourceSlideNames = source.slides.map((slide) => slide.name);
  const outputSlideNames = output.slides.map((slide) => slide.name);
  const sourceText = source.target?.texts || [];
  const outputText = output.target?.texts || [];
  const sourceParagraphs = source.richNotes?.paragraphs || [];
  const outputParagraphs = output.richNotes?.paragraphs || [];
  const sourceTargetRun = sourceParagraphs[fixture.targetRun.paragraphIndex]?.runs?.[fixture.targetRun.runIndex] || null;
  const outputTargetRun = outputParagraphs[fixture.targetRun.paragraphIndex]?.runs?.[fixture.targetRun.runIndex] || null;
  const sourceFirstRun = sourceParagraphs[0]?.runs?.[0] || null;
  const outputFirstRun = outputParagraphs[0]?.runs?.[0] || null;
  const sourceSecondParagraph = sourceParagraphs[1] || null;
  const outputSecondParagraph = outputParagraphs[1] || null;
  const topology = (paragraphs) => paragraphs.map((paragraph) => ({
    bulletCharacter: paragraph.bulletCharacter,
    autoNumber: paragraph.autoNumber,
    runCount: paragraph.runs.length,
  }));
  const commandText = commands.join("\n");
  const operation = audit?.operation && typeof audit.operation === "object" ? audit.operation : {};
  return [
    check("pptx-rich-notes-machine:canonical-fixture", "machine", source.slides.length === 2
      && source.target?.title === fixture.originalTitle
      && source.target?.background === fixture.targetBackground
      && source.targetNotes === fixture.originalNotes
      && source.untouched?.background === fixture.untouchedBackground
      && sourceText.includes(fixture.supportingText)
      && source.richNotes?.present === true
      && source.richNotes?.shapeCount === 1
      && source.richNotes?.bodyShapeCount === 1
      && sourceParagraphs.length === 2
      && sourceParagraphs[0]?.bulletCharacter === "•"
      && sourceParagraphs[0]?.runs.length === 2
      && sourceFirstRun?.text === fixture.originalNotesParagraphs[0].runs[0].text
      && sourceFirstRun?.style.bold === true
      && sourceFirstRun?.style.fontFamily === "Aptos"
      && sourceTargetRun?.text === fixture.targetRun.expectedText
      && sourceTargetRun?.style.italic === true
      && sourceTargetRun?.style.bold === false
      && sourceTargetRun?.style.color === "#7C2D12"
      && sourceTargetRun?.style.fontSize !== null
      && sourceSecondParagraph?.autoNumber?.type === "arabicPeriod"
      && sourceSecondParagraph?.autoNumber?.startAt === 2
      && sourceSecondParagraph?.runs.length === 1
      && sourceSecondParagraph?.runs[0]?.text === fixture.originalNotesParagraphs[1].runs[0].text, {
      sourceTarget: source.target,
      sourceNotes: source.targetNotes,
      sourceRichNotes: source.richNotes,
      sourceSlides: sourceSlideNames,
    }),
    check("pptx-rich-notes-machine:title-and-target-run-edited", "machine", output.target?.title === fixture.replacementTitle
      && output.targetNotes === fixture.replacementNotes
      && outputTargetRun?.text === fixture.targetRun.replacementText
      && outputTargetRun?.style.bold === true
      && outputTargetRun?.style.italic === false
      && outputTargetRun?.style.color === "#0F766E"
      && outputTargetRun?.style.fontSize === sourceTargetRun?.style.fontSize
      && !outputText.includes(fixture.originalTitle)
      && !String(output.targetNotes || "").includes(fixture.targetRun.expectedText), {
      outputTarget: output.target,
      outputNotes: output.targetNotes,
      outputTargetRun,
    }),
    check("pptx-rich-notes-machine:fixed-topology-and-siblings-preserved", "machine", sameArray(sourceSlideNames, outputSlideNames)
      && output.target?.background === fixture.targetBackground
      && outputText.includes(fixture.supportingText)
      && output.untouched?.background === fixture.untouchedBackground
      && source.untouched?.path === output.untouched?.path
      && output.richNotes?.shapeCount === source.richNotes?.shapeCount
      && output.richNotes?.bodyShapeCount === source.richNotes?.bodyShapeCount
      && output.richNotes?.outsideBodySha256 === source.richNotes?.outsideBodySha256
      && sameValue(topology(outputParagraphs), topology(sourceParagraphs))
      && sameValue(outputFirstRun, sourceFirstRun)
      && sameValue(outputSecondParagraph, sourceSecondParagraph), {
      sourceSlides: sourceSlideNames,
      outputSlides: outputSlideNames,
      sourceRichNotes: source.richNotes,
      outputRichNotes: output.richNotes,
      sourceTarget: source.target,
      outputTarget: output.target,
      sourceUntouched: source.untouched,
      outputUntouched: output.untouched,
    }),
    check("pptx-rich-notes-machine:only-bounded-parts-changed", "machine", sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
    }),
    check("pptx-rich-notes-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("pptx-rich-notes-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("pptx-rich-notes-visual:target-change-and-untouched-slide-stable", "visual", visual.targetChanged && visual.untouchedStable, {
      visual: evidence.visual,
    }),
    gate("pptx-rich-notes-security:fixed-topology-and-package-preservation", "security", sameArray(source.paths, output.paths)
      && source.target?.path === output.target?.path
      && source.targetNotesPath === output.targetNotesPath
      && source.partHashes[source.untouched?.path] === output.partHashes[output.untouched?.path]
      && output.richNotes?.shapeCount === source.richNotes?.shapeCount
      && output.richNotes?.bodyShapeCount === source.richNotes?.bodyShapeCount
      && output.richNotes?.outsideBodySha256 === source.richNotes?.outsideBodySha256
      && sameValue(topology(outputParagraphs), topology(sourceParagraphs))
      && sameValue(outputFirstRun, sourceFirstRun)
      && sameValue(outputSecondParagraph, sourceSecondParagraph)
      && sameArray(changedPaths, expectedChangedPaths), {
      sourcePaths: source.paths,
      outputPaths: output.paths,
      changedPaths,
      targetPath: { source: source.target?.path, output: output.target?.path },
      notesPath: { source: source.targetNotesPath, output: output.targetNotesPath },
      untouchedPath: { source: source.untouched?.path, output: output.untouched?.path },
    }),
    gate("pptx-rich-notes-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("pptx-rich-notes-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("pptx-rich-notes-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("pptx-rich-notes-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("pptx-rich-notes-trace:fixed-topology-operation", "trace", operation.type === "title-and-rich-speaker-notes-run-edit"
      && operation.paragraphIndex === fixture.targetRun.paragraphIndex
      && operation.runIndex === fixture.targetRun.runIndex
      && operation.expectedRun?.text === fixture.targetRun.expectedText
      && operation.replacementRun?.text === fixture.targetRun.replacementText, {
      operation: audit?.operation || null,
    }),
    check("pptx-rich-notes-trace:typed-roundtrip", "trace", usedTypedRichNotesRoundTrip(commandText), {
      expected: "public PresentationFile importPptx/exportPptx calls or the integrity-protected published rich-speaker-notes workflow",
    }),
    check("pptx-rich-notes-trace:second-import", "trace", audit?.validation?.reimport?.ok === true
      && audit?.validation?.reimport?.richNotesFixedTopology === true
      && audit?.validation?.reimport?.targetRunExact === true
      && audit?.validation?.reimport?.notesIdPreserved === true, {
      validation: audit?.validation?.reimport || null,
    }),
  ];
}

/**
 * Grade the narrow non-visual mutation separately from title/notes. An Open
 * XML SDK save may canonicalize the target SlidePart, so this oracle checks
 * the semantic p:cSld name and requires every *other* part to remain byte
 * identical. Native page pixels must consequently stay stable on both slides.
 */
export function gradePptxSlideNameEvidence({ evidence, audit, commands }) {
  const fixture = PPTX_SLIDE_NAME_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const visual = stableVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const sourceTargets = source.slides.filter((slide) => slide.name === fixture.expectedName);
  const sourceTarget = sourceTargets.length === 1 ? sourceTargets[0] : null;
  const outputTarget = sourceTarget
    ? output.slides.find((slide) => slide.path === sourceTarget.path) || null
    : null;
  const sourceUntouched = source.slides.find((slide) => slide.name === fixture.untouchedSlideName) || null;
  const outputUntouched = sourceUntouched
    ? output.slides.find((slide) => slide.path === sourceUntouched.path) || null
    : null;
  const sourceSlideNames = source.slides.map((slide) => slide.name);
  const expectedSlideNames = sourceSlideNames.map((name) => name === fixture.expectedName ? fixture.replacementName : name);
  const outputSlideNames = output.slides.map((slide) => slide.name);
  const changedPaths = packageChanges(source, output);
  const expectedChangedPaths = sourceTarget ? [sourceTarget.path] : [];
  const commandText = commands.join("\n");
  const operation = audit?.operation && typeof audit.operation === "object" ? audit.operation : {};
  return [
    check("pptx-name-machine:canonical-fixture", "machine", sourceTargets.length === 1
      && source.slides.length === 2
      && sourceTarget?.title === PPTX_TITLE_NOTES_FIXTURE.originalTitle
      && sourceTarget?.background === PPTX_TITLE_NOTES_FIXTURE.targetBackground
      && source.targetNotes === PPTX_TITLE_NOTES_FIXTURE.originalNotes
      && sourceUntouched?.background === PPTX_TITLE_NOTES_FIXTURE.untouchedBackground
      && sourceTarget?.texts.includes(PPTX_TITLE_NOTES_FIXTURE.supportingText), {
      sourceTargets,
      sourceNotes: source.targetNotes,
      sourceSlideNames,
    }),
    check("pptx-name-machine:native-name-edited", "machine", outputTarget?.name === fixture.replacementName, {
      sourceTarget,
      outputTarget,
    }),
    check("pptx-name-machine:semantic-content-and-order-preserved", "machine", sameArray(outputSlideNames, expectedSlideNames)
      && sourceTarget?.path === outputTarget?.path
      && sourceTarget?.title === outputTarget?.title
      && sameArray(sourceTarget?.texts || [], outputTarget?.texts || [])
      && sourceTarget?.background === outputTarget?.background
      && source.targetNotes === output.targetNotes
      && sourceUntouched?.path === outputUntouched?.path
      && sourceUntouched?.title === outputUntouched?.title
      && sameArray(sourceUntouched?.texts || [], outputUntouched?.texts || [])
      && sourceUntouched?.background === outputUntouched?.background, {
      sourceSlideNames,
      expectedSlideNames,
      outputSlideNames,
      sourceTarget,
      outputTarget,
      sourceUntouched,
      outputUntouched,
    }),
    check("pptx-name-machine:only-target-slide-part-changed", "machine", sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
    }),
    check("pptx-name-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("pptx-name-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("pptx-name-visual:all-pages-pixel-stable", "visual", visual.targetStable && visual.untouchedStable, {
      visual: evidence.visual,
    }),
    gate("pptx-name-security:fixed-topology-and-non-target-byte-preservation", "security", sameArray(source.paths, output.paths)
      && sourceTarget?.path === outputTarget?.path
      && sourceUntouched?.path === outputUntouched?.path
      && source.partHashes[sourceUntouched?.path] === output.partHashes[outputUntouched?.path]
      && sameArray(changedPaths, expectedChangedPaths), {
      sourcePaths: source.paths,
      outputPaths: output.paths,
      changedPaths,
      targetPath: { source: sourceTarget?.path, output: outputTarget?.path },
      untouchedPath: { source: sourceUntouched?.path, output: outputUntouched?.path },
    }),
    gate("pptx-name-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("pptx-name-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("pptx-name-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("pptx-name-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("pptx-name-trace:source-bound-name-operation", "trace", /slide.*name|name.*slide/i.test(auditOperation(audit))
      && operation.sourcePart === sourceTarget?.path
      && operation.expectedName === fixture.expectedName
      && operation.replacementName === fixture.replacementName
      && operation.nativeAttribute === "p:cSld/@name", {
      operation: audit?.operation || null,
      expected: {
        sourcePart: sourceTarget?.path,
        expectedName: fixture.expectedName,
        replacementName: fixture.replacementName,
        nativeAttribute: "p:cSld/@name",
      },
    }),
    check("pptx-name-trace:typed-roundtrip", "trace", usedTypedSlideNameRoundTrip(commandText), {
      expected: "public PresentationFile importPptx/exportPptx calls or the integrity-protected published slide-name workflow",
    }),
    check("pptx-name-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

async function readAudit(workspace) {
  try {
    return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8"));
  } catch {
    return null;
  }
}

export async function gradePptxCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  if (!pptxGradedCaseIds.has(item.id)) return { supported: false };
  const isSlideNameCase = item.id === "pptx-source-bound-slide-name-edit";
  const isClosedLeafCloneCase = item.id === "pptx-closed-leaf-slide-clone";
  const fixture = isClosedLeafCloneCase
    ? PPTX_CLOSED_LEAF_CLONE_FIXTURE
    : isSlideNameCase ? PPTX_SLIDE_NAME_FIXTURE : PPTX_RICH_NOTES_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.presentationName);
  const outputPath = path.join(workspace, "outputs", isClosedLeafCloneCase
    ? "release-review-with-copy.pptx"
    : isSlideNameCase ? "launch-review-renamed.pptx" : "rich-notes-review-updated.pptx");
  const inspect = isClosedLeafCloneCase
    ? inspectClosedLeafClonePptx
    : isSlideNameCase ? inspectTitleNotesPptx : inspectRichNotesPptx;
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspect(sourcePath),
      inspect(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("pptx-machine:readable-output", "machine", false, { error: error.message }),
      gate("pptx-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "pptx-source"),
    renderOfficeFile(outputPath, "pptx-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler presentation rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }
  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = isClosedLeafCloneCase
    ? gradePptxClosedLeafCloneEvidence({ evidence, audit, commands, item })
    : isSlideNameCase
      ? gradePptxSlideNameEvidence({ evidence, audit, commands, item })
      : gradePptxRichNotesEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}
