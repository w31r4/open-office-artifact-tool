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
const INLINE_ACTIONS = new Set([
  "ppaction://hlinkshowjump?jump=nextslide",
  "ppaction://hlinkshowjump?jump=previousslide",
  "ppaction://hlinkshowjump?jump=firstslide",
  "ppaction://hlinkshowjump?jump=lastslide",
  "ppaction://hlinkshowjump?jump=endshow",
]);
const SLIDE_JUMP_ACTION = "ppaction://hlinksldjump";
const CUSTOM_SHOW_ACTION = /^ppaction:\/\/customshow\?id=([0-9]+)(?:&return=(true|false))?$/i;
const MAX_CLOSED_CHART_PARTS = 256;
const MAX_CLOSED_OLE_WORKBOOK_PARTS = 64;
const MAX_CLOSED_DIAGRAM_PARTS = 256;
const MAX_CLOSED_INK_CONTENT_PARTS = 256;
const XLSX_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const INK_CONTENT_TYPE = "application/inkml+xml";
const CUSTOM_XML_RELATIONSHIP_TYPES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml",
  "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml",
]);
const CLOSED_DIAGRAM_PARTS = [
  { attribute: "dm", suffix: "diagramData", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml" },
  { attribute: "lo", suffix: "diagramLayout", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml" },
  { attribute: "qs", suffix: "diagramQuickStyle", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml" },
  { attribute: "cs", suffix: "diagramColors", contentType: "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml" },
];
const OOXML_RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
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

function semanticXmlAttributes(tag) {
  return Object.fromEntries(Object.entries(xmlAttributes(tag)).filter(([name]) => !/^xmlns(?::|$)/i.test(name)));
}

function hasStandardInkMlRoot(xml) {
  const root = String(xml).match(/^\s*(?:<\?xml\b[^>]*>\s*)?<(?:(?<prefix>[A-Za-z_][\w.-]*):)?ink\b[^>]*>/i);
  if (!root) return false;
  const namespaceAttribute = root.groups?.prefix ? `xmlns:${root.groups.prefix}` : "xmlns";
  return xmlAttributes(root[0])[namespaceAttribute] === "http://www.w3.org/2003/InkML";
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
  return resolveRelationshipTargetFromPart("ppt/presentation.xml", target);
}

function resolvePackageRelationshipTargetFromPart(sourcePart, target) {
  if (typeof sourcePart !== "string" || sourcePart.startsWith("/") || sourcePart.includes("\\") || sourcePart.split("/").includes("..")) {
    throw new Error("Invalid OPC relationship source part: " + JSON.stringify(sourcePart));
  }
  if (typeof target !== "string" || !target || /[\\?#]/.test(target) || /%[0-9a-f]{2}/i.test(target)) {
    throw new Error("Unsafe OPC relationship target: " + JSON.stringify(target));
  }
  const partPath = target.startsWith("/")
    ? target.replace(/^\/+/, "")
    : path.posix.normalize(path.posix.join(sourcePart ? path.posix.dirname(sourcePart) : ".", target));
  if (!partPath || partPath === "." || partPath.startsWith("../") || partPath.split("/").includes("..")) {
    throw new Error("Unsafe OPC relationship target: " + JSON.stringify(target));
  }
  return partPath.replace(/^\.\//, "");
}

function resolveRelationshipTargetFromPart(sourcePart, target) {
  if (typeof sourcePart !== "string" || !sourcePart.startsWith("ppt/") || !sourcePart.endsWith(".xml")) {
    throw new Error("Invalid PPTX relationship source part: " + JSON.stringify(sourcePart));
  }
  const partPath = resolvePackageRelationshipTargetFromPart(sourcePart, target);
  if (!partPath.startsWith("ppt/")) {
    throw new Error("Unsafe PPTX relationship target: " + JSON.stringify(target));
  }
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

async function inspectCanonicalCustomShows(zip) {
  const [presentationXml, relationshipsXml] = await Promise.all([
    requiredZipBytes(zip, "ppt/presentation.xml").then((bytes) => Buffer.from(bytes).toString("utf8")),
    requiredZipBytes(zip, "ppt/_rels/presentation.xml.rels").then((bytes) => Buffer.from(bytes).toString("utf8")),
  ]);
  const slidePartByRelationshipId = new Map();
  for (const match of relationshipsXml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type?.toLowerCase().endsWith("/slide")) continue;
    if (attributes.TargetMode?.toLowerCase() === "external" || !attributes.Target) {
      throw new Error("Custom-show membership references a non-internal presentation slide relationship.");
    }
    slidePartByRelationshipId.set(attributes.Id, resolveRelationshipTarget(attributes.Target));
  }

  const listMatches = [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?custShowLst\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?custShowLst\s*>/gi)];
  const listOpeningCount = [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?custShowLst\b/gi)].length;
  if (!listMatches.length) {
    if (listOpeningCount) {
      throw new Error("Canonical duplicate workflow found a non-canonical custom-show list.");
    }
    return { shows: [], byNativeId: new Map(), fingerprint: "[]" };
  }
  if (listMatches.length !== 1 || listOpeningCount !== 1 || Object.keys(semanticXmlAttributes(listMatches[0][1])).length) {
    throw new Error("Canonical duplicate workflow requires exactly one unextended custom-show list.");
  }

  const showPattern = /<(?:[A-Za-z_][\w.-]*:)?custShow\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?custShow\s*>/gi;
  const listBody = listMatches[0][2];
  const showMatches = [...listBody.matchAll(showPattern)];
  if (!showMatches.length || listBody.replace(showPattern, "").trim()) {
    throw new Error("Canonical duplicate workflow found a non-canonical custom-show child.");
  }
  const names = new Set();
  const byNativeId = new Map();
  const shows = [];
  for (const match of showMatches) {
    const attributes = semanticXmlAttributes(match[1]);
    if (Object.keys(attributes).length !== 2 || !("name" in attributes) || !("id" in attributes)) {
      throw new Error("Canonical custom shows require exactly name and id attributes.");
    }
    const name = unescapeXml(attributes.name);
    const normalizedName = name.toLowerCase();
    const nativeId = Number(attributes.id);
    if (!name || names.has(normalizedName) || !/^[0-9]+$/.test(attributes.id) || !Number.isInteger(nativeId) || nativeId < 0 || nativeId > 0xffffffff || byNativeId.has(nativeId)) {
      throw new Error("Canonical custom shows require unique non-empty names and unsigned 32-bit native IDs.");
    }

    const slideList = match[2].match(/^\s*<(?:[A-Za-z_][\w.-]*:)?sldLst\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?sldLst\s*>\s*$/i);
    if (!slideList || Object.keys(semanticXmlAttributes(slideList[1])).length) throw new Error("Canonical custom shows require one unextended slide list.");
    const entryPattern = /<(?:[A-Za-z_][\w.-]*:)?sld\b([^>]*)\/\s*>/gi;
    const entryMatches = [...slideList[2].matchAll(entryPattern)];
    if (!entryMatches.length || slideList[2].replace(entryPattern, "").trim()) {
      throw new Error("Canonical custom shows require one or more relationship-only slide entries.");
    }
    const slideParts = entryMatches.map((entry) => {
      const entryAttributes = semanticXmlAttributes(entry[1]);
      if (Object.keys(entryAttributes).length !== 1 || !("r:id" in entryAttributes)) {
        throw new Error("Canonical custom-show slide entries require exactly one r:id attribute.");
      }
      const slidePart = slidePartByRelationshipId.get(entryAttributes["r:id"]);
      if (!slidePart || !zip.file(slidePart)) {
        throw new Error("Canonical custom-show membership references an unresolved SlidePart.");
      }
      return slidePart;
    });
    const show = { name, nativeId, slideParts };
    names.add(normalizedName);
    byNativeId.set(nativeId, show);
    shows.push(show);
  }
  return { shows, byNativeId, fingerprint: JSON.stringify(shows) };
}

function relationshipPartPath(partPath) {
  return path.posix.join(path.posix.dirname(partPath), "_rels", path.posix.basename(partPath) + ".rels");
}

function relationshipTypeMatches(entry, suffix) {
  return entry.type.toLowerCase().endsWith("/" + suffix.toLowerCase());
}

function exactlyOneRelationship(entries, suffix, label) {
  const matches = entries.filter((entry) => relationshipTypeMatches(entry, suffix));
  if (matches.length !== 1) {
    throw new Error(label + " must contain exactly one " + suffix + " relationship; found " + matches.length + ".");
  }
  return matches[0];
}

async function relationshipEntriesForPart(zip, sourcePart, { required = true } = {}) {
  const relationshipPart = relationshipPartPath(sourcePart);
  const file = zip.file(relationshipPart);
  if (!file) {
    if (!required) return [];
    throw new Error("PPTX part " + sourcePart + " has no relationship part.");
  }
  const xml = await file.async("text");
  const ids = new Set();
  const entries = [];
  for (const match of xml.matchAll(/<Relationship\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (!attributes.Id || !attributes.Type || ids.has(attributes.Id)) {
      throw new Error("PPTX relationship part " + relationshipPart + " has a missing or duplicate relationship identity.");
    }
    ids.add(attributes.Id);
    const external = attributes.TargetMode?.toLowerCase() === "external";
    if (!external && !attributes.Target) {
      throw new Error("PPTX relationship " + JSON.stringify(attributes.Id) + " in " + relationshipPart + " has no internal target.");
    }
    entries.push({
      id: attributes.Id,
      type: attributes.Type,
      target: attributes.Target || null,
      external,
      targetPart: external ? null : resolveRelationshipTargetFromPart(sourcePart, attributes.Target),
    });
  }
  return entries;
}

async function inspectCanonicalRunHyperlinks(zip, slidePart, customShows) {
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  if (/<(?:[A-Za-z_][\w.-]*:)?hlinkHover\b/i.test(slideXml)) {
    throw new Error("Canonical duplicate workflow does not accept hover hyperlinks in " + slidePart + ".");
  }
  const relationships = (await relationshipEntriesForPart(zip, slidePart))
    .filter((entry) => relationshipTypeMatches(entry, "hyperlink") || relationshipTypeMatches(entry, "slide"));
  const relationshipById = new Map(relationships.map((entry) => [entry.id, entry]));
  const usedRelationshipIds = new Set();
  const clicks = [];
  const customShowActions = [];
  for (const match of slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?hlinkClick\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    const relationshipId = attributes["r:id"] || "";
    const action = unescapeXml(attributes.action || "");
    if (!relationshipId) {
      const customShow = CUSTOM_SHOW_ACTION.exec(action);
      if (customShow) {
        const nativeId = Number(customShow[1]);
        const target = customShows.byNativeId.get(nativeId);
        if (!target || !Number.isInteger(nativeId) || nativeId > 0xffffffff) {
          throw new Error("Canonical duplicate workflow found a custom-show run action with an unresolved native ID in " + slidePart + ".");
        }
        customShowActions.push({
          nativeId,
          name: target.name,
          returnToSlide: customShow[2] ? customShow[2].toLowerCase() === "true" : null,
        });
      } else if (!INLINE_ACTIONS.has(action)) {
        throw new Error("Canonical duplicate workflow found an unsupported relationship-free click action in " + slidePart + ".");
      }
    } else {
      const relationship = relationshipById.get(relationshipId);
      if (!relationship) {
        throw new Error("Canonical duplicate workflow found an unresolved run hyperlink " + JSON.stringify(relationshipId) + " in " + slidePart + ".");
      }
      if (relationshipTypeMatches(relationship, "hyperlink")) {
        if (!relationship.external || action) throw new Error("External run hyperlinks must be external and have no click action.");
      } else if (relationshipTypeMatches(relationship, "slide")) {
        if (relationship.external || action.toLowerCase() !== SLIDE_JUMP_ACTION) {
          throw new Error("Internal run hyperlinks must target a SlidePart through ppaction://hlinksldjump.");
        }
        if (!/^ppt\/slides\/slide\d+\.xml$/i.test(relationship.targetPart || "") || !zip.file(relationship.targetPart)) {
          throw new Error("Internal run hyperlink target is not a retained canonical SlidePart.");
        }
      }
      usedRelationshipIds.add(relationshipId);
    }
    clicks.push(Object.fromEntries(Object.entries(attributes).sort(([left], [right]) => left.localeCompare(right))));
  }
  if (usedRelationshipIds.size !== relationshipById.size || [...relationshipById.keys()].some((id) => !usedRelationshipIds.has(id))) {
    throw new Error("Canonical duplicate workflow found an orphan or unmodeled hyperlink relationship in " + slidePart + ".");
  }
  const relationshipFingerprint = relationships
    .map((entry) => ({
      id: entry.id,
      type: entry.type,
      external: entry.external,
      target: entry.external ? unescapeXml(entry.target) : entry.targetPart,
    }))
    .sort((left, right) => left.id.localeCompare(right.id));
  return {
    clicks,
    relationships: relationshipFingerprint,
    relationshipCount: relationships.length,
    actionOnlyCount: clicks.filter((click) => !(click["r:id"] || "")).length,
    customShowCount: customShowActions.length,
    customShowActions,
    fingerprint: JSON.stringify({ clicks, relationships: relationshipFingerprint }),
  };
}

function requireInternalRelationship(entry, label) {
  if (entry.external || !entry.targetPart) {
    throw new Error(label + " must be an internal PPTX relationship.");
  }
  return entry;
}

async function assertNoChildRelationshipGraph(zip, partPath, label) {
  const entries = await relationshipEntriesForPart(zip, partPath, { required: false });
  if (entries.length) throw new Error(label + " must not have a child relationship graph.");
}

async function inspectNotesLeaf(zip, sourcePart, relationship) {
  const notesRelationship = requireInternalRelationship(relationship, "Speaker-notes leaf");
  const notesPart = notesRelationship.targetPart;
  if (!/^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(notesPart)) {
    throw new Error("Speaker-notes leaf must target one canonical NotesSlide part.");
  }
  await requiredZipBytes(zip, notesPart);
  const entries = await relationshipEntriesForPart(zip, notesPart);
  if (entries.length !== 2) {
    throw new Error("Speaker-notes leaf must contain exactly notesMaster and slide relationships.");
  }
  const notesMaster = requireInternalRelationship(exactlyOneRelationship(entries, "notesMaster", "Speaker-notes leaf"), "Speaker-notes master");
  const slide = requireInternalRelationship(exactlyOneRelationship(entries, "slide", "Speaker-notes leaf"), "Speaker-notes back-reference");
  if (!/^ppt\/notesMasters\/notesMaster\d+\.xml$/i.test(notesMaster.targetPart) || !zip.file(notesMaster.targetPart)) {
    throw new Error("Speaker-notes leaf must target one existing canonical NotesMaster part.");
  }
  if (slide.targetPart !== sourcePart) {
    throw new Error("Speaker-notes leaf back-reference must point at its source SlidePart.");
  }
  return {
    relationship: notesRelationship,
    part: notesPart,
    relationshipPart: relationshipPartPath(notesPart),
    notesMaster,
    slide,
  };
}

async function inspectLegacyCommentsLeaf(zip, relationship) {
  const commentsRelationship = requireInternalRelationship(relationship, "Legacy-comments leaf");
  const commentsPart = commentsRelationship.targetPart;
  if (!/^ppt\/comments\/comment\d+\.xml$/i.test(commentsPart)) {
    throw new Error("Legacy-comments leaf must target one canonical SlideComments part.");
  }
  await requiredZipBytes(zip, commentsPart);
  await assertNoChildRelationshipGraph(zip, commentsPart, "Legacy-comments leaf");
  const presentationRelationships = await relationshipEntriesForPart(zip, "ppt/presentation.xml");
  const authorCatalog = requireInternalRelationship(
    exactlyOneRelationship(presentationRelationships, "commentAuthors", "Presentation comment-author catalog"),
    "Presentation comment-author catalog",
  );
  if (authorCatalog.targetPart !== "ppt/commentAuthors.xml") {
    throw new Error("Legacy-comments leaf must use the canonical presentation-wide comment-author catalog.");
  }
  await requiredZipBytes(zip, authorCatalog.targetPart);
  await assertNoChildRelationshipGraph(zip, authorCatalog.targetPart, "Presentation comment-author catalog");
  return {
    relationship: commentsRelationship,
    part: commentsPart,
    authorCatalog,
  };
}

async function inspectClosedLeaves(zip, sourcePart, { allowClosedLeaves }) {
  const relationships = await relationshipEntriesForPart(zip, sourcePart);
  const notesRelationships = relationships.filter((entry) => relationshipTypeMatches(entry, "notesSlide"));
  const commentsRelationships = relationships.filter((entry) => relationshipTypeMatches(entry, "comments"));
  if (!allowClosedLeaves) {
    if (notesRelationships.length) {
      throw new Error("This duplicate workflow intentionally accepts no speaker-notes leaf unless allowClosedLeaves is explicitly true.");
    }
    if (commentsRelationships.length) {
      throw new Error("This duplicate workflow intentionally accepts no legacy-comments leaf unless allowClosedLeaves is explicitly true.");
    }
    return { profile: "canonical-inline-leaves-without-notes-or-comments", notes: null, comments: null };
  }
  if (notesRelationships.length > 1 || commentsRelationships.length > 1) {
    throw new Error("Closed-leaf duplication accepts at most one NotesSlide and one legacy-comments relationship.");
  }
  const notes = notesRelationships.length ? await inspectNotesLeaf(zip, sourcePart, notesRelationships[0]) : null;
  const comments = commentsRelationships.length ? await inspectLegacyCommentsLeaf(zip, commentsRelationships[0]) : null;
  return {
    profile: notes || comments
      ? "canonical-inline-leaves-with-closed-relationship-leaves"
      : "canonical-inline-leaves-without-notes-or-comments",
    notes,
    comments,
  };
}

async function inspectClosedChartParts(zip, slidePart) {
  const relationships = (await relationshipEntriesForPart(zip, slidePart))
    .filter((entry) => relationshipTypeMatches(entry, "chart"));
  if (relationships.length > MAX_CLOSED_CHART_PARTS) {
    throw new Error("Closed-chart duplication exceeds the " + MAX_CLOSED_CHART_PARTS + "-part budget.");
  }
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  const referenceIds = [];
  for (const match of slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?chart\b[^>]*>/gi)) {
    const relationshipId = xmlAttributes(match[0])["r:id"];
    if (!relationshipId) throw new Error("A chart reference in " + slidePart + " has no r:id.");
    referenceIds.push(relationshipId);
  }
  if (new Set(referenceIds).size !== referenceIds.length) {
    throw new Error("Closed-chart duplication requires every chart frame to own one unique relationship.");
  }
  const relationshipById = new Map(relationships.map((entry) => [entry.id, entry]));
  if (referenceIds.length !== relationships.length || referenceIds.some((id) => !relationshipById.has(id))) {
    throw new Error("Closed-chart duplication found an orphan or unmodeled ChartPart relationship in " + slidePart + ".");
  }
  const charts = [];
  for (const relationshipId of referenceIds) {
    const relationship = requireInternalRelationship(relationshipById.get(relationshipId), "Chart leaf");
    if (!/^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i.test(relationship.targetPart)) {
      throw new Error("Chart leaf must target one numbered ChartPart under ppt/charts or ppt/slides/charts.");
    }
    const bytes = await requiredZipBytes(zip, relationship.targetPart);
    await assertNoChildRelationshipGraph(zip, relationship.targetPart, "Chart leaf");
    charts.push({
      relationship,
      part: relationship.targetPart,
      sha256: sha256(bytes),
      bytes: bytes.length,
    });
  }
  return {
    count: charts.length,
    charts,
    fingerprint: JSON.stringify(charts.map((chart) => ({
      id: chart.relationship.id,
      type: chart.relationship.type,
      sha256: chart.sha256,
    }))),
  };
}

async function contentTypeForPart(zip, partPath) {
  const xml = Buffer.from(await requiredZipBytes(zip, "[Content_Types].xml")).toString("utf8");
  for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Override\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (String(attributes.PartName || "").replace(/^\/+/, "") === partPath) return attributes.ContentType || null;
  }
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Default\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    if (String(attributes.Extension || "").toLowerCase() === extension) return attributes.ContentType || null;
  }
  return null;
}

function relationshipSourcePartFromPath(relationshipPart) {
  if (relationshipPart === "_rels/.rels") return "";
  const match = /^(.*\/)?_rels\/([^/]+)\.rels$/i.exec(relationshipPart);
  if (!match) throw new Error("Invalid PPTX relationship-part path " + relationshipPart + ".");
  return `${match[1] || ""}${match[2]}`;
}

async function packageRelationshipInboundCount(zip, targetPart) {
  let count = 0;
  for (const relationshipPart of packagePartPaths(zip).filter((partPath) => partPath === "_rels/.rels" || /(?:^|\/)_rels\/[^/]+\.rels$/i.test(partPath))) {
    const sourcePart = relationshipSourcePartFromPath(relationshipPart);
    const xml = Buffer.from(await requiredZipBytes(zip, relationshipPart)).toString("utf8");
    for (const match of xml.matchAll(/<Relationship\b[^>]*>/gi)) {
      const attributes = xmlAttributes(match[0]);
      if (!String(attributes.Type || "").toLowerCase().endsWith("/package") || String(attributes.TargetMode || "").toLowerCase() === "external") continue;
      if (!attributes.Target) throw new Error("PPTX package relationship has no internal target in " + relationshipPart + ".");
      const resolved = resolvePackageRelationshipTargetFromPart(sourcePart, attributes.Target);
      if (resolved === targetPart) count++;
    }
  }
  return count;
}

async function inspectClosedOleWorkbookParts(zip, slidePart) {
  const relationships = await relationshipEntriesForPart(zip, slidePart);
  const packageRelationships = relationships.filter((entry) => relationshipTypeMatches(entry, "package"));
  if (packageRelationships.length > MAX_CLOSED_OLE_WORKBOOK_PARTS) {
    throw new Error("Embedded-XLSX OLE duplication exceeds the " + MAX_CLOSED_OLE_WORKBOOK_PARTS + "-part budget.");
  }
  const packageRelationshipById = new Map(packageRelationships.map((entry) => [entry.id, entry]));
  const imageRelationshipById = new Map(relationships.filter((entry) => relationshipTypeMatches(entry, "image")).map((entry) => [entry.id, entry]));
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  const frames = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?graphicFrame\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?graphicFrame\s*>/gi)]
    .filter((match) => /<(?:[A-Za-z_][\w.-]*:)?oleObj\b/i.test(match[0]));
  const allOleObjectCount = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?oleObj\b/gi)].length;
  if (frames.length !== allOleObjectCount || frames.length !== packageRelationships.length) {
    throw new Error("Embedded-XLSX OLE duplication found an orphan, nested, or non-canonical OLE/package relationship graph in " + slidePart + ".");
  }

  const workbooks = [];
  const usedPackageRelationshipIds = new Set();
  for (const frameMatch of frames) {
    const frame = frameMatch[0];
    const groupPrefix = slideXml.slice(0, frameMatch.index);
    const groupDepth = [...groupPrefix.matchAll(/<(\/)?(?:[A-Za-z_][\w.-]*:)?grpSp\b[^>]*>/gi)]
      .reduce((depth, match) => depth + (match[1] ? -1 : /\/\s*>$/.test(match[0]) ? 0 : 1), 0);
    if (groupDepth !== 0) throw new Error("Embedded-XLSX OLE duplication accepts only a top-level graphicFrame.");
    const oleTags = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?oleObj\b[^>]*>/gi)];
    const embeds = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?embed\b[^>]*\/?>/gi)];
    const links = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?link\b/gi)];
    const pictures = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?pic\b/gi)];
    const blips = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?blip\b[^>]*>/gi)];
    if (oleTags.length !== 1 || embeds.length !== 1 || links.length || pictures.length !== 1 || blips.length !== 1) {
      throw new Error("Embedded-XLSX OLE duplication requires one embed payload and one preview picture with no linked OLE object.");
    }
    const packageRelationshipId = xmlAttributes(oleTags[0][0])["r:id"];
    const blipAttributes = xmlAttributes(blips[0][0]);
    const previewRelationshipId = blipAttributes["r:embed"];
    const relationshipAttributes = [...frame.matchAll(/\br:(?:id|embed|link)\s*=\s*(["'])[^"']*\1/gi)];
    if (!packageRelationshipId || !previewRelationshipId || blipAttributes["r:link"] || relationshipAttributes.length !== 2) {
      throw new Error("Embedded-XLSX OLE frame must contain exactly its package r:id and preview r:embed bindings.");
    }
    const packageRelationship = packageRelationshipById.get(packageRelationshipId);
    const previewRelationship = imageRelationshipById.get(previewRelationshipId);
    if (!packageRelationship || packageRelationship.external || !packageRelationship.targetPart || !previewRelationship || previewRelationship.external || !previewRelationship.targetPart) {
      throw new Error("Embedded-XLSX OLE frame has an unresolved or external package/preview relationship.");
    }
    if (!usedPackageRelationshipIds.add(packageRelationshipId)) {
      throw new Error("Embedded-XLSX OLE package relationship is referenced by more than one frame.");
    }
    if (!packageRelationship.targetPart.toLowerCase().endsWith(".xlsx") || await contentTypeForPart(zip, packageRelationship.targetPart) !== XLSX_CONTENT_TYPE) {
      throw new Error("Embedded-XLSX OLE package must be one internal XLSX part with the standard SpreadsheetML content type.");
    }
    if (!/^ppt\/media\/[^/]+\.[a-z0-9]+$/i.test(previewRelationship.targetPart) || !String(await contentTypeForPart(zip, previewRelationship.targetPart)).toLowerCase().startsWith("image/")) {
      throw new Error("Embedded-XLSX OLE preview must be one internal PPTX ImagePart.");
    }
    await assertNoChildRelationshipGraph(zip, packageRelationship.targetPart, "Embedded-XLSX OLE package");
    if (await packageRelationshipInboundCount(zip, packageRelationship.targetPart) !== 1) {
      throw new Error("Embedded-XLSX OLE package must have exactly one inbound package relationship.");
    }
    const [workbookBytes, previewBytes] = await Promise.all([
      requiredZipBytes(zip, packageRelationship.targetPart),
      requiredZipBytes(zip, previewRelationship.targetPart),
    ]);
    if (!workbookBytes.length || !previewBytes.length) throw new Error("Embedded-XLSX OLE package and preview must be non-empty.");
    workbooks.push({
      relationship: packageRelationship,
      part: packageRelationship.targetPart,
      sha256: sha256(workbookBytes),
      bytes: workbookBytes.length,
      previewRelationship,
      previewPart: previewRelationship.targetPart,
      previewSha256: sha256(previewBytes),
    });
  }
  if (usedPackageRelationshipIds.size !== packageRelationshipById.size) {
    throw new Error("Embedded-XLSX OLE duplication found an orphan package relationship in " + slidePart + ".");
  }
  return {
    count: workbooks.length,
    workbooks,
    fingerprint: JSON.stringify(workbooks.map((workbook) => ({
      relationshipId: workbook.relationship.id,
      relationshipType: workbook.relationship.type,
      sourceSha256: workbook.sha256,
      previewRelationshipId: workbook.previewRelationship.id,
      previewRelationshipType: workbook.previewRelationship.type,
      previewPart: workbook.previewPart,
      previewSha256: workbook.previewSha256,
    }))),
  };
}

async function inspectClosedDiagramParts(zip, slidePart) {
  const relationships = await relationshipEntriesForPart(zip, slidePart);
  const diagramRelationships = relationships.filter((entry) =>
    CLOSED_DIAGRAM_PARTS.some((definition) => relationshipTypeMatches(entry, definition.suffix)));
  if (diagramRelationships.length > MAX_CLOSED_DIAGRAM_PARTS) {
    throw new Error("SmartArt duplication exceeds the " + MAX_CLOSED_DIAGRAM_PARTS + "-part budget.");
  }
  const relationshipById = new Map(diagramRelationships.map((entry) => [entry.id, entry]));
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  const relationshipPrefixes = new Set([...slideXml.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])([^"']+)\2/gi)]
    .filter((match) => OOXML_RELATIONSHIP_NAMESPACES.has(match[3]))
    .map((match) => match[1]));
  const frames = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?graphicFrame\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?graphicFrame\s*>/gi)]
    .filter((match) => /<(?:[A-Za-z_][\w.-]*:)?relIds\b/i.test(match[0]) && /drawingml\/2006\/diagram/i.test(match[0]));
  const relationshipRootCount = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?relIds\b/gi)].length;
  if (frames.length !== relationshipRootCount || diagramRelationships.length !== frames.length * CLOSED_DIAGRAM_PARTS.length) {
    throw new Error("SmartArt duplication found an orphan, nested, or non-canonical four-part diagram graph in " + slidePart + ".");
  }

  const usedRelationshipIds = new Set();
  const diagrams = [];
  for (const frameMatch of frames) {
    const frame = frameMatch[0];
    const groupPrefix = slideXml.slice(0, frameMatch.index);
    const groupDepth = [...groupPrefix.matchAll(/<(\/)?(?:[A-Za-z_][\w.-]*:)?grpSp\b[^>]*>/gi)]
      .reduce((depth, match) => depth + (match[1] ? -1 : /\/\s*>$/.test(match[0]) ? 0 : 1), 0);
    if (groupDepth !== 0) throw new Error("SmartArt duplication accepts only a top-level graphicFrame.");
    const relationshipRoots = [...frame.matchAll(/<(?:[A-Za-z_][\w.-]*:)?relIds\b[^>]*\/?\s*>/gi)];
    const relationshipAttributes = [...frame.matchAll(/\b([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)\s*=\s*(["'])[^"']*\3/gi)]
      .filter((match) => relationshipPrefixes.has(match[1]));
    if (relationshipRoots.length !== 1 || relationshipAttributes.length !== CLOSED_DIAGRAM_PARTS.length) {
      throw new Error("SmartArt frame must contain exactly one dgm:relIds root and its four relationship bindings.");
    }
    const attributes = xmlAttributes(relationshipRoots[0][0]);
    const roots = [];
    for (const definition of CLOSED_DIAGRAM_PARTS) {
      const matchingAttributes = Object.entries(attributes).filter(([name]) => {
        const [prefix, localName] = name.split(":");
        return localName === definition.attribute && relationshipPrefixes.has(prefix);
      });
      const relationshipId = matchingAttributes.length === 1 ? matchingAttributes[0][1] : null;
      const relationship = relationshipById.get(relationshipId);
      if (!relationshipId || !relationship || !relationshipTypeMatches(relationship, definition.suffix) ||
          !usedRelationshipIds.add(relationshipId)) {
        throw new Error("SmartArt frame has a missing, mistyped, duplicate, or shared r:" + definition.attribute + " relationship.");
      }
      requireInternalRelationship(relationship, "SmartArt " + definition.suffix + " leaf");
      if (!/^ppt\/(?:graphics|diagrams|slides\/diagrams)\/[^/]+\.xml$/i.test(relationship.targetPart)) {
        throw new Error("SmartArt " + definition.suffix + " must target one XML part under a canonical PPTX diagram directory.");
      }
      if (await contentTypeForPart(zip, relationship.targetPart) !== definition.contentType) {
        throw new Error("SmartArt " + definition.suffix + " has an unexpected content type.");
      }
      await assertNoChildRelationshipGraph(zip, relationship.targetPart, "SmartArt " + definition.suffix + " leaf");
      const bytes = await requiredZipBytes(zip, relationship.targetPart);
      if (!bytes.length) throw new Error("SmartArt " + definition.suffix + " part is empty.");
      roots.push({
        attribute: definition.attribute,
        relationship,
        part: relationship.targetPart,
        contentType: definition.contentType,
        sha256: sha256(bytes),
        bytes: bytes.length,
      });
    }
    diagrams.push({ roots });
  }
  if (usedRelationshipIds.size !== relationshipById.size) {
    throw new Error("SmartArt duplication found an orphan diagram relationship in " + slidePart + ".");
  }
  return {
    count: diagrams.length,
    partCount: usedRelationshipIds.size,
    diagrams,
    fingerprint: JSON.stringify(diagrams.map((diagram) => diagram.roots.map((root) => ({
      attribute: root.attribute,
      id: root.relationship.id,
      type: root.relationship.type,
      sha256: root.sha256,
    })))),
  };
}

async function inspectClosedInkContentParts(zip, slidePart) {
  const relationships = await relationshipEntriesForPart(zip, slidePart);
  const inkRelationships = relationships.filter((entry) => CUSTOM_XML_RELATIONSHIP_TYPES.has(entry.type));
  if (inkRelationships.length > MAX_CLOSED_INK_CONTENT_PARTS) {
    throw new Error("InkML duplication exceeds the " + MAX_CLOSED_INK_CONTENT_PARTS + "-part budget.");
  }
  const relationshipById = new Map(inkRelationships.map((entry) => [entry.id, entry]));
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  const relationshipPrefixes = new Set([...slideXml.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])([^"']+)\2/gi)]
    .filter((match) => OOXML_RELATIONSHIP_NAMESPACES.has(match[3]))
    .map((match) => match[1]));
  const elements = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?contentPart\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?contentPart\s*>/gi)];
  const openingCount = [...slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?contentPart\b/gi)].length;
  if (elements.length !== openingCount || elements.length !== inkRelationships.length) {
    throw new Error("InkML duplication found an orphan, nested, self-closing, or non-canonical contentPart relationship graph in " + slidePart + ".");
  }

  const usedRelationshipIds = new Set();
  const contents = [];
  for (const elementMatch of elements) {
    const element = elementMatch[0];
    const groupPrefix = slideXml.slice(0, elementMatch.index);
    const groupDepth = [...groupPrefix.matchAll(/<(\/)?(?:[A-Za-z_][\w.-]*:)?grpSp\b[^>]*>/gi)]
      .reduce((depth, match) => depth + (match[1] ? -1 : /\/\s*>$/.test(match[0]) ? 0 : 1), 0);
    if (groupDepth !== 0) throw new Error("InkML duplication accepts only a top-level p:contentPart.");
    const relationshipAttributes = [...element.matchAll(/\b([A-Za-z_][\w.-]*):([A-Za-z_][\w.-]*)\s*=\s*(["'])([^"']*)\3/gi)]
      .filter((match) => relationshipPrefixes.has(match[1]));
    if (relationshipAttributes.length !== 1 || relationshipAttributes[0][2] !== "id") {
      throw new Error("InkML contentPart must contain exactly one standard OOXML relationship binding.");
    }
    if ([...element.matchAll(/<(?:[A-Za-z_][\w.-]*:)?nvContentPartPr\b/gi)].length !== 1 ||
        [...element.matchAll(/<(?:[A-Za-z_][\w.-]*:)?xfrm\b/gi)].length !== 1 ||
        /<(?:[A-Za-z_][\w.-]*:)?extLst\b/i.test(element)) {
      throw new Error("InkML contentPart must contain one bounded non-visual record and transform with no extension list.");
    }
    const relationshipId = unescapeXml(relationshipAttributes[0][4]);
    const relationship = relationshipById.get(relationshipId);
    if (!relationship || !CUSTOM_XML_RELATIONSHIP_TYPES.has(relationship.type) || !usedRelationshipIds.add(relationshipId)) {
      throw new Error("InkML contentPart has a missing, mistyped, duplicate, or shared customXml relationship.");
    }
    requireInternalRelationship(relationship, "InkML contentPart");
    if (!/^ppt\/customXml\/[^/]+\.xml$/i.test(relationship.targetPart)) {
      throw new Error("InkML contentPart must target one XML part under ppt/customXml.");
    }
    if (await contentTypeForPart(zip, relationship.targetPart) !== INK_CONTENT_TYPE) {
      throw new Error("InkML contentPart has an unexpected content type.");
    }
    await assertNoChildRelationshipGraph(zip, relationship.targetPart, "InkML content part");
    const bytes = await requiredZipBytes(zip, relationship.targetPart);
    const xml = Buffer.from(bytes).toString("utf8");
    if (!bytes.length || !hasStandardInkMlRoot(xml)) {
      throw new Error("InkML content part must contain a non-empty standard InkML root.");
    }
    contents.push({
      relationship,
      part: relationship.targetPart,
      contentType: INK_CONTENT_TYPE,
      sha256: sha256(bytes),
      bytes: bytes.length,
    });
  }
  if (usedRelationshipIds.size !== relationshipById.size) {
    throw new Error("InkML duplication found an orphan customXml relationship in " + slidePart + ".");
  }
  return {
    count: contents.length,
    contents,
    fingerprint: JSON.stringify(contents.map((content) => ({
      id: content.relationship.id,
      type: content.relationship.type,
      sha256: content.sha256,
    }))),
  };
}

function modelDiagramBindings(slide) {
  return (slide.nativeObjects?.items || [])
    .filter((object) => object.nativeKind === "diagram")
    .map((object) => {
      if (!/^<(?:[A-Za-z_][\w.-]*:)?graphicFrame(?:\s|>)/.test(String(object.rawXml || "").trimStart())) {
        throw new Error("Selected slide contains a SmartArt object outside the top-level graphicFrame clone profile.");
      }
      const references = new Map((object.relationshipReferences || [])
        .filter((reference) => OOXML_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")))
        .map((reference) => [String(reference.attribute || "").split(":").at(-1), reference.id]));
      const roots = new Map((object.rootRelationships || []).map((relationship) => [relationship.id, relationship]));
      const parts = new Map((object.parts || []).map((part) => [part.path, part]));
      if (references.size !== 4 || roots.size !== 4 || parts.size !== 4) {
        throw new Error("Selected slide contains a SmartArt object outside the closed four-part model profile.");
      }
      return CLOSED_DIAGRAM_PARTS.map((definition) => {
        const relationshipId = references.get(definition.attribute);
        const relationship = roots.get(relationshipId);
        const partPath = relationship && !String(relationship.targetMode || "").toLowerCase().includes("external")
          ? resolveRelationshipTargetFromPart(object.sourcePart, relationship.target)
          : null;
        const part = partPath ? parts.get(partPath) : null;
        if (!relationshipId || !relationship || !relationship.type.endsWith("/" + definition.suffix) ||
            !part || part.contentType !== definition.contentType || part.relationships.length ||
            !/^[0-9a-f]{64}$/i.test(part.sourceSha256 || "")) {
          throw new Error("Selected slide SmartArt model does not match its closed " + definition.attribute + " binding.");
        }
        return { relationshipId, partPath, contentType: part.contentType, sourceSha256: part.sourceSha256 };
      });
    });
}

function assertModelDiagramBindings(bindings, inspected) {
  if (bindings.length !== inspected.count) throw new Error("Imported model and independent OPC inspection disagree on SmartArt count.");
  const inspectedByKey = new Map(inspected.diagrams.map((diagram) => [
    diagram.roots.map((root) => root.relationship.id).sort().join("\0"),
    diagram,
  ]));
  for (const binding of bindings) {
    const key = binding.map((root) => root.relationshipId).sort().join("\0");
    const diagram = inspectedByKey.get(key);
    if (!diagram) throw new Error("Imported model and independent OPC inspection disagree on a SmartArt relationship set.");
    const byRelationshipId = new Map(diagram.roots.map((root) => [root.relationship.id, root]));
    for (const root of binding) {
      const inspectedRoot = byRelationshipId.get(root.relationshipId);
      if (!inspectedRoot || root.partPath !== inspectedRoot.part || root.contentType !== inspectedRoot.contentType || root.sourceSha256 !== inspectedRoot.sha256) {
        throw new Error("Imported model and independent OPC inspection disagree on a SmartArt part binding.");
      }
    }
  }
}

function modelInkContentBindings(slide) {
  return (slide.nativeObjects?.items || [])
    .filter((object) => object.nativeKind === "contentPart")
    .map((object) => {
      if (!/^<(?:[A-Za-z_][\w.-]*:)?contentPart(?:\s|>)/.test(String(object.rawXml || "").trimStart())) {
        throw new Error("Selected slide contains an InkML object outside the top-level contentPart clone profile.");
      }
      const references = (object.relationshipReferences || []).filter((reference) =>
        OOXML_RELATIONSHIP_NAMESPACES.has(String(reference.namespaceUri || "")));
      if (references.length !== 1 || String(references[0].attribute || "").split(":").at(-1) !== "id") {
        throw new Error("Selected slide InkML model must expose exactly one standard relationship reference.");
      }
      const relationshipId = references[0].id;
      const relationship = (object.rootRelationships || []).find((candidate) => candidate.id === relationshipId);
      const part = (object.parts || [])[0];
      if ((object.rootRelationships || []).length !== 1 || (object.parts || []).length !== 1 ||
          !relationship || !CUSTOM_XML_RELATIONSHIP_TYPES.has(relationship.type) ||
          String(relationship.targetMode || "").toLowerCase() === "external" ||
          !part || part.contentType !== INK_CONTENT_TYPE || !part.path ||
          !Array.isArray(part.relationships) || part.relationships.length ||
          !/^[0-9a-f]{64}$/i.test(part.sourceSha256 || "")) {
        throw new Error("Selected slide InkML model does not match one closed CustomXmlPart binding.");
      }
      return {
        relationshipId,
        relationshipType: relationship.type,
        partPath: part.path,
        contentType: part.contentType,
        sourceSha256: part.sourceSha256,
      };
    });
}

function assertModelInkContentBindings(bindings, inspected, expectedParts = new Map()) {
  if (bindings.length !== inspected.count) throw new Error("Imported model and independent OPC inspection disagree on InkML contentPart count.");
  const byRelationshipId = new Map(bindings.map((binding) => [binding.relationshipId, binding]));
  for (const content of inspected.contents) {
    const binding = byRelationshipId.get(content.relationship.id);
    const expectedPart = expectedParts.get(content.relationship.id) || content.part;
    if (!binding || binding.relationshipType !== content.relationship.type || binding.partPath !== expectedPart ||
        binding.contentType !== INK_CONTENT_TYPE || binding.sourceSha256 !== content.sha256) {
      throw new Error("Imported model and independent OPC inspection disagree on an InkML content-part binding.");
    }
  }
}

function closedLeafSemanticSnapshot(slide) {
  return JSON.stringify({
    speakerNotes: typeof slide.speakerNotes?.text === "string" ? slide.speakerNotes.text : null,
    legacyComments: (slide.comments?.items || []).map((thread) => ({
      nativeFormat: thread.nativeFormat || null,
      targetId: thread.targetId || null,
      position: thread.position || null,
      comments: (thread.comments || []).map((comment) => ({
        author: comment.author || null,
        text: comment.text || null,
        created: comment.created || null,
      })),
    })),
  });
}

function modelOleWorkbookBindings(slide) {
  return (slide.nativeObjects?.items || [])
    .filter((object) => object.nativeKind === "oleObject")
    .map((object) => {
      if (!object.oleWorkbook) throw new Error("Selected slide contains an OLE object outside the uniquely bound embedded-XLSX clone profile.");
      return {
        relationshipId: object.oleWorkbook.relationshipId,
        partPath: object.oleWorkbook.partPath,
        sourceSha256: object.oleWorkbook.sourceSha256,
        contentType: object.oleWorkbook.contentType,
      };
    })
    .sort((left, right) => left.relationshipId.localeCompare(right.relationshipId));
}

function assertModelOleWorkbookBindings(bindings, inspected, expectedParts = new Map()) {
  if (bindings.length !== inspected.count) throw new Error("Imported model and independent OPC inspection disagree on embedded-XLSX OLE count.");
  const byRelationshipId = new Map(bindings.map((binding) => [binding.relationshipId, binding]));
  for (const workbook of inspected.workbooks) {
    const binding = byRelationshipId.get(workbook.relationship.id);
    const expectedPart = expectedParts.get(workbook.relationship.id) || workbook.part;
    if (!binding || binding.partPath !== expectedPart || binding.sourceSha256 !== workbook.sha256 || binding.contentType !== XLSX_CONTENT_TYPE) {
      throw new Error("Imported model and independent OPC inspection disagree on an embedded-XLSX OLE source binding.");
    }
  }
}

function canonicalCloneSnapshot(slide, slidePartById = new Map()) {
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
    const oleObject = value.kind === "nativeObject" && value.nativeKind === "oleObject";
    const result = {};
    for (const [key, item] of Object.entries(value)) {
      if (key === "id" || key === "layoutId") continue;
      if ((key === "startTargetId" || key === "endTargetId") && typeof item === "string") {
        result[key] = locationById.get(item) || "unresolved:" + item;
      } else if (key === "slideId" && typeof item === "string") {
        result[key] = slidePartById.get(item) || "unresolved:" + item;
      } else if (oleObject && key === "embeddedWorkbook" && item && typeof item === "object") {
        result[key] = { ...normalize(item), partPath: "<independent-clone-local-xlsx>" };
      } else {
        result[key] = normalize(item);
      }
    }
    return result;
  };
  collectIds(proto, "slide");
  return JSON.stringify(normalize(proto));
}

function canonicalCustomShowSnapshot(presentation, slidePartById) {
  return JSON.stringify((presentation.customShows?.items || []).map((show) => ({
    name: show.name,
    nativeId: show.nativeId,
    slideParts: show.slideIds.map((slideId) => slidePartById.get(slideId) || "unresolved:" + slideId),
  })));
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

async function assertClosedLeafClonePackage(sourceZip, outputZip, sourceLeaves, cloneLeaves, sourcePart, clonePart) {
  if (Boolean(sourceLeaves.notes) !== Boolean(cloneLeaves.notes) || Boolean(sourceLeaves.comments) !== Boolean(cloneLeaves.comments)) {
    throw new Error("PPTX duplicate did not retain the exact closed-leaf profile of its source SlidePart.");
  }
  const validation = { speakerNotes: null, legacyComments: null };
  if (sourceLeaves.notes) {
    const sourceNotes = sourceLeaves.notes;
    const cloneNotes = cloneLeaves.notes;
    if (sourceNotes.relationship.id !== cloneNotes.relationship.id || sourceNotes.relationship.type !== cloneNotes.relationship.type) {
      throw new Error("PPTX duplicate changed the source SlidePart NotesSlide relationship identity.");
    }
    if (sourceNotes.part === cloneNotes.part || sourceZip.file(cloneNotes.part)) {
      throw new Error("PPTX duplicate did not allocate a distinct NotesSlide part for the clone.");
    }
    if (sourceNotes.notesMaster.id !== cloneNotes.notesMaster.id ||
        sourceNotes.notesMaster.type !== cloneNotes.notesMaster.type ||
        sourceNotes.notesMaster.targetPart !== cloneNotes.notesMaster.targetPart) {
      throw new Error("PPTX duplicate did not retain the immutable NotesMaster relationship.");
    }
    if (sourceNotes.slide.id !== cloneNotes.slide.id || sourceNotes.slide.type !== cloneNotes.slide.type ||
        sourceNotes.slide.targetPart !== sourcePart || cloneNotes.slide.targetPart !== clonePart) {
      throw new Error("PPTX duplicate did not repoint the cloned NotesSlide back-reference at the clone SlidePart.");
    }
    const [sourceNotesBytes, cloneNotesBytes] = await Promise.all([
      requiredZipBytes(sourceZip, sourceNotes.part),
      requiredZipBytes(outputZip, cloneNotes.part),
    ]);
    if (!Buffer.from(sourceNotesBytes).equals(Buffer.from(cloneNotesBytes))) {
      throw new Error("PPTX duplicate changed NotesSlide XML instead of verbatim-copying the accepted leaf.");
    }
    validation.speakerNotes = {
      sourcePart: sourceNotes.part,
      clonePart: cloneNotes.part,
      sourceRelationshipPart: sourceNotes.relationshipPart,
      cloneRelationshipPart: cloneNotes.relationshipPart,
      notesMasterPart: sourceNotes.notesMaster.targetPart,
      notesXmlByteIdentical: true,
      notesMasterShared: true,
      cloneBackReferencePointsAtClone: true,
    };
  }
  if (sourceLeaves.comments) {
    const sourceComments = sourceLeaves.comments;
    const cloneComments = cloneLeaves.comments;
    if (sourceComments.relationship.id !== cloneComments.relationship.id || sourceComments.relationship.type !== cloneComments.relationship.type) {
      throw new Error("PPTX duplicate changed the source SlidePart legacy-comments relationship identity.");
    }
    if (sourceComments.part === cloneComments.part || sourceZip.file(cloneComments.part)) {
      throw new Error("PPTX duplicate did not allocate a distinct SlideComments part for the clone.");
    }
    if (sourceComments.authorCatalog.id !== cloneComments.authorCatalog.id ||
        sourceComments.authorCatalog.type !== cloneComments.authorCatalog.type ||
        sourceComments.authorCatalog.targetPart !== cloneComments.authorCatalog.targetPart) {
      throw new Error("PPTX duplicate did not retain the immutable comment-author catalog relationship.");
    }
    const [sourceCommentsBytes, cloneCommentsBytes] = await Promise.all([
      requiredZipBytes(sourceZip, sourceComments.part),
      requiredZipBytes(outputZip, cloneComments.part),
    ]);
    if (!Buffer.from(sourceCommentsBytes).equals(Buffer.from(cloneCommentsBytes))) {
      throw new Error("PPTX duplicate changed SlideComments XML instead of verbatim-copying the accepted leaf.");
    }
    validation.legacyComments = {
      sourcePart: sourceComments.part,
      clonePart: cloneComments.part,
      commentAuthorsPart: sourceComments.authorCatalog.targetPart,
      commentsXmlByteIdentical: true,
      commentAuthorsShared: true,
    };
  }
  return validation;
}

async function assertClosedChartClonePackage(sourceZip, outputZip, sourceCharts, cloneCharts) {
  if (sourceCharts.count !== cloneCharts.count) {
    throw new Error("PPTX duplicate changed the number of closed ChartPart leaves.");
  }
  const cloneByRelationshipId = new Map(cloneCharts.charts.map((chart) => [chart.relationship.id, chart]));
  const validation = [];
  for (const sourceChart of sourceCharts.charts) {
    const cloneChart = cloneByRelationshipId.get(sourceChart.relationship.id);
    if (!cloneChart || cloneChart.relationship.type !== sourceChart.relationship.type) {
      throw new Error("PPTX duplicate changed a source SlidePart chart relationship identity or type.");
    }
    if (sourceChart.part === cloneChart.part || sourceZip.file(cloneChart.part)) {
      throw new Error("PPTX duplicate must allocate a distinct ChartPart for every cloned chart.");
    }
    const [sourceBytes, cloneBytes] = await Promise.all([
      requiredZipBytes(sourceZip, sourceChart.part),
      requiredZipBytes(outputZip, cloneChart.part),
    ]);
    if (!Buffer.from(sourceBytes).equals(Buffer.from(cloneBytes))) {
      throw new Error("PPTX duplicate changed ChartPart XML instead of byte-copying the accepted closed leaf.");
    }
    validation.push({
      relationshipId: sourceChart.relationship.id,
      sourcePart: sourceChart.part,
      clonePart: cloneChart.part,
      chartXmlByteIdentical: true,
      sourceSha256: sourceChart.sha256,
    });
  }
  return {
    count: validation.length,
    independentParts: true,
    allPayloadsByteIdentical: true,
    parts: validation,
  };
}

async function assertClosedOleWorkbookClonePackage(sourceZip, outputZip, sourceOleWorkbooks, cloneOleWorkbooks) {
  if (sourceOleWorkbooks.count !== cloneOleWorkbooks.count) {
    throw new Error("PPTX duplicate changed the number of embedded-XLSX OLE leaves.");
  }
  const cloneByRelationshipId = new Map(cloneOleWorkbooks.workbooks.map((workbook) => [workbook.relationship.id, workbook]));
  const validation = [];
  for (const sourceWorkbook of sourceOleWorkbooks.workbooks) {
    const cloneWorkbook = cloneByRelationshipId.get(sourceWorkbook.relationship.id);
    if (!cloneWorkbook || cloneWorkbook.relationship.type !== sourceWorkbook.relationship.type) {
      throw new Error("PPTX duplicate changed an OLE package relationship identity or type.");
    }
    if (sourceWorkbook.part === cloneWorkbook.part || sourceZip.file(cloneWorkbook.part)) {
      throw new Error("PPTX duplicate must allocate a distinct EmbeddedPackagePart for every cloned XLSX workbook.");
    }
    if (sourceWorkbook.sha256 !== cloneWorkbook.sha256 || sourceWorkbook.bytes !== cloneWorkbook.bytes) {
      throw new Error("PPTX duplicate changed embedded XLSX bytes instead of copying the accepted closed package.");
    }
    if (sourceWorkbook.previewRelationship.id !== cloneWorkbook.previewRelationship.id ||
        sourceWorkbook.previewRelationship.type !== cloneWorkbook.previewRelationship.type ||
        sourceWorkbook.previewPart !== cloneWorkbook.previewPart ||
        sourceWorkbook.previewSha256 !== cloneWorkbook.previewSha256) {
      throw new Error("PPTX duplicate did not retain the exact immutable OLE preview ImagePart binding.");
    }
    const [sourceBytes, cloneBytes] = await Promise.all([
      requiredZipBytes(sourceZip, sourceWorkbook.part),
      requiredZipBytes(outputZip, cloneWorkbook.part),
    ]);
    if (!Buffer.from(sourceBytes).equals(Buffer.from(cloneBytes))) {
      throw new Error("PPTX duplicate did not byte-copy the accepted embedded XLSX package.");
    }
    validation.push({
      relationshipId: sourceWorkbook.relationship.id,
      sourcePart: sourceWorkbook.part,
      clonePart: cloneWorkbook.part,
      contentType: XLSX_CONTENT_TYPE,
      sourceSha256: sourceWorkbook.sha256,
      workbookBytesByteIdentical: true,
      independentPackagePart: true,
      previewRelationshipId: sourceWorkbook.previewRelationship.id,
      previewPart: sourceWorkbook.previewPart,
      previewShared: true,
    });
  }
  return {
    count: validation.length,
    independentParts: true,
    allPayloadsByteIdentical: true,
    previewPartsShared: true,
    parts: validation,
  };
}

async function assertClosedDiagramClonePackage(sourceZip, outputZip, sourceDiagrams, cloneDiagrams) {
  if (sourceDiagrams.count !== cloneDiagrams.count || sourceDiagrams.partCount !== cloneDiagrams.partCount) {
    throw new Error("PPTX duplicate changed the number of closed SmartArt frames or diagram parts.");
  }
  const cloneRootsByRelationshipId = new Map(cloneDiagrams.diagrams.flatMap((diagram) =>
    diagram.roots.map((root) => [root.relationship.id, root])));
  const validation = [];
  for (const sourceDiagram of sourceDiagrams.diagrams) {
    for (const sourceRoot of sourceDiagram.roots) {
      const cloneRoot = cloneRootsByRelationshipId.get(sourceRoot.relationship.id);
      if (!cloneRoot || cloneRoot.attribute !== sourceRoot.attribute || cloneRoot.relationship.type !== sourceRoot.relationship.type ||
          cloneRoot.contentType !== sourceRoot.contentType) {
        throw new Error("PPTX duplicate changed a SmartArt relationship identity, role, type, or content type.");
      }
      if (sourceRoot.part === cloneRoot.part || sourceZip.file(cloneRoot.part)) {
        throw new Error("PPTX duplicate must allocate a distinct part for every cloned SmartArt root.");
      }
      const [sourceBytes, cloneBytes] = await Promise.all([
        requiredZipBytes(sourceZip, sourceRoot.part),
        requiredZipBytes(outputZip, cloneRoot.part),
      ]);
      if (!Buffer.from(sourceBytes).equals(Buffer.from(cloneBytes)) || sourceRoot.sha256 !== cloneRoot.sha256) {
        throw new Error("PPTX duplicate changed SmartArt XML instead of byte-copying the accepted closed part.");
      }
      validation.push({
        attribute: sourceRoot.attribute,
        relationshipId: sourceRoot.relationship.id,
        relationshipType: sourceRoot.relationship.type,
        sourcePart: sourceRoot.part,
        clonePart: cloneRoot.part,
        contentType: sourceRoot.contentType,
        sourceSha256: sourceRoot.sha256,
        diagramXmlByteIdentical: true,
        independentPart: true,
      });
    }
  }
  return {
    count: sourceDiagrams.count,
    partCount: validation.length,
    independentParts: true,
    allPayloadsByteIdentical: true,
    parts: validation,
  };
}

async function assertClosedInkContentClonePackage(sourceZip, outputZip, sourceContents, cloneContents) {
  if (sourceContents.count !== cloneContents.count) {
    throw new Error("PPTX duplicate changed the number of closed InkML content parts.");
  }
  const cloneByRelationshipId = new Map(cloneContents.contents.map((content) => [content.relationship.id, content]));
  const validation = [];
  for (const sourceContent of sourceContents.contents) {
    const cloneContent = cloneByRelationshipId.get(sourceContent.relationship.id);
    if (!cloneContent || cloneContent.relationship.type !== sourceContent.relationship.type ||
        cloneContent.contentType !== sourceContent.contentType) {
      throw new Error("PPTX duplicate changed an InkML relationship identity, type, or content type.");
    }
    if (sourceContent.part === cloneContent.part || sourceZip.file(cloneContent.part) ||
        !/^ppt\/customXml\/item\d+\.xml$/i.test(cloneContent.part)) {
      throw new Error("PPTX duplicate must allocate a distinct canonical CustomXmlPart for every cloned InkML object.");
    }
    const [sourceBytes, cloneBytes] = await Promise.all([
      requiredZipBytes(sourceZip, sourceContent.part),
      requiredZipBytes(outputZip, cloneContent.part),
    ]);
    if (!Buffer.from(sourceBytes).equals(Buffer.from(cloneBytes)) || sourceContent.sha256 !== cloneContent.sha256) {
      throw new Error("PPTX duplicate changed InkML XML instead of byte-copying the accepted closed part.");
    }
    validation.push({
      relationshipId: sourceContent.relationship.id,
      relationshipType: sourceContent.relationship.type,
      sourcePart: sourceContent.part,
      clonePart: cloneContent.part,
      contentType: sourceContent.contentType,
      sourceSha256: sourceContent.sha256,
      inkXmlByteIdentical: true,
      independentPart: true,
    });
  }
  return {
    count: validation.length,
    independentParts: true,
    allPayloadsByteIdentical: true,
    parts: validation,
  };
}

async function assertDuplicatePackageScope(sourceBytes, outputBytes, sourceIndex, { allowClosedLeaves }) {
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
  const [sourceCustomShows, outputCustomShows] = await Promise.all([
    inspectCanonicalCustomShows(sourceZip),
    inspectCanonicalCustomShows(outputZip),
  ]);
  if (sourceCustomShows.fingerprint !== outputCustomShows.fingerprint) {
    throw new Error("PPTX duplicate changed custom-show identity or membership instead of retaining the presentation-wide catalog.");
  }
  const [sourceLeaves, cloneLeaves] = await Promise.all([
    inspectClosedLeaves(sourceZip, sourcePart, { allowClosedLeaves }),
    inspectClosedLeaves(outputZip, clonePart, { allowClosedLeaves }),
  ]);
  const [sourceRunHyperlinks, cloneRunHyperlinks] = await Promise.all([
    inspectCanonicalRunHyperlinks(sourceZip, sourcePart, sourceCustomShows),
    inspectCanonicalRunHyperlinks(outputZip, clonePart, outputCustomShows),
  ]);
  const [sourceCharts, cloneCharts] = await Promise.all([
    inspectClosedChartParts(sourceZip, sourcePart),
    inspectClosedChartParts(outputZip, clonePart),
  ]);
  const [sourceOleWorkbooks, cloneOleWorkbooks] = await Promise.all([
    inspectClosedOleWorkbookParts(sourceZip, sourcePart),
    inspectClosedOleWorkbookParts(outputZip, clonePart),
  ]);
  const [sourceDiagrams, cloneDiagrams] = await Promise.all([
    inspectClosedDiagramParts(sourceZip, sourcePart),
    inspectClosedDiagramParts(outputZip, clonePart),
  ]);
  const [sourceInkContents, cloneInkContents] = await Promise.all([
    inspectClosedInkContentParts(sourceZip, sourcePart),
    inspectClosedInkContentParts(outputZip, clonePart),
  ]);
  if (sourceRunHyperlinks.fingerprint !== cloneRunHyperlinks.fingerprint) {
    throw new Error("PPTX duplicate changed the canonical run-hyperlink XML/relationship graph.");
  }
  const closedLeaves = await assertClosedLeafClonePackage(sourceZip, outputZip, sourceLeaves, cloneLeaves, sourcePart, clonePart);
  const chartParts = await assertClosedChartClonePackage(sourceZip, outputZip, sourceCharts, cloneCharts);
  const oleWorkbookParts = await assertClosedOleWorkbookClonePackage(sourceZip, outputZip, sourceOleWorkbooks, cloneOleWorkbooks);
  const diagramParts = await assertClosedDiagramClonePackage(sourceZip, outputZip, sourceDiagrams, cloneDiagrams);
  const inkContentParts = await assertClosedInkContentClonePackage(sourceZip, outputZip, sourceInkContents, cloneInkContents);
  const newParts = outputParts.filter((partPath) => !sourceZip.file(partPath));
  const expectedNewParts = [clonePart, relationshipPartPath(clonePart)];
  expectedNewParts.push(...cloneCharts.charts.map((chart) => chart.part));
  expectedNewParts.push(...cloneOleWorkbooks.workbooks.map((workbook) => workbook.part));
  expectedNewParts.push(...cloneDiagrams.diagrams.flatMap((diagram) => diagram.roots.map((root) => root.part)));
  expectedNewParts.push(...cloneInkContents.contents.map((content) => content.part));
  if (cloneLeaves.notes) expectedNewParts.push(cloneLeaves.notes.part, cloneLeaves.notes.relationshipPart);
  if (cloneLeaves.comments) expectedNewParts.push(cloneLeaves.comments.part);
  expectedNewParts.sort();
  if (new Set(expectedNewParts).size !== expectedNewParts.length) {
    throw new Error("PPTX duplicate calculated an ambiguous closed-leaf package delta.");
  }
  if (JSON.stringify(newParts) !== JSON.stringify(expectedNewParts)) {
    throw new Error("PPTX duplicate created an unexpected package part outside the selected source-bound clone profile.");
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
    profile: sourceInkContents.count
      ? sourceDiagrams.count || sourceCharts.count || sourceOleWorkbooks.count || sourceLeaves.notes || sourceLeaves.comments
        ? "canonical-inline-leaves-with-closed-inkml-and-other-relationship-leaves"
        : "canonical-inline-leaves-with-closed-inkml-leaves"
      : sourceDiagrams.count
      ? sourceCharts.count || sourceOleWorkbooks.count || sourceLeaves.notes || sourceLeaves.comments
        ? "canonical-inline-leaves-with-closed-smartart-and-other-relationship-leaves"
        : "canonical-inline-leaves-with-closed-smartart-leaves"
      : sourceOleWorkbooks.count
      ? sourceCharts.count || sourceLeaves.notes || sourceLeaves.comments
        ? "canonical-inline-leaves-with-closed-relationship-chart-or-ole-leaves"
        : "canonical-inline-leaves-with-closed-ole-workbook-leaves"
      : sourceCharts.count
        ? sourceLeaves.notes || sourceLeaves.comments
          ? "canonical-inline-leaves-with-closed-relationship-and-chart-leaves"
          : "canonical-inline-leaves-with-closed-chart-leaves"
        : sourceLeaves.profile,
    outputSlideParts,
    runHyperlinks: {
      relationshipCount: sourceRunHyperlinks.relationshipCount,
      actionOnlyCount: sourceRunHyperlinks.actionOnlyCount,
      customShowCount: sourceRunHyperlinks.customShowCount,
      customShowActions: sourceRunHyperlinks.customShowActions,
      exactSourceGraphRetained: true,
    },
    customShows: {
      count: sourceCustomShows.shows.length,
      shows: sourceCustomShows.shows,
      exactSourceMembershipRetained: true,
    },
    chartParts,
    oleWorkbookParts,
    diagramParts,
    inkContentParts,
    closedLeaves,
  };
}

export async function duplicatePptxSlide({ inputPath, outputPath, auditPath, expectedName, allowClosedLeaves = false }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const sourceName = requiredText(expectedName, "expectedName");
  if (typeof allowClosedLeaves !== "boolean") throw new TypeError("allowClosedLeaves must be a boolean when supplied.");

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
  const sourceCustomShows = await inspectCanonicalCustomShows(sourceZip);
  const sourceLeaves = await inspectClosedLeaves(sourceZip, sourcePart, { allowClosedLeaves });
  const sourceCharts = await inspectClosedChartParts(sourceZip, sourcePart);
  const sourceOleWorkbooks = await inspectClosedOleWorkbookParts(sourceZip, sourcePart);
  const sourceDiagrams = await inspectClosedDiagramParts(sourceZip, sourcePart);
  const sourceInkContents = await inspectClosedInkContentParts(sourceZip, sourcePart);
  const sourceRunHyperlinks = await inspectCanonicalRunHyperlinks(sourceZip, sourcePart, sourceCustomShows);
  assertModelOleWorkbookBindings(modelOleWorkbookBindings(target), sourceOleWorkbooks);
  assertModelDiagramBindings(modelDiagramBindings(target), sourceDiagrams);
  assertModelInkContentBindings(modelInkContentBindings(target), sourceInkContents);
  if (slideNameFromXml(Buffer.from(await requiredZipBytes(sourceZip, sourcePart)).toString("utf8"), sourcePart) !== sourceName) {
    throw new Error("The selected model slide does not match its source SlidePart p:cSld/@name.");
  }
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const sourceSlidePartById = new Map(presentation.slides.items.map((slide, index) => [slide.id, sourceSlideParts[index]]));
  const sourceSemantic = canonicalCloneSnapshot(target, sourceSlidePartById);
  const sourceClosedLeafSemantics = closedLeafSemanticSnapshot(target);
  const sourceCustomShowSemantics = canonicalCustomShowSnapshot(presentation, sourceSlidePartById);
  const clone = target.duplicate();
  if (presentation.slides.items.indexOf(clone) !== sourceIndex + 1 || clone.name !== target.name) {
    throw new Error("slide.duplicate did not create an adjacent same-name pending clone.");
  }
  if (canonicalCloneSnapshot(clone, sourceSlidePartById) !== sourceSemantic) {
    throw new Error("Pending slide clone changed the canonical source semantics before export.");
  }
  if (canonicalCustomShowSnapshot(presentation, sourceSlidePartById) !== sourceCustomShowSemantics) {
    throw new Error("slide.duplicate changed custom-show membership before export.");
  }

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertDuplicatePackageScope(source, output, sourceIndex, { allowClosedLeaves });
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const expectedNames = [...originalNames];
    expectedNames.splice(sourceIndex + 1, 0, sourceName);
    if (JSON.stringify(reimported.slides.items.map((slide) => slide.name)) !== JSON.stringify(expectedNames)) {
      throw new Error("PPTX export changed source names/order or did not retain the adjacent clone.");
    }
    const retained = reimported.slides.items[sourceIndex];
    const roundTripClone = reimported.slides.items[sourceIndex + 1];
    const outputZip = await JSZip.loadAsync(output);
    const [retainedOleWorkbooks, roundTripCloneOleWorkbooks, retainedDiagrams, roundTripCloneDiagrams, retainedInkContents, roundTripCloneInkContents] = await Promise.all([
      inspectClosedOleWorkbookParts(outputZip, packageScope.sourcePart),
      inspectClosedOleWorkbookParts(outputZip, packageScope.clonePart),
      inspectClosedDiagramParts(outputZip, packageScope.sourcePart),
      inspectClosedDiagramParts(outputZip, packageScope.clonePart),
      inspectClosedInkContentParts(outputZip, packageScope.sourcePart),
      inspectClosedInkContentParts(outputZip, packageScope.clonePart),
    ]);
    assertModelOleWorkbookBindings(modelOleWorkbookBindings(retained), retainedOleWorkbooks);
    assertModelOleWorkbookBindings(modelOleWorkbookBindings(roundTripClone), roundTripCloneOleWorkbooks);
    assertModelDiagramBindings(modelDiagramBindings(retained), retainedDiagrams);
    assertModelDiagramBindings(modelDiagramBindings(roundTripClone), roundTripCloneDiagrams);
    assertModelInkContentBindings(modelInkContentBindings(retained), retainedInkContents);
    assertModelInkContentBindings(modelInkContentBindings(roundTripClone), roundTripCloneInkContents);
    const outputSlidePartById = new Map(reimported.slides.items.map((slide, index) => [slide.id, packageScope.outputSlideParts[index]]));
    if (canonicalCloneSnapshot(retained, outputSlidePartById) !== sourceSemantic || canonicalCloneSnapshot(roundTripClone, outputSlidePartById) !== sourceSemantic) {
      throw new Error("PPTX duplicate did not preserve source and clone semantic structure after reimport.");
    }
    if (closedLeafSemanticSnapshot(retained) !== sourceClosedLeafSemantics || closedLeafSemanticSnapshot(roundTripClone) !== sourceClosedLeafSemantics) {
      throw new Error("PPTX duplicate did not preserve source and clone closed-leaf semantics after reimport.");
    }
    if (canonicalCustomShowSnapshot(reimported, outputSlidePartById) !== sourceCustomShowSemantics) {
      throw new Error("PPTX duplicate changed custom-show identity or membership after reimport.");
    }
    // Compare both slides inside the same reimport identity domain. Inserting
    // the clone can renumber public model slide IDs even though the internal
    // hyperlink still targets the same retained SlidePart; SVG anchor hrefs
    // are non-visual but would otherwise make a pre/post raw comparison lie.
    const [retainedRender, cloneRender] = await Promise.all([modelSvg(retained), modelSvg(roundTripClone)]);
    if (cloneRender.visualSha256 !== retainedRender.visualSha256) throw new Error("PPTX duplicate changed the clone model SVG render.");
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
        scope: packageScope.profile,
        allowClosedLeaves,
        closedLeaves: {
          speakerNotes: Boolean(sourceLeaves.notes),
          legacyComments: Boolean(sourceLeaves.comments),
        },
        runHyperlinks: {
          relationshipCount: sourceRunHyperlinks.relationshipCount,
          actionOnlyCount: sourceRunHyperlinks.actionOnlyCount,
          customShowCount: sourceRunHyperlinks.customShowCount,
          customShowActions: sourceRunHyperlinks.customShowActions,
        },
        customShows: {
          count: sourceCustomShows.shows.length,
          shows: sourceCustomShows.shows,
        },
        chartParts: {
          count: sourceCharts.count,
          sourceParts: sourceCharts.charts.map((chart) => chart.part),
          relationshipIds: sourceCharts.charts.map((chart) => chart.relationship.id),
        },
        oleWorkbookParts: {
          count: sourceOleWorkbooks.count,
          sourceParts: sourceOleWorkbooks.workbooks.map((workbook) => workbook.part),
          relationshipIds: sourceOleWorkbooks.workbooks.map((workbook) => workbook.relationship.id),
          previewParts: sourceOleWorkbooks.workbooks.map((workbook) => workbook.previewPart),
        },
        diagramParts: {
          count: sourceDiagrams.count,
          partCount: sourceDiagrams.partCount,
          sourceParts: sourceDiagrams.diagrams.flatMap((diagram) => diagram.roots.map((root) => root.part)),
          relationshipIds: sourceDiagrams.diagrams.flatMap((diagram) => diagram.roots.map((root) => root.relationship.id)),
        },
        inkContentParts: {
          count: sourceInkContents.count,
          sourceParts: sourceInkContents.contents.map((content) => content.part),
          relationshipIds: sourceInkContents.contents.map((content) => content.relationship.id),
        },
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
          runHyperlinks: packageScope.runHyperlinks,
          customShows: packageScope.customShows,
          chartParts: packageScope.chartParts,
          oleWorkbookParts: packageScope.oleWorkbookParts,
          diagramParts: packageScope.diagramParts,
          inkContentParts: packageScope.inkContentParts,
          closedLeaves: packageScope.closedLeaves,
        },
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          sourceAndCloneSemanticsEqual: true,
          sourceAndCloneClosedLeavesEqual: true,
          sourceAndCloneOleWorkbookBindingsIndependent: packageScope.oleWorkbookParts.independentParts,
          sourceAndCloneDiagramBindingsIndependent: packageScope.diagramParts.independentParts,
          sourceAndCloneInkContentBindingsIndependent: packageScope.inkContentParts.independentParts,
          customShowMembershipRetained: true,
          sourceAndCloneNames: [retained.name, roundTripClone.name],
        },
        modelRender: {
          ok: true,
          renderer: "model-svg",
          sourceRawSha256: retainedRender.rawSha256,
          cloneRawSha256: cloneRender.rawSha256,
          visualSha256: retainedRender.visualSha256,
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
  const values = [];
  let allowClosedLeaves = false;
  for (const value of argv) {
    if (value === "--allow-closed-leaves") {
      if (allowClosedLeaves) throw new Error("--allow-closed-leaves may be supplied only once.");
      allowClosedLeaves = true;
      continue;
    }
    values.push(value);
  }
  if (values.length < 3 || values.length > 4) {
    throw new Error("Usage: openchestnut-slide-duplicate-workflow.mjs <input.pptx> <output.pptx> <audit.json> [unique source name] [--allow-closed-leaves]");
  }
  const [inputPath, outputPath, auditPath, expectedName = "Clone source"] = values;
  return { inputPath, outputPath, auditPath, expectedName, allowClosedLeaves };
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
    customShowActionCount: result.audit.operation.runHyperlinks.customShowCount,
    customShowCount: result.audit.operation.customShows.count,
    chartPartCount: result.audit.operation.chartParts.count,
    oleWorkbookPartCount: result.audit.operation.oleWorkbookParts.count,
    smartArtPartCount: result.audit.operation.diagramParts.partCount,
    inkContentPartCount: result.audit.operation.inkContentParts.count,
    closedLeaves: result.audit.operation.closedLeaves,
  }));
}
