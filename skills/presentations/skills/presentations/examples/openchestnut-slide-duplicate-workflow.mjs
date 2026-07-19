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
const MAX_CLOSED_CHART_PARTS = 256;

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
  return resolveRelationshipTargetFromPart("ppt/presentation.xml", target);
}

function resolveRelationshipTargetFromPart(sourcePart, target) {
  if (typeof sourcePart !== "string" || !sourcePart.startsWith("ppt/") || !sourcePart.endsWith(".xml")) {
    throw new Error("Invalid PPTX relationship source part: " + JSON.stringify(sourcePart));
  }
  if (typeof target !== "string" || !target || /[\\?#]/.test(target) || /%[0-9a-f]{2}/i.test(target)) {
    throw new Error("Unsafe PPTX relationship target: " + JSON.stringify(target));
  }
  const partPath = target.startsWith("/")
    ? target.replace(/^\/+/, "")
    : path.posix.normalize(path.posix.join(path.posix.dirname(sourcePart), target));
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) {
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

async function inspectCanonicalRunHyperlinks(zip, slidePart) {
  const slideXml = Buffer.from(await requiredZipBytes(zip, slidePart)).toString("utf8");
  if (/<(?:[A-Za-z_][\w.-]*:)?hlinkHover\b/i.test(slideXml)) {
    throw new Error("Canonical duplicate workflow does not accept hover hyperlinks in " + slidePart + ".");
  }
  const relationships = (await relationshipEntriesForPart(zip, slidePart))
    .filter((entry) => relationshipTypeMatches(entry, "hyperlink") || relationshipTypeMatches(entry, "slide"));
  const relationshipById = new Map(relationships.map((entry) => [entry.id, entry]));
  const usedRelationshipIds = new Set();
  const clicks = [];
  for (const match of slideXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?hlinkClick\b[^>]*>/gi)) {
    const attributes = xmlAttributes(match[0]);
    const relationshipId = attributes["r:id"] || "";
    const action = unescapeXml(attributes.action || "");
    if (!relationshipId) {
      if (!INLINE_ACTIONS.has(action)) {
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
    const result = {};
    for (const [key, item] of Object.entries(value)) {
      if (key === "id" || key === "layoutId") continue;
      if ((key === "startTargetId" || key === "endTargetId") && typeof item === "string") {
        result[key] = locationById.get(item) || "unresolved:" + item;
      } else if (key === "slideId" && typeof item === "string") {
        result[key] = slidePartById.get(item) || "unresolved:" + item;
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
  const [sourceLeaves, cloneLeaves] = await Promise.all([
    inspectClosedLeaves(sourceZip, sourcePart, { allowClosedLeaves }),
    inspectClosedLeaves(outputZip, clonePart, { allowClosedLeaves }),
  ]);
  const [sourceRunHyperlinks, cloneRunHyperlinks] = await Promise.all([
    inspectCanonicalRunHyperlinks(sourceZip, sourcePart),
    inspectCanonicalRunHyperlinks(outputZip, clonePart),
  ]);
  const [sourceCharts, cloneCharts] = await Promise.all([
    inspectClosedChartParts(sourceZip, sourcePart),
    inspectClosedChartParts(outputZip, clonePart),
  ]);
  if (sourceRunHyperlinks.fingerprint !== cloneRunHyperlinks.fingerprint) {
    throw new Error("PPTX duplicate changed the canonical run-hyperlink XML/relationship graph.");
  }
  const closedLeaves = await assertClosedLeafClonePackage(sourceZip, outputZip, sourceLeaves, cloneLeaves, sourcePart, clonePart);
  const chartParts = await assertClosedChartClonePackage(sourceZip, outputZip, sourceCharts, cloneCharts);
  const newParts = outputParts.filter((partPath) => !sourceZip.file(partPath));
  const expectedNewParts = [clonePart, relationshipPartPath(clonePart)];
  expectedNewParts.push(...cloneCharts.charts.map((chart) => chart.part));
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
    profile: sourceCharts.count
      ? sourceLeaves.notes || sourceLeaves.comments
        ? "canonical-inline-leaves-with-closed-relationship-and-chart-leaves"
        : "canonical-inline-leaves-with-closed-chart-leaves"
      : sourceLeaves.profile,
    outputSlideParts,
    runHyperlinks: {
      relationshipCount: sourceRunHyperlinks.relationshipCount,
      actionOnlyCount: sourceRunHyperlinks.actionOnlyCount,
      exactSourceGraphRetained: true,
    },
    chartParts,
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
  const sourceLeaves = await inspectClosedLeaves(sourceZip, sourcePart, { allowClosedLeaves });
  const sourceCharts = await inspectClosedChartParts(sourceZip, sourcePart);
  const sourceRunHyperlinks = await inspectCanonicalRunHyperlinks(sourceZip, sourcePart);
  if (slideNameFromXml(Buffer.from(await requiredZipBytes(sourceZip, sourcePart)).toString("utf8"), sourcePart) !== sourceName) {
    throw new Error("The selected model slide does not match its source SlidePart p:cSld/@name.");
  }
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const sourceSlidePartById = new Map(presentation.slides.items.map((slide, index) => [slide.id, sourceSlideParts[index]]));
  const sourceSemantic = canonicalCloneSnapshot(target, sourceSlidePartById);
  const sourceClosedLeafSemantics = closedLeafSemanticSnapshot(target);
  const clone = target.duplicate();
  if (presentation.slides.items.indexOf(clone) !== sourceIndex + 1 || clone.name !== target.name) {
    throw new Error("slide.duplicate did not create an adjacent same-name pending clone.");
  }
  if (canonicalCloneSnapshot(clone, sourceSlidePartById) !== sourceSemantic) {
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
    const packageScope = await assertDuplicatePackageScope(source, output, sourceIndex, { allowClosedLeaves });
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const expectedNames = [...originalNames];
    expectedNames.splice(sourceIndex + 1, 0, sourceName);
    if (JSON.stringify(reimported.slides.items.map((slide) => slide.name)) !== JSON.stringify(expectedNames)) {
      throw new Error("PPTX export changed source names/order or did not retain the adjacent clone.");
    }
    const retained = reimported.slides.items[sourceIndex];
    const roundTripClone = reimported.slides.items[sourceIndex + 1];
    const outputSlidePartById = new Map(reimported.slides.items.map((slide, index) => [slide.id, packageScope.outputSlideParts[index]]));
    if (canonicalCloneSnapshot(retained, outputSlidePartById) !== sourceSemantic || canonicalCloneSnapshot(roundTripClone, outputSlidePartById) !== sourceSemantic) {
      throw new Error("PPTX duplicate did not preserve source and clone semantic structure after reimport.");
    }
    if (closedLeafSemanticSnapshot(retained) !== sourceClosedLeafSemantics || closedLeafSemanticSnapshot(roundTripClone) !== sourceClosedLeafSemantics) {
      throw new Error("PPTX duplicate did not preserve source and clone closed-leaf semantics after reimport.");
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
        },
        chartParts: {
          count: sourceCharts.count,
          sourceParts: sourceCharts.charts.map((chart) => chart.part),
          relationshipIds: sourceCharts.charts.map((chart) => chart.relationship.id),
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
          chartParts: packageScope.chartParts,
          closedLeaves: packageScope.closedLeaves,
        },
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          sourceAndCloneSemanticsEqual: true,
          sourceAndCloneClosedLeavesEqual: true,
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
    chartPartCount: result.audit.operation.chartParts.count,
    closedLeaves: result.audit.operation.closedLeaves,
  }));
}
