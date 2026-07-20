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

function requiredContent(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(label + " must be a non-empty string.");
  return value;
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
  const base = "https://openchestnut.invalid/" + sourcePart;
  const resolved = new URL(target, base);
  if (resolved.origin !== "https://openchestnut.invalid") throw new Error("Unexpected external PPTX relationship target.");
  const partPath = resolved.pathname.replace(/^\/+/, "");
  if (!partPath.startsWith("ppt/") || partPath.split("/").includes("..")) {
    throw new Error("Unsafe PPTX relationship target: " + JSON.stringify(target));
  }
  return partPath;
}

async function relationshipEntries(zip, sourcePart) {
  const relsPath = relationshipPartPath(sourcePart);
  const xml = await zip.file(relsPath)?.async("text");
  if (!xml) return [];
  return [...xml.matchAll(/<Relationship\b[^>]*>/gi)].map((match) => {
    const attributes = xmlAttributes(match[0]);
    const external = attributes.TargetMode?.toLowerCase() === "external";
    if (!attributes.Id || !attributes.Type || !attributes.Target) throw new Error("Malformed PPTX relationship in " + relsPath + ".");
    return {
      id: attributes.Id,
      type: attributes.Type,
      target: attributes.Target,
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
  if (matches.length !== 1) throw new Error(label + " must have exactly one " + suffix + " relationship; found " + matches.length + ".");
  if (matches[0].external || !matches[0].targetPart) throw new Error(label + " " + suffix + " relationship must be internal.");
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
    const relationshipId = xmlAttributes(match[0])["r:id"];
    const partPath = slideById.get(relationshipId);
    if (!partPath || !zip.file(partPath)) throw new Error("Presentation slide list contains an unresolved SlidePart.");
    return partPath;
  });
  if (!paths.length || new Set(paths).size !== paths.length) throw new Error("Presentation slide list must contain distinct SlideParts.");
  return paths;
}

async function notesMasterState(zip) {
  const presentationXml = await zip.file("ppt/presentation.xml")?.async("text");
  if (!presentationXml) throw new Error("PPTX is missing ppt/presentation.xml.");
  const relationships = await relationshipEntries(zip, "ppt/presentation.xml");
  const masterRelationships = relationships.filter((entry) => relationshipType(entry, "notesMaster"));
  const masterIds = [...presentationXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?notesMasterId\b[^>]*>/gi)]
    .map((match) => xmlAttributes(match[0])["r:id"]);
  if (masterRelationships.length !== masterIds.length || masterRelationships.length > 1) {
    throw new Error("PPTX NotesMaster list and relationship graph must contain zero or one matching entry.");
  }
  if (!masterRelationships.length) return null;
  const relationship = masterRelationships[0];
  if (relationship.external || !relationship.targetPart || masterIds[0] !== relationship.id || !zip.file(relationship.targetPart)) {
    throw new Error("PPTX NotesMaster entry is unresolved or external.");
  }
  return {
    relationshipId: relationship.id,
    partPath: relationship.targetPart,
    sha256: sha256(await zip.file(relationship.targetPart).async("uint8array")),
  };
}

function visibleSlideSnapshot(slide) {
  const value = structuredClone(slide.toProto());
  delete value.notes;
  return JSON.stringify(value);
}

async function modelSvg(slide) {
  const rendered = await slide.export({ format: "svg" });
  if (!/<svg\b/i.test(await rendered.text())) throw new Error("Presentation model render did not produce SVG.");
  return { bytes: rendered.bytes.length, sha256: sha256(rendered.bytes) };
}

async function auditAddedNotesGraph(sourceBytes, outputBytes, targetIndex) {
  const sourceZip = await JSZip.loadAsync(sourceBytes);
  const outputZip = await JSZip.loadAsync(outputBytes);
  const sourceSlides = await orderedSlidePartPaths(sourceZip);
  const outputSlides = await orderedSlidePartPaths(outputZip);
  if (JSON.stringify(sourceSlides) !== JSON.stringify(outputSlides)) throw new Error("Adding speaker notes changed slide package identity or order.");
  const targetPart = sourceSlides[targetIndex];
  if (!targetPart) throw new Error("Target slide index is outside the source package slide list.");

  const sourceSlideRelationships = await relationshipEntries(sourceZip, targetPart);
  if (sourceSlideRelationships.some((entry) => relationshipType(entry, "notesSlide"))) {
    throw new Error("Selected source SlidePart already has a NotesSlide relationship.");
  }
  const outputSlideRelationships = await relationshipEntries(outputZip, targetPart);
  const notesRelationship = exactlyOne(outputSlideRelationships, "notesSlide", "Target SlidePart");
  if (!outputZip.file(notesRelationship.targetPart)) throw new Error("Added NotesSlide relationship points at a missing part.");

  const sourceMaster = await notesMasterState(sourceZip);
  const outputMaster = await notesMasterState(outputZip);
  if (!outputMaster) throw new Error("Added NotesSlide has no presentation NotesMaster.");
  const notesRelationships = await relationshipEntries(outputZip, notesRelationship.targetPart);
  if (notesRelationships.length !== 2) throw new Error("Added NotesSlide must have exactly two internal relationships.");
  const masterRelationship = exactlyOne(notesRelationships, "notesMaster", "Added NotesSlide");
  const slideBackReference = exactlyOne(notesRelationships, "slide", "Added NotesSlide");
  if (masterRelationship.targetPart !== outputMaster.partPath || slideBackReference.targetPart !== targetPart) {
    throw new Error("Added NotesSlide does not point at the canonical NotesMaster and its owning SlidePart.");
  }

  let masterPolicy;
  if (sourceMaster) {
    if (sourceMaster.partPath !== outputMaster.partPath || sourceMaster.relationshipId !== outputMaster.relationshipId || sourceMaster.sha256 !== outputMaster.sha256) {
      throw new Error("Existing NotesMaster was not reused byte-for-byte with stable package identity.");
    }
    masterPolicy = "reused-byte-identical";
  } else {
    const masterRelationships = await relationshipEntries(outputZip, outputMaster.partPath);
    if (masterRelationships.length !== 1) throw new Error("New canonical NotesMaster must have exactly one relationship.");
    const theme = exactlyOne(masterRelationships, "theme", "New NotesMaster");
    if (!outputZip.file(theme.targetPart)) throw new Error("New NotesMaster theme relationship points at a missing part.");
    masterPolicy = "created-canonical-shared-theme";
  }

  return {
    ok: true,
    targetSlidePart: targetPart,
    notesSlidePart: notesRelationship.targetPart,
    notesMasterPart: outputMaster.partPath,
    notesMasterPolicy: masterPolicy,
    sourceHadNotesMaster: Boolean(sourceMaster),
    notesSlideRelationshipCount: notesRelationships.length,
    slideBackReferenceVerified: true,
  };
}

export async function addPptxSpeakerNotes({ inputPath, outputPath, auditPath, slideName, notes }) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const targetSlideName = requiredText(slideName, "slideName");
  const nextNotes = requiredContent(notes, "notes");

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const candidates = presentation.slides.items.filter((slide) => slide.name === targetSlideName);
  if (candidates.length !== 1) throw new Error("Expected exactly one imported slide named " + JSON.stringify(targetSlideName) + "; found " + candidates.length + ".");
  const target = candidates[0];
  const targetIndex = presentation.slides.items.indexOf(target);
  const capability = target.speakerNotes.capability;
  if (!capability.sourceBound || capability.partPresent || capability.editable || !capability.addable || target.speakerNotes.text) {
    throw new Error("Selected slide is not an imported, notes-absent slide whose package graph explicitly permits speaker-notes creation.");
  }
  if (presentation.resolve(target.speakerNotes.id) !== target.speakerNotes) throw new Error("Speaker-notes object cannot be resolved by its stable model ID.");
  const originalNames = presentation.slides.items.map((slide) => slide.name);
  const visibleSnapshot = visibleSlideSnapshot(target);
  const sourceRender = await modelSvg(target);
  target.addNotes(nextNotes);

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageGraph = await auditAddedNotesGraph(source, output, targetIndex);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (JSON.stringify(reimported.slides.items.map((slide) => slide.name)) !== JSON.stringify(originalNames)) {
      throw new Error("PPTX export changed slide count, order, or slide names.");
    }
    const roundTrip = reimported.slides.items[targetIndex];
    if (visibleSlideSnapshot(roundTrip) !== visibleSnapshot) throw new Error("Adding speaker notes changed visible target-slide semantics.");
    if (roundTrip.speakerNotes.text !== nextNotes) throw new Error("Reimported speaker notes do not exactly match the requested text.");
    const outputCapability = roundTrip.speakerNotes.capability;
    if (!outputCapability.sourceBound || !outputCapability.partPresent || !outputCapability.editable || outputCapability.addable) {
      throw new Error("Reimported speaker-notes capability does not reflect the newly created editable NotesSlide.");
    }
    const outputRender = await modelSvg(roundTrip);
    if (outputRender.sha256 !== sourceRender.sha256) throw new Error("Adding nonvisual speaker notes changed the target slide's model SVG render.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    const sourceAfter = await fs.readFile(sourcePath);
    if (sha256(sourceAfter) !== sourceHash) throw new Error("Source presentation changed during the transaction.");

    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-speaker-notes-add",
        slideId: target.id,
        slideName: target.name,
        slideIndex: targetIndex,
        notesCharacters: nextNotes.length,
      },
      precondition: { capability },
      warnings: [],
      validation: {
        package: packageGraph,
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          slideNamesPreserved: true,
          visibleSemanticsPreserved: true,
          speakerNotesExact: true,
          capability: outputCapability,
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
    slideName = "Speaker notes target",
    notes = "Lead with the evidence.\nClose with the requested decision.",
  ] = argv;
  return { inputPath, outputPath, auditPath, slideName, notes };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await addPptxSpeakerNotes(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    notesSlidePart: result.audit.validation.package.notesSlidePart,
  }));
}
