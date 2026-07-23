import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const require = createRequire(import.meta.url);

const DEFAULTS = Object.freeze({
  slideName: "Go-no-go decision",
  titleShapeName: "approval-title",
  expectedTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  paragraphIndex: 0,
  runIndex: 1,
  expectedRunText: "the pending legal condition.",
  replacementRunText: "the approved control set.",
  expectedRunStyle: Object.freeze({ italic: true, fontSize: 18, color: "#7c2d12" }),
  replacementRunStyle: Object.freeze({ bold: true, italic: false, fontSize: 18, color: "#0f766e" }),
});

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

function requiredIndex(value, label) {
  if (!Number.isInteger(value) || value < 0) throw new TypeError(label + " must be a non-negative integer.");
  return value;
}

function snapshot(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
}

function canonical(value) {
  if (Array.isArray(value)) return value.map(canonical);
  if (value && typeof value === "object") {
    return Object.fromEntries(Object.keys(value).sort().map((key) => [key, canonical(value[key])]));
  }
  return value;
}

function sameValue(left, right) {
  return JSON.stringify(canonical(left)) === JSON.stringify(canonical(right));
}

function targetSnapshot(slide, title) {
  return {
    slideId: slide.id,
    slideName: slide.name,
    slidePosition: snapshot(slide.position),
    background: snapshot(slide.background),
    titleId: title.id,
    titleName: title.name,
    titleGeometry: title.geometry,
    titlePosition: snapshot(title.position),
    notesId: slide.speakerNotes?.id,
  };
}

function sameSnapshot(left, right) {
  return sameValue(left, right);
}

function resolveTarget(presentation, config) {
  const slides = presentation.slides.items.filter((slide) => slide.name === config.slideName);
  if (slides.length !== 1) throw new Error("Expected exactly one slide named " + JSON.stringify(config.slideName) + "; found " + slides.length + ".");
  const slide = slides[0];
  const titles = slide.shapes.items.filter((shape) => shape.name === config.titleShapeName);
  if (titles.length !== 1) throw new Error("Expected exactly one title shape named " + JSON.stringify(config.titleShapeName) + "; found " + titles.length + ".");
  const title = titles[0];
  if (title.text.value !== config.expectedTitle) throw new Error("The selected title does not contain the expected original text; refusing an ambiguous slide edit.");
  const notes = slide.speakerNotes;
  const capability = snapshot(notes?.capability);
  if (!notes || !capability?.sourceBound || !capability.partPresent || !capability.editable) {
    throw new Error("The selected slide does not contain an editable source-bound speaker-notes part.");
  }
  const paragraphs = snapshot(notes.textFrame.paragraphs);
  const run = paragraphs[config.paragraphIndex]?.runs?.[config.runIndex];
  if (!run || run.text !== config.expectedRunText || !sameValue(run.style, config.expectedRunStyle)) {
    throw new Error("The selected rich speaker-notes run does not match the declared text/style contract; refusing to flatten or guess its topology.");
  }
  return { slide, title, notes, paragraphs, capability };
}

async function renderModel(slide) {
  const preview = await slide.export({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Presentation model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

/**
 * Edit exactly one existing ordinary rich-notes run in an imported PPTX.
 *
 * This is intentionally not a general notes reflow or replace-all helper.
 * It snapshots the imported paragraph/run tree, requires the selected run's
 * text and direct style to match exactly, edits only that run, and requires
 * the full reimported tree to equal the expected fixed-topology result.
 */
export async function editPptxRichSpeakerNotes({
  inputPath,
  outputPath,
  auditPath,
  slideName = DEFAULTS.slideName,
  titleShapeName = DEFAULTS.titleShapeName,
  expectedTitle = DEFAULTS.expectedTitle,
  replacementTitle = DEFAULTS.replacementTitle,
  paragraphIndex = DEFAULTS.paragraphIndex,
  runIndex = DEFAULTS.runIndex,
  expectedRunText = DEFAULTS.expectedRunText,
  replacementRunText = DEFAULTS.replacementRunText,
  expectedRunStyle = DEFAULTS.expectedRunStyle,
  replacementRunStyle = DEFAULTS.replacementRunStyle,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const config = {
    slideName: requiredText(slideName, "slideName"),
    titleShapeName: requiredText(titleShapeName, "titleShapeName"),
    expectedTitle: requiredText(expectedTitle, "expectedTitle"),
    replacementTitle: requiredText(replacementTitle, "replacementTitle"),
    paragraphIndex: requiredIndex(paragraphIndex, "paragraphIndex"),
    runIndex: requiredIndex(runIndex, "runIndex"),
    expectedRunText: requiredText(expectedRunText, "expectedRunText"),
    replacementRunText: requiredText(replacementRunText, "replacementRunText"),
    expectedRunStyle: snapshot(expectedRunStyle),
    replacementRunStyle: snapshot(replacementRunStyle),
  };
  if (!config.expectedRunStyle || !config.replacementRunStyle || typeof config.expectedRunStyle !== "object" || typeof config.replacementRunStyle !== "object") {
    throw new TypeError("expectedRunStyle and replacementRunStyle must be plain style records.");
  }

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const originalSlideNames = presentation.slides.items.map((slide) => slide.name);
  const target = resolveTarget(presentation, config);
  const identity = targetSnapshot(target.slide, target.title);
  const expectedParagraphs = snapshot(target.paragraphs);
  expectedParagraphs[config.paragraphIndex].runs[config.runIndex] = {
    ...expectedParagraphs[config.paragraphIndex].runs[config.runIndex],
    text: config.replacementRunText,
    style: snapshot(config.replacementRunStyle),
  };

  target.title.text.set(config.replacementTitle);
  target.notes.textFrame.paragraphs = expectedParagraphs;

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (!sameValue(reimported.slides.items.map((slide) => slide.name), originalSlideNames)) {
      throw new Error("PPTX export changed slide count, order, or slide names.");
    }
    const roundTrip = resolveTarget(reimported, {
      ...config,
      expectedTitle: config.replacementTitle,
      expectedRunText: config.replacementRunText,
      expectedRunStyle: config.replacementRunStyle,
    });
    if (!sameSnapshot(identity, targetSnapshot(roundTrip.slide, roundTrip.title))) {
      throw new Error("PPTX export changed the target slide identity, title-shape identity/geometry, direct background, or speaker-notes identity.");
    }
    if (!sameValue(roundTrip.paragraphs, expectedParagraphs)) {
      throw new Error("PPTX export changed the rich speaker-notes paragraph/run topology or a non-target style.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Presentation verification failed: " + verification.ndjson);
    const modelRender = await renderModel(roundTrip.slide);
    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "title-and-rich-speaker-notes-run-edit",
        slideId: identity.slideId,
        slideName: identity.slideName,
        titleId: identity.titleId,
        titleName: identity.titleName,
        notesId: identity.notesId,
        paragraphIndex: config.paragraphIndex,
        runIndex: config.runIndex,
        expectedRun: { text: config.expectedRunText, style: config.expectedRunStyle },
        replacementRun: { text: config.replacementRunText, style: config.replacementRunStyle },
      },
      warnings: [],
      validation: {
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          slideNamesPreserved: true,
          slideIdPreserved: roundTrip.slide.id === identity.slideId,
          titleIdPreserved: roundTrip.title.id === identity.titleId,
          titleExact: roundTrip.title.text.value === config.replacementTitle,
          notesIdPreserved: roundTrip.slide.speakerNotes.id === identity.notesId,
          richNotesFixedTopology: true,
          paragraphCount: roundTrip.paragraphs.length,
          runCounts: roundTrip.paragraphs.map((paragraph) => paragraph.runs.length),
          targetRunExact: sameValue(roundTrip.paragraphs[config.paragraphIndex]?.runs?.[config.runIndex], expectedParagraphs[config.paragraphIndex]?.runs?.[config.runIndex]),
          directBackgroundPreserved: sameValue(roundTrip.slide.background, identity.background),
        },
        verify: { ok: verification.ok },
        modelRender: { ok: true, ...modelRender },
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
    slideName = DEFAULTS.slideName,
    titleShapeName = DEFAULTS.titleShapeName,
    expectedTitle = DEFAULTS.expectedTitle,
    replacementTitle = DEFAULTS.replacementTitle,
    expectedRunText = DEFAULTS.expectedRunText,
    replacementRunText = DEFAULTS.replacementRunText,
  ] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    slideName,
    titleShapeName,
    expectedTitle,
    replacementTitle,
    expectedRunText,
    replacementRunText,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxRichSpeakerNotes(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    slideId: result.audit.operation.slideId,
    notesId: result.audit.operation.notesId,
  }));
}
