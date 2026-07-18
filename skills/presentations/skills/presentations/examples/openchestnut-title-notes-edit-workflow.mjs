import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

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

function snapshot(value) {
  if (value === undefined) return undefined;
  return JSON.parse(JSON.stringify(value));
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
  return JSON.stringify(left) === JSON.stringify(right);
}

function resolveTarget(presentation, slideName, titleShapeName, expectedTitle, expectedNotes) {
  const slides = presentation.slides.items.filter((slide) => slide.name === slideName);
  if (slides.length !== 1) throw new Error("Expected exactly one slide named " + JSON.stringify(slideName) + "; found " + slides.length + ".");
  const slide = slides[0];
  const titles = slide.shapes.items.filter((shape) => shape.name === titleShapeName);
  if (titles.length !== 1) throw new Error("Expected exactly one title shape named " + JSON.stringify(titleShapeName) + "; found " + titles.length + ".");
  const title = titles[0];
  if (title.text.value !== expectedTitle) throw new Error("The selected title does not contain the expected original text; refusing an ambiguous slide edit.");
  if (!slide.speakerNotes || slide.speakerNotes.text !== expectedNotes) {
    throw new Error("The selected slide does not contain exactly the expected plain-text speaker notes; rich, absent, or changed notes are source-bound.");
  }
  return { slide, title };
}

async function renderModel(slide) {
  const preview = await slide.export({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Presentation model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

export async function editPptxTitleAndNotes({
  inputPath,
  outputPath,
  auditPath,
  slideName,
  titleShapeName,
  expectedTitle,
  replacementTitle,
  expectedNotes,
  replacementNotes,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original presentation remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from the source and PPTX output paths.");
  const targetSlideName = requiredText(slideName, "slideName");
  const targetShapeName = requiredText(titleShapeName, "titleShapeName");
  const originalTitle = requiredText(expectedTitle, "expectedTitle");
  const nextTitle = requiredText(replacementTitle, "replacementTitle");
  const originalNotes = requiredText(expectedNotes, "expectedNotes");
  const nextNotes = requiredText(replacementNotes, "replacementNotes");

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const originalSlideNames = presentation.slides.items.map((slide) => slide.name);
  const { slide, title } = resolveTarget(presentation, targetSlideName, targetShapeName, originalTitle, originalNotes);
  const identity = targetSnapshot(slide, title);
  title.text.set(nextTitle);
  slide.speakerNotes.textFrame.setText(nextNotes);

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    if (JSON.stringify(reimported.slides.items.map((candidate) => candidate.name)) !== JSON.stringify(originalSlideNames)) {
      throw new Error("PPTX export changed slide count, order, or slide names.");
    }
    const roundTrip = resolveTarget(reimported, targetSlideName, targetShapeName, nextTitle, nextNotes);
    if (!sameSnapshot(identity, targetSnapshot(roundTrip.slide, roundTrip.title))) {
      throw new Error("PPTX export changed the target slide identity, title-shape identity/geometry, direct background, or speaker-notes identity.");
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
        type: "title-and-speaker-notes-text-edit",
        slideId: identity.slideId,
        slideName: identity.slideName,
        titleId: identity.titleId,
        titleName: identity.titleName,
      },
      warnings: [],
      validation: {
        reimport: {
          ok: true,
          slideCount: reimported.slides.count,
          slideNamesPreserved: true,
          slideIdPreserved: roundTrip.slide.id === identity.slideId,
          titleIdPreserved: roundTrip.title.id === identity.titleId,
          titleExact: roundTrip.title.text.value === nextTitle,
          speakerNotesExact: roundTrip.slide.speakerNotes.text === nextNotes,
          notesIdPreserved: roundTrip.slide.speakerNotes.id === identity.notesId,
          directBackgroundPreserved: JSON.stringify(roundTrip.slide.background) === JSON.stringify(identity.background),
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
    slideName = "Go-no-go decision",
    titleShapeName = "approval-title",
    expectedTitle = "Decision: hold for legal review",
    replacementTitle = "Decision: approve controlled rollout",
    expectedNotes = "Lead with the pending legal condition.\nClose with the accountable owner.",
    replacementNotes = "Lead with the approved controls.\nClose with the accountable rollout owner.",
  ] = argv;
  return { inputPath, outputPath, auditPath, slideName, titleShapeName, expectedTitle, replacementTitle, expectedNotes, replacementNotes };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxTitleAndNotes(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    slideId: result.audit.operation.slideId,
  }));
}
