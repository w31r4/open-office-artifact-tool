import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import {
  FileBlob,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
  verifyArtifact,
  visualQaArtifact,
} from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import { normalizeOpenChestnutCodecName, presentationOpenChestnutConfig } from "../../shared/open-chestnut-compat.mjs";
import { addOpenChestnutNativeGraphFixture } from "./open-chestnut-native-fixture.mjs";
import {
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../../shared/visual-baselines.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function commandExists(command) {
  const result = spawnSync(process.platform === "win32" ? "where" : "which", [command], { encoding: "utf8", shell: false });
  return result.status === 0;
}

export function nativePresentationRenderStatus() {
  const commands = { soffice: commandExists("soffice"), pdftoppm: commandExists("pdftoppm"), pdfinfo: commandExists("pdfinfo") };
  return { available: Object.values(commands).every(Boolean), commands };
}

function addSlideShape(slide, config = {}) {
  const shape = slide.shapes.add({ ...config, text: config.text || "" });
  if (config.textStyle) shape.text.style = { fontFamily: slide.presentation.theme.fonts.minor, ...config.textStyle };
  if (config.paragraphStyles || config.inheritedParagraphStyles) shape.text.inheritedParagraphStyles = config.paragraphStyles || config.inheritedParagraphStyles;
  return shape;
}

async function presentationFixtureChartConfig(config = {}) {
  const workbookConfig = config.externalData?.workbook;
  if (!workbookConfig?.sheets) return config;
  const workbook = Workbook.create({ dateSystem: workbookConfig.dateSystem });
  for (const sheetConfig of workbookConfig.sheets) {
    const sheet = workbook.worksheets.add(sheetConfig.name);
    if (sheetConfig.values) sheet.getRange(sheetConfig.range || "A1").values = sheetConfig.values;
  }
  const workbookFile = await SpreadsheetFile.exportXlsx(workbook);
  return { ...config, externalData: { ...config.externalData, workbook: workbookFile } };
}

async function addFixtureSlide(presentation, config = {}) {
  const slide = presentation.slides.add({ name: config.name, layoutId: config.layoutId, notes: config.notes });
  if (config.background) slide.background.fill = config.background;
  const byName = new Map();
  const remember = (item, name) => { if (name) byName.set(name, item); return item; };
  if (config.applyLayoutPlaceholders) {
    const layout = presentation.layouts.getItem(config.layoutId);
    assert.ok(layout, `Missing presentation fixture layout ${config.layoutId}`);
    for (const placeholder of layout.apply(slide)) remember(placeholder, placeholder.name);
  }
  for (const shape of config.shapes || []) remember(addSlideShape(slide, shape), shape.name);
  for (const group of config.groups || []) {
    const created = slide.groups.add(group);
    for (const element of created.allElements()) remember(element, element.name);
  }
  for (const table of config.tables || []) remember(slide.tables.add(table), table.name);
  for (const chart of config.charts || []) {
    const resolvedChart = await presentationFixtureChartConfig(chart);
    remember(slide.charts.add(resolvedChart.chartType || resolvedChart.type || "bar", resolvedChart), chart.name);
  }
  for (const image of config.images || []) remember(slide.images.add(image), image.name);
  for (const connector of config.connectors || []) {
    const from = byName.get(connector.fromName) || connector.from || connector.start;
    const to = byName.get(connector.toName) || connector.to || connector.end;
    remember(slide.connectors.add({ ...connector, from, to }), connector.name);
  }
  for (const comment of config.comments || []) {
    const targetElement = byName.get(comment.targetName) || comment.targetId;
    assert.ok(targetElement, `Missing presentation fixture comment target ${comment.targetName || comment.targetId}`);
    const target = comment.targetTextRange
      ? slide.resolve(`${typeof targetElement === "string" ? targetElement : targetElement.id}/text`)
      : targetElement;
    assert.ok(target, `Missing presentation fixture text range for ${comment.targetName || comment.targetId}`);
    const thread = slide.comments.addThread(target, comment.text || "", comment);
    for (const reply of comment.replies || []) thread.addReply(reply.text || reply, typeof reply === "object" ? reply : {});
    if (comment.resolved) thread.resolve();
  }
  return slide;
}

export async function createPresentationFromFixture(fixture = {}) {
  const presentation = Presentation.create({ slideSize: fixture.slideSize || { width: 1280, height: 720 }, theme: fixture.theme || {}, master: fixture.master || {}, masters: fixture.masters, commentFormat: fixture.commentFormat });
  if (fixture.theme?.colors) presentation.theme.setColors(fixture.theme.colors);
  if (fixture.theme?.fonts) presentation.theme.setFonts(fixture.theme.fonts);
  if (fixture.theme?.textStyles) presentation.theme.setTextStyles(fixture.theme.textStyles);
  if (fixture.theme?.colorMap) presentation.theme.setColorMap(fixture.theme.colorMap);
  for (const layout of fixture.layouts || []) presentation.layouts.add(layout);
  for (const slide of fixture.slides || []) await addFixtureSlide(presentation, slide);
  for (const customShow of fixture.customShows || []) {
    const indexes = customShow.slideIndexes || [];
    const slides = indexes.map((index) => presentation.slides.getItem(Number(index) - 1));
    assert.ok(slides.length > 0 && slides.every(Boolean), `Invalid presentation fixture custom show slideIndexes for ${customShow.name}`);
    presentation.customShows.add({ name: customShow.name, nativeId: customShow.nativeId, slides });
  }
  const inspectKind = fixture.qa?.inspectKind || "deck,theme,slideMaster,layout,slide,textbox,shape,table,chart,image,connector,notes,comment,textRange";
  for (const expected of fixture.expectInspect || []) {
    assert.match(presentation.inspect({ kind: expected.kind || inspectKind, maxChars: fixture.qa?.maxChars || 30_000 }).ndjson, new RegExp(expected.pattern));
  }
  return presentation;
}

function packageImageBytes(image = {}) {
  const match = /^data:[^;,]+;base64,([A-Za-z0-9+/=\s]+)$/.exec(String(image.dataUrl || ""));
  if (!match) throw new Error("packageDrawing.image requires a base64 dataUrl.");
  return Buffer.from(match[1].replace(/\s+/g, ""), "base64");
}

async function packageChartXml(chart = {}) {
  const source = Presentation.create();
  const resolvedChart = await presentationFixtureChartConfig(chart);
  source.slides.add().charts.add(resolvedChart.chartType || resolvedChart.type || "bar", resolvedChart);
  const zip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(source)).arrayBuffer()));
  const xml = await zip.file("ppt/charts/chart1.xml")?.async("text");
  if (!xml) throw new Error("Could not generate packageDrawing chart XML through the public presentation API.");
  return xml;
}

async function applyFixturePackageDrawing(pptx, fixture = {}) {
  const drawing = fixture.packageDrawing;
  if (!drawing) return pptx;
  const slideIndex = Number(drawing.slideIndex ?? 0);
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= (fixture.slides || []).length) throw new RangeError("packageDrawing.slideIndex is out of range.");
  const source = `ppt/slides/slide${slideIndex + 1}.xml`;
  const patches = [];
  if (drawing.image) patches.push({
    path: drawing.image.partPath || `ppt/review/media/image-${slideIndex + 1}.png`,
    bytes: packageImageBytes(drawing.image),
    recipe: {
      kind: "image",
      source,
      id: drawing.image.relationshipId,
      sourceReference: { objectId: drawing.image.objectId, name: drawing.image.name, alt: drawing.image.alt, position: drawing.image.position },
    },
  });
  if (drawing.chart) patches.push({
    path: drawing.chart.partPath || `ppt/review/charts/chart-${slideIndex + 1}.xml`,
    xml: await packageChartXml(drawing.chart),
    recipe: {
      kind: "chart",
      source,
      id: drawing.chart.relationshipId,
      sourceReference: { objectId: drawing.chart.objectId, name: drawing.chart.name, alt: drawing.chart.alt, position: drawing.chart.position },
    },
  });
  return patches.length ? PresentationFile.patchPptx(pptx, patches) : pptx;
}

async function applyFixturePackageReview(pptx, fixture = {}) {
  const review = fixture.packageReview;
  if (!review) return pptx;
  const slideIndex = Number(review.slideIndex ?? 0);
  if (!Number.isInteger(slideIndex) || slideIndex < 0 || slideIndex >= (fixture.slides || []).length) throw new RangeError("packageReview.slideIndex is out of range.");
  const number = slideIndex + 1;
  const source = `ppt/slides/slide${number}.xml`;
  const zip = await JSZip.loadAsync(new Uint8Array(await pptx.arrayBuffer()));
  const notesSource = `ppt/notesSlides/notesSlide${number}.xml`;
  const commentsSource = `ppt/comments/comment${number}.xml`;
  const authorsSource = "ppt/commentAuthors.xml";
  const [notesXml, commentsXml, authorsXml] = await Promise.all([
    zip.file(notesSource)?.async("text"),
    zip.file(commentsSource)?.async("text"),
    zip.file(authorsSource)?.async("text"),
  ]);
  if (!notesXml || !commentsXml || !authorsXml) throw new Error("packageReview requires the target slide to contain notes and comments with authors.");
  const notesPath = review.notesPartPath || `ppt/review/notes/slide-${number}.xml`;
  const commentsPath = review.commentsPartPath || `ppt/review/comments/slide-${number}.xml`;
  const authorsPath = review.authorsPartPath || "ppt/review/comments/authors.xml";
  return PresentationFile.patchPptx(pptx, [
    { path: notesSource, remove: true },
    { path: commentsSource, remove: true },
    { path: authorsSource, remove: true },
    { path: notesPath, xml: notesXml, recipe: { kind: "notesSlide", source, id: review.notesRelationshipId } },
    { path: commentsPath, xml: commentsXml, recipe: { kind: "comments", source, id: review.commentsRelationshipId } },
    { path: authorsPath, xml: authorsXml, recipe: { kind: "commentAuthors", source: "ppt/presentation.xml", id: review.authorsRelationshipId } },
  ]);
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", shell: false });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/m.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

async function renderNativeSlides(pptxBlob, outputDir, slideCount, options = {}) {
  const pdf = await createLibreOfficeRenderer({ timeoutMs: options.nativeTimeout ?? 60_000 })({
    input: pptxBlob,
    inputType: PPTX_MIME,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "presentation",
  });
  const pdfPath = path.join(outputDir, "native-render.pdf");
  await pdf.save(pdfPath);
  const pageCount = pdfPageCount(pdfPath);
  if (pageCount !== slideCount) throw new Error(`Native PPTX render produced ${pageCount} pages for ${slideCount} slides.`);
  const pagesDir = path.join(outputDir, "native-slides");
  await fs.mkdir(pagesDir, { recursive: true });
  const poppler = createPopplerRenderer({ dpi: options.dpi ?? 144, timeoutMs: options.nativeTimeout ?? 60_000 });
  const pages = [];
  const qaLines = [];
  const baselineSet = await prepareNumberedVisualBaselines(options.baselineDir, "native-slide", options);
  for (let slideIndex = 0; slideIndex < pageCount; slideIndex += 1) {
    const png = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "presentation", pageIndex: slideIndex });
    const slidePath = path.join(pagesDir, `slide-${slideIndex + 1}.png`);
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `native-slide-${slideIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `native-slide-${slideIndex + 1}.png`);
    const qa = await runPngVisualQa({ export: () => png }, { baselinePath, diffPath, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration, minBytes: options.minBytes, maxChars: options.maxChars });
    await png.save(slidePath);
    qaLines.push(qa.ndjson);
    pages.push({ slide: slideIndex + 1, path: slidePath, diffPath: qa.diffPath, bytes: png.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, pageCount, { artifactKind: "presentation", baselineKind: "native" });
  if (issue) qaLines.push(issue);
  const qaPath = path.join(outputDir, "native-visual-qa.ndjson");
  await fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8");
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pdfPath, qaPath, pageCount, baselinePageCount, pageCountMatches, pages };
}

async function renderModelSlides(presentation, outputDir, options = {}) {
  const slidesDir = path.join(outputDir, "model-slides");
  const layoutsDir = path.join(outputDir, "layouts");
  await Promise.all([fs.mkdir(slidesDir, { recursive: true }), fs.mkdir(layoutsDir, { recursive: true })]);
  const renderer = createPlaywrightRenderer({ viewport: options.viewport || { width: 1280, height: 720 }, deviceScaleFactor: options.deviceScaleFactor ?? 1, timeout: options.timeout ?? 30_000 });
  const slides = [];
  const qaLines = [];
  const baselineSet = await prepareNumberedVisualBaselines(options.baselineDir, "model-slide", options);
  for (let slideIndex = 0; slideIndex < presentation.slides.count; slideIndex += 1) {
    const slide = presentation.slides.getItem(slideIndex);
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `model-slide-${slideIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `model-slide-${slideIndex + 1}.png`);
    const qa = await runPngVisualQa(presentation, {
      renderer,
      renderOptions: { slide },
      baselinePath,
      diffPath,
      writeBaseline: options.writeBaseline,
      pixelThreshold: options.pixelThreshold,
      diffAlignment: options.diffAlignment,
      diffPalette: options.diffPalette,
      pixelRegistration: options.pixelRegistration,
      minBytes: options.minBytes,
      maxChars: options.maxChars,
    });
    const slidePath = path.join(slidesDir, `slide-${slideIndex + 1}.png`);
    const layoutPath = path.join(layoutsDir, `slide-${slideIndex + 1}.json`);
    await Promise.all([qa.blob.save(slidePath), fs.writeFile(layoutPath, await (await slide.export({ format: "layout" })).text(), "utf8")]);
    qaLines.push(qa.ndjson);
    slides.push({ slide: slideIndex + 1, path: slidePath, layoutPath, diffPath: qa.diffPath, bytes: qa.blob.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, presentation.slides.count, { artifactKind: "presentation", baselineKind: "model" });
  if (issue) qaLines.push(issue);
  const montageArtifact = { export: () => presentation.export({ format: "montage", columns: options.montageColumns || 2, scale: options.montageScale || 0.32 }) };
  const montageQa = await visualQaArtifact(montageArtifact, { format: "png", renderer, minBytes: options.minBytes ?? 100, maxChars: options.maxChars ?? 20_000 });
  const montagePath = path.join(outputDir, "model-montage.png");
  const qaPath = path.join(outputDir, "model-visual-qa.ndjson");
  await Promise.all([montageQa.blob.save(montagePath), fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n${montageQa.ndjson}\n`, "utf8")]);
  return { status: "passed", ok: pageCountMatches && slides.every((slide) => slide.ok) && montageQa.ok, pageCount: presentation.slides.count, baselinePageCount, pageCountMatches, slides, montagePath, montage: { bytes: montageQa.blob.bytes.length, hash: montageQa.summary.hash, ok: montageQa.ok }, qaPath };
}

export async function verifyPresentationFile(inputPath, options = {}) {
  const absoluteInput = path.resolve(inputPath);
  const outputDir = path.resolve(options.outputDir || path.join(path.dirname(absoluteInput), `${path.basename(absoluteInput, path.extname(absoluteInput))}-qa`));
  await fs.mkdir(outputDir, { recursive: true });
  const loaded = await FileBlob.load(absoluteInput);
  const pptxBlob = new FileBlob(loaded.bytes, { type: PPTX_MIME, name: path.basename(absoluteInput) });
  const presentation = await PresentationFile.importPptx(pptxBlob);
  const maxChars = options.maxChars ?? 30_000;
  const inspect = presentation.inspect({ kind: options.inspectKind || "deck,theme,layout,slide,textbox,shape,table,chart,image,connector,notes,comment,textRange", maxChars });
  const packageInspect = await PresentationFile.inspectPptx(pptxBlob, { includeText: options.includePackageText === true, maxChars });
  const verify = verifyArtifact(presentation, { maxChars });
  const baselineDir = options.baselineDir ? path.resolve(options.baselineDir) : undefined;
  const modelRender = await renderModelSlides(presentation, outputDir, { ...options, maxChars, baselineDir });
  const paths = {
    inspect: path.join(outputDir, "inspect.ndjson"),
    packageInspect: path.join(outputDir, "package-inspect.ndjson"),
    verify: path.join(outputDir, "verify.ndjson"),
    summary: path.join(outputDir, "summary.json"),
  };
  const verifyNdjson = verify.ndjson || JSON.stringify({ kind: "verificationSummary", artifactKind: "presentation", ok: verify.ok, issues: verify.issues?.length || 0 });
  await Promise.all([
    fs.writeFile(paths.inspect, inspect.ndjson, "utf8"),
    fs.writeFile(paths.packageInspect, packageInspect.ndjson, "utf8"),
    fs.writeFile(paths.verify, `${verifyNdjson}\n`, "utf8"),
  ]);

  const requestedNative = String(options.nativeRender ?? "auto").toLowerCase();
  const nativeStatus = nativePresentationRenderStatus();
  let nativeRender = { status: "skipped", reason: "native render disabled" };
  if (requestedNative !== "off" && requestedNative !== "false") {
    if (nativeStatus.available) nativeRender = await renderNativeSlides(pptxBlob, outputDir, presentation.slides.count, { ...options, maxChars, baselineDir });
    else if (requestedNative === "required" || requestedNative === "true") throw new Error(`Native presentation render requires soffice, pdftoppm, and pdfinfo: ${JSON.stringify(nativeStatus.commands)}`);
    else nativeRender = { status: "skipped", reason: "native render commands unavailable", commands: nativeStatus.commands };
  }
  const summary = { input: absoluteInput, outputDir, slides: presentation.slides.count, packageOk: packageInspect.ok, verifyOk: verify.ok, baselineDir, writeBaseline: Boolean(options.writeBaseline), modelRender, nativeRender, files: paths };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  const visualFailed = modelRender.ok === false || (nativeRender.status === "passed" && nativeRender.ok === false);
  if (options.failOnIssues !== false && (!packageInspect.ok || !verify.ok || visualFailed)) throw new Error(`Presentation QA failed: package=${packageInspect.ok}, semantic=${verify.ok}, visual=${!visualFailed}, native=${nativeRender.status}. See ${outputDir}`);
  return { presentation, inspect, packageInspect, verify, modelRender, nativeRender, summary };
}

export async function runPresentationFixture(fixturePath, options = {}) {
  const absoluteFixture = path.resolve(fixturePath);
  const fixture = JSON.parse(await fs.readFile(absoluteFixture, "utf8"));
  const outputDir = path.resolve(options.outputDir || path.join("tmp", "presentation-skill", fixture.name || "fixture"));
  await fs.mkdir(outputDir, { recursive: true });
  const presentation = await createPresentationFromFixture(fixture);
  const pptxPath = path.join(outputDir, fixture.outputName || `${fixture.name || "presentation"}.pptx`);
  let pptx = await PresentationFile.exportPptx(presentation);
  pptx = await applyFixturePackageDrawing(pptx, fixture);
  pptx = await applyFixturePackageReview(pptx, fixture);
  const roundtripCodec = normalizeOpenChestnutCodecName(options.roundtripCodec || fixture.roundtripCodec || "none");
  if (!new Set(["none", "open-chestnut"]).has(roundtripCodec)) throw new Error(`Unsupported presentation roundtrip codec ${roundtripCodec}; expected none or open-chestnut.`);
  if (roundtripCodec === "open-chestnut") {
    const openChestnut = presentationOpenChestnutConfig(fixture);
    if (openChestnut?.nativeGraphFixture) {
      const embeddedWorkbook = Workbook.create();
      embeddedWorkbook.worksheets.add("Embedded").getRange("A1").values = [[openChestnut.embeddedWorkbook?.sourceValue || "OpenChestnut source workbook"]];
      const embeddedWorkbookFile = await SpreadsheetFile.exportXlsx(embeddedWorkbook);
      pptx = new FileBlob(await addOpenChestnutNativeGraphFixture(new Uint8Array(await pptx.arrayBuffer()), embeddedWorkbookFile.bytes), { type: PPTX_MIME });
    }
    const imported = await PresentationFile.importPptx(pptx, { codec: "open-chestnut" });
    for (const expected of openChestnut?.nativeObjects || []) {
      const object = imported.slides.items.flatMap((slide) => slide.nativeObjects.items).find((item) => item.nativeKind === expected.nativeKind);
      assert.ok(object, `Missing OpenChestnut native object ${expected.nativeKind}`);
      assert.equal(object.rootRelationships.length, Number(expected.relationships), `${expected.nativeKind} root relationship count`);
      assert.equal(object.parts.length, Number(expected.parts), `${expected.nativeKind} preserved part count`);
      assert.equal(object.editable, expected.editable === true, `${expected.nativeKind} placement editability`);
      assert.equal(imported.resolve(object.id), object, `${expected.nativeKind} must resolve by stable import ID`);
      const inspected = imported.inspect({ kind: "nativeObject", maxChars: fixture.qa?.maxChars || 30_000 }).ndjson;
      assert.match(inspected, new RegExp(`"nativeKind":"${expected.nativeKind}"`));
      assert.match(inspected, new RegExp(`"preservedParts":${Number(expected.parts)}`));
    }
    const edit = openChestnut?.edit;
    if (edit) {
      if (edit.clearMasterBackground) imported.master.clearBackground();
      else if (edit.masterBackground) imported.master.setBackground(edit.masterBackground);
      if (edit.clearLayoutBackground) imported.layouts.items[0].clearBackground();
      else if (edit.layoutBackground) imported.layouts.items[0].setBackground(edit.layoutBackground);
      if (edit.masterPlaceholder) {
        const placeholder = imported.master.placeholders.find((item) => item.type === edit.masterPlaceholder.type && item.idx === Number(edit.masterPlaceholder.idx));
        assert.ok(placeholder, `Missing OpenChestnut master placeholder ${edit.masterPlaceholder.type}:${edit.masterPlaceholder.idx}`);
        if (Object.hasOwn(edit.masterPlaceholder, "text")) placeholder.text = edit.masterPlaceholder.text;
        if (edit.masterPlaceholder.textBodyProperties) placeholder.textBodyProperties = edit.masterPlaceholder.textBodyProperties;
      }
      if (edit.layoutPlaceholder) {
        const placeholder = imported.layouts.items[0].placeholders.find((item) => item.type === edit.layoutPlaceholder.type && item.idx === Number(edit.layoutPlaceholder.idx));
        assert.ok(placeholder, `Missing OpenChestnut layout placeholder ${edit.layoutPlaceholder.type}:${edit.layoutPlaceholder.idx}`);
        if (Object.hasOwn(edit.layoutPlaceholder, "text")) placeholder.text = edit.layoutPlaceholder.text;
        if (edit.layoutPlaceholder.textBodyProperties) placeholder.textBodyProperties = edit.layoutPlaceholder.textBodyProperties;
      }
      if (edit.masterTextParagraphStyles) imported.master.textParagraphStyles = edit.masterTextParagraphStyles;
      for (const nativeEdit of edit.nativeObjects || []) {
        const object = imported.slides.items.flatMap((item) => item.nativeObjects.items).find((item) => item.nativeKind === nativeEdit.nativeKind);
        assert.ok(object, `Missing OpenChestnut editable native object ${nativeEdit.nativeKind}`);
        object.setName(nativeEdit.name ?? object.name);
        if (nativeEdit.position) object.setPosition(nativeEdit.position);
        if (nativeEdit.embeddedWorkbookValue) {
          const replacement = Workbook.create();
          replacement.worksheets.add("Embedded").getRange("A1").values = [[nativeEdit.embeddedWorkbookValue]];
          object.replaceEmbeddedWorkbook(await SpreadsheetFile.exportXlsx(replacement));
        }
      }
      const slide = imported.slides.getItem(Number(edit.slideIndex || 0));
      const shape = slide?.shapes.items.find((item) => item.name === edit.shapeName || item.id === edit.shapeId);
      assert.ok(shape, `Missing OpenChestnut editable shape ${edit.shapeName || edit.shapeId}`);
      shape.text.set(edit.text ?? shape.text.value);
      if (edit.textBodyProperties) shape.text.bodyProperties = edit.textBodyProperties;
      if (edit.paragraphStyles || edit.inheritedParagraphStyles) shape.text.inheritedParagraphStyles = edit.paragraphStyles || edit.inheritedParagraphStyles;
    }
    pptx = await PresentationFile.exportPptx(imported, { codec: "open-chestnut" });
    if (openChestnut?.edit?.nativeObjects?.length) {
      const edited = await PresentationFile.importPptx(pptx, { codec: "open-chestnut" });
      for (const expected of openChestnut.edit.nativeObjects) {
        const object = edited.slides.items.flatMap((slide) => slide.nativeObjects.items).find((item) => item.nativeKind === expected.nativeKind);
        assert.ok(object, `Missing edited OpenChestnut native object ${expected.nativeKind}`);
        assert.equal(object.name, expected.name, `${expected.nativeKind} edited name`);
        assert.deepEqual(object.position, expected.position, `${expected.nativeKind} edited position`);
        assert.equal(object.editable, true, `${expected.nativeKind} remains placement-editable`);
        if (expected.embeddedWorkbookValue) {
          const workbook = await SpreadsheetFile.importXlsx(object.getEmbeddedWorkbook());
          assert.equal(workbook.worksheets.getItem(0).getRange("A1").values[0][0], expected.embeddedWorkbookValue, `${expected.nativeKind} embedded workbook replacement`);
        }
      }
    }
  }
  await pptx.save(pptxPath);
  const qa = await verifyPresentationFile(pptxPath, {
    outputDir: path.join(outputDir, "qa"),
    nativeRender: options.nativeRender ?? fixture.qa?.nativeRender ?? "auto",
    baselineDir: options.baselineDir,
    writeBaseline: options.writeBaseline,
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    inspectKind: fixture.qa?.inspectKind,
    maxChars: fixture.qa?.maxChars,
  });
  return { fixture, pptxPath, qa, roundtripCodec };
}
