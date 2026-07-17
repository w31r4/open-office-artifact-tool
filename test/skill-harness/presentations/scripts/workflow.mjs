import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

import {
  FileBlob,
  Presentation,
  PresentationFile,
  verifyArtifact,
  visualQaArtifact,
} from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
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

function addGroupShape(slide, group, config = {}) {
  const shape = group.shapes.add({ ...config, text: config.text || "" });
  if (config.textStyle) shape.text.style = { fontFamily: slide.presentation.theme.fonts.minor, ...config.textStyle };
  if (config.paragraphStyles || config.inheritedParagraphStyles) shape.text.inheritedParagraphStyles = config.paragraphStyles || config.inheritedParagraphStyles;
  return shape;
}

function addFixtureGroup(slide, owner, config, byName, remember) {
  const { shapes = [], tables = [], charts = [], images = [], groups = [], connectors = [], children, nativeObjects, ...groupConfig } = config;
  if (children != null || nativeObjects != null) throw new Error(`Presentation fixture group ${config.name || "group"} uses unsupported raw children or nativeObjects.`);
  const group = remember(owner.groups.add(groupConfig), config.name);
  for (const shape of shapes) remember(addGroupShape(slide, group, shape), shape.name);
  for (const table of tables) remember(group.tables.add(table), table.name);
  for (const chart of charts) remember(group.charts.add(chart.chartType || chart.type || "bar", chart), chart.name);
  for (const image of images) remember(group.images.add(image), image.name);
  for (const nested of groups) addFixtureGroup(slide, group, nested, byName, remember);
  for (const connector of connectors) {
    const from = byName.get(connector.fromName) || connector.from || connector.start;
    const to = byName.get(connector.toName) || connector.to || connector.end;
    remember(group.connectors.add({ ...connector, from, to }), connector.name);
  }
  return group;
}

async function addFixtureSlide(presentation, config = {}) {
  for (const field of ["layoutId", "notes", "applyLayoutPlaceholders", "comments"]) {
    if (config[field] != null) throw new Error(`Presentation fixture ${config.name || "slide"} uses unsupported 0.2 field ${field}.`);
  }
  const slide = presentation.slides.add({ name: config.name, ...(config.background ? { background: config.background } : {}) });
  const byName = new Map();
  const remember = (item, name) => { if (name) byName.set(name, item); return item; };
  for (const shape of config.shapes || []) remember(addSlideShape(slide, shape), shape.name);
  for (const table of config.tables || []) remember(slide.tables.add(table), table.name);
  for (const chart of config.charts || []) {
    remember(slide.charts.add(chart.chartType || chart.type || "bar", chart), chart.name);
  }
  for (const image of config.images || []) remember(slide.images.add(image), image.name);
  for (const group of config.groups || []) addFixtureGroup(slide, slide, group, byName, remember);
  for (const connector of config.connectors || []) {
    const from = byName.get(connector.fromName) || connector.from || connector.start;
    const to = byName.get(connector.toName) || connector.to || connector.end;
    remember(slide.connectors.add({ ...connector, from, to }), connector.name);
  }
  for (const comment of config.legacyComments || []) {
    const allowed = new Set(["text", "author", "created", "position"]);
    const unexpected = Object.keys(comment || {}).filter((key) => !allowed.has(key));
    if (unexpected.length || typeof comment?.text !== "string" || !comment.text.trim() || typeof comment?.author !== "string" || !comment.author.trim()) {
      throw new Error(`Presentation fixture ${config.name || "slide"} legacy comment must have only non-empty text and author plus optional created/position fields.`);
    }
    const position = comment.position;
    if (!position || !Number.isFinite(position.x) || !Number.isFinite(position.y) || (position.unit != null && position.unit !== "px" && position.unit !== "emu")) {
      throw new Error(`Presentation fixture ${config.name || "slide"} legacy comment requires position { x, y, unit?: 'px'|'emu' }.`);
    }
    slide.comments.addThread(undefined, comment.text, {
      author: comment.author,
      ...(comment.created == null ? {} : { created: comment.created }),
      position: { x: Number(position.x), y: Number(position.y), ...(position.unit == null ? {} : { unit: position.unit }) },
    });
  }
  return slide;
}

export async function createPresentationFromFixture(fixture = {}) {
  for (const field of ["theme", "master", "masters", "layouts", "customShows", "commentFormat", "packageReview", "openChestnut"]) {
    if (fixture[field] != null) throw new Error(`Presentation fixture ${fixture.name || "fixture"} uses unsupported 0.2 field ${field}.`);
  }
  const presentation = Presentation.create({ slideSize: fixture.slideSize || { width: 1280, height: 720 } });
  for (const slide of fixture.slides || []) await addFixtureSlide(presentation, slide);
  const inspectKind = fixture.qa?.inspectKind || "deck,slide,textbox,shape,table,chart,image,connector,textRange";
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
  source.slides.add().charts.add(chart.chartType || chart.type || "bar", chart);
  const zip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(source)).arrayBuffer()));
  const chartPath = Object.keys(zip.files).find((name) => /\/charts\/chart1\.xml$/.test(name));
  const xml = chartPath ? await zip.file(chartPath)?.async("text") : undefined;
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

function fixtureItem(items, edit, kind) {
  const item = items.find((candidate) => candidate.name === edit.sourceName || candidate.id === edit.id);
  assert.ok(item, `Missing presentation fixture ${kind} ${edit.sourceName || edit.id}`);
  return item;
}

function applyPresentationFixtureEdits(presentation, edit = {}) {
  if (!edit || Object.keys(edit).length === 0) return;
  const slide = presentation.slides.getItem(Number(edit.slideIndex || 0));
  assert.ok(slide, `Missing presentation fixture edit slide ${edit.slideIndex || 0}`);
  for (const change of edit.shapes || []) {
    const shape = fixtureItem(slide.shapes.items, change, "shape");
    if (Object.hasOwn(change, "name")) shape.name = change.name;
    if (Object.hasOwn(change, "text")) shape.text.set(change.text);
    if (Object.hasOwn(change, "position")) shape.position = { ...change.position };
    if (Object.hasOwn(change, "transform")) shape.transform = change.transform == null ? undefined : { ...change.transform };
    if (Object.hasOwn(change, "shadow")) shape.shadow = change.shadow == null ? undefined : { ...change.shadow };
  }
  for (const change of edit.tables || []) {
    const table = fixtureItem(slide.tables.items, change, "table");
    if (Object.hasOwn(change, "name")) table.name = change.name;
    if (Object.hasOwn(change, "position")) table.position = { ...change.position };
    for (const cell of change.cells || []) table.cells.set(Number(cell.row), Number(cell.column), cell.value);
  }
  for (const change of edit.images || []) {
    const image = fixtureItem(slide.images.items, change, "image");
    if (Object.hasOwn(change, "name")) image.name = change.name;
    if (Object.hasOwn(change, "alt")) image.alt = change.alt;
    if (Object.hasOwn(change, "position")) image.position = { ...change.position };
    if (Object.hasOwn(change, "fit")) image.fit = change.fit;
    if (Object.hasOwn(change, "crop")) image.crop = change.crop == null ? undefined : { ...change.crop };
    if (Object.hasOwn(change, "transform")) image.transform = change.transform == null ? undefined : { ...change.transform };
    if (Object.hasOwn(change, "dataUrl")) image.dataUrl = change.dataUrl;
  }
  for (const change of edit.connectors || []) {
    const connector = fixtureItem(slide.connectors.items, change, "connector");
    if (Object.hasOwn(change, "name")) connector.name = change.name;
    if (Object.hasOwn(change, "start")) connector.start = { ...change.start };
    if (Object.hasOwn(change, "end")) connector.end = { ...change.end };
    if (change.line) connector.line = { ...connector.line, ...change.line };
    for (const arrow of ["startArrow", "endArrow"]) {
      if (!Object.hasOwn(change, arrow)) continue;
      if (change[arrow] == null) delete connector.line[arrow];
      else connector.line[arrow] = change[arrow];
    }
  }
  for (const change of edit.charts || []) {
    const chart = fixtureItem(slide.charts.items, change, "chart");
    if (Object.hasOwn(change, "name")) chart.name = change.name;
    if (Object.hasOwn(change, "title")) chart.title = change.title;
    if (Object.hasOwn(change, "position")) chart.position = { ...change.position };
    if (change.seriesValues) {
      assert.equal(change.seriesValues.length, chart.series.length, `${chart.name || chart.id} edit must preserve series topology`);
      change.seriesValues.forEach((values, index) => {
        assert.equal(values.length, chart.series[index].values.length, `${chart.name || chart.id} edit must preserve point topology`);
        chart.series[index].values = [...values];
      });
    }
  }
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
  const imported = await PresentationFile.importPptx(pptx);
  applyPresentationFixtureEdits(imported, fixture.edit);
  pptx = await PresentationFile.exportPptx(imported);
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
  return { fixture, pptxPath, qa };
}
