import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

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
  return shape;
}

function addFixtureSlide(presentation, config = {}) {
  const slide = presentation.slides.add({ name: config.name, layoutId: config.layoutId, notes: config.notes });
  if (config.background) slide.background.fill = config.background;
  const byName = new Map();
  const remember = (item, name) => { if (name) byName.set(name, item); return item; };
  for (const shape of config.shapes || []) remember(addSlideShape(slide, shape), shape.name);
  for (const table of config.tables || []) remember(slide.tables.add(table), table.name);
  for (const chart of config.charts || []) remember(slide.charts.add(chart.chartType || chart.type || "bar", chart), chart.name);
  for (const image of config.images || []) remember(slide.images.add(image), image.name);
  for (const connector of config.connectors || []) {
    const from = byName.get(connector.fromName) || connector.from || connector.start;
    const to = byName.get(connector.toName) || connector.to || connector.end;
    remember(slide.connectors.add({ ...connector, from, to }), connector.name);
  }
  for (const comment of config.comments || []) {
    const target = byName.get(comment.targetName) || comment.targetId;
    assert.ok(target, `Missing presentation fixture comment target ${comment.targetName || comment.targetId}`);
    const thread = slide.comments.addThread(target, comment.text || "", comment);
    for (const reply of comment.replies || []) thread.addReply(reply.text || reply, typeof reply === "object" ? reply : {});
    if (comment.resolved) thread.resolve();
  }
  return slide;
}

export function createPresentationFromFixture(fixture = {}) {
  const presentation = Presentation.create({ slideSize: fixture.slideSize || { width: 1280, height: 720 }, theme: fixture.theme || {} });
  if (fixture.theme?.colors) presentation.theme.setColors(fixture.theme.colors);
  if (fixture.theme?.fonts) presentation.theme.setFonts(fixture.theme.fonts);
  for (const layout of fixture.layouts || []) presentation.layouts.add(layout);
  for (const slide of fixture.slides || []) addFixtureSlide(presentation, slide);
  const inspectKind = fixture.qa?.inspectKind || "deck,theme,layout,slide,textbox,shape,table,chart,image,connector,notes,comment,textRange";
  for (const expected of fixture.expectInspect || []) {
    assert.match(presentation.inspect({ kind: expected.kind || inspectKind, maxChars: fixture.qa?.maxChars || 30_000 }).ndjson, new RegExp(expected.pattern));
  }
  return presentation;
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", shell: false });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/m.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

async function optionalBaseline(baselinePath) {
  if (!baselinePath) return undefined;
  try { return await FileBlob.load(baselinePath); } catch (error) { if (error.code === "ENOENT") return undefined; throw error; }
}

async function numberedBaselineFiles(baselineDir, prefix) {
  if (!baselineDir) return [];
  try {
    const pattern = new RegExp(`^${prefix}-\\d+\\.png$`);
    return (await fs.readdir(baselineDir)).filter((name) => pattern.test(name)).map((name) => path.join(baselineDir, name));
  } catch (error) {
    if (error.code === "ENOENT") return [];
    throw error;
  }
}

async function prepareBaselineSet(baselineDir, prefix, writeBaseline) {
  if (!baselineDir) return [];
  const existing = await numberedBaselineFiles(baselineDir, prefix);
  if (writeBaseline) {
    await fs.mkdir(baselineDir, { recursive: true });
    await Promise.all(existing.map((filePath) => fs.unlink(filePath)));
    return [];
  }
  return existing;
}

async function runPngQa(artifact, options = {}) {
  const baseline = options.writeBaseline ? undefined : await optionalBaseline(options.baselinePath);
  const qa = await visualQaArtifact(artifact, {
    ...options.renderOptions,
    format: "png",
    renderer: options.renderer,
    baseline,
    pixelDiff: Boolean(baseline),
    pixelThreshold: options.pixelThreshold ?? 0,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    minBytes: options.minBytes ?? 100,
    maxChars: options.maxChars ?? 20_000,
  });
  if (options.writeBaseline && options.baselinePath) {
    await fs.mkdir(path.dirname(options.baselinePath), { recursive: true });
    await qa.blob.save(options.baselinePath);
  }
  if (qa.diffBlob && options.diffPath) {
    await fs.mkdir(path.dirname(options.diffPath), { recursive: true });
    await qa.diffBlob.save(options.diffPath);
    qa.diffPath = options.diffPath;
  }
  return qa;
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
  const existingBaselines = await prepareBaselineSet(options.baselineDir, "native-slide", options.writeBaseline);
  for (let slideIndex = 0; slideIndex < pageCount; slideIndex += 1) {
    const png = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "presentation", pageIndex: slideIndex });
    const slidePath = path.join(pagesDir, `slide-${slideIndex + 1}.png`);
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `native-slide-${slideIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `native-slide-${slideIndex + 1}.png`);
    const qa = await runPngQa({ export: () => png }, { baselinePath, diffPath, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, minBytes: options.minBytes, maxChars: options.maxChars });
    await png.save(slidePath);
    qaLines.push(qa.ndjson);
    pages.push({ slide: slideIndex + 1, path: slidePath, diffPath: qa.diffPath, bytes: png.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const baselinePageCount = options.baselineDir && !options.writeBaseline ? existingBaselines.length : undefined;
  const pageCountMatches = baselinePageCount == null || baselinePageCount === 0 || baselinePageCount === pageCount;
  if (!pageCountMatches) qaLines.push(JSON.stringify({ kind: "visualPageCountDiff", artifactKind: "presentation", severity: "warning", pageCount, baselinePageCount, baselineKind: "native" }));
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
  const existingBaselines = await prepareBaselineSet(options.baselineDir, "model-slide", options.writeBaseline);
  for (let slideIndex = 0; slideIndex < presentation.slides.count; slideIndex += 1) {
    const slide = presentation.slides.getItem(slideIndex);
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `model-slide-${slideIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `model-slide-${slideIndex + 1}.png`);
    const qa = await runPngQa(presentation, {
      renderer,
      renderOptions: { slide },
      baselinePath,
      diffPath,
      writeBaseline: options.writeBaseline,
      pixelThreshold: options.pixelThreshold,
      diffAlignment: options.diffAlignment,
      diffPalette: options.diffPalette,
      minBytes: options.minBytes,
      maxChars: options.maxChars,
    });
    const slidePath = path.join(slidesDir, `slide-${slideIndex + 1}.png`);
    const layoutPath = path.join(layoutsDir, `slide-${slideIndex + 1}.json`);
    await Promise.all([qa.blob.save(slidePath), fs.writeFile(layoutPath, await (await slide.export({ format: "layout" })).text(), "utf8")]);
    qaLines.push(qa.ndjson);
    slides.push({ slide: slideIndex + 1, path: slidePath, layoutPath, diffPath: qa.diffPath, bytes: qa.blob.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const baselinePageCount = options.baselineDir && !options.writeBaseline ? existingBaselines.length : undefined;
  const pageCountMatches = baselinePageCount == null || baselinePageCount === 0 || baselinePageCount === presentation.slides.count;
  if (!pageCountMatches) qaLines.push(JSON.stringify({ kind: "visualPageCountDiff", artifactKind: "presentation", severity: "warning", pageCount: presentation.slides.count, baselinePageCount, baselineKind: "model" }));
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
  const presentation = createPresentationFromFixture(fixture);
  const pptxPath = path.join(outputDir, fixture.outputName || `${fixture.name || "presentation"}.pptx`);
  await (await PresentationFile.exportPptx(presentation)).save(pptxPath);
  const qa = await verifyPresentationFile(pptxPath, {
    outputDir: path.join(outputDir, "qa"),
    nativeRender: options.nativeRender ?? fixture.qa?.nativeRender ?? "auto",
    baselineDir: options.baselineDir,
    writeBaseline: options.writeBaseline,
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    inspectKind: fixture.qa?.inspectKind,
    maxChars: fixture.qa?.maxChars,
  });
  return { fixture, pptxPath, qa };
}
