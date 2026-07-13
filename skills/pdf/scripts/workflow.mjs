import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

import {
  FileBlob,
  PdfArtifact,
  PdfFile,
  verifyArtifact,
} from "open-office-artifact-tool";
import { createPdfjsParser } from "open-office-artifact-tool/pdf/pdfjs";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import {
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../../shared/visual-baselines.mjs";

const PDF_MIME = "application/pdf";

function commandExists(command) {
  const result = spawnSync(process.platform === "win32" ? "where" : "which", [command], { encoding: "utf8", shell: false });
  return result.status === 0;
}

export function nativePdfRenderStatus() {
  const commands = { pdftoppm: commandExists("pdftoppm"), pdfinfo: commandExists("pdfinfo") };
  return { available: Object.values(commands).every(Boolean), commands };
}

export function createPdfFromFixture(fixture = {}) {
  const pdf = PdfArtifact.create({ metadata: { fixture: fixture.name || "pdf-fixture", ...(fixture.metadata || {}) }, pages: fixture.pages || [] });
  for (const [index, flow] of (fixture.flows || []).entries()) pdf.addFlowText(flow.text || "", { id: flow.id || flow.name || `fixture-flow-${index + 1}`, ...(flow.options || flow) });
  const inspectKind = fixture.qa?.inspectKind || "page,text,textItem,readingOrder,region,table,image,chart";
  for (const expected of fixture.expectInspect || []) assert.match(pdf.inspect({ kind: expected.kind || inspectKind, maxChars: fixture.qa?.maxChars || 30_000 }).ndjson, new RegExp(expected.pattern));
  return pdf;
}

function pdfInfo(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", shell: false });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/m.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return { pages, text: result.stdout };
}

async function renderModelPages(pdf, outputDir, options = {}) {
  const pagesDir = path.join(outputDir, "model-pages");
  const layoutsDir = path.join(outputDir, "layouts");
  await Promise.all([fs.mkdir(pagesDir, { recursive: true }), fs.mkdir(layoutsDir, { recursive: true })]);
  const renderer = createPlaywrightRenderer({ viewport: options.viewport || { width: 900, height: 1100 }, deviceScaleFactor: options.deviceScaleFactor ?? 1, timeout: options.timeout ?? 30_000 });
  const pages = [];
  const qaLines = [];
  const baselineSet = await prepareNumberedVisualBaselines(options.baselineDir, "model-page", options);
  for (let pageIndex = 0; pageIndex < pdf.pages.length; pageIndex += 1) {
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `model-page-${pageIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `model-page-${pageIndex + 1}.png`);
    const qa = await runPngVisualQa(pdf, { renderer, renderOptions: { pageIndex }, baselinePath, diffPath, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration, minBytes: options.minBytes, maxChars: options.maxChars });
    const pagePath = path.join(pagesDir, `page-${pageIndex + 1}.png`);
    const layoutPath = path.join(layoutsDir, `page-${pageIndex + 1}.json`);
    await Promise.all([qa.blob.save(pagePath), fs.writeFile(layoutPath, await (await pdf.render({ format: "layout", pageIndex })).text(), "utf8")]);
    qaLines.push(qa.ndjson);
    pages.push({ page: pageIndex + 1, path: pagePath, layoutPath, diffPath: qa.diffPath, bytes: qa.blob.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, pdf.pages.length, { artifactKind: "pdf", baselineKind: "model" });
  if (issue) qaLines.push(issue);
  const qaPath = path.join(outputDir, "model-visual-qa.ndjson");
  await fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8");
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pageCount: pdf.pages.length, baselinePageCount, pageCountMatches, pages, qaPath };
}

async function renderNativePages(pdfBlob, pdfPath, outputDir, expectedPages, options = {}) {
  const info = pdfInfo(pdfPath);
  if (info.pages !== expectedPages) throw new Error(`PDF reports ${info.pages} pages for ${expectedPages} modeled pages.`);
  const pagesDir = path.join(outputDir, "native-pages");
  await fs.mkdir(pagesDir, { recursive: true });
  const renderer = createPopplerRenderer({ dpi: options.dpi ?? 144, timeoutMs: options.nativeTimeout ?? 60_000 });
  const pages = [];
  const qaLines = [];
  const baselineSet = await prepareNumberedVisualBaselines(options.baselineDir, "native-page", options);
  for (let pageIndex = 0; pageIndex < info.pages; pageIndex += 1) {
    const png = await renderer({ input: pdfBlob, inputType: PDF_MIME, outputType: "image/png", format: "png", artifactKind: "pdf", pageIndex });
    const pagePath = path.join(pagesDir, `page-${pageIndex + 1}.png`);
    const baselinePath = options.baselineDir ? path.join(options.baselineDir, `native-page-${pageIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `native-page-${pageIndex + 1}.png`);
    const qa = await runPngVisualQa({ render: () => png }, { baselinePath, diffPath, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration, minBytes: options.minBytes, maxChars: options.maxChars });
    await png.save(pagePath);
    qaLines.push(qa.ndjson);
    pages.push({ page: pageIndex + 1, path: pagePath, diffPath: qa.diffPath, bytes: png.bytes.length, hash: qa.summary.hash, baselineCompared: Boolean(qa.summary.baselineHash), pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, info.pages, { artifactKind: "pdf", baselineKind: "native" });
  if (issue) qaLines.push(issue);
  const qaPath = path.join(outputDir, "native-visual-qa.ndjson");
  const infoPath = path.join(outputDir, "pdfinfo.txt");
  await Promise.all([fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8"), fs.writeFile(infoPath, info.text, "utf8")]);
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pageCount: info.pages, baselinePageCount, pageCountMatches, pages, qaPath, infoPath };
}

async function parseWithPdfjs(pdfBlob, expectedPages, outputDir, options = {}) {
  const requested = String(options.pdfjs ?? options.pdfjsParse ?? "auto").toLowerCase();
  if (requested === "off" || requested === "false") return { status: "skipped", reason: "PDF.js parsing disabled" };
  try {
    const parsed = await PdfFile.importPdf(pdfBlob, { parser: createPdfjsParser(), preferParser: true, parserName: "pdfjs" });
    if (parsed.pages.length !== expectedPages) throw new Error(`PDF.js parsed ${parsed.pages.length} pages for ${expectedPages} modeled pages.`);
    const verify = verifyArtifact(parsed, { maxChars: options.maxChars ?? 30_000 });
    const inspect = parsed.inspect({ kind: "page,text,textItem,region,table,image,chart", maxChars: options.maxChars ?? 30_000 });
    const text = parsed.extractText();
    const tables = parsed.extractTables();
    const paths = { inspect: path.join(outputDir, "pdfjs-inspect.ndjson"), verify: path.join(outputDir, "pdfjs-verify.ndjson"), text: path.join(outputDir, "pdfjs-text.txt"), tables: path.join(outputDir, "pdfjs-tables.json") };
    await Promise.all([fs.writeFile(paths.inspect, inspect.ndjson, "utf8"), fs.writeFile(paths.verify, `${verify.ndjson}\n`, "utf8"), fs.writeFile(paths.text, text, "utf8"), fs.writeFile(paths.tables, `${JSON.stringify(tables, null, 2)}\n`, "utf8")]);
    if (!verify.ok) throw new Error(`PDF.js parsed model failed verification. See ${paths.verify}`);
    return { status: "passed", pdf: parsed, verify, inspect, text, tables, paths };
  } catch (error) {
    if (requested === "required" || requested === "true") throw error;
    return { status: "skipped", reason: error.message };
  }
}

function expectedTaggedTableStructure(tables = []) {
  return tables.reduce((summary, table) => {
    const values = table.values || [];
    const rows = Math.max(1, values.length);
    const cells = Array.isArray(table.cells) && table.cells.length ? table.cells : Array.from({ length: rows }, (_, row) => Array.from({ length: Math.max(1, ...values.map((entry) => entry.length)) }, (_, column) => ({ role: row === 0 ? "TH" : "TD", row, column }))).flat();
    summary.tables += 1;
    summary.rows += rows;
    summary.headers += cells.filter((cell) => cell.role === "TH").length;
    summary.dataCells += cells.filter((cell) => cell.role === "TD").length;
    summary.rowSpans += cells.filter((cell) => Number(cell.rowSpan) > 1).length;
    summary.columnSpans += cells.filter((cell) => Number(cell.columnSpan) > 1).length;
    summary.headerAssociations += cells.filter((cell) => Array.isArray(cell.effectiveHeaders) && cell.effectiveHeaders.length).length;
    return summary;
  }, { tables: 0, rows: 0, headers: 0, dataCells: 0, rowSpans: 0, columnSpans: 0, headerAssociations: 0 });
}

export async function verifyPdfFile(inputPath, options = {}) {
  const absoluteInput = path.resolve(inputPath);
  const outputDir = path.resolve(options.outputDir || path.join(path.dirname(absoluteInput), `${path.basename(absoluteInput, path.extname(absoluteInput))}-qa`));
  await fs.mkdir(outputDir, { recursive: true });
  const loaded = await FileBlob.load(absoluteInput);
  const pdfBlob = new FileBlob(loaded.bytes, { type: PDF_MIME, name: path.basename(absoluteInput) });
  const pdf = await PdfFile.importPdf(pdfBlob);
  const maxChars = options.maxChars ?? 30_000;
  const inspect = pdf.inspect({ kind: options.inspectKind || "page,text,textItem,readingOrder,region,table,image,chart", maxChars });
  const fileInspect = await PdfFile.inspectPdf(pdfBlob, { maxObjects: options.maxObjects, maxChars });
  const verify = verifyArtifact(pdf, { maxChars });
  const extractedText = pdf.extractText();
  const extractedTables = pdf.extractTables();
  const baselineDir = options.baselineDir ? path.resolve(options.baselineDir) : undefined;
  const modelRender = await renderModelPages(pdf, outputDir, { ...options, maxChars, baselineDir });
  const paths = { inspect: path.join(outputDir, "inspect.ndjson"), fileInspect: path.join(outputDir, "file-inspect.ndjson"), verify: path.join(outputDir, "verify.ndjson"), text: path.join(outputDir, "text.txt"), tables: path.join(outputDir, "tables.json"), summary: path.join(outputDir, "summary.json") };
  await Promise.all([fs.writeFile(paths.inspect, inspect.ndjson, "utf8"), fs.writeFile(paths.fileInspect, fileInspect.ndjson, "utf8"), fs.writeFile(paths.verify, `${verify.ndjson}\n`, "utf8"), fs.writeFile(paths.text, extractedText, "utf8"), fs.writeFile(paths.tables, `${JSON.stringify(extractedTables, null, 2)}\n`, "utf8")]);

  const nativeStatus = nativePdfRenderStatus();
  const requestedNative = String(options.nativeRender ?? "auto").toLowerCase();
  let nativeRender = { status: "skipped", reason: "native render disabled" };
  if (requestedNative !== "off" && requestedNative !== "false") {
    if (nativeStatus.available) nativeRender = await renderNativePages(pdfBlob, absoluteInput, outputDir, pdf.pages.length, { ...options, maxChars, baselineDir });
    else if (requestedNative === "required" || requestedNative === "true") throw new Error(`Native PDF render requires pdftoppm and pdfinfo: ${JSON.stringify(nativeStatus.commands)}`);
    else nativeRender = { status: "skipped", reason: "Poppler commands unavailable", commands: nativeStatus.commands };
  }
  const pdfjs = await parseWithPdfjs(pdfBlob, pdf.pages.length, outputDir, { ...options, maxChars });
  const expectedTableStructure = expectedTaggedTableStructure(extractedTables);
  const tableStructurePassed = fileInspect.summary.tableStructures >= expectedTableStructure.tables
    && fileInspect.summary.tableRows >= expectedTableStructure.rows
    && fileInspect.summary.tableHeaders >= expectedTableStructure.headers
    && fileInspect.summary.tableDataCells >= expectedTableStructure.dataCells
    && fileInspect.summary.rowSpans >= expectedTableStructure.rowSpans
    && fileInspect.summary.columnSpans >= expectedTableStructure.columnSpans
    && fileInspect.summary.headerAssociations >= expectedTableStructure.headerAssociations;
  const expectedReadingOrderIds = pdf.pages.flatMap((page, pageIndex) => page.readingOrderRecords(pageIndex).map((record) => record.targetId));
  const readingOrderPassed = !fileInspect.summary.hasEmbeddedModel || JSON.stringify(fileInspect.summary.readingOrderIds) === JSON.stringify(expectedReadingOrderIds);
  const modeledFigures = pdf.pages.flatMap((page) => [...page.images, ...page.charts]);
  const expectedFigureAccessibility = { figures: modeledFigures.filter((figure) => !figure.decorative).length, artifacts: modeledFigures.filter((figure) => figure.decorative).length };
  const figureAccessibilityPassed = !fileInspect.summary.hasEmbeddedModel || (fileInspect.summary.figures === expectedFigureAccessibility.figures
    && fileInspect.summary.figureAltTexts === expectedFigureAccessibility.figures
    && fileInspect.summary.missingFigureAltTexts === 0
    && fileInspect.summary.artifacts >= expectedFigureAccessibility.artifacts);
  const accessibility = { requireTagged: options.requireTagged === true, tagged: fileInspect.summary.tagged, expectedTableStructure, tableStructurePassed, expectedReadingOrderIds, readingOrderPassed, expectedFigureAccessibility, figureAccessibilityPassed };
  const summary = { input: absoluteInput, outputDir, pages: pdf.pages.length, verifyOk: verify.ok, file: fileInspect.summary, accessibility, extractedTextChars: extractedText.length, extractedTables: extractedTables.length, baselineDir, writeBaseline: Boolean(options.writeBaseline), modelRender, nativeRender, pdfjs: { status: pdfjs.status, reason: pdfjs.reason, pages: pdfjs.pdf?.pages.length, textChars: pdfjs.text?.length, tables: pdfjs.tables?.length, paths: pdfjs.paths }, files: paths };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  const visualFailed = modelRender.ok === false || (nativeRender.status === "passed" && nativeRender.ok === false);
  const parserFailed = (String(options.pdfjs ?? options.pdfjsParse ?? "auto").toLowerCase() === "required" || String(options.pdfjs ?? options.pdfjsParse ?? "auto").toLowerCase() === "true") && pdfjs.status !== "passed";
  const accessibilityFailed = options.requireTagged === true && (!fileInspect.summary.tagged || !tableStructurePassed || !readingOrderPassed || !figureAccessibilityPassed);
  if (options.failOnIssues !== false && (!verify.ok || visualFailed || parserFailed || accessibilityFailed)) throw new Error(`PDF QA failed: semantic=${verify.ok}, visual=${!visualFailed}, native=${nativeRender.status}, pdfjs=${pdfjs.status}, tagged=${fileInspect.summary.tagged}, tableStructure=${tableStructurePassed}, readingOrder=${readingOrderPassed}, figures=${figureAccessibilityPassed}. See ${outputDir}`);
  return { pdf, inspect, fileInspect, verify, extractedText, extractedTables, modelRender, nativeRender, pdfjs, summary };
}

export async function runPdfFixture(fixturePath, options = {}) {
  const absoluteFixture = path.resolve(fixturePath);
  const fixture = JSON.parse(await fs.readFile(absoluteFixture, "utf8"));
  const outputDir = path.resolve(options.outputDir || path.join("tmp", "pdf-skill", fixture.name || "fixture"));
  await fs.mkdir(outputDir, { recursive: true });
  const pdf = createPdfFromFixture(fixture);
  const pdfPath = path.join(outputDir, fixture.outputName || `${fixture.name || "artifact"}.pdf`);
  const fixtureExport = fixture.export || {};
  await (await PdfFile.exportPdf(pdf, { ...fixtureExport, font: options.font ?? fixtureExport.font, maxFontBytes: options.maxFontBytes ?? fixtureExport.maxFontBytes, subsetFont: options.subsetFont ?? fixtureExport.subsetFont })).save(pdfPath);
  const qa = await verifyPdfFile(pdfPath, { outputDir: path.join(outputDir, "qa"), nativeRender: options.nativeRender ?? fixture.qa?.nativeRender ?? "auto", pdfjs: options.pdfjs ?? fixture.qa?.pdfjs ?? "auto", requireTagged: options.requireTagged ?? fixture.qa?.requireTagged ?? true, baselineDir: options.baselineDir, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration, inspectKind: fixture.qa?.inspectKind, maxChars: fixture.qa?.maxChars });
  for (const expected of fixture.expectText || []) assert.match(qa.extractedText, new RegExp(expected));
  for (const expected of fixture.expectPdfjsText || []) if (qa.pdfjs.status === "passed") assert.match(qa.pdfjs.text, new RegExp(expected));
  return { fixture, pdfPath, qa };
}
