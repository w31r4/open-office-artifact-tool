import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";

import {
  FileBlob,
  SpreadsheetFile,
  Workbook,
  verifyArtifact,
  visualQaArtifact,
} from "open-office-artifact-tool";
import { createPlaywrightRenderer } from "open-office-artifact-tool/renderers/playwright";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";
import { exportXlsxWithOpenXmlWasm, importXlsxWithOpenXmlWasm } from "open-office-artifact-tool/codecs/openxml-wasm";
import {
  loadVisualBaseline,
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../../shared/visual-baselines.mjs";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME = "text/csv";
const TSV_MIME = "text/tab-separated-values";

const EXTENSION_BY_FORMAT = {
  svg: "svg",
  png: "png",
  webp: "webp",
  jpg: "jpg",
  jpeg: "jpg",
  pdf: "pdf",
};

function rendererForFormat(format, options = {}) {
  if (format === "svg") return undefined;
  return createPlaywrightRenderer({
    viewport: options.viewport,
    deviceScaleFactor: options.deviceScaleFactor ?? 1,
    timeout: options.timeout ?? 30_000,
  });
}

function commandExists(command) {
  return spawnSync(process.platform === "win32" ? "where" : "which", [command], { encoding: "utf8", shell: false }).status === 0;
}

export function nativeSpreadsheetRenderStatus() {
  const commands = { soffice: commandExists("soffice"), pdftoppm: commandExists("pdftoppm"), pdfinfo: commandExists("pdfinfo") };
  return { available: Object.values(commands).every(Boolean), commands };
}

function safeFileSegment(value) {
  return String(value || "sheet").replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "sheet";
}

function pdfPageCount(pdfPath) {
  const result = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8" });
  if (result.status !== 0) throw new Error(`pdfinfo failed for ${pdfPath}: ${result.stderr || result.stdout}`);
  const pages = Number(/^Pages:\s+(\d+)/m.exec(result.stdout)?.[1]);
  if (!Number.isInteger(pages) || pages < 1) throw new Error(`pdfinfo did not report a valid page count for ${pdfPath}.`);
  return pages;
}

async function renderNativePages(inputBlob, outputDir, options = {}) {
  const pdf = await createLibreOfficeRenderer({ timeoutMs: options.nativeTimeout ?? 60_000 })({ input: inputBlob, inputType: options.inputType || inputBlob.type || XLSX_MIME, outputType: "application/pdf", format: "pdf", artifactKind: "workbook" });
  const pdfPath = path.join(outputDir, "native-render.pdf");
  await pdf.save(pdfPath);
  const pageCount = pdfPageCount(pdfPath);
  const pagesDir = path.join(outputDir, "native-pages");
  await fs.mkdir(pagesDir, { recursive: true });
  const baselineDir = options.baselineDir;
  const baselineSet = await prepareNumberedVisualBaselines(baselineDir, "native-page", options);
  const poppler = createPopplerRenderer({ dpi: options.dpi ?? 150, timeoutMs: options.nativeTimeout ?? 60_000 });
  const pages = [];
  const qaLines = [];
  for (let pageIndex = 0; pageIndex < pageCount; pageIndex += 1) {
    const png = await poppler({ input: pdf, inputType: "application/pdf", outputType: "image/png", format: "png", artifactKind: "workbook", pageIndex });
    const pagePath = path.join(pagesDir, `page-${pageIndex + 1}.png`);
    const baselinePath = baselineDir ? path.join(baselineDir, `native-page-${pageIndex + 1}.png`) : undefined;
    const diffPath = path.join(outputDir, "diffs", `native-page-${pageIndex + 1}.png`);
    const qa = await runPngVisualQa({ render: () => png }, { baselinePath, diffPath, writeBaseline: options.writeBaseline, pixelThreshold: options.pixelThreshold, diffAlignment: options.diffAlignment, diffPalette: options.diffPalette, pixelRegistration: options.pixelRegistration, minBytes: options.minBytes ?? 100, maxChars: options.maxChars ?? 16_000 });
    await png.save(pagePath);
    qaLines.push(qa.ndjson);
    pages.push({ page: pageIndex + 1, path: pagePath, diffPath: qa.diffPath, baselinePath, baselineCompared: Boolean(qa.summary.baselineHash), bytes: png.bytes.length, hash: qa.summary.hash, pixelDiff: qa.summary.pixelDiff, ok: qa.ok });
  }
  const { baselinePageCount, pageCountMatches, issue } = visualBaselineCountResult(baselineSet, pageCount, { artifactKind: "workbook", baselineKind: "native" });
  if (issue) qaLines.push(issue);
  const qaPath = path.join(outputDir, "native-visual-qa.ndjson");
  await fs.writeFile(qaPath, `${qaLines.filter(Boolean).join("\n")}\n`, "utf8");
  return { status: "passed", ok: pageCountMatches && pages.every((page) => page.ok), pdfPath, qaPath, pageCount, baselinePageCount, pageCountMatches, pages };
}

function applyRangeOperation(sheet, operation = {}) {
  if (!operation.range) throw new Error(`Spreadsheet fixture range operation on ${sheet.name} is missing range.`);
  const range = sheet.getRange(operation.range);
  if (operation.values) range.values = operation.values;
  if (operation.formulas) range.formulas = operation.formulas;
  if (operation.format) range.format = operation.format;
  if (operation.autofitColumns === true) range.format.autofitColumns();
  if (operation.autofitRows === true) range.format.autofitRows();
  if (operation.dataValidation) range.dataValidation = operation.dataValidation;
  for (const rule of operation.conditionalFormats || []) {
    range.conditionalFormats.add(rule.ruleType, rule.config || {});
  }
  if (operation.fillDown === true) range.fillDown();
  if (operation.fillRight === true) range.fillRight();
  if (operation.unmerge === true) range.unmerge();
  if (operation.merge === true || operation.merge === "across") range.merge(operation.merge === "across");
  return range;
}

export function createWorkbookFromFixture(fixture = {}) {
  const workbook = Workbook.create({ dateSystem: fixture.dateSystem, date1904: fixture.date1904, theme: fixture.theme });
  for (const sheetFixture of fixture.sheets || []) {
    const sheet = workbook.worksheets.add(sheetFixture.name);
    if (sheetFixture.showGridLines != null) sheet.showGridLines = Boolean(sheetFixture.showGridLines);
    if (sheetFixture.freezePanes?.rows != null) sheet.freezePanes.freezeRows(sheetFixture.freezePanes.rows);
    if (sheetFixture.freezePanes?.columns != null) sheet.freezePanes.freezeColumns(sheetFixture.freezePanes.columns);
    for (const operation of sheetFixture.ranges || []) applyRangeOperation(sheet, operation);
    for (const table of sheetFixture.tables || []) {
      const created = sheet.tables.add(table.range, table.hasHeaders !== false, table.name);
      if (table.style) created.style = table.style;
      if (table.showTotals != null) created.showTotals = table.showTotals;
      if (table.showBandedColumns != null) created.showBandedColumns = table.showBandedColumns;
    }
    for (const pivotFixture of sheetFixture.pivots || sheetFixture.pivotTables || []) sheet.pivotTables.add(pivotFixture);
    for (const chartFixture of sheetFixture.charts || []) {
      const source = chartFixture.sourceRange ? sheet.getRange(chartFixture.sourceRange) : chartFixture;
      const created = sheet.charts.add(chartFixture.chartType || chartFixture.type || "bar", source);
      if (chartFixture.name) created.name = chartFixture.name;
      if (chartFixture.title) created.title = chartFixture.title;
      if (chartFixture.hasLegend != null) created.hasLegend = Boolean(chartFixture.hasLegend);
      if (chartFixture.topLeft && chartFixture.bottomRight) created.setPosition(chartFixture.topLeft, chartFixture.bottomRight);
      else if (chartFixture.position) created.position = { ...chartFixture.position };
    }
    for (const imageFixture of sheetFixture.images || []) sheet.images.add(imageFixture);
  }
  if (fixture.commentSelf) workbook.comments.setSelf(fixture.commentSelf);
  for (const commentFixture of fixture.comments || []) {
    const sheet = workbook.worksheets.getItem(commentFixture.sheet);
    assert.ok(sheet, `Missing comment sheet ${commentFixture.sheet}`);
    const thread = workbook.comments.addThread({ cell: sheet.getRange(commentFixture.address) }, commentFixture.text || "", commentFixture);
    for (const reply of commentFixture.replies || []) thread.addReply(reply.text || "", reply);
    if (commentFixture.resolved) thread.resolve();
  }
  workbook.recalculate();
  for (const expectation of fixture.expectations || []) {
    const sheet = workbook.worksheets.getItem(expectation.sheet);
    assert.ok(sheet, `Missing expected sheet ${expectation.sheet}`);
    assert.deepEqual(sheet.getRange(expectation.range).values, expectation.values, `${expectation.sheet}!${expectation.range} did not match fixture values`);
  }
  return workbook;
}

export async function verifyWorkbookFile(inputPath, options = {}) {
  const absoluteInput = path.resolve(inputPath);
  const outputDir = path.resolve(options.outputDir || path.join(path.dirname(absoluteInput), `${path.basename(absoluteInput, path.extname(absoluteInput))}-qa`));
  await fs.mkdir(outputDir, { recursive: true });

  const requestedFormat = String(options.inputFormat || path.extname(absoluteInput).slice(1) || "xlsx").toLowerCase();
  if (!new Set(["xlsx", "csv", "tsv"]).has(requestedFormat)) throw new Error(`Unsupported spreadsheet input format: ${requestedFormat}. Expected xlsx, csv, or tsv.`);
  const inputType = requestedFormat === "csv" ? CSV_MIME : requestedFormat === "tsv" ? TSV_MIME : XLSX_MIME;
  const inputBlob = await FileBlob.load(absoluteInput, { type: inputType });
  const packageInspect = requestedFormat === "xlsx"
    ? await SpreadsheetFile.inspectXlsx(inputBlob, { includeText: options.includePackageText === true, maxChars: options.maxChars ?? 16_000 })
    : await SpreadsheetFile.inspectDelimited(inputBlob, { delimiter: requestedFormat === "tsv" ? "\t" : ",", maxChars: options.maxChars ?? 16_000 });
  const workbook = requestedFormat === "xlsx"
    ? await SpreadsheetFile.importXlsx(inputBlob)
    : await SpreadsheetFile.importDelimited(inputBlob, { delimiter: requestedFormat === "tsv" ? "\t" : ",", sheetName: options.sheetName || "Sheet1", coerceTypes: options.coerceTypes === true });
  workbook.recalculate();
  const sheetName = options.sheetName || workbook.worksheets.getItemAt(0)?.name;
  if (!sheetName) throw new Error("Workbook verification requires at least one worksheet.");
  const range = options.range;
  const inspect = workbook.inspect({
    kind: options.inspectKind || "workbook,sheet,table,formula,style,computedStyle,drawing,thread",
    sheetName,
    range,
    maxChars: options.maxChars ?? 16_000,
  });
  const verify = verifyArtifact(workbook, { maxChars: options.maxChars ?? 16_000 });
  const layoutBlob = await workbook.render({ format: "layout", sheetName, range });
  const renderFormat = String(options.renderFormat || "svg").toLowerCase();
  const renderer = rendererForFormat(renderFormat, options);
  const previewExtension = EXTENSION_BY_FORMAT[renderFormat];
  if (!previewExtension) throw new Error(`Unsupported spreadsheet preview format: ${renderFormat}`);
  const baselineDir = options.baselineDir ? path.resolve(options.baselineDir) : undefined;
  const baselinePath = baselineDir ? path.join(baselineDir, `${safeFileSegment(sheetName)}.${previewExtension}`) : undefined;
  const baseline = await loadVisualBaseline(baselinePath, options);
  const visualQa = await visualQaArtifact(workbook, {
    format: renderFormat,
    renderer,
    sheetName,
    range,
    baseline,
    pixelDiff: Boolean(baseline && ["png", "webp", "jpeg", "jpg"].includes(renderFormat)),
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    minBytes: options.minBytes ?? 20,
    maxChars: options.maxChars ?? 16_000,
  });
  if (options.writeBaseline && baselinePath) {
    await fs.mkdir(baselineDir, { recursive: true });
    await visualQa.blob.save(baselinePath);
  }
  const paths = {
    inspect: path.join(outputDir, "inspect.ndjson"),
    packageInspect: path.join(outputDir, "package-inspect.ndjson"),
    verify: path.join(outputDir, "verify.ndjson"),
    layout: path.join(outputDir, "layout.json"),
    visualQa: path.join(outputDir, "visual-qa.ndjson"),
    preview: path.join(outputDir, `preview.${previewExtension}`),
    summary: path.join(outputDir, "summary.json"),
  };
  if (visualQa.diffBlob) paths.diff = path.join(outputDir, `diff-${safeFileSegment(sheetName)}.png`);
  const verifyNdjson = verify.ndjson || JSON.stringify({
    kind: "verificationSummary",
    artifactKind: "workbook",
    ok: verify.ok,
    issues: verify.issues?.length || 0,
  });
  await Promise.all([
    fs.writeFile(paths.inspect, inspect.ndjson, "utf8"),
    fs.writeFile(paths.packageInspect, packageInspect.ndjson, "utf8"),
    fs.writeFile(paths.verify, `${verifyNdjson}\n`, "utf8"),
    fs.writeFile(paths.layout, await layoutBlob.text(), "utf8"),
    fs.writeFile(paths.visualQa, visualQa.ndjson, "utf8"),
    visualQa.blob.save(paths.preview),
    ...(paths.diff ? [visualQa.diffBlob.save(paths.diff)] : []),
  ]);
  const sheetRenders = [{
    sheetName,
    range: range || null,
    preview: paths.preview,
    layout: paths.layout,
    diffPath: paths.diff,
    baselinePath,
    baselineCompared: Boolean(baseline),
    pixelDiff: visualQa.summary.pixelDiff,
    hash: visualQa.summary.hash,
    ok: visualQa.ok,
  }];
  if (options.allSheets === true) {
    const sheetsDir = path.join(outputDir, "sheets");
    await fs.mkdir(sheetsDir, { recursive: true });
    for (const targetSheet of workbook.worksheets) {
      if (targetSheet.name === sheetName) continue;
      const segment = safeFileSegment(targetSheet.name);
      const targetBaselinePath = baselineDir ? path.join(baselineDir, `${segment}.${previewExtension}`) : undefined;
      const targetBaseline = await loadVisualBaseline(targetBaselinePath, options);
      const targetLayout = await workbook.render({ format: "layout", sheetName: targetSheet.name });
      const targetQa = await visualQaArtifact(workbook, {
        format: renderFormat,
        renderer,
        sheetName: targetSheet.name,
        baseline: targetBaseline,
        pixelDiff: Boolean(targetBaseline && ["png", "webp", "jpeg", "jpg"].includes(renderFormat)),
        pixelThreshold: options.pixelThreshold,
        diffAlignment: options.diffAlignment,
        diffPalette: options.diffPalette,
        pixelRegistration: options.pixelRegistration,
        minBytes: options.minBytes ?? 20,
        maxChars: options.maxChars ?? 16_000,
      });
      const targetPreviewPath = path.join(sheetsDir, `${segment}.${previewExtension}`);
      const targetLayoutPath = path.join(sheetsDir, `${segment}.layout.json`);
      const targetQaPath = path.join(sheetsDir, `${segment}.visual-qa.ndjson`);
      const targetDiffPath = targetQa.diffBlob ? path.join(sheetsDir, `${segment}.diff.png`) : undefined;
      await Promise.all([targetQa.blob.save(targetPreviewPath), fs.writeFile(targetLayoutPath, await targetLayout.text(), "utf8"), fs.writeFile(targetQaPath, targetQa.ndjson, "utf8"), ...(targetDiffPath ? [targetQa.diffBlob.save(targetDiffPath)] : [])]);
      if (options.writeBaseline && targetBaselinePath) await targetQa.blob.save(targetBaselinePath);
      sheetRenders.push({
        sheetName: targetSheet.name,
        range: null,
        preview: targetPreviewPath,
        layout: targetLayoutPath,
        visualQa: targetQaPath,
        diffPath: targetDiffPath,
        baselinePath: targetBaselinePath,
        baselineCompared: Boolean(targetBaseline),
        pixelDiff: targetQa.summary.pixelDiff,
        hash: targetQa.summary.hash,
        ok: targetQa.ok,
      });
    }
  }
  const requestedNative = String(options.nativeRender ?? "auto").toLowerCase();
  const nativeStatus = nativeSpreadsheetRenderStatus();
  let nativeRender = { status: "skipped", reason: "native render disabled" };
  if (requestedNative !== "off" && requestedNative !== "false") {
    if (nativeStatus.available) nativeRender = await renderNativePages(inputBlob, outputDir, { ...options, baselineDir, inputType });
    else if (requestedNative === "required" || requestedNative === "true") throw new Error(`Native spreadsheet render requires soffice, pdftoppm, and pdfinfo: ${JSON.stringify(nativeStatus.commands)}`);
    else nativeRender = { status: "skipped", reason: "native render commands unavailable", commands: nativeStatus.commands };
  }
  const summary = {
    input: absoluteInput,
    inputFormat: requestedFormat,
    inputType,
    outputDir,
    sheetName,
    range: range || null,
    renderFormat,
    baselineDir,
    baselinePath,
    writeBaseline: Boolean(options.writeBaseline),
    baselineCompared: Boolean(baseline),
    pixelDiff: visualQa.summary.pixelDiff,
    allSheets: options.allSheets === true,
    sheetRenders,
    nativeRender,
    packageOk: packageInspect.ok,
    verifyOk: verify.ok,
    visualQaOk: sheetRenders.every((item) => item.ok),
    renderHash: visualQa.summary.hash,
    files: paths,
  };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  if (options.failOnIssues !== false && (!packageInspect.ok || !verify.ok || sheetRenders.some((item) => !item.ok) || (nativeRender.status === "passed" && nativeRender.ok === false))) {
    throw new Error(`Spreadsheet QA failed: package=${packageInspect.ok}, semantic=${verify.ok}, visual=${sheetRenders.every((item) => item.ok)}, native=${nativeRender.status}. See ${outputDir}`);
  }
  return { workbook, inspect, packageInspect, verify, visualQa, layoutBlob, summary };
}

export async function runSpreadsheetFixture(fixturePath, options = {}) {
  const absoluteFixture = path.resolve(fixturePath);
  const fixture = JSON.parse(await fs.readFile(absoluteFixture, "utf8"));
  const outputDir = path.resolve(options.outputDir || path.join("tmp", "spreadsheet-skill", fixture.name || "fixture"));
  await fs.mkdir(outputDir, { recursive: true });
  const workbook = createWorkbookFromFixture(fixture);
  const workbookPath = path.join(outputDir, fixture.outputName || `${fixture.name || "workbook"}.xlsx`);
  const codec = String(options.codec || fixture.codec || "javascript").toLowerCase();
  if (!new Set(["javascript", "openxml-wasm"]).has(codec)) throw new Error(`Unsupported spreadsheet fixture codec ${codec}; expected javascript or openxml-wasm.`);
  let file = codec === "openxml-wasm"
    ? await exportXlsxWithOpenXmlWasm(workbook)
    : await SpreadsheetFile.exportXlsx(workbook);
  const roundtripCodec = String(options.roundtripCodec || fixture.roundtripCodec || "none").toLowerCase();
  if (!new Set(["none", "openxml-wasm"]).has(roundtripCodec)) throw new Error(`Unsupported spreadsheet roundtrip codec ${roundtripCodec}; expected none or openxml-wasm.`);
  if (roundtripCodec === "openxml-wasm") {
    const imported = await importXlsxWithOpenXmlWasm(file);
    file = await exportXlsxWithOpenXmlWasm(imported, { recalculate: false });
  }
  await file.save(workbookPath);
  const qa = await verifyWorkbookFile(workbookPath, {
    outputDir: path.join(outputDir, "qa"),
    sheetName: options.sheetName || fixture.qa?.sheetName,
    range: options.range || fixture.qa?.range,
    renderFormat: options.renderFormat || fixture.qa?.renderFormat || "svg",
    inspectKind: fixture.qa?.inspectKind,
    maxChars: fixture.qa?.maxChars,
    baselineDir: options.baselineDir,
    writeBaseline: options.writeBaseline,
    pixelThreshold: options.pixelThreshold,
    diffAlignment: options.diffAlignment,
    diffPalette: options.diffPalette,
    pixelRegistration: options.pixelRegistration,
    allSheets: options.allSheets,
    nativeRender: options.nativeRender ?? fixture.qa?.nativeRender ?? "auto",
  });
  return { fixture, workbookPath, qa, codec, roundtripCodec };
}
