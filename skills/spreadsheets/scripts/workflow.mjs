import assert from "node:assert/strict";
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

async function optionalBaseline(baselinePath) {
  if (!baselinePath) return undefined;
  try { return await FileBlob.load(baselinePath); } catch (error) { if (error.code === "ENOENT") return undefined; throw error; }
}

function safeFileSegment(value) {
  return String(value || "sheet").replace(/[^A-Za-z0-9._-]+/g, "-").replace(/^-+|-+$/g, "") || "sheet";
}

function applyRangeOperation(sheet, operation = {}) {
  if (!operation.range) throw new Error(`Spreadsheet fixture range operation on ${sheet.name} is missing range.`);
  const range = sheet.getRange(operation.range);
  if (operation.values) range.values = operation.values;
  if (operation.formulas) range.formulas = operation.formulas;
  if (operation.format) range.format = operation.format;
  if (operation.dataValidation) range.dataValidation = operation.dataValidation;
  for (const rule of operation.conditionalFormats || []) {
    range.conditionalFormats.add(rule.ruleType, rule.config || {});
  }
  return range;
}

export function createWorkbookFromFixture(fixture = {}) {
  const workbook = Workbook.create();
  for (const sheetFixture of fixture.sheets || []) {
    const sheet = workbook.worksheets.add(sheetFixture.name);
    for (const operation of sheetFixture.ranges || []) applyRangeOperation(sheet, operation);
    for (const table of sheetFixture.tables || []) {
      const created = sheet.tables.add(table.range, table.hasHeaders !== false, table.name);
      if (table.style) created.style = table.style;
      if (table.showTotals != null) created.showTotals = table.showTotals;
      if (table.showBandedColumns != null) created.showBandedColumns = table.showBandedColumns;
    }
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

  const xlsxBlob = await FileBlob.load(absoluteInput);
  const packageInspect = await SpreadsheetFile.inspectXlsx(xlsxBlob, { includeText: options.includePackageText === true, maxChars: options.maxChars ?? 16_000 });
  const workbook = await SpreadsheetFile.importXlsx(xlsxBlob);
  workbook.recalculate();
  const sheetName = options.sheetName || workbook.worksheets.getItemAt(0)?.name;
  if (!sheetName) throw new Error("Workbook verification requires at least one worksheet.");
  const range = options.range;
  const inspect = workbook.inspect({
    kind: options.inspectKind || "workbook,sheet,table,formula,style,computedStyle,drawing",
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
  const baseline = options.writeBaseline ? undefined : await optionalBaseline(baselinePath);
  const visualQa = await visualQaArtifact(workbook, {
    format: renderFormat,
    renderer,
    sheetName,
    range,
    baseline,
    pixelDiff: Boolean(baseline && ["png", "webp", "jpeg", "jpg"].includes(renderFormat)),
    pixelThreshold: options.pixelThreshold,
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
  ]);
  const sheetRenders = [{
    sheetName,
    range: range || null,
    preview: paths.preview,
    layout: paths.layout,
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
      const targetBaseline = options.writeBaseline ? undefined : await optionalBaseline(targetBaselinePath);
      const targetLayout = await workbook.render({ format: "layout", sheetName: targetSheet.name });
      const targetQa = await visualQaArtifact(workbook, {
        format: renderFormat,
        renderer,
        sheetName: targetSheet.name,
        baseline: targetBaseline,
        pixelDiff: Boolean(targetBaseline && ["png", "webp", "jpeg", "jpg"].includes(renderFormat)),
        pixelThreshold: options.pixelThreshold,
        minBytes: options.minBytes ?? 20,
        maxChars: options.maxChars ?? 16_000,
      });
      const targetPreviewPath = path.join(sheetsDir, `${segment}.${previewExtension}`);
      const targetLayoutPath = path.join(sheetsDir, `${segment}.layout.json`);
      const targetQaPath = path.join(sheetsDir, `${segment}.visual-qa.ndjson`);
      await Promise.all([targetQa.blob.save(targetPreviewPath), fs.writeFile(targetLayoutPath, await targetLayout.text(), "utf8"), fs.writeFile(targetQaPath, targetQa.ndjson, "utf8")]);
      if (options.writeBaseline && targetBaselinePath) await targetQa.blob.save(targetBaselinePath);
      sheetRenders.push({
        sheetName: targetSheet.name,
        range: null,
        preview: targetPreviewPath,
        layout: targetLayoutPath,
        visualQa: targetQaPath,
        baselinePath: targetBaselinePath,
        baselineCompared: Boolean(targetBaseline),
        pixelDiff: targetQa.summary.pixelDiff,
        hash: targetQa.summary.hash,
        ok: targetQa.ok,
      });
    }
  }
  const summary = {
    input: absoluteInput,
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
    packageOk: packageInspect.ok,
    verifyOk: verify.ok,
    visualQaOk: sheetRenders.every((item) => item.ok),
    renderHash: visualQa.summary.hash,
    files: paths,
  };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  if (options.failOnIssues !== false && (!packageInspect.ok || !verify.ok || sheetRenders.some((item) => !item.ok))) {
    throw new Error(`Spreadsheet QA failed: package=${packageInspect.ok}, semantic=${verify.ok}, visual=${sheetRenders.every((item) => item.ok)}. See ${outputDir}`);
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
  await (await SpreadsheetFile.exportXlsx(workbook)).save(workbookPath);
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
    allSheets: options.allSheets,
  });
  return { fixture, workbookPath, qa };
}
