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

  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(absoluteInput));
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
  const visualQa = await visualQaArtifact(workbook, {
    format: renderFormat,
    renderer,
    sheetName,
    range,
    minBytes: options.minBytes ?? 20,
    maxChars: options.maxChars ?? 16_000,
  });

  const previewExtension = EXTENSION_BY_FORMAT[renderFormat];
  if (!previewExtension) throw new Error(`Unsupported spreadsheet preview format: ${renderFormat}`);
  const paths = {
    inspect: path.join(outputDir, "inspect.ndjson"),
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
    fs.writeFile(paths.verify, `${verifyNdjson}\n`, "utf8"),
    fs.writeFile(paths.layout, await layoutBlob.text(), "utf8"),
    fs.writeFile(paths.visualQa, visualQa.ndjson, "utf8"),
    visualQa.blob.save(paths.preview),
  ]);
  const summary = {
    input: absoluteInput,
    outputDir,
    sheetName,
    range: range || null,
    renderFormat,
    verifyOk: verify.ok,
    visualQaOk: visualQa.ok,
    renderHash: visualQa.summary.hash,
    files: paths,
  };
  await fs.writeFile(paths.summary, `${JSON.stringify(summary, null, 2)}\n`, "utf8");
  if (options.failOnIssues !== false && (!verify.ok || !visualQa.ok)) {
    throw new Error(`Spreadsheet QA failed: semantic=${verify.ok}, visual=${visualQa.ok}. See ${outputDir}`);
  }
  return { workbook, inspect, verify, visualQa, layoutBlob, summary };
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
  });
  return { fixture, workbookPath, qa };
}
