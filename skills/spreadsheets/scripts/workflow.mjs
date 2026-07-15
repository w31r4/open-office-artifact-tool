import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import JSZip from "jszip";

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
import { exportXlsxWithOpenChestnut, importXlsxWithOpenChestnut } from "open-office-artifact-tool/codecs/open-chestnut";
import { normalizeOpenChestnutCodecName } from "../../shared/open-chestnut-compat.mjs";
import {
  loadVisualBaseline,
  prepareNumberedVisualBaselines,
  runPngVisualQa,
  visualBaselineCountResult,
} from "../../shared/visual-baselines.mjs";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const CSV_MIME = "text/csv";
const TSV_MIME = "text/tab-separated-values";

function fixtureXmlEscape(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function sourceQuerySortStateXml(sortState) {
  if (!sortState) return "";
  const reference = String(sortState.reference || "");
  const conditions = Array.isArray(sortState.conditions) ? sortState.conditions : [];
  if (!reference || conditions.length === 0) throw new Error("sourceQueryTableFixture.sortState requires a reference and at least one condition.");
  const conditionXml = conditions.map((condition) => {
    const conditionReference = String(condition?.reference || "");
    if (!conditionReference) throw new Error("sourceQueryTableFixture.sortState conditions require references.");
    const descending = condition.descending ? ' descending="1"' : "";
    if (condition.kind === "icon" || condition.iconSet) {
      const iconSet = String(condition.iconSet || "");
      if (!iconSet) throw new Error("sourceQueryTableFixture icon sorts require iconSet.");
      const iconId = condition.iconId == null ? "" : ` iconId="${Number(condition.iconId)}"`;
      return `<x:sortCondition ref="${fixtureXmlEscape(conditionReference)}"${descending} sortBy="icon" iconSet="${fixtureXmlEscape(iconSet)}"${iconId}/>`;
    }
    const customList = condition.customList == null ? "" : ` customList="${fixtureXmlEscape(condition.customList)}"`;
    return `<x:sortCondition ref="${fixtureXmlEscape(conditionReference)}"${descending}${customList}/>`;
  }).join("");
  const caseSensitive = sortState.caseSensitive ? ' caseSensitive="1"' : "";
  const sortMethod = sortState.sortMethod == null ? "" : ` sortMethod="${fixtureXmlEscape(sortState.sortMethod)}"`;
  const columnSort = sortState.columnSort == null ? "" : ` columnSort="${sortState.columnSort ? 1 : 0}"`;
  return `<x:sortState ref="${fixtureXmlEscape(reference)}"${caseSensitive}${sortMethod}${columnSort}>${conditionXml}<x:extLst><x:ext uri="{A1E10EA8-3B88-4BE3-9884-625AB42E9DDC}"><fixture:sortOpaque value="kept"/></x:ext></x:extLst></x:sortState>`;
}

async function attachSourceQueryTableFixture(file, config = {}) {
  const tablePartPath = String(config.tablePartPath || "xl/tables/table1.xml");
  const queryPartPath = String(config.queryPartPath || "xl/queryTables/queryTable1.xml");
  if (!/^xl\/tables\/table[1-9][0-9]*\.xml$/.test(tablePartPath) || !/^xl\/queryTables\/queryTable[1-9][0-9]*\.xml$/.test(queryPartPath))
    throw new Error("sourceQueryTableFixture requires canonical numbered table/queryTable part paths.");
  const bytes = new Uint8Array(await file.arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  if (!zip.file(tablePartPath)) throw new Error(`sourceQueryTableFixture cannot find ${tablePartPath}.`);
  if (zip.file("xl/connections.xml") || zip.file(queryPartPath)) throw new Error("sourceQueryTableFixture refuses to replace an existing connection or QueryTable part.");
  const contentTypes = await zip.file("[Content_Types].xml")?.async("text");
  const workbookRelationships = await zip.file("xl/_rels/workbook.xml.rels")?.async("text");
  if (!contentTypes?.includes("</Types>") || !workbookRelationships?.includes("</Relationships>"))
    throw new Error("sourceQueryTableFixture requires readable content types and workbook relationships.");
  const tableName = path.posix.basename(tablePartPath);
  const tableRelationshipPath = `${path.posix.dirname(tablePartPath)}/_rels/${tableName}.rels`;
  if (zip.file(tableRelationshipPath)) throw new Error(`sourceQueryTableFixture refuses to replace ${tableRelationshipPath}.`);
  const connectionId = Number(config.connectionId ?? 7);
  if (!Number.isInteger(connectionId) || connectionId <= 0) throw new Error("sourceQueryTableFixture.connectionId must be a positive integer.");
  const fields = Array.isArray(config.fields) && config.fields.length ? config.fields.map(String) : ["Key", "Value"];
  const queryTarget = path.posix.relative(path.posix.dirname(tablePartPath), queryPartPath);
  zip.file("[Content_Types].xml", contentTypes.replace(
    "</Types>",
    `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/><Override PartName="/${fixtureXmlEscape(queryPartPath)}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>`,
  ));
  zip.file("xl/_rels/workbook.xml.rels", workbookRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>',
  ));
  zip.file(tableRelationshipPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="${fixtureXmlEscape(queryTarget)}"/></Relationships>`);
  zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture"><x:connection id="${connectionId}" name="${fixtureXmlEscape(config.connectionName || "Fixture connection")}" description="Read-only fixture source" type="5" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="0" saveData="1" savePassword="0" credentials="integrated"><x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="SELECT fixture fields" commandType="2"/><x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="kept"/></x:ext></x:extLst></x:connection></x:connections>`);
  const queryFields = fields.map((name, index) => index === 0
    ? `<x:queryTableField id="1" name="${fixtureXmlEscape(name)}" dataBound="1" tableColumnId="1" fillFormulas="0" clipped="0"><x:extLst><x:ext uri="{71C44015-E485-449B-93BE-190C959F820F}"><fixture:fieldOpaque value="kept"/></x:ext></x:extLst></x:queryTableField>`
    : `<x:queryTableField id="${index + 1}" name="${fixtureXmlEscape(name)}" dataBound="1" tableColumnId="${index + 1}"/>`).join("");
  const deletedFieldNames = Array.isArray(config.deletedFieldNames) ? config.deletedFieldNames.map(String) : [];
  const deletedFieldsXml = deletedFieldNames.length
    ? `<x:queryTableDeletedFields count="${deletedFieldNames.length}">${deletedFieldNames.map((name) => `<x:deletedField name="${fixtureXmlEscape(name)}"/>`).join("")}</x:queryTableDeletedFields>`
    : "";
  const sortStateXml = sourceQuerySortStateXml(config.sortState);
  zip.file(queryPartPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture" name="${fixtureXmlEscape(config.queryName || "Fixture query")}" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" firstBackgroundRefresh="0" refreshOnLoad="0" growShrinkType="insertClear" fillFormulas="0" removeDataOnSave="0" disableEdit="0" preserveFormatting="1" adjustColumnWidth="1" intermediate="0" connectionId="${connectionId}"><x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="${fields.length + 1}" unboundColumnsLeft="0" unboundColumnsRight="0"><x:queryTableFields count="${fields.length}">${queryFields}</x:queryTableFields>${deletedFieldsXml}${sortStateXml}</x:queryTableRefresh><x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:opaque value="kept"/></x:ext></x:extLst></x:queryTable>`);
  return new FileBlob(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: XLSX_MIME, name: "source-query-table-fixture.xlsx" });
}

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

function fixtureCellAddress(row, column) {
  let number = column + 1;
  let label = "";
  while (number > 0) {
    number -= 1;
    label = String.fromCharCode(65 + number % 26) + label;
    number = Math.floor(number / 26);
  }
  return `${label}${row + 1}`;
}

function applyFormulaMetadata(sheet, range, metadata = {}) {
  const kind = String(metadata.kind || metadata.type || "");
  const rangeAddress = `${fixtureCellAddress(range.bounds.top, range.bounds.left)}${range.bounds.top === range.bounds.bottom && range.bounds.left === range.bounds.right ? "" : `:${fixtureCellAddress(range.bounds.bottom, range.bounds.right)}`}`;
  const reference = String(metadata.reference || (kind === "shared" ? rangeAddress : ""));
  if (kind === "shared") {
    if (!Number.isInteger(metadata.sharedIndex) || metadata.sharedIndex < 0) throw new Error(`Spreadsheet fixture ${sheet.name}!${rangeAddress} shared formula metadata requires a non-negative sharedIndex.`);
    for (let row = range.bounds.top; row <= range.bounds.bottom; row += 1) {
      for (let column = range.bounds.left; column <= range.bounds.right; column += 1) {
        Object.assign(sheet.store.get(fixtureCellAddress(row, column)), { formulaType: "shared", sharedIndex: metadata.sharedIndex, sharedRef: reference });
      }
    }
    return;
  }
  if (kind === "array") {
    if (!reference) throw new Error(`Spreadsheet fixture ${sheet.name}!${rangeAddress} array formula metadata requires reference.`);
    Object.assign(sheet.store.get(fixtureCellAddress(range.bounds.top, range.bounds.left)), { formulaType: "array", arrayRef: reference });
    return;
  }
  throw new Error(`Spreadsheet fixture ${sheet.name}!${rangeAddress} formulaMetadata.kind must be shared or array.`);
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
  if (operation.formulaMetadata) applyFormulaMetadata(sheet, range, operation.formulaMetadata);
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
    if (sheetFixture.sortState) sheet.sortState = { ...sheetFixture.sortState, conditions: (sheetFixture.sortState.conditions || []).map((condition) => ({ ...condition })) };
    for (const operation of sheetFixture.ranges || []) applyRangeOperation(sheet, operation);
    for (const table of sheetFixture.tables || []) {
      const created = sheet.tables.add({ range: table.range, name: table.name, hasHeaders: table.hasHeaders !== false, columnNames: table.columnNames, columnDefinitions: table.columnDefinitions, filters: table.filters, sortState: table.sortState });
      if (table.style) created.style = table.style;
      if (table.showTotals != null) created.showTotals = table.showTotals;
      if (table.showFilterButton != null) created.showFilterButton = table.showFilterButton;
      if (table.showFirstColumn != null) created.showFirstColumn = table.showFirstColumn;
      if (table.showLastColumn != null) created.showLastColumn = table.showLastColumn;
      if (table.showRowStripes != null) created.showRowStripes = table.showRowStripes;
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
  const codec = normalizeOpenChestnutCodecName(options.codec || fixture.codec || "javascript");
  if (!new Set(["javascript", "open-chestnut"]).has(codec)) throw new Error(`Unsupported spreadsheet fixture codec ${codec}; expected javascript or open-chestnut.`);
  let file = codec === "open-chestnut"
    ? await exportXlsxWithOpenChestnut(workbook)
    : await SpreadsheetFile.exportXlsx(workbook);
  let sourceQueryTable;
  let sourceConnections;
  if (fixture.sourceQueryTableFixture) {
    if (codec !== "open-chestnut") throw new Error("sourceQueryTableFixture requires codec=open-chestnut.");
    file = await attachSourceQueryTableFixture(file, fixture.sourceQueryTableFixture);
    const imported = await importXlsxWithOpenChestnut(file);
    const sourceQuery = fixture.sourceQueryTableFixture;
    const sheet = imported.worksheets.getItem(sourceQuery.sheet);
    const table = sheet?.tables.getItemOrNullObject(sourceQuery.table);
    if (!sheet || !table || table.isNullObject || !table.queryTable)
      throw new Error(`sourceQueryTableFixture could not resolve ${sourceQuery.sheet}!${sourceQuery.table}.`);
    Object.assign(table.queryTable, sourceQuery.edit || {});
    if (sourceQuery.connectionEdit) {
      const connection = imported.connections.find((candidate) => candidate.connectionId === Number(sourceQuery.connectionId));
      if (!connection) throw new Error(`sourceQueryTableFixture.connectionEdit cannot resolve connection ${sourceQuery.connectionId}.`);
      const { connectionId: _identity, type: _type, refreshedVersion: _version, ...changes } = sourceQuery.connectionEdit;
      Object.assign(connection, changes);
    }
    if (sourceQuery.refreshEdit) {
      if (!table.queryTable.refresh) throw new Error("sourceQueryTableFixture.refreshEdit requires a recognized queryTableRefresh profile.");
      const { fields: fieldEdits = [], ...refreshEdit } = sourceQuery.refreshEdit;
      Object.assign(table.queryTable.refresh, refreshEdit);
      const editedIds = new Set();
      for (const patch of fieldEdits) {
        const id = Number(patch?.id);
        if (!Number.isInteger(id) || id <= 0 || editedIds.has(id)) throw new Error("sourceQueryTableFixture.refreshEdit fields require unique positive source IDs.");
        editedIds.add(id);
        const field = table.queryTable.refresh.fields.find((candidate) => candidate.id === id);
        if (!field) throw new Error(`sourceQueryTableFixture.refreshEdit cannot resolve query field ${id}.`);
        const { id: _identity, tableColumnId: _binding, ...changes } = patch;
        Object.assign(field, changes);
      }
    }
    file = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
    sourceQueryTable = { sheet: sourceQuery.sheet, table: sourceQuery.table, query: structuredClone(table.queryTable) };
    sourceConnections = structuredClone(imported.connections);
  }
  const roundtripCodec = normalizeOpenChestnutCodecName(options.roundtripCodec || fixture.roundtripCodec || "none");
  if (!new Set(["none", "open-chestnut"]).has(roundtripCodec)) throw new Error(`Unsupported spreadsheet roundtrip codec ${roundtripCodec}; expected none or open-chestnut.`);
  if (roundtripCodec === "open-chestnut") {
    const imported = await importXlsxWithOpenChestnut(file);
    file = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
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
  return { fixture, workbookPath, qa, codec, roundtripCodec, sourceQueryTable, sourceConnections };
}
