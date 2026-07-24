import crypto from "node:crypto";
import { constants as fsConstants } from "node:fs";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

async function assertNewFile(filePath, label) {
  try {
    await fs.access(filePath);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists; refusing to overwrite it.`);
}

async function publishNoOverwrite(temporaryPath, finalPath, label) {
  try {
    await fs.copyFile(temporaryPath, finalPath, fsConstants.COPYFILE_EXCL);
  } catch (error) {
    if (error?.code === "EEXIST") throw new Error(`${label} already exists; refusing to overwrite it.`);
    throw error;
  }
}

function locatePivot(workbook, worksheetName, pivotName) {
  const sheet = workbook.worksheets.items.find((candidate) => candidate.name === worksheetName);
  if (!sheet) throw new Error(`Worksheet ${JSON.stringify(worksheetName)} was not found.`);
  const matches = sheet.pivotTables.items.filter((candidate) => candidate.name === pivotName);
  if (matches.length !== 1) throw new Error(`Expected exactly one PivotTable ${JSON.stringify(pivotName)} on ${JSON.stringify(worksheetName)}; found ${matches.length}.`);
  return matches[0];
}

function pivotProjection(pivot) {
  const record = pivot.inspectRecord();
  return {
    sheet: record.sheet,
    name: record.name,
    sourceRange: record.sourceRange,
    sourceSheet: record.sourceSheet,
    targetRange: record.targetRange,
    rowFields: record.rowFields,
    columnFields: record.columnFields,
    valueFields: record.valueFields,
    groupFields: record.groupFields,
    calculatedFields: record.calculatedFields,
    filters: record.filters,
    refreshPolicy: record.refreshPolicy,
    rowGrandTotals: record.rowGrandTotals,
    columnGrandTotals: record.columnGrandTotals,
    values: record.values,
    sourceCapabilities: pivot.sourceCapabilities,
  };
}

function expectedRefreshOnLoadHardening(projection, capabilities) {
  return {
    ...projection,
    refreshPolicy: { ...projection.refreshPolicy, refreshOnLoad: false },
    sourceCapabilities: capabilities,
  };
}

function readRefreshOnLoadRoot(xml, label) {
  const root = /<(?:[\w.-]+:)?pivotCacheDefinition\b[^>]*>/i.exec(xml);
  if (!root || root.index == null) throw new Error(`${label} is not a Pivot cache-definition root.`);
  const attributes = root[0];
  const matches = [...attributes.matchAll(/\s+refreshOnLoad="([^"]*)"/gi)];
  if (matches.length !== 1) throw new Error(`${label} must carry exactly one explicit refreshOnLoad attribute.`);
  const value = String(matches[0][1]).toLowerCase();
  if (!new Set(["1", "true", "0", "false"]).has(value)) throw new Error(`${label} has an invalid refreshOnLoad lexical value.`);
  const normalizedRoot = attributes.replace(matches[0][0], ' refreshOnLoad="__pivot_refresh__"');
  return {
    enabled: value === "1" || value === "true",
    normalized: `${xml.slice(0, root.index)}${normalizedRoot}${xml.slice(root.index + attributes.length)}`,
  };
}

async function readPivotPackage(bytes, label) {
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const cachePaths = paths.filter((name) => /^pivotCache\/pivotCacheDefinition\d*\.xml$/i.test(name));
  if (cachePaths.length !== 1) throw new Error(`${label} must contain exactly one Pivot cache-definition part; found ${cachePaths.length}.`);
  const cacheDefinitionPath = cachePaths[0];
  const cacheXml = await zip.file(cacheDefinitionPath).async("text");
  return { zip, paths, cacheDefinitionPath, cache: readRefreshOnLoadRoot(cacheXml, `${label} ${cacheDefinitionPath}`) };
}

async function assertOnlyCacheRefreshChanged(source, output) {
  if (!sameJson(source.paths, output.paths) || source.cacheDefinitionPath !== output.cacheDefinitionPath) {
    throw new Error("The exported package changed Pivot part topology or its cache-definition locator.");
  }
  if (!source.cache.enabled || output.cache.enabled || source.cache.normalized !== output.cache.normalized) {
    throw new Error("The cache definition changed more than refreshOnLoad=true to false.");
  }
  for (const partPath of source.paths) {
    if (partPath === source.cacheDefinitionPath) continue;
    const [before, after] = await Promise.all([
      source.zip.file(partPath).async("uint8array"),
      output.zip.file(partPath).async("uint8array"),
    ]);
    if (!Buffer.from(before).equals(Buffer.from(after))) {
      throw new Error(`Only ${source.cacheDefinitionPath} may change during PivotTable refresh-on-load hardening; ${partPath} changed.`);
    }
  }
  return {
    pathsPreserved: true,
    onlyCacheDefinitionChanged: true,
    cacheDefinitionPath: source.cacheDefinitionPath,
    refreshOnLoadResidualPreserved: true,
  };
}

async function renderAllSheets(workbook) {
  const sheets = [];
  for (const sheet of workbook.worksheets.items) {
    const preview = await workbook.render({ sheetName: sheet.name, autoCrop: "all", format: "svg" });
    const svg = await preview.text();
    if (!/<svg\b/i.test(svg)) throw new Error(`Model render for sheet ${sheet.name} did not produce SVG.`);
    sheets.push({ sheet: sheet.name, bytes: preview.bytes.length, renderer: "model-svg" });
  }
  return sheets;
}

/**
 * Disable only one eligible imported PivotTable cache's refresh-on-load flag.
 * This does not refresh data or make a PivotTable generally editable.
 */
export async function hardenXlsxPivotRefreshOnLoad({
  inputPath,
  outputPath,
  auditPath,
  worksheetName,
  pivotName,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const sheetName = requiredText(worksheetName, "worksheetName");
  const name = requiredText(pivotName, "pivotName");
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original workbook remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from source and XLSX output paths.");
  await Promise.all([
    assertNewFile(finalPath, "XLSX output"),
    assertNewFile(finalAuditPath, "Audit output"),
  ]);

  const source = await fs.readFile(sourcePath);
  const sourcePackage = await readPivotPackage(source, "Source workbook");
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const sourcePivot = locatePivot(workbook, sheetName, name);
  const sourceProjection = pivotProjection(sourcePivot);
  if (sourceProjection.refreshPolicy.refreshOnLoad !== true || !sourceProjection.sourceCapabilities.sourceBound || !sourceProjection.sourceCapabilities.refreshOnLoadHardenable) {
    throw new Error(`PivotTable ${JSON.stringify(name)} does not prove the explicit, uniquely owned refreshOnLoad=true capability required for this workflow.`);
  }

  sourcePivot.disableRefreshOnLoad();
  const inMemoryProjection = pivotProjection(sourcePivot);
  if (!sameJson(inMemoryProjection, expectedRefreshOnLoadHardening(sourceProjection, sourceProjection.sourceCapabilities))) {
    throw new Error("The in-memory PivotTable hardening changed more than the validated refreshOnLoad switch.");
  }

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  try {
    await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
    const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const outputPackage = await readPivotPackage(output, "Output workbook");
    const packageValidation = await assertOnlyCacheRefreshChanged(sourcePackage, outputPackage);
    const reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const outputPivot = locatePivot(reimported, sheetName, name);
    const outputProjection = pivotProjection(outputPivot);
    const expectedOutput = expectedRefreshOnLoadHardening(sourceProjection, { sourceBound: true, refreshOnLoadHardenable: false });
    if (!sameJson(outputProjection, expectedOutput)) throw new Error("Second import did not preserve the validated PivotTable semantics and capability withdrawal.");
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed: ${verification.ndjson}`);
    const renders = await renderAllSheets(reimported);
    const sourceAfter = await fs.readFile(sourcePath);
    if (!source.equals(sourceAfter)) throw new Error("The source workbook changed during the transaction; refusing to publish output.");
    const audit = {
      schema: "open-office-artifact-tool.xlsx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "pivot-refresh-on-load-hardening",
        pivot: {
          worksheetName: sheetName,
          name,
          targetRange: sourceProjection.targetRange,
          previousRefreshOnLoad: true,
          refreshOnLoad: false,
          cacheDefinitionPath: sourcePackage.cacheDefinitionPath,
        },
      },
      warnings: [
        "This disables only the Pivot cache refresh-on-load request. Manual, macro, external-data, and other host-triggered refreshes remain outside this operation.",
      ],
      validation: {
        package: packageValidation,
        reimport: {
          ok: true,
          pivotProjectionPreserved: true,
          refreshOnLoadDisabled: outputProjection.refreshPolicy.refreshOnLoad === false,
          refreshOnLoadHardenableWithdrawn: outputProjection.sourceCapabilities.refreshOnLoadHardenable === false,
        },
        verify: { ok: verification.ok },
        modelRender: { ok: true, sheets: renders },
      },
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, "utf8");
    await publishNoOverwrite(temporaryPath, finalPath, "XLSX output");
    try {
      await publishNoOverwrite(temporaryAuditPath, finalAuditPath, "Audit output");
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } finally {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, worksheetName, pivotName] = argv;
  return { inputPath, outputPath, auditPath, worksheetName, pivotName };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await hardenXlsxPivotRefreshOnLoad(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    pivot: result.audit.operation.pivot,
  }));
}
