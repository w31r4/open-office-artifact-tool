import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const require = createRequire(import.meta.url);
const REVENUE_RANGE = "B5:B7";
const GROSS_PROFIT_RANGE = "C4:C7";
const REVENUE_FORMULAS = ["=B4*(1+$B$9)", "=B5*(1+$B$9)", "=B6*(1+$B$9)"];

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function requiredFiniteNumber(value, label) {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new TypeError(`${label} must be a finite number.`);
  return value;
}

function closeEnough(actual, expected, tolerance = 1e-9) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
}

function normalizeFormula(value) {
  return String(value || "").replace(/^=/, "");
}

function scalar(sheet, address) {
  return sheet.getRange(address).values?.[0]?.[0];
}

function snapshot(value) {
  return JSON.parse(JSON.stringify(value));
}

function expectedRevenue(growth) {
  let prior = 100;
  return REVENUE_FORMULAS.map(() => {
    prior *= 1 + growth;
    return prior;
  });
}

function assertFormulaSnapshot(sheet, label) {
  const revenue = sheet.getRange(REVENUE_RANGE).formulas.flat();
  if (!revenue.every((formula, index) => normalizeFormula(formula) === normalizeFormula(REVENUE_FORMULAS[index]))) {
    throw new Error(`${label} changed one or more revenue formulas; refusing a lossy formula rewrite.`);
  }
  const grossProfit = sheet.getRange(GROSS_PROFIT_RANGE).formulas.flat();
  if (!grossProfit.every((formula) => /^=B\d+\*\$B\$10$/i.test(String(formula)))) {
    throw new Error(`${label} changed one or more gross-profit formulas; refusing a lossy formula rewrite.`);
  }
}

function assertRecalculatedRevenue(sheet, growth, label) {
  const expected = expectedRevenue(growth);
  const actual = sheet.getRange(REVENUE_RANGE).values.flat();
  if (!actual.every((value, index) => closeEnough(value, expected[index], 1e-7))) {
    throw new Error(`${label} did not recalculate the revenue chain from the requested growth assumption.`);
  }
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

function resolveWorkbook(workbook, targetSheetName, canarySheetName, growthAddress, marginAddress, expectedGrowth, expectedMargin) {
  const sheetNames = workbook.worksheets.items.map((sheet) => sheet.name);
  if (JSON.stringify(sheetNames) !== JSON.stringify([targetSheetName, canarySheetName])) {
    throw new Error(`Expected the fixed workbook sheet order ${JSON.stringify([targetSheetName, canarySheetName])}; found ${JSON.stringify(sheetNames)}.`);
  }
  const target = workbook.worksheets.getItem(targetSheetName);
  const canary = workbook.worksheets.getItem(canarySheetName);
  if (!closeEnough(scalar(target, growthAddress), expectedGrowth)) {
    throw new Error(`The selected growth assumption does not match the expected source value ${expectedGrowth}; refusing an ambiguous edit.`);
  }
  if (!closeEnough(scalar(target, marginAddress), expectedMargin)) {
    throw new Error(`The protected gross-margin assumption does not match ${expectedMargin}; refusing an unsafe edit.`);
  }
  assertFormulaSnapshot(target, "The selected workbook");
  return { target, canary, sheetNames };
}

export async function updateXlsxGrowthAssumption({
  inputPath,
  outputPath,
  auditPath,
  targetSheetName = "Forecast",
  canarySheetName = "Approved Baseline",
  growthAddress = "B9",
  marginAddress = "B10",
  expectedGrowth = 0.08,
  replacementGrowth = 0.1,
  expectedMargin = 0.6,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must be distinct from inputPath so the original workbook remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must be distinct from source and XLSX output paths.");
  const sheetName = requiredText(targetSheetName, "targetSheetName");
  const baselineName = requiredText(canarySheetName, "canarySheetName");
  const growthCell = requiredText(growthAddress, "growthAddress").toUpperCase();
  const marginCell = requiredText(marginAddress, "marginAddress").toUpperCase();
  const originalGrowth = requiredFiniteNumber(expectedGrowth, "expectedGrowth");
  const nextGrowth = requiredFiniteNumber(replacementGrowth, "replacementGrowth");
  const margin = requiredFiniteNumber(expectedMargin, "expectedMargin");
  if (nextGrowth <= -1) throw new RangeError("replacementGrowth must be greater than -1.");

  const source = await fs.readFile(sourcePath);
  const workbook = await SpreadsheetFile.importXlsx(new FileBlob(source, { type: XLSX_MIME, name: path.basename(sourcePath) }));
  const { target, canary, sheetNames } = resolveWorkbook(workbook, sheetName, baselineName, growthCell, marginCell, originalGrowth, margin);
  const identity = {
    targetId: target.id,
    canaryId: canary.id,
    sheetNames,
    canaryValues: snapshot(canary.getRange("A1:C5").values),
    canaryFormulas: snapshot(canary.getRange("A1:C5").formulas),
    protectedMargin: scalar(target, marginCell),
  };

  target.getRange(growthCell).values = [[nextGrowth]];
  workbook.recalculate();
  assertFormulaSnapshot(target, "The edited workbook");
  assertRecalculatedRevenue(target, nextGrowth, "The edited workbook");
  if (!closeEnough(scalar(target, marginCell), identity.protectedMargin)) {
    throw new Error("The edit changed the protected gross-margin assumption before export.");
  }

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const reimported = await SpreadsheetFile.importXlsx(new FileBlob(output, { type: XLSX_MIME, name: path.basename(finalPath) }));
    const roundTrip = resolveWorkbook(reimported, sheetName, baselineName, growthCell, marginCell, nextGrowth, margin);
    assertFormulaSnapshot(roundTrip.target, "The exported workbook");
    assertRecalculatedRevenue(roundTrip.target, nextGrowth, "The exported workbook");
    if (roundTrip.target.id !== identity.targetId || roundTrip.canary.id !== identity.canaryId) {
      throw new Error("XLSX export changed target or canary worksheet identity.");
    }
    if (JSON.stringify(roundTrip.canary.getRange("A1:C5").values) !== JSON.stringify(identity.canaryValues)
      || JSON.stringify(roundTrip.canary.getRange("A1:C5").formulas) !== JSON.stringify(identity.canaryFormulas)) {
      throw new Error("XLSX export changed the protected Approved Baseline sheet.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Workbook verification failed: ${verification.ndjson}`);
    const renders = await renderAllSheets(reimported);
    const audit = {
      schema: "open-office-artifact-tool.xlsx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "growth-assumption-update",
        target: { sheet: sheetName, cell: growthCell, previous: originalGrowth, replacement: nextGrowth },
        protected: { sheet: sheetName, cell: marginCell, value: margin },
        canarySheet: baselineName,
      },
      warnings: [],
      validation: {
        reimport: {
          ok: true,
          sheetNamesPreserved: JSON.stringify(roundTrip.sheetNames) === JSON.stringify(identity.sheetNames),
          targetSheetIdPreserved: roundTrip.target.id === identity.targetId,
          canarySheetIdPreserved: roundTrip.canary.id === identity.canaryId,
          formulasPreserved: true,
          growthExact: closeEnough(scalar(roundTrip.target, growthCell), nextGrowth),
          grossMarginPreserved: closeEnough(scalar(roundTrip.target, marginCell), margin),
          revisedRevenue: roundTrip.target.getRange(REVENUE_RANGE).values.flat(),
          canaryValuesPreserved: true,
        },
        verify: { ok: verification.ok },
        modelRender: { ok: true, sheets: renders },
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
    targetSheetName = "Forecast",
    canarySheetName = "Approved Baseline",
    growthAddress = "B9",
    marginAddress = "B10",
    expectedGrowth = "0.08",
    replacementGrowth = "0.1",
    expectedMargin = "0.6",
  ] = argv;
  return {
    inputPath,
    outputPath,
    auditPath,
    targetSheetName,
    canarySheetName,
    growthAddress,
    marginAddress,
    expectedGrowth: Number(expectedGrowth),
    replacementGrowth: Number(replacementGrowth),
    expectedMargin: Number(expectedMargin),
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await updateXlsxGrowthAssumption(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    revisedRevenue: result.audit.validation.reimport.revisedRevenue,
  }));
}
