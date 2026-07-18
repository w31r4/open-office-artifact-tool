import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";
import { XLSX_GROWTH_UPDATE_FIXTURE, generateOfficeInput } from "../scripts/agent-eval-office-fixtures.mjs";
import { replyAndResolveThreadedComment } from "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-threaded-comment-reply-workflow.mjs";
import { updateXlsxGrowthAssumption } from "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-growth-assumption-edit-workflow.mjs";
import { createWorkbookFromFixture, nativeSpreadsheetRenderStatus, runSpreadsheetFixture, verifyWorkbookFile } from "./skill-harness/spreadsheets/scripts/workflow.mjs";

const repoRoot = path.resolve(new URL("..", import.meta.url).pathname);
const fixtureDir = path.join(repoRoot, "test", "skill-harness", "spreadsheets", "fixtures");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-spreadsheet-skill-"));

async function runFixture(name) {
  const result = await runSpreadsheetFixture(path.join(fixtureDir, `${name}.json`), {
    outputDir: path.join(outputDir, name),
    renderFormat: "svg",
    allSheets: true,
    nativeRender: "off",
  });
  assert.equal(result.fixture.name, name);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.renderFormat, "svg");
  assert.ok(["disabled", "skipped"].includes(result.qa.summary.nativeRender.status));
  assert.ok((await fs.stat(result.workbookPath)).size > 0);
  for (const filePath of Object.values(result.qa.summary.files)) assert.ok((await fs.stat(filePath)).size > 0, `Expected non-empty QA artifact ${filePath}`);
  return result;
}

try {
  assert.throws(
    () => createWorkbookFromFixture({
      sheets: [{ name: "Dynamic", ranges: [{ range: "A1:A2", formulas: [["=SEQUENCE(2)"], [null]], formulaMetadata: { kind: "dynamicArray", reference: "A1:A2" } }] }],
    }),
    /must be shared or array.*dynamic arrays are import-only and read-only/i,
  );
  const formulaResult = await runFixture("formula-summary");
  const formulaBlob = await FileBlob.load(formulaResult.workbookPath);
  const formulaWorkbook = await SpreadsheetFile.importXlsx(formulaBlob);
  const summary = formulaWorkbook.worksheets.getItem("Summary");
  assert.ok(summary);
  assert.equal(summary.showGridLines, false);
  assert.deepEqual(summary.freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
  assert.deepEqual(summary.getRange("E2:E4").values, [[0.4], [0.4166666666666667], [0.4]]);
  assert.deepEqual(summary.getRange("E2:E4").formulas, [["=(B2-C2)/B2"], ["=(B3-C3)/B3"], ["=(B4-C4)/B4"]]);
  assert.deepEqual(summary.mergedRanges, ["A6:E6"]);
  assert.equal(summary.tables.items[0].name, "SummaryTable");
  assert.equal(summary.images.items.length, 1);
  assert.equal(summary.images.items[0].alt, "Green status marker");
  assert.equal(summary.charts.items.length, 1);
  assert.equal(summary.charts.items[0].type, "line");
  assert.equal(summary.dataValidations.items.length, 1);
  assert.deepEqual(summary.dataValidations.items[0].rule, { type: "list", values: ["Planned", "In progress", "Done"] });
  assert.equal(summary.conditionalFormattings.items.length, 2);
  assert.deepEqual(summary.conditionalFormattings.items.map((item) => item.ruleType), ["cellIs", "colorScale"]);
  assert.equal(formulaWorkbook.comments.threads.length, 1);
  const [review] = formulaWorkbook.comments.threads;
  assert.equal(review.comments.length, 2);
  assert.equal(review.comments[0].text, "Check the modeled margin.");
  assert.equal(review.comments[1].text, "Confirmed against the modeled source data.");
  assert.equal(review.resolved, true);

  const formulaZip = await JSZip.loadAsync(await fs.readFile(formulaResult.workbookPath));
  const summaryXml = await formulaZip.file("xl/worksheets/sheet1.xml").async("text");
  assert.match(summaryXml, /<x:dataValidations count="1">/);
  assert.equal((summaryXml.match(/<x:conditionalFormatting\b/g) || []).length, 2);
  const threadedPath = Object.keys(formulaZip.files).find((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
  const personPath = Object.keys(formulaZip.files).find((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name));
  assert.ok(threadedPath);
  assert.ok(personPath);
  const threadedXml = await formulaZip.file(threadedPath).async("text");
  assert.match(threadedXml, /parentId="\{11111111-1111-4111-8111-111111111111\}"/);
  assert.match(threadedXml, /Check the modeled margin\./);
  assert.match(threadedXml, /Confirmed against the modeled source data\./);
  assert.match(await formulaZip.file(personPath).async("text"), /displayName="Reviewer"/);
  assert.match(await formulaZip.file(personPath).async("text"), /displayName="Lead reviewer"/);

  const threadedEvalInput = path.join(outputDir, "threaded-review-input.xlsx");
  const threadedEvalOutput = path.join(outputDir, "threaded-review-output.xlsx");
  const threadedEvalAudit = path.join(outputDir, "threaded-review-audit.json");
  await generateOfficeInput("xlsx-threaded-review", threadedEvalInput);
  const threadedResult = await replyAndResolveThreadedComment({
    inputPath: threadedEvalInput,
    outputPath: threadedEvalOutput,
    auditPath: threadedEvalAudit,
    sheetName: "Forecast",
    cell: "F19",
    reply: "Approved after sensitivity review",
    author: "Board secretary",
  });
  assert.equal(threadedResult.audit.provider.actual, "open-chestnut");
  assert.equal(threadedResult.audit.savePolicy.strategy, "rewrite");
  const threadedRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(threadedEvalOutput));
  const threadedThread = threadedRoundTrip.comments.threads.find((thread) => thread.target.sheetName === "Forecast" && thread.target.address === "F19");
  assert.ok(threadedThread);
  assert.equal(threadedThread.comments.length, 3);
  assert.equal(threadedThread.comments[2].text, "Approved after sensitivity review");
  assert.equal(threadedThread.comments[2].parentId, threadedThread.comments[0].id);
  assert.equal(threadedThread.resolved, true);
  assert.equal(threadedThread.comments.every((comment) => comment.done), true);
  assert.equal(JSON.parse(await fs.readFile(threadedEvalAudit, "utf8")).validation.reimport.ok, true);

  const growthEvalInput = path.join(outputDir, XLSX_GROWTH_UPDATE_FIXTURE.workbookName);
  const growthEvalOutput = path.join(outputDir, "operating-plan-updated.xlsx");
  const growthEvalAudit = path.join(outputDir, "operating-plan-audit.json");
  await generateOfficeInput("xlsx-growth-update", growthEvalInput);
  const growthSourceBefore = await fs.readFile(growthEvalInput);
  const growthResult = await updateXlsxGrowthAssumption({
    inputPath: growthEvalInput,
    outputPath: growthEvalOutput,
    auditPath: growthEvalAudit,
  });
  assert.equal(growthResult.audit.provider.actual, "open-chestnut");
  assert.equal(growthResult.audit.savePolicy.strategy, "rewrite");
  assert.equal(growthResult.audit.validation.reimport.ok, true);
  assert.deepEqual(await fs.readFile(growthEvalInput), growthSourceBefore);
  const growthRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(growthEvalOutput));
  const growthForecast = growthRoundTrip.worksheets.getItem(XLSX_GROWTH_UPDATE_FIXTURE.targetSheetName);
  const growthBaseline = growthRoundTrip.worksheets.getItem(XLSX_GROWTH_UPDATE_FIXTURE.canarySheetName);
  assert.ok(growthForecast);
  assert.ok(growthBaseline);
  assert.equal(growthForecast.getRange(XLSX_GROWTH_UPDATE_FIXTURE.growthAddress).values[0][0], XLSX_GROWTH_UPDATE_FIXTURE.replacementGrowth);
  assert.equal(growthForecast.getRange(XLSX_GROWTH_UPDATE_FIXTURE.marginAddress).values[0][0], XLSX_GROWTH_UPDATE_FIXTURE.grossMargin);
  assert.deepEqual(growthForecast.getRange("B5:B7").formulas.flat(), XLSX_GROWTH_UPDATE_FIXTURE.revenueFormulas);
  assert.ok(growthForecast.getRange("B5:B7").values.flat().every((value, index) => Math.abs(value - XLSX_GROWTH_UPDATE_FIXTURE.revisedRevenue[index]) < 1e-7));
  assert.equal(growthBaseline.getRange("A1").values[0][0], XLSX_GROWTH_UPDATE_FIXTURE.canaryText);
  assert.deepEqual(JSON.parse(await fs.readFile(growthEvalAudit, "utf8")).validation.reimport.revisedRevenue, growthResult.audit.validation.reimport.revisedRevenue);

  const financialFixture = await runFixture("financial-returns");
  const financialFixtureWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(financialFixture.workbookPath));
  const fixtureReturns = financialFixtureWorkbook.worksheets.getItem("Returns");
  const fixtureChecks = financialFixtureWorkbook.worksheets.getItem("Checks");
  assert.ok(Math.abs(fixtureReturns.getRange("B4").values[0][0] - 0.1709368633949911) < 1e-9);
  assert.ok(Math.abs(fixtureReturns.getRange("B7").values[0][0] - 0.14400168352963139) < 1e-9);
  assert.ok(Math.abs(fixtureReturns.getRange("B8").values[0][0] + 8884.878867834168) < 1e-9);
  assert.deepEqual(fixtureChecks.getRange("E2:E8").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  assert.equal(fixtureChecks.conditionalFormattings.items.length, 2);

  const loanFixture = await runFixture("loan-amortization");
  const loanFixtureWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(loanFixture.workbookPath));
  const fixtureAmortization = loanFixtureWorkbook.worksheets.getItem("Amortization");
  const fixtureLoanChecks = loanFixtureWorkbook.worksheets.getItem("Checks");
  assert.deepEqual(fixtureAmortization.getRange("C2:E2").values, [[-8884.878867834168, -1000, -7884.878867834168]]);
  assert.equal(fixtureAmortization.getRange("D2").formulas[0][0], "=IPMT('Inputs'!$B$7,A2,'Inputs'!$B$8,'Inputs'!$B$2,0,'Inputs'!$B$6)");
  assert.ok(Math.abs(fixtureAmortization.getRange("F13").values[0][0]) < 1e-7);
  assert.equal(fixtureLoanChecks.getRange("B7").formulas[0][0], "=RATE('Inputs'!$B$8,'Amortization'!$C$2,'Inputs'!$B$2,0,'Inputs'!$B$6,'Inputs'!$B$7)");
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B7").values[0][0] - 0.01) < 1e-10);
  assert.equal(fixtureLoanChecks.getRange("B8").formulas[0][0], "=PV('Inputs'!$B$7,'Inputs'!$B$8,'Amortization'!$C$2,0,'Inputs'!$B$6)");
  assert.equal(fixtureLoanChecks.getRange("B9").formulas[0][0], "=FV('Inputs'!$B$7,'Inputs'!$B$8,'Amortization'!$C$2,'Inputs'!$B$2,'Inputs'!$B$6)");
  assert.equal(fixtureLoanChecks.getRange("B10").formulas[0][0], "=NPER('Inputs'!$B$7,'Amortization'!$C$2,'Inputs'!$B$2,0,'Inputs'!$B$6)");
  assert.equal(fixtureLoanChecks.getRange("B11").formulas[0][0], "=CUMIPMT('Inputs'!$B$7,'Inputs'!$B$8,'Inputs'!$B$2,1,'Inputs'!$B$8,'Inputs'!$B$6)");
  assert.equal(fixtureLoanChecks.getRange("B12").formulas[0][0], "=CUMPRINC('Inputs'!$B$7,'Inputs'!$B$8,'Inputs'!$B$2,1,'Inputs'!$B$8,'Inputs'!$B$6)");
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B8").values[0][0] - 100000) < 1e-7);
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B9").values[0][0]) < 1e-7);
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B10").values[0][0] - 12) < 1e-10);
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B11").values[0][0] + 6618.54641401005) < 1e-8);
  assert.ok(Math.abs(fixtureLoanChecks.getRange("B12").values[0][0] + 100000) < 1e-8);
  assert.deepEqual(fixtureLoanChecks.getRange("E2:E13").values, Array.from({ length: 12 }, () => ["OK"]));
  assert.equal(fixtureLoanChecks.conditionalFormattings.items.length, 2);

  const assetFixture = await runFixture("asset-depreciation");
  const assetFixtureWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(assetFixture.workbookPath));
  const fixtureDepreciation = assetFixtureWorkbook.worksheets.getItem("Depreciation");
  const fixtureAssetChecks = assetFixtureWorkbook.worksheets.getItem("Checks");
  assert.deepEqual(fixtureDepreciation.getRange("C2:E2").values, [[18000, 36900, 40000]]);
  assert.equal(fixtureDepreciation.getRange("D2").formulas[0][0], "=DB('Inputs'!$B$2,'Inputs'!$B$3,'Inputs'!$B$4,A2,'Inputs'!$B$5)");
  assert.equal(fixtureDepreciation.getRange("E2").formulas[0][0], "=DDB('Inputs'!$B$2,'Inputs'!$B$3,'Inputs'!$B$4,A2,'Inputs'!$B$6)");
  assert.equal(fixtureDepreciation.getRange("F6").values[0][0], 10000);
  assert.deepEqual(fixtureAssetChecks.getRange("E2:E7").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  assert.equal(fixtureAssetChecks.conditionalFormattings.items.length, 2);

  const basicResult = await runFixture("open-chestnut-basic");
  const basicWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(basicResult.workbookPath));
  const dashboard = basicWorkbook.worksheets.getItem("Dashboard");
  assert.deepEqual(dashboard.getRange("C2:C4").values, [[20], [40], [60]]);
  assert.deepEqual(dashboard.mergedRanges, ["A6:D6"]);
  assert.deepEqual(dashboard.charts.items.map((chart) => chart.type), ["bar", "line", "pie", "area", "doughnut"]);
  assert.equal(dashboard.charts.items[1].title, "Edited line trend");
  assert.equal(dashboard.images.items[0].alt, "Edited OpenChestnut marker");
  assert.equal(dashboard.dataValidations.items.length, 1);
  assert.deepEqual(dashboard.conditionalFormattings.items.map((item) => item.ruleType), ["containsText", "colorScale"]);
  assert.equal(basicWorkbook.definedNames.getItem("ActualValues").refersTo, "Dashboard!$B$2:$B$4");
  const basicZip = await JSZip.loadAsync(await fs.readFile(basicResult.workbookPath));
  assert.equal(Object.keys(basicZip.files).filter((name) => /\/charts\/chart\d+\.xml$/i.test(name)).length, 5);
  const chartXml = await Promise.all(Object.keys(basicZip.files).filter((name) => /\/charts\/chart\d+\.xml$/i.test(name)).map((name) => basicZip.file(name).async("text")));
  assert.equal(chartXml.filter((xml) => /<c:areaChart>/.test(xml)).length, 1);
  assert.equal(chartXml.filter((xml) => /<c:doughnutChart>/.test(xml)).length, 1);

  const sparklineResult = await runFixture("open-chestnut-sparklines");
  const sparklineWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(sparklineResult.workbookPath));
  const trends = sparklineWorkbook.worksheets.getItem("Trends");
  assert.deepEqual(trends.sparklineGroups.items.map((group) => group.type), ["line", "column", "stacked"]);
  assert.equal(trends.sparklineGroups.items[0].sparklineCount, 3);
  assert.equal(trends.sparklineGroups.items[0].seriesColor, "#F97316");
  assert.equal(trends.sparklineGroups.items[0].axis.manualMax, 120);
  const sparklineZip = await JSZip.loadAsync(await fs.readFile(sparklineResult.workbookPath));
  const sparklineXml = await sparklineZip.file("xl/worksheets/sheet1.xml").async("text");
  assert.equal((sparklineXml.match(/<x14:sparklineGroup\b/g) || []).length, 3);
  assert.equal((sparklineXml.match(/<x14:sparkline\b/g) || []).length, 9);

  const { createDataTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-data-table-workflow.mjs"
  );
  const dataTablePath = path.join(outputDir, "openchestnut-data-table-workflow.xlsx");
  const dataTableResult = await createDataTableWorkbook(dataTablePath);
  assert.equal(dataTableResult.verification.ok, true);
  const dataTableWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(dataTablePath));
  assert.equal(dataTableWorkbook.worksheets.getItem("Scenario Analysis").dataTables.__getDefinitions().length, 2);
  const dataTableZip = await JSZip.loadAsync(await fs.readFile(dataTablePath));
  const dataTableXml = await dataTableZip.file("xl/worksheets/sheet1.xml").async("text");
  assert.equal((dataTableXml.match(/<x:f\b[^>]*t="dataTable"/g) || []).length, 2);

  const { createPivotTableWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-pivot-table-workflow.mjs"
  );
  const pivotTablePath = path.join(outputDir, "openchestnut-pivot-table-workflow.xlsx");
  const pivotTableResult = await createPivotTableWorkbook(pivotTablePath);
  assert.equal(pivotTableResult.verification.ok, true);
  assert.match(pivotTableResult.inspection.ndjson, /"kind":"pivotTable"/);
  const pivotTableWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(pivotTablePath));
  assert.ok(Math.abs(pivotTableWorkbook.worksheets.getItem("Pivot Summary").getRange("A1:A5").format.columnWidthPx - 112) <= 1);
  assert.ok(Math.abs(pivotTableWorkbook.worksheets.getItem("Pivot Summary").getRange("B1:C5").format.columnWidthPx - 88) <= 1);
  assert.deepEqual(pivotTableWorkbook.worksheets.getItem("Pivot Summary").pivotTables.items[0].computedValues(), [
    ["Region", "Direct", "Partner", "Grand Total"],
    ["East", 120, 80, 200],
    ["West", 150, 90, 240],
    ["North", 110, 70, 180],
    ["Grand Total", 380, 240, 620],
  ]);
  const pivotTableZip = await JSZip.loadAsync(await fs.readFile(pivotTablePath));
  assert.equal(Object.keys(pivotTableZip.files).filter((name) => /pivotTables\/pivotTable.*\.xml$/i.test(name)).length, 1);
  assert.equal(Object.keys(pivotTableZip.files).filter((name) => /pivotCache\/pivotCacheRecords.*\.xml$/i.test(name)).length, 1);
  const pivotNativeStatus = nativeSpreadsheetRenderStatus();
  const pivotQa = await verifyWorkbookFile(pivotTablePath, {
    outputDir: path.join(outputDir, "pivot-native-qa"),
    sheetName: "Pivot Summary",
    renderFormat: "svg",
    allSheets: true,
    nativeRender: pivotNativeStatus.available ? "required" : "off",
  });
  if (pivotNativeStatus.available) {
    assert.equal(pivotQa.summary.nativeRender.status, "passed");
    assert.equal(pivotQa.summary.nativeRender.pageCount, 2, "the Data and Pivot Summary sheets must each fit on one native-rendered page");
  }

  const { createFinancialReturnsWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-financial-returns-workflow.mjs"
  );
  const financialReturnsPath = path.join(outputDir, "openchestnut-financial-returns-workflow.xlsx");
  const financialReturnsResult = await createFinancialReturnsWorkbook(financialReturnsPath);
  assert.equal(financialReturnsResult.verification.ok, true);
  assert.match(financialReturnsResult.inspection.ndjson, /XIRR/);
  assert.match(financialReturnsResult.inspection.ndjson, /MIRR/);
  const financialReturnsWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(financialReturnsPath));
  financialReturnsWorkbook.recalculate();
  const financialReturns = financialReturnsWorkbook.worksheets.getItem("Returns");
  const financialChecks = financialReturnsWorkbook.worksheets.getItem("Checks");
  assert.equal(financialReturns.getRange("B8").formulas[0][0], "=XIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$14:$B$18,'Inputs'!$B$7)");
  assert.equal(financialReturns.getRange("B9").formulas[0][0], "=MIRR('Inputs'!$C$14:$C$18,'Inputs'!$B$5,'Inputs'!$B$6)");
  assert.equal(financialReturns.getRange("B7").format.numberFormat, "$#,##0;[Red]($#,##0);-");
  assert.ok(Math.abs(financialReturns.getRange("B8").values[0][0] - 0.17083686863616765) < 1e-9);
  assert.ok(Math.abs(financialReturns.getRange("B9").values[0][0] - 0.14400168352963139) < 1e-9);
  assert.deepEqual(financialChecks.getRange("E4:E10").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  const financialReturnsQa = await verifyWorkbookFile(financialReturnsPath, {
    outputDir: path.join(outputDir, "openchestnut-financial-returns-native-qa"),
    sheetName: "Returns",
    renderFormat: "svg",
    nativeRender: "auto",
    allSheets: true,
  });
  if (nativeSpreadsheetRenderStatus().available) {
    assert.equal(financialReturnsQa.summary.nativeRender.status, "passed");
    assert.equal(financialReturnsQa.summary.nativeRender.ok, true);
  }

  const { createLoanAmortizationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-loan-amortization-workflow.mjs"
  );
  const loanAmortizationPath = path.join(outputDir, "openchestnut-loan-amortization-workflow.xlsx");
  const loanAmortizationResult = await createLoanAmortizationWorkbook(loanAmortizationPath);
  assert.equal(loanAmortizationResult.verification.ok, true);
  assert.match(loanAmortizationResult.inspection.ndjson, /IPMT/);
  assert.match(loanAmortizationResult.inspection.ndjson, /PPMT/);
  assert.match(loanAmortizationResult.checksInspection.ndjson, /PV/);
  assert.match(loanAmortizationResult.checksInspection.ndjson, /FV/);
  assert.match(loanAmortizationResult.checksInspection.ndjson, /NPER/);
  const loanAmortizationWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(loanAmortizationPath));
  loanAmortizationWorkbook.recalculate();
  const loanAmortization = loanAmortizationWorkbook.worksheets.getItem("Amortization");
  const loanAmortizationChecks = loanAmortizationWorkbook.worksheets.getItem("Checks");
  assert.equal(loanAmortization.getRange("D5").formulas[0][0], "=IPMT('Inputs'!$B$10,A5,'Inputs'!$B$11,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortization.getRange("E5").format.numberFormat, "$#,##0;[Red]($#,##0);-");
  assert.ok(Math.abs(loanAmortization.getRange("F16").values[0][0]) < 1e-7);
  assert.equal(loanAmortizationChecks.getRange("B9").formulas[0][0], "=RATE('Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9,'Inputs'!$B$10)");
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B9").values[0][0] - 0.01) < 1e-10);
  assert.equal(loanAmortizationChecks.getRange("B10").formulas[0][0], "=PV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationChecks.getRange("B11").formulas[0][0], "=FV('Inputs'!$B$10,'Inputs'!$B$11,'Amortization'!$C$5,'Inputs'!$B$5,'Inputs'!$B$9)");
  assert.equal(loanAmortizationChecks.getRange("B12").formulas[0][0], "=NPER('Inputs'!$B$10,'Amortization'!$C$5,'Inputs'!$B$5,0,'Inputs'!$B$9)");
  assert.equal(loanAmortizationChecks.getRange("B13").formulas[0][0], "=CUMIPMT('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.equal(loanAmortizationChecks.getRange("B14").formulas[0][0], "=CUMPRINC('Inputs'!$B$10,'Inputs'!$B$11,'Inputs'!$B$5,1,'Inputs'!$B$11,'Inputs'!$B$9)");
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B10").values[0][0] - 100000) < 1e-7);
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B11").values[0][0]) < 1e-7);
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B12").values[0][0] - 12) < 1e-10);
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B13").values[0][0] + 6618.54641401005) < 1e-8);
  assert.ok(Math.abs(loanAmortizationChecks.getRange("B14").values[0][0] + 100000) < 1e-8);
  assert.deepEqual(loanAmortizationChecks.getRange("E4:E15").values, Array.from({ length: 12 }, () => ["OK"]));
  const loanAmortizationQa = await verifyWorkbookFile(loanAmortizationPath, {
    outputDir: path.join(outputDir, "openchestnut-loan-amortization-native-qa"),
    sheetName: "Amortization",
    renderFormat: "svg",
    nativeRender: "auto",
    allSheets: true,
  });
  if (nativeSpreadsheetRenderStatus().available) {
    assert.equal(loanAmortizationQa.summary.nativeRender.status, "passed");
    assert.equal(loanAmortizationQa.summary.nativeRender.ok, true);
  }

  const { createAssetDepreciationWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-asset-depreciation-workflow.mjs"
  );
  const assetDepreciationPath = path.join(outputDir, "openchestnut-asset-depreciation-workflow.xlsx");
  const assetDepreciationResult = await createAssetDepreciationWorkbook(assetDepreciationPath);
  assert.equal(assetDepreciationResult.verification.ok, true);
  assert.match(assetDepreciationResult.inspection.ndjson, /SLN/);
  assert.match(assetDepreciationResult.inspection.ndjson, /DDB/);
  const assetDepreciationWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(assetDepreciationPath));
  assetDepreciationWorkbook.recalculate();
  const assetDepreciation = assetDepreciationWorkbook.worksheets.getItem("Depreciation");
  const assetDepreciationChecks = assetDepreciationWorkbook.worksheets.getItem("Checks");
  assert.equal(assetDepreciation.getRange("D5").formulas[0][0], "=DB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$8)");
  assert.equal(assetDepreciation.getRange("E5").formulas[0][0], "=DDB('Inputs'!$B$5,'Inputs'!$B$6,'Inputs'!$B$7,A5,'Inputs'!$B$9)");
  assert.equal(assetDepreciation.getRange("F9").values[0][0], 10000);
  assert.deepEqual(assetDepreciationChecks.getRange("E4:E9").values, [["OK"], ["OK"], ["OK"], ["OK"], ["OK"], ["OK"]]);
  const assetDepreciationQa = await verifyWorkbookFile(assetDepreciationPath, {
    outputDir: path.join(outputDir, "openchestnut-asset-depreciation-native-qa"),
    sheetName: "Depreciation",
    renderFormat: "svg",
    nativeRender: "auto",
    allSheets: true,
  });
  if (nativeSpreadsheetRenderStatus().available) {
    assert.equal(assetDepreciationQa.summary.nativeRender.status, "passed");
    assert.equal(assetDepreciationQa.summary.nativeRender.ok, true);
  }

  const { createScatterWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-scatter-chart-workflow.mjs"
  );
  const scatterPath = path.join(outputDir, "openchestnut-scatter-chart-workflow.xlsx");
  const scatterResult = await createScatterWorkbook(scatterPath);
  assert.equal(scatterResult.verification.ok, true);
  const scatterWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(scatterPath));
  const scatterChart = scatterWorkbook.worksheets.getItem("Relationship Analysis").charts.items[0];
  assert.equal(scatterChart.type, "scatter");
  assert.deepEqual(scatterChart.series.items[0].xValues, [10, 20, 25, 34, 45]);
  const scatterZip = await JSZip.loadAsync(await fs.readFile(scatterPath));
  const scatterChartPath = Object.keys(scatterZip.files).find((name) => /\/charts\/chart\d+\.xml$/i.test(name));
  assert.ok(scatterChartPath);
  const scatterXml = await scatterZip.file(scatterChartPath).async("text");
  assert.match(scatterXml, /<c:scatterChart>/);
  assert.match(scatterXml, /<c:scatterStyle val="marker"/);
  assert.equal((scatterXml.match(/<a:ln><a:noFill\s*\/><\/a:ln>/g) || []).length, 2);
  assert.equal((scatterXml.match(/<c:valAx>/g) || []).length, 2);
  const scatterQa = await verifyWorkbookFile(scatterPath, {
    outputDir: path.join(outputDir, "openchestnut-scatter-native-qa"),
    sheetName: "Relationship Analysis",
    renderFormat: "svg",
    nativeRender: "auto",
  });
  if (nativeSpreadsheetRenderStatus().available) {
    assert.equal(scatterQa.summary.nativeRender.status, "passed");
    assert.equal(scatterQa.summary.nativeRender.pageCount, 1);
  }

  const { createBubbleWorkbook } = await import(
    "../skills/spreadsheets/skills/spreadsheets/examples/openchestnut-bubble-chart-workflow.mjs"
  );
  const bubblePath = path.join(outputDir, "openchestnut-bubble-chart-workflow.xlsx");
  const bubbleResult = await createBubbleWorkbook(bubblePath);
  assert.equal(bubbleResult.verification.ok, true);
  const bubbleWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(bubblePath));
  const bubbleChart = bubbleWorkbook.worksheets.getItem("Opportunity Analysis").charts.items[0];
  assert.equal(bubbleChart.type, "bubble");
  assert.deepEqual(bubbleChart.series.items[0].xValues, [10, 20, 25, 34, 45]);
  assert.deepEqual(bubbleChart.series.items[0].bubbleSizes, [4, 10, 12, 18, 27]);
  const bubbleZip = await JSZip.loadAsync(await fs.readFile(bubblePath));
  const bubbleChartPath = Object.keys(bubbleZip.files).find((name) => /\/charts\/chart\d+\.xml$/i.test(name));
  assert.ok(bubbleChartPath);
  const bubbleXml = await bubbleZip.file(bubbleChartPath).async("text");
  assert.match(bubbleXml, /<c:bubbleChart>/);
  assert.match(bubbleXml, /<c:bubbleSize>/);
  assert.equal((bubbleXml.match(/<c:valAx>/g) || []).length, 2);
  const bubbleQa = await verifyWorkbookFile(bubblePath, {
    outputDir: path.join(outputDir, "openchestnut-bubble-native-qa"),
    sheetName: "Opportunity Analysis",
    renderFormat: "svg",
    nativeRender: "auto",
  });
  if (nativeSpreadsheetRenderStatus().available) {
    assert.equal(bubbleQa.summary.nativeRender.status, "passed");
    assert.equal(bubbleQa.summary.nativeRender.pageCount, 1);
  }

  const queryResult = await runFixture("open-chestnut-query-table");
  assert.equal(queryResult.sourceQueryTable.sheet, "External Data");
  assert.equal(queryResult.sourceQueryTable.table, "ExternalSales");
  assert.equal(queryResult.sourceQueryTable.query.name, "Warehouse sales");
  assert.equal(queryResult.sourceConnections.length, 1);
  assert.equal(queryResult.sourceConnections[0].name, "Fixture warehouse");
  const queryZip = await JSZip.loadAsync(await fs.readFile(queryResult.workbookPath));
  assert.ok(queryZip.file("xl/connections.xml"));
  assert.ok(queryZip.file("xl/queryTables/queryTable1.xml"));
  assert.match(await queryZip.file("xl/queryTables/queryTable1.xml").async("text"), /fixture:opaque value="kept"/);

  const intersectionResult = await runFixture("structured-intersection");
  const intersectionWorkbook = await SpreadsheetFile.importXlsx(await FileBlob.load(intersectionResult.workbookPath));
  const intersection = intersectionWorkbook.worksheets.getItem("Intersection");
  assert.deepEqual(intersection.getRange("D2:D3").values, [[370], ["#NULL!"]]);
  assert.deepEqual(intersection.getRange("D2:D3").formulas, [
    ["=SUM(IntersectionTable[[Month]:[Revenue]] IntersectionTable[[Revenue]:[Cost]])"],
    ["=SUM(IntersectionTable[Month] IntersectionTable[Cost])"],
  ]);

  await assert.rejects(
    () => SpreadsheetFile.importXlsx(formulaBlob, { codec: "open-chestnut" }),
    /does not accept option codec/i,
  );
  await assert.rejects(
    () => SpreadsheetFile.exportXlsx(formulaWorkbook, { allowLossy: true }),
    /does not accept option allowLossy/i,
  );
  await assert.rejects(
    () => SpreadsheetFile.importXlsx(formulaBlob, { limits: { maxInputBytes: 16 } }),
    /input_budget_exceeded|exceeds max_input_bytes/i,
  );

  const csv = await SpreadsheetFile.exportCsv(formulaWorkbook, { sheetName: "Summary", range: "A1:E4" });
  const csvPath = path.join(outputDir, "summary.csv");
  await csv.save(csvPath);
  const csvQa = await verifyWorkbookFile(csvPath, {
    inputFormat: "csv",
    outputDir: path.join(outputDir, "csv-qa"),
    sheetName: "Summary",
    renderFormat: "svg",
    nativeRender: "off",
    coerceTypes: true,
  });
  assert.equal(csvQa.summary.packageOk, true);
  assert.equal(csvQa.summary.verifyOk, true);
  assert.equal(csvQa.summary.visualQaOk, true);

  console.log("spreadsheet skill tests passed");
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}
