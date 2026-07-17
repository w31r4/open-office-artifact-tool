import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";
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
  assert.equal((scatterXml.match(/<a:ln><a:noFill\/><\/a:ln>/g) || []).length, 2);
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
