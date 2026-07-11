import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob, SpreadsheetFile } from "open-office-artifact-tool";
import { runSpreadsheetFixture, verifyWorkbookFile } from "../skills/spreadsheets/scripts/workflow.mjs";

const repoRoot = path.resolve(new URL("..", import.meta.url).pathname);
const fixturePath = path.join(repoRoot, "skills", "spreadsheets", "fixtures", "formula-summary.json");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-spreadsheet-skill-"));
const baselineDir = path.join(outputDir, "baselines");

try {
  const result = await runSpreadsheetFixture(fixturePath, { outputDir });
  assert.equal(result.fixture.name, "formula-summary");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.renderFormat, "svg");
  for (const filePath of Object.values(result.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty skill output ${filePath}`);
  }
  const workbook = await SpreadsheetFile.importXlsx(await FileBlob.load(result.workbookPath));
  assert.deepEqual(workbook.worksheets.getItem("Summary").getRange("B2:D4").values, [
    [100, 60, 0.4],
    [120, 70, 0.4166666666666667],
    [150, 90, 0.4],
  ]);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /SummaryTable/);
  assert.match(await fs.readFile(result.qa.summary.files.inspect, "utf8"), /Inputs!B2/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /xl\/workbook\.xml/);
  assert.equal(result.qa.packageInspect.records[0].sheets, 2);
  assert.match(await fs.readFile(result.qa.summary.files.preview, "utf8"), /<svg/);

  const secondQa = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "second-qa"),
    sheetName: "Inputs",
    range: "A1:C4",
    renderFormat: "svg",
  });
  assert.equal(secondQa.summary.verifyOk, true);
  assert.equal(secondQa.summary.packageOk, true);
  assert.equal(secondQa.summary.sheetName, "Inputs");

  const baselineWrite = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    sheetName: "Summary",
    range: "A1:D4",
    renderFormat: "png",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.baselinePath)).size > 100);
  const baselineCompare = await verifyWorkbookFile(result.workbookPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    sheetName: "Summary",
    range: "A1:D4",
    renderFormat: "png",
    baselineDir,
  });
  assert.equal(baselineCompare.summary.baselineCompared, true);
  assert.equal(baselineCompare.summary.pixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "spreadsheets", "SKILL.md"), "utf8");
  assert.match(skillText, /open-office-artifact-tool/);
  assert.match(skillText, /verify-workbook\.mjs/);
  assert.match(skillText, /baseline-dir/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("spreadsheet skill smoke ok");
