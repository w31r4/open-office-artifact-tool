import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DocumentFile, FileBlob } from "open-office-artifact-tool";
import {
  nativeDocumentRenderStatus,
  runDocumentFixture,
  verifyDocumentFile,
} from "../skills/documents/scripts/workflow.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const fixturePath = path.join(repoRoot, "skills", "documents", "fixtures", "business-brief.json");
const outputDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-document-skill-"));
const baselineDir = path.join(outputDir, "baselines");

try {
  const result = await runDocumentFixture(fixturePath, { outputDir, nativeRender: "off" });
  assert.equal(result.fixture.name, "business-brief");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.packageOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.nativeRender.status, "skipped");
  for (const filePath of Object.values(result.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty document skill output ${filePath}`);
  }
  const imported = await DocumentFile.importDocx(await FileBlob.load(result.docxPath));
  const inspect = imported.inspect({ kind: "paragraph,listItem,table,comment,header,hyperlink,citation,image,field,section", maxChars: 20_000 }).ndjson;
  assert.match(inspect, /Office artifact readiness brief/);
  assert.match(inspect, /readiness-table/);
  assert.match(inspect, /native render review/);
  assert.match(inspect, /Opening section evidence/);
  assert.equal(imported.headers.find((item) => item.name === "opening-header")?.sectionIndex, 0);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/document\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.preview, "utf8"), /<svg/);
  const nativePreferred = await verifyDocumentFile(result.docxPath, { outputDir: path.join(outputDir, "native-preferred"), preferNative: true, nativeRender: "off" });
  assert.equal(nativePreferred.summary.verifyOk, true);
  assert.match(nativePreferred.inspect.ndjson, /Office artifact readiness brief/);
  const nativePreferredDocument = await DocumentFile.importDocx(await FileBlob.load(result.docxPath), { preferNative: true });
  assert.equal(nativePreferredDocument.headers.find((item) => item.text === "Opening section evidence")?.sectionIndex, 0);
  assert.equal(nativePreferredDocument.headers.find((item) => item.text === "Clean-room document workflow")?.sectionIndex, 1);

  const nativeStatus = nativeDocumentRenderStatus();
  const baselineWrite = await verifyDocumentFile(result.docxPath, {
    outputDir: path.join(outputDir, "baseline-write"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "off",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(baselineWrite.summary.writeBaseline, true);
  assert.ok((await fs.stat(baselineWrite.summary.modelBaselinePath)).size > 100);
  const baselineCompare = await verifyDocumentFile(result.docxPath, {
    outputDir: path.join(outputDir, "baseline-compare"),
    previewFormat: "png",
    nativeRender: nativeStatus.available ? "required" : "off",
    baselineDir,
  });
  assert.equal(baselineCompare.summary.modelBaselineCompared, true);
  assert.equal(baselineCompare.summary.modelPixelDiff.changed, false);
  assert.equal(baselineCompare.summary.visualQaOk, true);
  if (nativeStatus.available) {
    assert.equal(baselineCompare.summary.nativeRender.status, "passed");
    assert.equal(baselineCompare.summary.nativeRender.ok, true);
    assert.equal(baselineCompare.summary.nativeRender.pageCountMatches, true);
    assert.ok(baselineCompare.summary.nativeRender.pageCount >= 1);
    for (const page of baselineCompare.summary.nativeRender.pages) {
      assert.equal(page.baselineCompared, true);
      assert.equal(page.pixelDiff.changed, false);
      assert.equal(page.ok, true);
      assert.ok((await fs.stat(page.path)).size > 100);
    }
  }

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "SKILL.md"), "utf8");
  assert.match(skillText, /LibreOffice PDF plus Poppler page PNGs/);
  assert.match(skillText, /baseline-dir/);
  assert.match(skillText, /preferNative/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");
