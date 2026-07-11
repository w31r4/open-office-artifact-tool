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

try {
  const result = await runDocumentFixture(fixturePath, { outputDir, nativeRender: "off" });
  assert.equal(result.fixture.name, "business-brief");
  assert.equal(result.qa.summary.verifyOk, true);
  assert.equal(result.qa.summary.visualQaOk, true);
  assert.equal(result.qa.summary.nativeRender.status, "skipped");
  for (const filePath of Object.values(result.qa.summary.files)) {
    const stat = await fs.stat(filePath);
    assert.ok(stat.isFile() && stat.size > 0, `Expected non-empty document skill output ${filePath}`);
  }
  const imported = await DocumentFile.importDocx(await FileBlob.load(result.docxPath));
  const inspect = imported.inspect({ kind: "paragraph,listItem,table,comment,hyperlink,citation,image,field", maxChars: 20_000 }).ndjson;
  assert.match(inspect, /Office artifact readiness brief/);
  assert.match(inspect, /readiness-table/);
  assert.match(inspect, /native render review/);
  assert.match(await fs.readFile(result.qa.summary.files.packageInspect, "utf8"), /word\/document\.xml/);
  assert.match(await fs.readFile(result.qa.summary.files.preview, "utf8"), /<svg/);

  const nativeStatus = nativeDocumentRenderStatus();
  if (nativeStatus.available) {
    const nativeQa = await verifyDocumentFile(result.docxPath, {
      outputDir: path.join(outputDir, "native-qa"),
      previewFormat: "png",
      nativeRender: "required",
    });
    assert.equal(nativeQa.summary.nativeRender.status, "passed");
    assert.ok(nativeQa.summary.nativeRender.pageCount >= 1);
    for (const page of nativeQa.summary.nativeRender.pages) {
      const stat = await fs.stat(page.path);
      assert.ok(stat.size > 100);
      assert.equal(page.ok, true);
    }
  }

  const packageJson = JSON.parse(await fs.readFile(path.join(repoRoot, "package.json"), "utf8"));
  assert.ok(packageJson.files.includes("skills/**"));
  const skillText = await fs.readFile(path.join(repoRoot, "skills", "documents", "SKILL.md"), "utf8");
  assert.match(skillText, /LibreOffice PDF plus Poppler page PNGs/);
} finally {
  await fs.rm(outputDir, { recursive: true, force: true });
}

console.log("document skill smoke ok");

