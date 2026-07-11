import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { PdfFile } from "../src/index.mjs";
import { nativePdfRenderStatus, runPdfFixture, verifyPdfFile } from "../skills/pdf/scripts/workflow.mjs";

const root = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-skill-test-"));
const baselineDir = path.join(root, "baselines");
const nativeStatus = nativePdfRenderStatus();

try {
  const first = await runPdfFixture("skills/pdf/fixtures/qa-report.json", { outputDir: path.join(root, "fixture"), nativeRender: nativeStatus.available ? "required" : "auto", pdfjs: "required", baselineDir, writeBaseline: true });
  assert.equal(first.qa.verify.ok, true);
  assert.equal(first.qa.pdf.pages.length, 3);
  assert.equal(first.qa.fileInspect.summary.pages, 3);
  assert.ok(first.qa.fileInspect.summary.objects >= 30);
  assert.equal(first.qa.fileInspect.summary.hasEmbeddedModel, true);
  assert.equal(first.qa.fileInspect.summary.hasEof, true);
  assert.equal(first.qa.fileInspect.summary.tagged, true);
  assert.equal(first.qa.fileInspect.summary.language, "en-US");
  assert.ok(first.qa.fileInspect.summary.structureElements >= 10);
  assert.equal(first.qa.fileInspect.summary.markedContentItems, first.qa.fileInspect.summary.structureElements);
  const untaggedPath = path.join(root, "untagged.pdf");
  await (await PdfFile.exportPdf(first.qa.pdf, { tagged: false })).save(untaggedPath);
  await assert.rejects(() => verifyPdfFile(untaggedPath, { outputDir: path.join(root, "untagged-qa"), nativeRender: "off", pdfjs: "off", requireTagged: true }), /tagged=false/);
  assert.equal(first.qa.modelRender.pages.length, 3);
  assert.equal(first.qa.pdfjs.status, "passed");
  assert.equal(first.qa.pdfjs.pdf.pages.length, 3);
  assert.match(first.qa.pdfjs.text, /Office artifact QA report/);
  assert.match(first.qa.pdfjs.text, /PDF\.js and Poppler agree/);
  assert.match(first.qa.pdfjs.text, /Automatically paginated appendix/);
  assert.ok(first.qa.pdfjs.pdf.pages.flatMap((page) => page.images).some((image) => /^data:image\/png;base64,/.test(image.dataUrl || "")));
  assert.match(first.qa.inspect.ndjson, /qa-gates/);
  assert.match(first.qa.inspect.ndjson, /Verified workflow/);
  assert.match(first.qa.inspect.ndjson, /Independent checks accumulate/);
  assert.equal(first.qa.extractedTables.length, 2);
  assert.equal(first.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(first.qa.nativeRender.pageCount, 3);

  const compared = await verifyPdfFile(first.pdfPath, { outputDir: path.join(root, "compare"), nativeRender: nativeStatus.available ? "required" : "auto", pdfjs: "required", baselineDir });
  assert.equal(compared.verify.ok, true);
  assert.equal(compared.pdfjs.status, "passed");
  assert.equal(compared.modelRender.baselinePageCount, 3);
  assert.equal(compared.modelRender.pageCountMatches, true);
  assert.ok(compared.modelRender.pages.every((page) => page.baselineCompared && page.pixelDiff?.changed === false && page.ok));
  if (nativeStatus.available) {
    assert.equal(compared.nativeRender.baselinePageCount, 3);
    assert.equal(compared.nativeRender.pageCountMatches, true);
    assert.ok(compared.nativeRender.pages.every((page) => page.baselineCompared && page.pixelDiff?.changed === false && page.ok));
  }
  for (const page of compared.modelRender.pages) assert.ok((await fs.stat(page.path)).size > 100);
  if (nativeStatus.available) for (const page of compared.nativeRender.pages) assert.ok((await fs.stat(page.path)).size > 100);
  console.log("pdf skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
