import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import {
  nativePresentationRenderStatus,
  runPresentationFixture,
  verifyPresentationFile,
} from "../skills/presentations/scripts/workflow.mjs";

const root = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-presentation-skill-test-"));
const baselineDir = path.join(root, "baselines");
const nativeStatus = nativePresentationRenderStatus();

try {
  const first = await runPresentationFixture("skills/presentations/fixtures/agent-readiness.json", {
    outputDir: path.join(root, "fixture"),
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(first.qa.verify.ok, true);
  assert.equal(first.qa.presentation.slides.count, 3);
  assert.equal(first.qa.presentation.layouts.items[0].id, "pptx-layout-2147483649");
  assert.equal(first.qa.presentation.layouts.items[0].masterId, "pptx-master-2147483648");
  assert.equal(first.qa.modelRender.slides.length, 3);
  assert.equal(first.qa.modelRender.montage.ok, true);
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/slides/slide1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/charts/chart1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/media/image1.png"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/notesSlides/notesSlide1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/comments/comment2.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/commentAuthors.xml"));
  const importedSkillThread = first.qa.presentation.slides.items[1].comments.items[0];
  assert.deepEqual(importedSkillThread.comments.map((comment) => comment.author), ["QA Agent", "Maintainer"]);
  assert.equal(first.qa.summary.packageOk, true);
  assert.match(first.qa.inspect.ndjson, /PPTX workflows now close the QA loop/);
  assert.match(first.qa.inspect.ndjson, /evidence-table/);
  assert.match(first.qa.inspect.ndjson, /Evidence produced per gate/);
  assert.match(first.qa.inspect.ndjson, /Verified clean-room workflow/);
  assert.match(first.qa.inspect.ndjson, /Inspect every rendered slide/);
  assert.match(first.qa.inspect.ndjson, /chart-to-semantic/);
  assert.equal(first.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(first.qa.nativeRender.pageCount, 3);

  const compared = await verifyPresentationFile(first.pptxPath, {
    outputDir: path.join(root, "compare"),
    nativeRender: nativeStatus.available ? "required" : "auto",
    baselineDir,
  });
  assert.equal(compared.verify.ok, true);
  assert.equal(compared.modelRender.baselinePageCount, 3);
  assert.equal(compared.modelRender.pageCountMatches, true);
  assert.ok(compared.modelRender.slides.every((slide) => slide.baselineCompared && slide.pixelDiff?.changed === false && slide.ok));
  if (nativeStatus.available) {
    assert.equal(compared.nativeRender.baselinePageCount, 3);
    assert.equal(compared.nativeRender.pageCountMatches, true);
    assert.ok(compared.nativeRender.pages.every((slide) => slide.baselineCompared && slide.pixelDiff?.changed === false && slide.ok));
  }
  for (const slide of compared.modelRender.slides) assert.ok((await fs.stat(slide.path)).size > 100);
  if (nativeStatus.available) for (const slide of compared.nativeRender.pages) assert.ok((await fs.stat(slide.path)).size > 100);
  console.log("presentation skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
