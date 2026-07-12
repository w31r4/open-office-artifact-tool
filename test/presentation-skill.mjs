import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

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
  assert.equal(first.qa.presentation.layouts.items[0].id, "pptx-layout-2147483650");
  assert.equal(first.qa.presentation.layouts.items[0].masterId, "pptx-master-2147483648");
  assert.equal(first.qa.presentation.master.name, "Agent Readiness Master");
  assert.deepEqual(first.qa.presentation.master.background, { fill: "bg1", mode: "reference", index: 1001 });
  assert.equal(first.qa.presentation.master.placeholders[0].style.color, "accent1");
  assert.equal(first.qa.presentation.masters.count, 2);
  assert.equal(first.qa.presentation.masters.items[1].name, "Verification Master");
  assert.equal(first.qa.presentation.masters.items[1].theme?.name, "Verification Theme");
  assert.equal(first.qa.presentation.masters.items[1].effectiveTheme().colors.accent1, "#ede9fe");
  assert.equal(first.qa.presentation.masters.items[1].effectiveTheme().fonts.major, "Georgia");
  assert.equal(first.qa.presentation.layouts.items[1].id, "pptx-layout-2147483651");
  assert.equal(first.qa.presentation.layouts.items[1].masterId, "pptx-master-2147483649");
  assert.deepEqual(first.qa.presentation.layouts.items[0].background, { fill: "#ffffff", mode: "solid" });
  assert.equal(first.qa.presentation.theme.colors.accent6, "#dc2626");
  assert.equal(first.qa.presentation.theme.fonts.majorEastAsia, "PingFang SC");
  assert.equal(first.qa.presentation.theme.textStyles.title.fontSize, 42);
  assert.equal(first.qa.presentation.theme.textStyles.body.color, "tx2");
  assert.equal(first.qa.presentation.theme.colorMap.accent1, "accent2");
  assert.deepEqual(first.qa.presentation.slides.items[0].effectiveBackground(), { fill: "accent1", mode: "reference", index: 1001 });
  assert.equal(first.qa.presentation.slides.items[0].effectiveTheme().name, "Verification Theme");
  assert.equal(first.qa.modelRender.slides.length, 3);
  assert.equal(first.qa.modelRender.montage.ok, true);
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/slides/slide1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/charts/chart1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/media/image1.png"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/notesSlides/notesSlide1.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/comments/comment2.xml"));
  assert.ok(first.qa.packageInspect.parts.some((part) => part.path === "ppt/commentAuthors.xml"));
  const fixtureZip = await JSZip.loadAsync(await fs.readFile(first.pptxPath));
  const fixtureThemeXml = await fixtureZip.file("ppt/theme/theme1.xml").async("text");
  const fixtureSecondThemeXml = await fixtureZip.file("ppt/theme/theme2.xml").async("text");
  const fixtureMasterXml = await fixtureZip.file("ppt/slideMasters/slideMaster1.xml").async("text");
  const fixtureSecondMasterXml = await fixtureZip.file("ppt/slideMasters/slideMaster2.xml").async("text");
  const fixturePresentationXml = await fixtureZip.file("ppt/presentation.xml").async("text");
  assert.match(fixtureThemeXml, /<a:accent6><a:srgbClr val="DC2626"\/><\/a:accent6>/);
  assert.match(fixtureSecondThemeXml, /name="Verification Theme"/);
  assert.match(fixtureSecondThemeXml, /<a:accent1><a:srgbClr val="EDE9FE"\/><\/a:accent1>/);
  assert.equal((/<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/.exec(fixtureThemeXml)?.[1].match(/<a:solidFill>/g) || []).length, 3);
  assert.match(fixtureMasterXml, /<p:clrMap[^>]*accent1="accent2"/);
  assert.match(fixtureMasterXml, /<p:titleStyle>[\s\S]*?<a:defRPr sz="4200"/);
  assert.match(fixtureMasterXml, /<p:cSld name="Agent Readiness Master"><p:bg><p:bgRef idx="1001">/);
  assert.match(fixtureMasterXml, /<p:ph type="title" idx="1"\/>/);
  assert.match(fixtureSecondMasterXml, /<p:cSld name="Verification Master"><p:bg><p:bgRef idx="1001">/);
  assert.match(fixtureSecondMasterXml, /<p:clrMap[^>]*accent1="accent1"/);
  assert.equal((fixturePresentationXml.match(/<p:sldMasterId\b/g) || []).length, 2);
  assert.match(await fixtureZip.file("ppt/slideLayouts/_rels/slideLayout2.xml.rels").async("text"), /slideMaster2\.xml/);
  assert.match(await fixtureZip.file("ppt/slideMasters/_rels/slideMaster2.xml.rels").async("text"), /theme2\.xml/);
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

  const packageDrawing = await runPresentationFixture("skills/presentations/fixtures/package-drawing.json", {
    outputDir: path.join(root, "package-drawing"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(packageDrawing.qa.summary.packageOk, true);
  assert.equal(packageDrawing.qa.verify.ok, true);
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/media/agent-status.png"));
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/charts/agent-readiness.xml"));
  const drawingSlide = packageDrawing.qa.presentation.slides.items[0];
  assert.equal(drawingSlide.images.items.find((item) => item.name === "Agent status")?.alt, "Green package status");
  assert.equal(drawingSlide.charts.items.find((item) => item.name === "Agent readiness chart")?.series[0].values[1], 100);
  assert.match(packageDrawing.qa.inspect.ndjson, /Agent readiness chart/);
  assert.match(packageDrawing.qa.inspect.ndjson, /Agent status/);
  assert.equal(packageDrawing.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const packageReview = await runPresentationFixture("skills/presentations/fixtures/package-notes-comments.json", {
    outputDir: path.join(root, "package-notes-comments"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(packageReview.qa.summary.packageOk, true);
  assert.equal(packageReview.qa.packageInspect.records[0].semanticValidation, true);
  assert.equal(packageReview.qa.packageInspect.records[0].semanticIssues, 0);
  assert.ok(packageReview.qa.packageInspect.parts.some((part) => part.path === "ppt/review/notes/agent-notes.xml"));
  assert.ok(packageReview.qa.packageInspect.parts.some((part) => part.path === "ppt/review/comments/agent-review.xml"));
  assert.ok(packageReview.qa.packageInspect.parts.some((part) => part.path === "ppt/review/comments/agent-authors.xml"));
  assert.match(packageReview.qa.presentation.slides.items[0].speakerNotes.text, /author identity and package relationships/);
  assert.deepEqual(packageReview.qa.presentation.slides.items[0].comments.items[0].comments.map((comment) => comment.author), ["QA Agent", "Maintainer"]);
  assert.equal(packageReview.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");

  const modernComments = await runPresentationFixture("skills/presentations/fixtures/modern-comments.json", {
    outputDir: path.join(root, "modern-comments"),
    nativeRender: nativeStatus.available ? "required" : "auto",
  });
  assert.equal(modernComments.qa.summary.packageOk, true);
  assert.equal(modernComments.qa.verify.ok, true);
  assert.equal(modernComments.qa.presentation.commentFormat, "modern");
  assert.ok(modernComments.qa.packageInspect.parts.some((part) => part.path === "ppt/authors.xml"));
  assert.ok(modernComments.qa.packageInspect.parts.some((part) => part.path === "ppt/comments/comment1.xml"));
  const modernSkillZip = await JSZip.loadAsync(await fs.readFile(modernComments.pptxPath));
  assert.match(await modernSkillZip.file("ppt/authors.xml").async("text"), /p188:authorLst/);
  const modernSkillCommentsXml = await modernSkillZip.file("ppt/comments/comment1.xml").async("text");
  const modernSkillSlideXml = await modernSkillZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(modernSkillCommentsXml, /p188:replyLst/);
  assert.match(modernSkillCommentsXml, /oac:txMkLst>[\s\S]*?<pc:sldMkLst>[\s\S]*?<oac:spMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:txMk cp="0" len="44"\/>/);
  assert.match(modernSkillCommentsXml, /oac:deMkLst>[\s\S]*?<oac:grpSpMk id="3" creationId="\{[0-9A-F-]+\}"\/>/);
  assert.match(modernSkillSlideXml, /<p:grpSp>[\s\S]*?<p:cNvPr id="3" name="review-evidence-group">/);
  assert.match(modernSkillSlideXml, /a16:creationId[^>]*id="\{[0-9A-F-]+\}"/);
  const modernSkillThread = modernComments.qa.presentation.slides.items[0].comments.items[0];
  const modernSkillTarget = modernComments.qa.presentation.slides.items[0].shapes.items.find((shape) => shape.name === "review-title");
  const modernSkillGroupThread = modernComments.qa.presentation.slides.items[0].comments.items[1];
  const modernSkillGroup = modernComments.qa.presentation.slides.items[0].groups.items[0];
  assert.equal(modernSkillThread.nativeFormat, "modern");
  assert.equal(modernSkillThread.resolved, true);
  assert.equal(modernSkillThread.targetId, `${modernSkillTarget.id}/text`);
  assert.equal(modernComments.qa.presentation.slides.items[0].resolve(modernSkillThread.targetId)?.kind, "textRange");
  assert.equal(modernSkillThread.nativeAnchor.type, "textRange");
  assert.equal(modernSkillThread.nativeAnchor.cp, 0);
  assert.equal(modernSkillThread.nativeAnchor.length, 44);
  assert.equal(modernSkillThread.nativeAnchor.moniker, "spMk");
  assert.equal(modernSkillThread.nativeAnchor.nativeId, modernSkillTarget.nativeId);
  assert.equal(modernSkillThread.nativeAnchor.creationId, modernSkillTarget.creationId);
  assert.equal(modernSkillGroupThread.targetId, modernSkillGroup.id);
  assert.equal(modernSkillGroupThread.nativeAnchor.moniker, "grpSpMk");
  assert.equal(modernSkillGroupThread.nativeAnchor.creationId, modernSkillGroup.creationId);
  assert.equal(modernSkillGroup.shapes.items[0].text.value.includes("text and group anchors"), true);
  assert.deepEqual(modernSkillThread.comments.map((comment) => comment.author), ["QA Agent", "Maintainer"]);
  assert.ok(modernSkillThread.comments.every((comment) => /^\{[0-9A-F-]+\}$/.test(comment.nativeId)));
  assert.equal(modernComments.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  const skillText = await fs.readFile("skills/presentations/SKILL.md", "utf8");
  assert.match(skillText, /PresentationFile\.patchPptx/);
  assert.match(skillText, /package-drawing\.json/);
  assert.match(skillText, /package-notes-comments\.json/);
  assert.match(skillText, /modern-comments\.json/);

  console.log("presentation skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
