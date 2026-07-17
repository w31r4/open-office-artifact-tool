import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import {
  nativePresentationRenderStatus,
  runPresentationFixture,
  verifyPresentationFile,
} from "./skill-harness/presentations/scripts/workflow.mjs";

const fixtureDir = path.join("test", "skill-harness", "presentations", "fixtures");

const root = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-presentation-skill-test-"));
const baselineDir = path.join(root, "baselines");
const nativeStatus = nativePresentationRenderStatus();
const nativeRender = nativeStatus.available ? "required" : "auto";

function itemByName(items, name) {
  const item = items.find((candidate) => candidate.name === name);
  assert.ok(item, "Missing presentation skill object " + name);
  return item;
}

try {
  const readiness = await runPresentationFixture(path.join(fixtureDir, "agent-readiness.json"), {
    outputDir: path.join(root, "agent-readiness"),
    nativeRender,
    baselineDir,
    writeBaseline: true,
  });
  assert.equal(Object.hasOwn(readiness, "roundtripCodec"), false);
  assert.equal(readiness.qa.verify.ok, true);
  assert.equal(readiness.qa.summary.packageOk, true);
  assert.equal(readiness.qa.presentation.slides.count, 3);
  assert.equal(readiness.qa.modelRender.slides.length, 3);
  assert.equal(readiness.qa.modelRender.montage.ok, true);
  assert.equal(readiness.qa.nativeRender.status, nativeStatus.available ? "passed" : "skipped");
  if (nativeStatus.available) assert.equal(readiness.qa.nativeRender.pageCount, 3);

  const workflowSlide = readiness.qa.presentation.slides.getItem(0);
  assert.deepEqual(workflowSlide.background, { fill: "#f1f5f9", mode: "solid" });
  assert.deepEqual(readiness.qa.presentation.slides.getItem(1).background, { fill: "#fff7ed", mode: "solid" });
  const authoredCard = itemByName(workflowSlide.shapes.items, "author-card");
  assert.equal(authoredCard.geometry, "roundRect");
  assert.deepEqual(authoredCard.shadow, { color: "#000000", blurRadius: 10, distance: 5, direction: 45, opacity: 0.2 });
  assert.equal(itemByName(workflowSlide.shapes.items, "workflow-title").geometry, "textbox");
  assert.equal(itemByName(workflowSlide.tables.items, "workflow-matrix").values[1][1], "Pass");
  const elbow = itemByName(workflowSlide.connectors.items, "verify-to-deliver");
  assert.equal(elbow.connectorType, "elbow");
  assert.equal(elbow.line.startArrow, "triangle");
  assert.equal(elbow.line.endArrow, "triangle");
  assert.ok(elbow.startTargetId && elbow.endTargetId);
  const charts = readiness.qa.presentation.slides.getItem(1).charts.items;
  assert.deepEqual(charts.map((chart) => chart.chartType), ["bar", "line", "pie"]);
  assert.equal(charts[0].dataLabels.showValue, true);
  assert.equal(charts[1].series[0].marker.symbol, "circle");
  assert.equal(charts[2].dataLabels.showCategoryName, true);
  const groupedWorkflowSlide = readiness.qa.presentation.slides.getItem(2);
  const nativeGroup = itemByName(groupedWorkflowSlide.groups.items, "native-agent-group");
  assert.deepEqual(nativeGroup.childFrame, { left: -80, top: 40, width: 1280, height: 540 });
  assert.deepEqual(nativeGroup.children.map((child) => child.layoutJson().kind), ["textbox", "textbox", "textbox", "groupShape", "connector"]);
  assert.equal(itemByName(nativeGroup.groups.items, "nested-qa-group").shapes.items[0].text.value, "Render + verify");
  assert.equal(itemByName(nativeGroup.connectors.items, "grouped-flow").line.endArrow, "triangle");
  assert.match(readiness.qa.inspect.ndjson, /OpenChestnut closes the presentation loop/);
  assert.match(readiness.qa.inspect.ndjson, /Coverage mix/);
  assert.match(readiness.qa.inspect.ndjson, /native-agent-group/);

  const readinessZip = await JSZip.loadAsync(await fs.readFile(readiness.pptxPath));
  const firstSlideXml = await readinessZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(firstSlideXml, /<a:srgbClr val="F1F5F9"/);
  assert.match(firstSlideXml, /<a:prstGeom prst="roundRect"[^>]*>/);
  assert.match(firstSlideXml, /<p:cNvSpPr txBox="1"\s*\/>/);
  assert.match(firstSlideXml, /<p:cxnSp>/);
  const groupedSlideXml = await readinessZip.file("ppt/slides/slide3.xml").async("text");
  assert.equal((groupedSlideXml.match(/<p:grpSp>/g) || []).length, 2);
  assert.match(groupedSlideXml, /<a:chOff x="-762000" y="381000"\s*\/>/);
  assert.equal(Object.keys(readinessZip.files).filter((name) => /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/.test(name)).length, 3);

  const compared = await verifyPresentationFile(readiness.pptxPath, {
    outputDir: path.join(root, "agent-readiness-compare"),
    nativeRender,
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

  const roundtrip = await runPresentationFixture(path.join(fixtureDir, "open-chestnut-preservation.json"), {
    outputDir: path.join(root, "roundtrip"),
    nativeRender,
  });
  assert.equal(roundtrip.qa.verify.ok, true);
  assert.equal(roundtrip.qa.summary.packageOk, true);
  const roundtripSlide = roundtrip.qa.presentation.slides.getItem(0);
  const editedCard = itemByName(roundtripSlide.shapes.items, "source-card");
  assert.equal(editedCard.text.value, "After edit");
  assert.equal(editedCard.shadow.opacity, 0.35);
  assert.equal(itemByName(roundtripSlide.connectors.items, "editable-connector").line.endArrow, undefined);
  assert.equal(itemByName(roundtripSlide.tables.items, "roundtrip-table").values[1][1], "After");
  const editedImage = itemByName(roundtripSlide.images.items, "edited-roundtrip-image");
  assert.equal(editedImage.alt, "Roundtrip status after edit");
  assert.deepEqual(editedImage.position, { left: 430, top: 320, width: 170, height: 170 });
  assert.equal(editedImage.fit, "stretch");
  assert.deepEqual(editedImage.crop, { left: 0, top: -0.5, right: 0, bottom: -0.5 });
  const roundtripZip = await JSZip.loadAsync(await fs.readFile(roundtrip.pptxPath));
  const roundtripSlideXml = await roundtripZip.file("ppt/slides/slide1.xml").async("text");
  assert.match(roundtripSlideXml, /<a:srcRect[^>]*t="-50000"/);
  assert.match(roundtripSlideXml, /<a:srcRect[^>]*b="-50000"/);
  const editedChart = itemByName(roundtripSlide.charts.items, "roundtrip-chart");
  assert.equal(editedChart.title, "After roundtrip");
  assert.deepEqual(editedChart.series[0].values, [8, 14, 6]);

  const packageDrawing = await runPresentationFixture(path.join(fixtureDir, "package-drawing.json"), {
    outputDir: path.join(root, "package-drawing"),
    nativeRender,
  });
  assert.equal(packageDrawing.qa.verify.ok, true);
  assert.equal(packageDrawing.qa.summary.packageOk, true);
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/media/agent-status.png"));
  assert.ok(packageDrawing.qa.packageInspect.parts.some((part) => part.path === "ppt/review/charts/agent-readiness.xml"));
  const drawingSlide = packageDrawing.qa.presentation.slides.getItem(0);
  assert.equal(itemByName(drawingSlide.images.items, "Agent status").alt, "Green package status");
  assert.equal(itemByName(drawingSlide.charts.items, "Agent readiness chart").series[0].values[1], 100);

  const coreReview = await runPresentationFixture(path.join(fixtureDir, "modern-comments.json"), {
    outputDir: path.join(root, "core-review"),
    nativeRender,
  });
  assert.equal(coreReview.qa.verify.ok, true);
  assert.equal(coreReview.qa.presentation.slides.getItem(0).comments.items.length, 0);
  assert.equal(itemByName(coreReview.qa.presentation.slides.getItem(0).images.items, "review-status").alt, "Green review status");
  assert.equal(itemByName(coreReview.qa.presentation.slides.getItem(0).connectors.items, "visual-to-delivery").line.endArrow, "triangle");

  const legacyComments = await runPresentationFixture(path.join(fixtureDir, "open-chestnut-legacy-comments.json"), {
    outputDir: path.join(root, "legacy-comments"),
    nativeRender,
  });
  assert.equal(legacyComments.qa.verify.ok, true);
  assert.equal(legacyComments.qa.summary.packageOk, true);
  const legacyComment = legacyComments.qa.presentation.slides.getItem(0).comments.items[0];
  assert.ok(legacyComment, "legacy comment fixture must survive canonical import/export");
  assert.equal(legacyComment.nativeFormat, "legacy");
  assert.equal(legacyComment.targetId, undefined);
  assert.equal(legacyComment.author, "Presentation Reviewer");
  assert.equal(legacyComment.comments.length, 1);
  assert.equal(legacyComment.comments[0].text, "Confirm the source before delivery.");
  assert.deepEqual(legacyComment.position, { x: 1040, y: 84, unit: "px" });
  assert.match(legacyComments.qa.inspect.ndjson, /Confirm the source before delivery\./);
  const legacyZip = await JSZip.loadAsync(await fs.readFile(legacyComments.pptxPath));
  const commentAuthorsXml = await legacyZip.file("ppt/commentAuthors.xml").async("text");
  const commentXml = await legacyZip.file("ppt/comments/comment1.xml").async("text");
  assert.match(commentAuthorsXml, /<p:cmAuthor[^>]*name="Presentation Reviewer"/);
  assert.match(commentXml, /<p:cm[^>]*authorId="0"[^>]*idx="0"/);
  assert.match(commentXml, /Confirm the source before delivery\./);

  const coreEvidence = await runPresentationFixture(path.join(fixtureDir, "package-notes-comments.json"), {
    outputDir: path.join(root, "core-evidence"),
    nativeRender,
  });
  assert.equal(coreEvidence.qa.verify.ok, true);
  const evidenceSlide = coreEvidence.qa.presentation.slides.getItem(0);
  const evidenceCopy = itemByName(evidenceSlide.shapes.items, "evidence-copy");
  assert.equal(evidenceCopy.geometry, "textbox");
  assert.equal(evidenceCopy.text.paragraphs[1].bulletCharacter, "•");
  assert.deepEqual(evidenceCopy.text.paragraphs[3].runs[0].link, {
    uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/",
  });
  assert.equal(itemByName(evidenceSlide.charts.items, "evidence-pie").chartType, "pie");

  const convergenceFiles = [
    "test/skill-harness/presentations/scripts/workflow.mjs",
    "test/skill-harness/presentations/scripts/run-fixture.mjs",
    ...[
      "agent-readiness.json",
      "modern-comments.json",
      "open-chestnut-legacy-comments.json",
      "open-chestnut-preservation.json",
      "package-drawing.json",
      "package-notes-comments.json",
    ].map((name) => path.join(fixtureDir, name)),
  ];
  for (const file of convergenceFiles) {
    const source = await fs.readFile(file, "utf8");
    assert.doesNotMatch(source, /\b(?:codec|initialCodec|roundtripCodec)\b/i, file + " must not expose an Office path selector");
  }
  const skillText = await fs.readFile("skills/presentations/skills/presentations/SKILL.md", "utf8");
  const quickStartText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/API_QUICK_START.md", "utf8");
  assert.match(skillText, /open-office-artifact-tool/);
  assert.match(quickStartText, /PresentationFile\.exportPptx/);
  assert.match(quickStartText, /open-office-artifact-tool/);
  assert.match(skillText, /slides_test\.py/);
  assert.match(skillText, /slide\.setBackground.*slide\.clearBackground/s);
  assert.match(skillText, /artifact_tool\/api\/references\/comments\.md/);
  const commentsReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/comments.md", "utf8");
  assert.match(commentsReferenceText, /Pass `undefined` as the target/);
  assert.match(commentsReferenceText, /one author, one text item, and one explicit\s+slide coordinate/is);
  assert.match(commentsReferenceText, /Modern threaded-comment graphs remain opaque and source-bound/);
  const slideReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/slide.spec.md", "utf8");
  assert.match(slideReferenceText, /never flattens the\s+inherited color/i);
  assert.match(slideReferenceText, /Gradient,\s+pattern, image.*opaque-preserved/is);
  const imageReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/images.spec.md", "utf8");
  assert.match(imageReferenceText, /signed normalized source edges in `-1\.\.1`/i);
  assert.match(imageReferenceText, /DrawingML `a:srcRect`/);
  assert.match(imageReferenceText, /PPTX has no native fit keyword/i);
  assert.match(imageReferenceText, /unsafe edits fail\s+closed/i);
  const groupingReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/grouping.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/grouping\.spec\.md/);
  assert.match(groupingReferenceText, /real `p:grpSp`/);
  assert.match(groupingReferenceText, /a:chOff\/a:chExt/);
  assert.match(groupingReferenceText, /presentation_group_topology_changed/);
  assert.match(groupingReferenceText, /one opaque, read-only native object/i);
  const oleWorkbookReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/ole-workbooks.spec.md", "utf8");
  assert.match(skillText, /artifact_tool\/api\/references\/ole-workbooks\.spec\.md/);
  assert.match(oleWorkbookReferenceText, /getEmbeddedWorkbook\(\).*replaceEmbeddedWorkbook/s);
  assert.match(oleWorkbookReferenceText, /preserving the OLE\s+shell, relationship topology, preview image/is);
  assert.match(oleWorkbookReferenceText, /Microsoft Open XML\s+SDK/);
  assert.match(oleWorkbookReferenceText, /shared, external, ambiguous, or non-XLSX/);
  assert.match(oleWorkbookReferenceText, /no lossy reconstruction or silent fallback/i);

  console.log("presentation skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
