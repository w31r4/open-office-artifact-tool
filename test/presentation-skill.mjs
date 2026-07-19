import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";
import {
  generateOfficeInput,
  PPTX_TITLE_NOTES_FIXTURE,
} from "../scripts/agent-eval-office-fixtures.mjs";
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
  assert.match(commentXml, /<p:cm[^>]*authorId="0"[^>]*idx="1"/);
  assert.match(commentXml, /Confirm the source before delivery\./);

  const modernCommentsDir = path.join(root, "modern-comments-workflow");
  const modernCommentsOutput = path.join(modernCommentsDir, "decision-review.pptx");
  const modernCommentsAudit = path.join(modernCommentsDir, "audit.json");
  const { createAndEditModernCommentThread } = await import(
    "../skills/presentations/skills/presentations/examples/openchestnut-modern-comment-workflow.mjs"
  );
  const modernCommentsResult = await createAndEditModernCommentThread({
    outputPath: modernCommentsOutput,
    auditPath: modernCommentsAudit,
  });
  assert.equal(modernCommentsResult.audit.provider.actual, "open-chestnut");
  assert.equal(modernCommentsResult.audit.provider.silentFallback, false);
  assert.equal(modernCommentsResult.audit.operation.type, "fixed-topology-modern-comment-text-status-edit");
  assert.equal(modernCommentsResult.audit.operation.replyCount, 1);
  assert.equal(modernCommentsResult.audit.validation.fixedIdentityPreserved, true);
  assert.equal(modernCommentsResult.audit.validation.package.ok, true);
  assert.equal(modernCommentsResult.audit.validation.modelRender.ok, true);
  const modernCommentsRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(modernCommentsOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "decision-review.pptx",
  }));
  assert.equal(modernCommentsRoundTrip.commentFormat, "modern");
  const modernThread = modernCommentsRoundTrip.slides.getItem(0).comments.items[0];
  assert.equal(modernThread.comments[0].text, "Customer evidence confirmed for delivery.");
  assert.equal(modernThread.comments[1].text, "Recorded in the decision log.");
  assert.equal(modernThread.comments[0].author, "Review Owner");
  assert.equal(modernThread.comments[1].author, "Evidence Owner");
  assert.equal(modernThread.resolved, true);
  assert.equal(modernThread.nativeAnchor.type, "textRange");
  assert.match(modernCommentsRoundTrip.inspect({ kind: "comment" }).ndjson, /Customer evidence confirmed/);
  const modernCommentsZip = await JSZip.loadAsync(await fs.readFile(modernCommentsOutput));
  assert.ok(Object.keys(modernCommentsZip.files).some((name) => /^ppt\/comments\/modernComment\d*\.xml$/.test(name)));
  assert.ok(modernCommentsZip.file("ppt/authors.xml"));

  const modernCommentsCliOutput = path.join(modernCommentsDir, "decision-review-cli.pptx");
  const modernCommentsCliAudit = path.join(modernCommentsDir, "cli-audit.json");
  const modernCommentsCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/openchestnut-modern-comment-workflow.mjs",
    modernCommentsCliOutput,
    modernCommentsCliAudit,
  ], { encoding: "utf8" });
  assert.equal(modernCommentsCli.status, 0, `modern-comment CLI failed\n${modernCommentsCli.stdout}\n${modernCommentsCli.stderr}`);
  assert.equal(JSON.parse(modernCommentsCli.stdout).threadId, "{11111111-1111-4111-8111-111111111111}");

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

  const titleNotesDir = path.join(root, "title-notes-workflow");
  const titleNotesInput = path.join(titleNotesDir, PPTX_TITLE_NOTES_FIXTURE.presentationName);
  const titleNotesOutput = path.join(titleNotesDir, "launch-review-updated.pptx");
  const titleNotesAudit = path.join(titleNotesDir, "audit.json");
  await generateOfficeInput("pptx-title-notes-review", titleNotesInput);
  const titleNotesSource = await fs.readFile(titleNotesInput);
  const { editPptxTitleAndNotes } = await import(
    "../skills/presentations/skills/presentations/examples/openchestnut-title-notes-edit-workflow.mjs"
  );
  const titleNotesResult = await editPptxTitleAndNotes({
    inputPath: titleNotesInput,
    outputPath: titleNotesOutput,
    auditPath: titleNotesAudit,
    slideName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    titleShapeName: PPTX_TITLE_NOTES_FIXTURE.titleShapeName,
    expectedTitle: PPTX_TITLE_NOTES_FIXTURE.originalTitle,
    replacementTitle: PPTX_TITLE_NOTES_FIXTURE.replacementTitle,
    expectedNotes: PPTX_TITLE_NOTES_FIXTURE.originalNotes,
    replacementNotes: PPTX_TITLE_NOTES_FIXTURE.replacementNotes,
  });
  assert.equal(titleNotesResult.audit.provider.actual, "open-chestnut");
  assert.equal(titleNotesResult.audit.validation.reimport.ok, true);
  assert.equal(titleNotesResult.audit.validation.modelRender.renderer, "model-svg");
  assert.deepEqual(await fs.readFile(titleNotesInput), titleNotesSource);
  const titleNotesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(titleNotesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "launch-review-updated.pptx",
  }));
  assert.deepEqual(titleNotesRoundTrip.slides.items.map((slide) => slide.name), [
    PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
  ]);
  const titleNotesSlide = titleNotesRoundTrip.slides.getItem(0);
  assert.equal(itemByName(titleNotesSlide.shapes.items, PPTX_TITLE_NOTES_FIXTURE.titleShapeName).text.value, PPTX_TITLE_NOTES_FIXTURE.replacementTitle);
  assert.equal(titleNotesSlide.speakerNotes.text, PPTX_TITLE_NOTES_FIXTURE.replacementNotes);
  assert.deepEqual(titleNotesSlide.background, { fill: "#f1f5f9", mode: "solid" });

  const slideNameDir = path.join(root, "slide-name-workflow");
  const slideNameInput = path.join(slideNameDir, PPTX_TITLE_NOTES_FIXTURE.presentationName);
  const slideNameOutput = path.join(slideNameDir, "launch-review-renamed.pptx");
  const slideNameAudit = path.join(slideNameDir, "audit.json");
  await generateOfficeInput("pptx-title-notes-review", slideNameInput);
  const slideNameSource = await fs.readFile(slideNameInput);
  const { editPptxSlideName } = await import(
    "../skills/presentations/skills/presentations/examples/openchestnut-slide-name-edit-workflow.mjs"
  );
  const slideNameResult = await editPptxSlideName({
    inputPath: slideNameInput,
    outputPath: slideNameOutput,
    auditPath: slideNameAudit,
    expectedName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    replacementName: "Go decision: controlled rollout",
  });
  assert.equal(slideNameResult.audit.provider.actual, "open-chestnut");
  assert.equal(slideNameResult.audit.operation.nativeAttribute, "p:cSld/@name");
  assert.equal(slideNameResult.audit.validation.package.targetNameVerified, true);
  assert.equal(slideNameResult.audit.validation.package.targetPartMayBeCanonicalized, true);
  assert.equal(slideNameResult.audit.validation.package.nonTargetPartsByteIdentical, true);
  assert.equal(slideNameResult.audit.validation.modelRender.byteIdentical, true);
  assert.deepEqual(await fs.readFile(slideNameInput), slideNameSource);
  const slideNameRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(slideNameOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "launch-review-renamed.pptx",
  }));
  assert.deepEqual(slideNameRoundTrip.slides.items.map((slide) => slide.name), [
    "Go decision: controlled rollout",
    PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
  ]);
  const slideNameTarget = slideNameRoundTrip.slides.getItem(0);
  assert.equal(itemByName(slideNameTarget.shapes.items, PPTX_TITLE_NOTES_FIXTURE.titleShapeName).text.value, PPTX_TITLE_NOTES_FIXTURE.originalTitle);
  assert.equal(slideNameTarget.speakerNotes.text, PPTX_TITLE_NOTES_FIXTURE.originalNotes);
  const missingOutput = path.join(slideNameDir, "should-not-exist.pptx");
  const missingAudit = path.join(slideNameDir, "should-not-exist.json");
  await assert.rejects(
    () => editPptxSlideName({
      inputPath: slideNameInput,
      outputPath: missingOutput,
      auditPath: missingAudit,
      expectedName: "Missing source slide",
      replacementName: "Never write this",
    }),
    /Expected exactly one imported slide named/,
  );
  assert.equal(await fs.access(missingOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(missingAudit).then(() => true, () => false), false);
  const slideNameCliOutput = path.join(slideNameDir, "launch-review-cli-renamed.pptx");
  const slideNameCliAudit = path.join(slideNameDir, "cli-audit.json");
  const slideNameCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/openchestnut-slide-name-edit-workflow.mjs",
    slideNameInput,
    slideNameCliOutput,
    slideNameCliAudit,
    PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
    "Go decision: CLI rollout",
  ], { encoding: "utf8" });
  assert.equal(slideNameCli.status, 0, `slide-name CLI failed\n${slideNameCli.stdout}\n${slideNameCli.stderr}`);
  assert.equal(JSON.parse(slideNameCli.stdout).sourcePart, "ppt/slides/slide1.xml");

  const customShowDir = path.join(root, "custom-show-workflow");
  const customShowInput = path.join(customShowDir, "custom-show-source.pptx");
  const customShowOutput = path.join(customShowDir, "custom-show-updated.pptx");
  const customShowAudit = path.join(customShowDir, "audit.json");
  await fs.mkdir(customShowDir, { recursive: true });
  const customShowFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const customShowSlides = [
    ["Overview", "#DBEAFE"],
    ["Evidence", "#DCFCE7"],
    ["Appendix", "#FEF3C7"],
  ].map(([name, fill], index) => {
    const slide = customShowFixture.slides.add({ name, background: { fill } });
    slide.shapes.add({ name: `${name.toLowerCase()}-title`, position: { left: 44, top: 44, width: 552, height: 72 }, text: `${index + 1}. ${name}`, fill: "#FFFFFF", line: { fill: "#0F172A", width: 1 } });
    return slide;
  });
  customShowFixture.customShows.add({ name: "Board route", nativeId: 7, slides: [customShowSlides[0], customShowSlides[2]] });
  customShowFixture.customShows.add({ name: "Review route", nativeId: 11, slides: [customShowSlides[1]] });
  await (await PresentationFile.exportPptx(customShowFixture)).save(customShowInput);
  const customShowSource = await fs.readFile(customShowInput);
  const { editPptxCustomShow } = await import(
    "../skills/presentations/skills/presentations/examples/openchestnut-custom-show-workflow.mjs"
  );
  const customShowResult = await editPptxCustomShow({
    inputPath: customShowInput,
    outputPath: customShowOutput,
    auditPath: customShowAudit,
    expectedName: "Board route",
    replacementName: "Executive route",
    orderedSlideNames: ["Appendix", "Overview", "Appendix"],
  });
  assert.equal(customShowResult.audit.provider.actual, "open-chestnut");
  assert.equal(customShowResult.audit.operation.nativeId, 7);
  assert.equal(customShowResult.audit.validation.package.onlyPresentationPartChanged, true);
  assert.equal(customShowResult.audit.validation.package.nativeIdPreserved, true);
  assert.equal(customShowResult.audit.validation.modelRender.slidesByteIdentical, true);
  assert.deepEqual(await fs.readFile(customShowInput), customShowSource);
  const customShowRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(customShowOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "custom-show-updated.pptx",
  }));
  assert.deepEqual(customShowRoundTrip.customShows.items.map((show) => [show.name, show.nativeId]), [["Executive route", 7], ["Review route", 11]]);
  assert.deepEqual(customShowRoundTrip.customShows.getItem("Executive route").slides.map((slide) => slide.name), ["Appendix", "Overview", "Appendix"]);
  const customShowBaselineDir = path.join(customShowDir, "baselines");
  await verifyPresentationFile(customShowInput, {
    outputDir: path.join(customShowDir, "source-qa"),
    nativeRender,
    baselineDir: customShowBaselineDir,
    writeBaseline: true,
  });
  const customShowReview = await verifyPresentationFile(customShowOutput, {
    outputDir: path.join(customShowDir, "output-qa"),
    nativeRender,
    baselineDir: customShowBaselineDir,
  });
  assert.ok(customShowReview.modelRender.slides.every((slide) => slide.pixelDiff?.changed === false));
  if (nativeStatus.available) assert.ok(customShowReview.nativeRender.pages.every((page) => page.pixelDiff?.changed === false));
  const customShowCliOutput = path.join(customShowDir, "custom-show-cli.pptx");
  const customShowCliAudit = path.join(customShowDir, "cli-audit.json");
  const customShowCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/openchestnut-custom-show-workflow.mjs",
    customShowInput,
    customShowCliOutput,
    customShowCliAudit,
    "Board route",
    "CLI route",
    "Evidence,Overview",
  ], { encoding: "utf8" });
  assert.equal(customShowCli.status, 0, `custom-show CLI failed\n${customShowCli.stdout}\n${customShowCli.stderr}`);
  assert.equal(JSON.parse(customShowCli.stdout).outputPath, path.resolve(customShowCliOutput));

  const duplicateDir = path.join(root, "slide-duplicate-workflow");
  const duplicateInput = path.join(duplicateDir, "connector-source.pptx");
  const duplicateOutput = path.join(duplicateDir, "connector-duplicate.pptx");
  const duplicateAudit = path.join(duplicateDir, "audit.json");
  await fs.mkdir(duplicateDir, { recursive: true });
  const duplicateFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const duplicateSourceSlide = duplicateFixture.slides.add({ name: "Clone connector source" });
  const duplicateGroup = duplicateSourceSlide.addGroup({
    name: "connector-cluster",
    position: { left: 48, top: 40, width: 320, height: 120 },
    childFrame: { left: 0, top: 0, width: 320, height: 120 },
  });
  const duplicateLeft = duplicateGroup.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
  const duplicateRight = duplicateGroup.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
  duplicateGroup.connectors.add({
    name: "join",
    from: duplicateLeft,
    to: duplicateRight,
    start: { x: 90, y: 41 },
    end: { x: 210, y: 41 },
    line: { fill: "#64748B", width: 1 },
  });
  const duplicateCanary = duplicateFixture.slides.add({ name: "Untouched canary" });
  duplicateCanary.shapes.add({
    name: "canary-copy",
    position: { left: 48, top: 40, width: 260, height: 72 },
    text: "Unchanged source slide",
  });
  duplicateSourceSlide.shapes.add({
    name: "clone-links",
    geometry: "textbox",
    position: { left: 48, top: 190, width: 480, height: 72 },
    fill: "transparent",
    line: { fill: "transparent", width: 0 },
    text: [{ runs: [
      { text: "Guide ", link: { uri: "https://example.com/clone-guide", tooltip: "Open guide" } },
      { text: "Canary ", link: { slideId: duplicateCanary.id } },
      { text: "Next", link: { action: "nextSlide" } },
    ] }],
  });
  duplicateSourceSlide.charts.add("bar", {
    name: "clone-pipeline-chart",
    position: { left: 390, top: 36, width: 218, height: 132 },
    title: "Pipeline",
    categories: ["Now", "Next"],
    series: [{ name: "Value", values: [42, 57], fill: "#2563EB" }],
    axes: { category: { title: "Stage" }, value: { title: "Value", min: 0, max: 80, majorUnit: 20 } },
    legend: false,
  });
  const duplicateSourcePptx = await PresentationFile.exportPptx(duplicateFixture);
  await duplicateSourcePptx.save(duplicateInput);
  const duplicateSourceBytes = await fs.readFile(duplicateInput);
  const { duplicatePptxSlide } = await import(
    "../skills/presentations/skills/presentations/examples/openchestnut-slide-duplicate-workflow.mjs"
  );
  const duplicateResult = await duplicatePptxSlide({
    inputPath: duplicateInput,
    outputPath: duplicateOutput,
    auditPath: duplicateAudit,
    expectedName: "Clone connector source",
  });
  assert.equal(duplicateResult.audit.provider.actual, "open-chestnut");
  assert.equal(duplicateResult.audit.operation.type, "source-bound-slide-duplicate");
  assert.equal(duplicateResult.audit.operation.sourcePart, "ppt/slides/slide1.xml");
  assert.equal(duplicateResult.audit.operation.clonePart, "ppt/slides/slide3.xml");
  assert.equal(duplicateResult.audit.operation.scope, "canonical-inline-leaves-with-closed-chart-leaves");
  assert.equal(duplicateResult.audit.validation.package.retainedSourcePartsByteIdentical, true);
  assert.deepEqual(duplicateResult.audit.operation.runHyperlinks, { relationshipCount: 2, actionOnlyCount: 1 });
  assert.deepEqual(duplicateResult.audit.validation.package.runHyperlinks, {
    relationshipCount: 2,
    actionOnlyCount: 1,
    exactSourceGraphRetained: true,
  });
  assert.equal(duplicateResult.audit.operation.chartParts.count, 1);
  assert.equal(duplicateResult.audit.validation.package.chartParts.count, 1);
  assert.equal(duplicateResult.audit.validation.package.chartParts.independentParts, true);
  assert.equal(duplicateResult.audit.validation.package.chartParts.allPayloadsByteIdentical, true);
  const duplicateChartAudit = duplicateResult.audit.validation.package.chartParts.parts[0];
  assert.match(duplicateChartAudit.sourcePart, /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i);
  assert.match(duplicateChartAudit.clonePart, /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i);
  assert.notEqual(duplicateChartAudit.sourcePart, duplicateChartAudit.clonePart);
  assert.deepEqual(duplicateResult.audit.validation.package.newPartPaths, [
    "ppt/slides/_rels/slide3.xml.rels",
    duplicateChartAudit.clonePart,
    "ppt/slides/slide3.xml",
  ].sort());
  assert.equal(duplicateResult.audit.validation.reimport.sourceAndCloneSemanticsEqual, true);
  assert.equal(duplicateResult.audit.validation.modelRender.visualEquivalent, true);
  assert.equal(duplicateResult.audit.validation.modelRender.identityAttributesIgnored, true);
  assert.deepEqual(await fs.readFile(duplicateInput), duplicateSourceBytes);
  const duplicateRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(duplicateOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "connector-duplicate.pptx",
  }));
  assert.deepEqual(duplicateRoundTrip.slides.items.map((slide) => slide.name), [
    "Clone connector source",
    "Clone connector source",
    "Untouched canary",
  ]);
  const duplicateSourceGroup = duplicateRoundTrip.slides.getItem(0).groups.items[0];
  const duplicateCloneGroup = duplicateRoundTrip.slides.getItem(1).groups.items[0];
  const duplicateCloneConnector = duplicateCloneGroup.connectors.items[0];
  assert.notEqual(duplicateCloneGroup.id, duplicateSourceGroup.id);
  assert.equal(duplicateCloneConnector.startTargetId, duplicateCloneGroup.shapes.items[0].id);
  assert.equal(duplicateCloneConnector.endTargetId, duplicateCloneGroup.shapes.items[1].id);
  assert.notEqual(duplicateCloneConnector.startTargetId, duplicateSourceGroup.shapes.items[0].id);
  const duplicateCloneLinks = duplicateRoundTrip.slides.getItem(1).shapes.items.find((shape) => shape.name === "clone-links").text.paragraphs[0].runs;
  assert.equal(duplicateCloneLinks[0].link.uri, "https://example.com/clone-guide");
  assert.equal(duplicateCloneLinks[1].link.slideId, duplicateRoundTrip.slides.getItem(2).id);
  assert.equal(duplicateCloneLinks[2].link.action, "nextSlide");
  const duplicateSourceChart = duplicateRoundTrip.slides.getItem(0).charts.items[0];
  const duplicateCloneChart = duplicateRoundTrip.slides.getItem(1).charts.items[0];
  assert.notEqual(duplicateCloneChart.id, duplicateSourceChart.id);
  assert.equal(duplicateSourceChart.title, "Pipeline");
  assert.equal(duplicateCloneChart.title, "Pipeline");
  assert.deepEqual(duplicateCloneChart.series[0].values, [42, 57]);
  const duplicateQa = await verifyPresentationFile(duplicateOutput, {
    outputDir: path.join(duplicateDir, "render-qa"),
    nativeRender,
  });
  assert.equal(duplicateQa.verify.ok, true);
  assert.equal(duplicateQa.packageInspect.ok, true);
  assert.equal(duplicateQa.modelRender.ok, true);
  if (nativeStatus.available) {
    assert.equal(duplicateQa.nativeRender.status, "passed");
    assert.deepEqual(
      await fs.readFile(duplicateQa.nativeRender.pages[0].path),
      await fs.readFile(duplicateQa.nativeRender.pages[1].path),
      "LibreOffice/Poppler must render the source and canonical hyperlink clone identically",
    );
  }

  duplicateCloneChart.title = "Updated clone pipeline";
  duplicateCloneChart.series[0].values[1] = 63;
  const duplicateEdited = await PresentationFile.exportPptx(duplicateRoundTrip);
  const duplicateEditedRoundTrip = await PresentationFile.importPptx(duplicateEdited);
  assert.equal(duplicateEditedRoundTrip.slides.getItem(0).charts.items[0].title, "Pipeline");
  assert.equal(duplicateEditedRoundTrip.slides.getItem(0).charts.items[0].series[0].values[1], 57);
  assert.equal(duplicateEditedRoundTrip.slides.getItem(1).charts.items[0].title, "Updated clone pipeline");
  assert.equal(duplicateEditedRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1], 63);

  const duplicateMissingOutput = path.join(duplicateDir, "missing-target.pptx");
  const duplicateMissingAudit = path.join(duplicateDir, "missing-target.json");
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: duplicateInput,
      outputPath: duplicateMissingOutput,
      auditPath: duplicateMissingAudit,
      expectedName: "Missing source slide",
    }),
    /Expected exactly one imported source slide named/,
  );
  assert.equal(await fs.access(duplicateMissingOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(duplicateMissingAudit).then(() => true, () => false), false);

  const orphanLinkInput = path.join(duplicateDir, "orphan-link-source.pptx");
  const orphanLinkOutput = path.join(duplicateDir, "orphan-link-output.pptx");
  const orphanLinkAudit = path.join(duplicateDir, "orphan-link-audit.json");
  const orphanLinkZip = await JSZip.loadAsync(duplicateSourceBytes);
  const orphanRelationships = await orphanLinkZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
  orphanLinkZip.file(
    "ppt/slides/_rels/slide1.xml.rels",
    orphanRelationships.replace(
      "</Relationships>",
      '<Relationship Id="rIdOrphanLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.com/orphan" TargetMode="External"/></Relationships>',
    ),
  );
  await fs.writeFile(orphanLinkInput, await orphanLinkZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: orphanLinkInput,
      outputPath: orphanLinkOutput,
      auditPath: orphanLinkAudit,
      expectedName: "Clone connector source",
    }),
    /orphan or unmodeled hyperlink relationship/,
  );
  assert.equal(await fs.access(orphanLinkOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(orphanLinkAudit).then(() => true, () => false), false);

  const notesInput = path.join(duplicateDir, "notes-source.pptx");
  const notesOutput = path.join(duplicateDir, "notes-output.pptx");
  const notesAudit = path.join(duplicateDir, "notes-audit.json");
  const notesFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  notesFixture.slides.add({ name: "Source with notes", notes: "This workflow must refuse notes." }).shapes.add({
    name: "notes-copy",
    position: { left: 48, top: 40, width: 260, height: 72 },
    text: "Source with notes",
  });
  await (await PresentationFile.exportPptx(notesFixture)).save(notesInput);
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: notesInput,
      outputPath: notesOutput,
      auditPath: notesAudit,
      expectedName: "Source with notes",
    }),
    /accepts no speaker-notes leaf/,
  );
  assert.equal(await fs.access(notesOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(notesAudit).then(() => true, () => false), false);

  const closedLeavesInput = path.join(duplicateDir, "closed-leaves-source.pptx");
  const closedLeavesOutput = path.join(duplicateDir, "closed-leaves-output.pptx");
  const closedLeavesAudit = path.join(duplicateDir, "closed-leaves-audit.json");
  const closedLeavesFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
  const closedLeavesSource = closedLeavesFixture.slides.add({
    name: "Closed leaves source",
    notes: "Open with the decision.\nClose with the owner.",
  });
  closedLeavesSource.shapes.add({
    name: "closed-leaves-title",
    position: { left: 48, top: 40, width: 300, height: 72 },
    text: "Closed leaves source",
  });
  closedLeavesSource.comments.addThread(undefined, "Confirm the original evidence before delivery.", {
    author: "Review Owner",
    created: "2026-07-18T03:05:00Z",
    position: { x: 360, y: 240 },
  });
  await (await PresentationFile.exportPptx(closedLeavesFixture)).save(closedLeavesInput);
  const closedLeavesSourceBytes = await fs.readFile(closedLeavesInput);
  const closedLeavesResult = await duplicatePptxSlide({
    inputPath: closedLeavesInput,
    outputPath: closedLeavesOutput,
    auditPath: closedLeavesAudit,
    expectedName: "Closed leaves source",
    allowClosedLeaves: true,
  });
  assert.equal(closedLeavesResult.audit.operation.scope, "canonical-inline-leaves-with-closed-relationship-leaves");
  assert.deepEqual(closedLeavesResult.audit.operation.closedLeaves, { speakerNotes: true, legacyComments: true });
  assert.deepEqual(closedLeavesResult.audit.validation.package.newPartPaths, [
    "ppt/comments/comment2.xml",
    "ppt/notesSlides/_rels/notesSlide2.xml.rels",
    "ppt/notesSlides/notesSlide2.xml",
    "ppt/slides/_rels/slide2.xml.rels",
    "ppt/slides/slide2.xml",
  ]);
  assert.equal(closedLeavesResult.audit.validation.package.retainedSourcePartsByteIdentical, true);
  assert.deepEqual(closedLeavesResult.audit.validation.package.closedLeaves.speakerNotes, {
    sourcePart: "ppt/notesSlides/notesSlide1.xml",
    clonePart: "ppt/notesSlides/notesSlide2.xml",
    sourceRelationshipPart: "ppt/notesSlides/_rels/notesSlide1.xml.rels",
    cloneRelationshipPart: "ppt/notesSlides/_rels/notesSlide2.xml.rels",
    notesMasterPart: "ppt/notesMasters/notesMaster1.xml",
    notesXmlByteIdentical: true,
    notesMasterShared: true,
    cloneBackReferencePointsAtClone: true,
  });
  assert.deepEqual(closedLeavesResult.audit.validation.package.closedLeaves.legacyComments, {
    sourcePart: "ppt/comments/comment1.xml",
    clonePart: "ppt/comments/comment2.xml",
    commentAuthorsPart: "ppt/commentAuthors.xml",
    commentsXmlByteIdentical: true,
    commentAuthorsShared: true,
  });
  assert.equal(closedLeavesResult.audit.validation.reimport.sourceAndCloneClosedLeavesEqual, true);
  assert.deepEqual(await fs.readFile(closedLeavesInput), closedLeavesSourceBytes);
  const closedLeavesRoundTrip = await PresentationFile.importPptx(new FileBlob(await fs.readFile(closedLeavesOutput), {
    type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    name: "closed-leaves-output.pptx",
  }));
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.name), ["Closed leaves source", "Closed leaves source"]);
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.speakerNotes.text), [
    "Open with the decision.\nClose with the owner.",
    "Open with the decision.\nClose with the owner.",
  ]);
  assert.deepEqual(closedLeavesRoundTrip.slides.items.map((slide) => slide.comments.items[0].comments[0].text), [
    "Confirm the original evidence before delivery.",
    "Confirm the original evidence before delivery.",
  ]);
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: closedLeavesInput,
      outputPath: path.join(duplicateDir, "invalid-option.pptx"),
      auditPath: path.join(duplicateDir, "invalid-option.json"),
      expectedName: "Closed leaves source",
      allowClosedLeaves: "yes",
    }),
    /allowClosedLeaves must be a boolean/,
  );

  const irregularLeavesInput = path.join(duplicateDir, "irregular-closed-leaves-source.pptx");
  const irregularLeavesOutput = path.join(duplicateDir, "irregular-closed-leaves-output.pptx");
  const irregularLeavesAudit = path.join(duplicateDir, "irregular-closed-leaves-audit.json");
  const irregularLeavesZip = await JSZip.loadAsync(closedLeavesSourceBytes);
  const sourceNotesRelationships = await irregularLeavesZip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels").async("text");
  irregularLeavesZip.file(
    "ppt/notesSlides/_rels/notesSlide1.xml.rels",
    sourceNotesRelationships.replace(
      "</Relationships>",
      '<Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="/ppt/customXml/item1.xml" Id="rIdUnexpected" /></Relationships>',
    ),
  );
  irregularLeavesZip.file("ppt/customXml/item1.xml", "<?xml version=\"1.0\" encoding=\"UTF-8\"?><openchestnut:unexpected xmlns:openchestnut=\"https://openchestnut.invalid\" />");
  await fs.writeFile(irregularLeavesInput, await irregularLeavesZip.generateAsync({ type: "nodebuffer" }));
  await assert.rejects(
    () => duplicatePptxSlide({
      inputPath: irregularLeavesInput,
      outputPath: irregularLeavesOutput,
      auditPath: irregularLeavesAudit,
      expectedName: "Closed leaves source",
      allowClosedLeaves: true,
    }),
    /exactly notesMaster and slide relationships/,
  );
  assert.equal(await fs.access(irregularLeavesOutput).then(() => true, () => false), false);
  assert.equal(await fs.access(irregularLeavesAudit).then(() => true, () => false), false);

  const duplicateCliOutput = path.join(duplicateDir, "connector-duplicate-cli.pptx");
  const duplicateCliAudit = path.join(duplicateDir, "cli-audit.json");
  const duplicateCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/openchestnut-slide-duplicate-workflow.mjs",
    duplicateInput,
    duplicateCliOutput,
    duplicateCliAudit,
    "Clone connector source",
  ], { encoding: "utf8" });
  assert.equal(duplicateCli.status, 0, `slide-duplicate CLI failed\n${duplicateCli.stdout}\n${duplicateCli.stderr}`);
  assert.equal(JSON.parse(duplicateCli.stdout).clonePart, "ppt/slides/slide3.xml");

  const closedLeavesCliOutput = path.join(duplicateDir, "closed-leaves-cli-output.pptx");
  const closedLeavesCliAudit = path.join(duplicateDir, "closed-leaves-cli-audit.json");
  const closedLeavesCli = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/examples/openchestnut-slide-duplicate-workflow.mjs",
    closedLeavesInput,
    closedLeavesCliOutput,
    closedLeavesCliAudit,
    "Closed leaves source",
    "--allow-closed-leaves",
  ], { encoding: "utf8" });
  assert.equal(closedLeavesCli.status, 0, `closed-leaves slide-duplicate CLI failed\n${closedLeavesCli.stdout}\n${closedLeavesCli.stderr}`);
  assert.deepEqual(JSON.parse(closedLeavesCli.stdout).closedLeaves, { speakerNotes: true, legacyComments: true });

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
  const starterRoot = path.join(root, "starter-fail-closed");
  const starterMap = path.join(starterRoot, "template-frame-map.json");
  const starterOutput = path.join(starterRoot, "template-starter.pptx");
  await fs.mkdir(starterRoot, { recursive: true });
  await fs.writeFile(starterMap, JSON.stringify({ outputSlides: [] }));
  const starterResult = spawnSync(process.execPath, [
    "skills/presentations/skills/presentations/template_following_scripts/prepare_template_starter_deck.mjs",
    "--workspace", starterRoot,
    "--pptx", readiness.pptxPath,
    "--map", starterMap,
    "--out", starterOutput,
  ], { encoding: "utf8" });
  assert.notEqual(starterResult.status, 0);
  assert.match(starterResult.stderr, /broad graph deletion[\s\S]*No output was written/);
  assert.deepEqual((await fs.readdir(starterRoot)).sort(), ["template-frame-map.json"]);
  assert.match(skillText, /open-office-artifact-tool/);
  assert.match(skillText, /openchestnut-title-notes-edit-workflow\.mjs/);
  assert.match(skillText, /openchestnut-modern-comment-workflow\.mjs/);
  assert.match(skillText, /openchestnut-slide-name-edit-workflow\.mjs/);
  assert.match(skillText, /openchestnut-slide-duplicate-workflow\.mjs/);
  assert.match(skillText, /--allow-closed-leaves/);
  assert.match(quickStartText, /PresentationFile\.exportPptx/);
  assert.match(quickStartText, /editPptxSlideName/);
  assert.match(quickStartText, /duplicatePptxSlide/);
  assert.match(quickStartText, /allowClosedLeaves:\s*true/);
  assert.match(quickStartText, /commentFormat:\s*"modern"/);
  assert.match(quickStartText, /open-office-artifact-tool/);
  assert.match(skillText, /slides_test\.py/);
  assert.match(skillText, /slide\.setBackground.*slide\.clearBackground/s);
  assert.match(skillText, /slide\.moveTo\(existingZeroBasedIndex\).*retained source.*p:sldIdLst.*slide\.delete\(\).*isolated.*layout relationship/is);
  assert.match(skillText, /starter-deck command below still needs a\s+broad imported-slide graph clone and broad graph delete semantics/is);
  assert.match(skillText, /slide\.duplicate\(\).*canonical shapes.*canonical inline fixed-grid tables.*recognized closed\s+literal-data charts.*canonical embedded rectangular images.*bounded canonical\s+straight\/elbow connectors.*new `SlidePart`.*every present\s+connector endpoint.*same copied `SlidePart`.*export plus reimport/is);
  assert.match(skillText, /recognized closed\s+literal-data charts.*unique internal relationship.*numbered `ChartPart`.*byte-copies.*distinct clone-local ChartPart.*ChartParts are independent.*advertises the ordinary fixed-topology\s+edit capability/is);
  assert.match(quickStartText, /recognized literal-data charts.*no child\/external\/hyperlink\/data relationship.*distinct byte-copied ChartPart/is);
  assert.match(skillText, /NotesSlide.*NotesMaster.*byte-for-byte.*back-reference.*clone/is);
  assert.match(skillText, /SlideCommentsPart.*CommentAuthorsPart.*byte-for-byte/is);
  assert.match(skillText, /artifact_tool\/api\/references\/comments\.md/);
  const commentsReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/comments.md", "utf8");
  assert.match(commentsReferenceText, /Pass `undefined` as the target/);
  assert.match(commentsReferenceText, /one author, one text item, and one explicit\s+slide coordinate/is);
  assert.match(commentsReferenceText, /Office 2021 modern threads/);
  assert.match(commentsReferenceText, /Only existing comment text and status are mutable/);
  assert.match(commentsReferenceText, /Reactions\/likes, task fields, extensions, rich text, nested replies/);
  assert.match(commentsReferenceText, /openchestnut-modern-comment-workflow\.mjs/);
  assert.match(commentsReferenceText, /slide\.duplicate\(\).*byte-cop(?:y|ied).*author\s+catalog/is);
  const slideReferenceText = await fs.readFile("skills/presentations/skills/presentations/artifact_tool/api/references/slide.spec.md", "utf8");
  assert.match(slideReferenceText, /never flattens the\s+inherited color/i);
  assert.match(slideReferenceText, /NotesSlide.*NotesMaster.*exactly those two\s+relationships.*byte-for-byte/is);
  assert.match(slideReferenceText, /canonical inline fixed-grid tables[\s\S]*cannot introduce a fill, link/i);
  assert.match(slideReferenceText, /recognized closed literal-data charts.*numbered\s+`ChartPart`.*no child, external, hyperlink, or data relationship.*distinct clone-local ChartPart/is);
  assert.match(slideReferenceText, /Gradient,\s+pattern, image.*opaque-preserved/is);
  assert.match(slideReferenceText, /p:cSld\/@name.*export\/reimport/is);
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
  const templateFollowingText = await fs.readFile("skills/presentations/skills/presentations/references/template-following.md", "utf8");
  assert.match(templateFollowingText, /source-preserving reordering.*isolated[\s>]+layout-only.*slide\.delete/is);
  assert.match(templateFollowingText, /broader OPC graph-clone milestone is unavailable/i);
  assert.match(templateFollowingText, /read-only path\/input preflight[\s\S]*then fails closed[\s\S]*not kept as dead\s+code/i);

  console.log("presentation skill smoke ok");
} finally {
  await fs.rm(root, { recursive: true, force: true });
}
