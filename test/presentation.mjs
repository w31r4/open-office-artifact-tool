import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { box, column, FileBlob, paragraph, Presentation, PresentationFile, row, run, rule, shape as composeShape } from "../src/index.mjs";

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
presentation.theme.setColors({ accent1: "#0ea5e9", accent2: "#f97316" }).setFonts({ major: "Aptos Display", minor: "Aptos" });
const titleLayout = presentation.layouts.add({
  id: "layout/title-content",
  name: "Title and Content",
  type: "titleAndContent",
  placeholders: [
    { id: "ph/title", type: "title", name: "Title Placeholder", position: { left: 72, top: 36, width: 720, height: 52 }, required: true, style: { fontSize: 28, bold: true } },
    { id: "ph/body", type: "body", name: "Body Placeholder", position: { left: 72, top: 520, width: 560, height: 80 } },
  ],
});
const slide = presentation.slides.add();
const [titlePlaceholder] = slide.applyLayout(titleLayout);
titlePlaceholder.text = "Quarterly plan template";
const themeLayoutSnapshot = presentation.inspect({ kind: "theme,layout,textbox", maxChars: 8000 }).ndjson;
assert.match(themeLayoutSnapshot, /Open Office Clean Room/);
assert.match(themeLayoutSnapshot, /Title and Content/);
assert.match(themeLayoutSnapshot, /Quarterly plan template/);
assert.equal(presentation.resolve("layout/title-content").name, "Title and Content");
assert.match(presentation.help("presentation.theme").ndjson, /theme colors/);
assert.match(presentation.help("presentation.layouts.add").ndjson, /slide layout/);
assert.match(presentation.help("slide.applyLayout").ndjson, /placeholder/);
const shape = slide.shapes.add({
  geometry: "roundRect",
  name: "summary-surface",
  position: { left: 72, top: 96, width: 420, height: 140 },
  fill: "white",
  line: { fill: "#cbd5e1", width: 1 },
});
shape.text = "Revenue outlook";
shape.text.style = { fontSize: 32, bold: true, color: "#0f172a" };

const snapshot = presentation.inspect({ kind: "deck,slide,textbox,layout", maxChars: 4000 }).ndjson;
assert.match(snapshot, /"kind":"textbox"/);
assert.match(snapshot, /Revenue outlook/);
const target = presentation.resolve(shape.id);
assert.equal(target.text.value, "Revenue outlook");
target.text.replace("outlook", "plan");
assert.match(presentation.inspect({ kind: "textbox" }).ndjson, /Revenue plan/);
const shapeTextRange = presentation.resolve(`${shape.id}/text`);
assert.equal(shapeTextRange.text, "Revenue plan");
const textRangeInspect = presentation.inspect({ kind: "textRange", target: `${shape.id}/text`, include: "text,parentId", maxChars: 4000 }).ndjson;
assert.match(textRangeInspect, /Revenue plan/);
assert.match(textRangeInspect, new RegExp(shape.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));

const composed = slide.compose(
  column({ name: "content-frame", width: "fill", height: "fill", gap: 16, padding: { x: 24, y: 20 } }, [
    paragraph({ id: "sh/stable-headline", name: "primary-heading", className: "text-slate-950 text-4xl font-bold leading-tight" }, [
      "Quarterly ",
      run({ textStyle: { bold: true, color: "sky-600" } }, ["readiness"]),
    ]),
    row({ name: "kpi-row", width: "fill", height: 120, gap: 12 }, [
      box({ name: "kpi-card-a", width: "fill", height: "fill", fill: "slate-50", line: { fill: "#e2e8f0", width: 1 }, padding: { x: 12, y: 10 } }, [
        paragraph({ name: "kpi-label-a", className: "text-slate-700 text-lg" }, ["Pipeline"]),
      ]),
      box({ name: "kpi-card-b", width: "fill", height: "fill", className: "bg-sky-50 rounded-xl", padding: { x: 12, y: 10 } }, [
        paragraph({ name: "kpi-label-b", style: "font: 700 22px Inter; color: #334155; leading: 1.2" }, ["Coverage"]),
      ]),
    ]),
    rule({ name: "section-rule", width: "fill", height: 2, stroke: "slate-900" }),
    composeShape({ name: "accent-pill", width: 180, height: 42, geometry: "roundRect", fill: "sky-100" }, ["Agent-ready"]),
  ]),
  { frame: { left: 80, top: 260, width: 720, height: 320 } },
);
assert.ok(composed.length >= 7);
const composedSnapshot = presentation.inspect({ kind: "textbox,shape", maxChars: 8000 }).ndjson;
assert.match(composedSnapshot, /primary-heading/);
assert.match(composedSnapshot, /Quarterly readiness/);
assert.match(composedSnapshot, /kpi-label-b/);
assert.equal(presentation.resolve("sh/stable-headline").text.value, "Quarterly readiness");
assert.match(presentation.help("slide.compose").ndjson, /Materialize/);

const cards = ["A", "B", "C"].map((label) => {
  const card = slide.shapes.add({ name: `auto-card-${label}`, position: { left: 0, top: 0, width: 100, height: 50 }, fill: "white" });
  card.text = label;
  return card;
});
slide.autoLayout(cards, {
  direction: "horizontal",
  frame: { left: 100, top: 600, width: 500, height: 80 },
  horizontalGap: "auto",
  align: "center",
});
assert.deepEqual(cards.map((card) => Math.round(card.position.left)), [100, 300, 500]);
assert.equal(Math.round(cards[0].position.top), 615);
assert.match(presentation.help("slide.autoLayout").ndjson, /horizontal or vertical flow/);

const nativeTable = slide.tables.add({
  name: "native-import-table",
  position: { left: 840, top: 96, width: 320, height: 120 },
  values: [["Metric", "Value"], ["ARR", "$12M"]],
  styleOptions: { headerRow: true },
});
const nativeChart = slide.charts.add("bar", {
  name: "native-import-chart",
  title: "Native Import Chart",
  position: { left: 840, top: 240, width: 320, height: 180 },
  categories: ["Q1", "Q2"],
  axes: { category: { title: "Quarter" }, value: { title: "Revenue" } },
  legend: { visible: true, position: "r" },
  dataLabels: { showValue: true },
  series: [
    { name: "Revenue", values: [10, 14], color: "#0ea5e9" },
    { name: "Forecast", values: [12, 16], color: "#f97316" },
  ],
});
const nativeImage = slide.images.add({
  name: "native-import-image",
  alt: "Native import logo",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  position: { left: 840, top: 450, width: 120, height: 90 },
});
const pieSlide = presentation.slides.add();
const pieChart = pieSlide.charts.add("pie", {
  name: "market-share-pie",
  title: "Market Share",
  position: { left: 120, top: 120, width: 360, height: 240 },
  categories: ["Product A", "Product B", "Other"],
  dataLabels: { showValue: true, showCategoryName: true },
  legend: { visible: true, position: "r" },
  series: [{ name: "Share", values: [45, 35, 20] }],
});
assert.equal(presentation.resolve(pieChart.id).chartType, "pie");
assert.match(presentation.inspect({ kind: "chart", target: pieChart.id, maxChars: 8000 }).ndjson, /Market Share/);
slide.addNotes("Speaker note: call out pipeline risk.");
const connector = slide.connectors.add({ name: "shape-to-table", from: shape, to: nativeTable, line: { fill: "#0284c7", width: 2, endArrow: "triangle" } });
const thread = slide.comments.addThread(shape, "Tighten this headline before shipping.", { author: "Reviewer" });
thread.addReply("Updated wording to plan.").resolve();
assert.match(presentation.inspect({ kind: "notes,comment,connector", maxChars: 8000 }).ndjson, /Speaker note/);
assert.match(presentation.inspect({ kind: "notes,comment,connector", maxChars: 8000 }).ndjson, /shape-to-table/);
assert.match(presentation.inspect({ kind: "notes,comment,connector", maxChars: 8000 }).ndjson, /Tighten this headline/);
assert.equal(presentation.resolve(connector.id).name, "shape-to-table");
assert.equal(presentation.resolve(thread.id).resolved, true);
assert.match(presentation.help("slide.addNotes").ndjson, /speaker notes/);
assert.match(presentation.help("slide.connectors.add").ndjson, /connector line/);
assert.match(presentation.help("slide.comments.addThread").ndjson, /Threaded comments|threaded comments/);
assert.equal(presentation.resolve(nativeTable.id).values[1][0], "ARR");
assert.equal(presentation.resolve(nativeChart.id).series[0].values[1], 14);
assert.equal(presentation.resolve(nativeChart.id).axes.value.title, "Revenue");
assert.equal(presentation.resolve(nativeChart.id).legend.visible, true);
assert.equal(presentation.resolve(nativeChart.id).dataLabels.showValue, true);
assert.match(presentation.inspect({ kind: "chart", target: nativeChart.id, maxChars: 8000 }).ndjson, /"axes"/);
assert.match(presentation.inspect({ kind: "chart", target: nativeChart.id, maxChars: 8000 }).ndjson, /Forecast/);
assert.match(presentation.help("slide.charts.add").ndjson, /bar\/line\/pie/);
assert.match(presentation.help("slide.charts.add").ndjson, /axes/);
assert.equal(presentation.resolve(nativeImage.id).alt, "Native import logo");
const targetedPresentationInspect = presentation.inspect({ kind: "table,chart,image,connector,comment", target: nativeImage.id, maxChars: 4000 }).ndjson;
assert.match(targetedPresentationInspect, /Native import logo/);
assert.doesNotMatch(targetedPresentationInspect, /Native Import Chart/);
const shapedPresentationInspect = presentation.inspect({ kind: "image", target: nativeImage.id, include: "alt,bbox", exclude: "dataUrl", maxChars: 4000 }).ndjson;
assert.match(shapedPresentationInspect, /Native import logo/);
assert.match(shapedPresentationInspect, /"bbox"/);
assert.doesNotMatch(shapedPresentationInspect, /"dataUrl"/);
const targetedCommentInspect = presentation.inspect({ kind: "comment,connector", target: thread.id, maxChars: 4000 }).ndjson;
assert.match(targetedCommentInspect, /Tighten this headline/);
assert.doesNotMatch(targetedCommentInspect, /shape-to-table/);

const qaClean = Presentation.create({ slideSize: { width: 400, height: 240 } });
const qaCleanSlide = qaClean.slides.add();
qaCleanSlide.shapes.add({ name: "clean-a", position: { left: 20, top: 20, width: 120, height: 60 }, text: "A" });
qaCleanSlide.shapes.add({ name: "clean-b", position: { left: 180, top: 20, width: 120, height: 60 }, text: "B" });
assert.equal(qaClean.validateLayout().ok, true);

const qaBroken = Presentation.create({ slideSize: { width: 320, height: 180 } });
const qaBrokenSlide = qaBroken.slides.add();
qaBrokenSlide.shapes.add({ name: "overlap-a", position: { left: 20, top: 20, width: 120, height: 80 }, text: "A" });
qaBrokenSlide.shapes.add({ name: "overlap-b", position: { left: 80, top: 40, width: 120, height: 80 }, text: "B" });
qaBrokenSlide.shapes.add({ name: "off-canvas", position: { left: 260, top: 120, width: 100, height: 80 }, text: "Off" });
qaBrokenSlide.shapes.add({ name: "overflow-text", position: { left: 20, top: 120, width: 80, height: 24 }, text: "This text is intentionally far too long for this tiny frame." });
qaBrokenSlide.tables.add({ name: "tiny-table", rows: 1, columns: 1, position: { left: 130, top: 125, width: 40, height: 14 }, values: [["Very long value"]] });
const brokenChart = qaBrokenSlide.charts.add("bar", { name: "bad-chart", title: "Bad Chart", position: { left: 30, top: 30, width: 90, height: 50 }, categories: ["A", "B", "C"], series: [{ name: "Mismatch", values: [1, "oops"] }] });
const raggedTable = qaBrokenSlide.tables.add({ name: "ragged-table", rows: 2, columns: 2, position: { left: 5, top: 150, width: 40, height: 20 }, values: [["A", "B"], ["only one"]] });
raggedTable.values[1] = ["only one"];
qaBrokenSlide.images.add({ name: "empty-image", position: { left: 200, top: 20, width: 40, height: 40 } });
qaBrokenSlide.images.add({ name: "bad-data-image", dataUrl: "data:not-base64", position: { left: 250, top: 20, width: 40, height: 40 } });
qaBrokenSlide.connectors.add({ name: "bad-connector", start: { x: 10, y: 10 }, end: { x: 400, y: 20 } });
qaBrokenSlide.comments.addThread("missing-shape", "Dangling comment target.");
qaBrokenSlide.shapes.add({ name: "empty-title", position: { left: 10, top: 10, width: 100, height: 30 }, placeholder: { type: "title", required: true } });
const qa = qaBroken.validateLayout({ minOverlapArea: 16, maxChars: 8000 });
assert.equal(qa.ok, false);
assert.match(qa.ndjson, /"type":"overlap"/);
assert.match(qa.ndjson, /"type":"offCanvas"/);
assert.match(qa.ndjson, /"type":"textOverflow"/);
assert.match(qa.ndjson, /"type":"tableTextOverflow"/);
assert.match(qa.ndjson, /"type":"connectorOffCanvas"/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /danglingComment/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /placeholderMissingContent/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /chartDataMismatch/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /chartDataNonNumeric/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /raggedTableRows/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /emptyImage/);
assert.match(qaBroken.verify({ maxChars: 12000 }).ndjson, /invalidImageDataUrl/);
assert.match(qaBroken.help("presentation.validateLayout").ndjson, /off-canvas/);

const layout = await slide.export({ format: "layout" });
assert.equal(layout.type, "application/vnd.open-office-artifact.layout+json");
const slideLayoutJson = JSON.parse(await layout.text());
assert.ok(slideLayoutJson.elements.some((element) => element.name === "summary-surface"));
const targetedSlideLayoutBlob = await slide.export({ format: "layout", target: nativeImage.id });
assert.equal(targetedSlideLayoutBlob.metadata.target, nativeImage.id);
const targetedSlideLayout = JSON.parse(await targetedSlideLayoutBlob.text());
assert.deepEqual(targetedSlideLayout.elements.map((element) => element.id), [nativeImage.id]);
const chartSearchLayout = JSON.parse(await (await presentation.export({ slide, format: "layout", search: "Native Import Chart" })).text());
assert.deepEqual(chartSearchLayout.elements.map((element) => element.id), [nativeChart.id]);
const contextSlideLayout = slide.layoutJson({ target: nativeImage.id, before: 1 });
assert.deepEqual(contextSlideLayout.elements.map((element) => element.id), [nativeChart.id, nativeImage.id]);
const commentSlideLayout = slide.layoutJson({ target: thread.id });
assert.deepEqual(commentSlideLayout.elements.map((element) => element.id), [shape.id]);
const textRangeSlideLayout = slide.layoutJson({ target: `${shape.id}/text` });
assert.deepEqual(textRangeSlideLayout.elements.map((element) => element.id), [shape.id]);
assert.match(presentation.help("presentation.export").ndjson, /target\/search-sliced layout JSON/);
const preview = await presentation.export({ slide, format: "svg" });
assert.equal(preview.type, "image/svg+xml");
const previewSvg = await preview.text();
assert.match(previewSvg, /Revenue plan/);
assert.match(previewSvg, /Quarter/);
assert.match(previewSvg, /Forecast/);
const pieSvg = await (await presentation.export({ slide: pieSlide, format: "svg" })).text();
assert.match(pieSvg, /Market Share/);
assert.match(pieSvg, /<path d="M /);
assert.match(pieSvg, /Product A/);
const montage = await presentation.export({ format: "montage", columns: 2, scale: 0.2 });
assert.equal(montage.type, "image/svg+xml");
const montageSvg = await montage.text();
assert.match(montageSvg, /data-slide="1"/);
assert.match(montageSvg, /Slide 1/);
assert.match(montageSvg, /Revenue plan/);
assert.equal(montage.metadata.format, "montage");
assert.match(presentation.help("presentation.export").ndjson, /montage/);

const pptx = await PresentationFile.exportPptx(presentation);
const packageInspect = await PresentationFile.inspectPptx(pptx, { includeText: true, maxChars: 16000 });
assert.equal(packageInspect.records[0].kind, "pptxPackage");
assert.equal(packageInspect.records[0].slides, 2);
assert.equal(packageInspect.ok, true);
assert.ok(packageInspect.parts.some((part) => part.path === "ppt/slides/slide1.xml"));
assert.ok(packageInspect.parts.some((part) => part.path === "ppt/charts/chart1.xml"));
assert.match(packageInspect.ndjson, /pptxPart/);
assert.ok(packageInspect.records[0].uncompressedBytes > 0);
assert.ok(packageInspect.records[0].relationshipReferences > 0);
assert.equal(packageInspect.records[0].relationshipReferenceIssues, 0);
assert.ok(packageInspect.parts.some((part) => part.path === "ppt/presentation.xml" && part.contentType.includes("presentationml.presentation.main+xml")));
const pptxReferenceZip = await JSZip.loadAsync(new Uint8Array(await pptx.arrayBuffer()));
const pptxSlideReferenceXml = await pptxReferenceZip.file("ppt/slides/slide1.xml").async("text");
const brokenPptxReferenceXml = pptxSlideReferenceXml.replace(/<\/p:sld>\s*$/, '<p:extLst><p:ext uri="urn:open-office:missing-reference"><a:blip xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:rel="http://purl.oclc.org/ooxml/officeDocument/relationships" rel:embed="rIdMissingSourceReference"/></p:ext></p:extLst></p:sld>');
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/slide1.xml", xml: brokenPptxReferenceXml }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const invalidReferencePptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/slide1.xml", xml: brokenPptxReferenceXml }], { validateResult: false });
const invalidReferencePptxInspect = await PresentationFile.inspectPptx(invalidReferencePptx);
assert.ok(invalidReferencePptxInspect.issues.some((issue) => issue.type === "relationshipReferenceIdNotFound" && issue.referenceAttribute === "rel:embed"));
const patchedPptx = await PresentationFile.patchPptx(pptx, [{
  path: "customXml/review.json",
  json: { status: "ok" },
  relationship: { source: "ppt/presentation.xml", id: "rIdReview", type: "urn:open-office:relationships/review" },
}]);
assert.equal(patchedPptx.type, pptx.type);
assert.equal(patchedPptx.metadata.patchedParts, 1);
assert.equal(patchedPptx.metadata.contentTypesUpdated, 1);
assert.equal(patchedPptx.metadata.relationshipsUpdated, 1);
assert.equal(patchedPptx.metadata.validated, true);
assert.equal(patchedPptx.metadata.validationIssues, 0);
const patchedPptxInspect = await PresentationFile.inspectPptx(patchedPptx, { includeText: true, maxChars: 16000 });
assert.match(patchedPptxInspect.ndjson, /review\.json/);
assert.equal(patchedPptxInspect.ok, true);
assert.ok(patchedPptxInspect.parts.some((part) => part.path === "customXml/review.json" && part.contentType === "application/json"));
const recipeChartPptx = await PresentationFile.patchPptx(pptx, [{
  path: "ppt/charts/review.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea/></c:chart></c:chartSpace>',
  recipe: { kind: "chart", source: "ppt/slides/slide1.xml", id: "rIdReviewChart" },
}]);
assert.equal(recipeChartPptx.metadata.recipesApplied, 1);
const recipeChartInspect = await PresentationFile.inspectPptx(recipeChartPptx);
assert.equal(recipeChartInspect.ok, true);
assert.ok(recipeChartInspect.parts.some((part) => part.path === "ppt/charts/review.xml" && part.contentType === "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"));
const recipeChartZip = await JSZip.loadAsync(new Uint8Array(await recipeChartPptx.arrayBuffer()));
assert.match(await recipeChartZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /Id="rIdReviewChart"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/chart"[^>]*Target="\.\.\/charts\/review\.xml"/);
const recipeSlideXml = '<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>';
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/duplicate-id.xml", xml: recipeSlideXml, recipe: { kind: "slide", source: "ppt/presentation.xml", sourceReference: { slideId: 256 } } }]), /slideId 256 already exists/);
const recipeSlidePptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/slideReview.xml", xml: recipeSlideXml, recipe: { kind: "slide", source: "ppt/presentation.xml", sourceReference: { slideId: 512 } } }]);
assert.equal(recipeSlidePptx.metadata.sourceReferencesUpdated, 1);
const recipeSlideZip = await JSZip.loadAsync(new Uint8Array(await recipeSlidePptx.arrayBuffer()));
const recipeSlideRels = await recipeSlideZip.file("ppt/_rels/presentation.xml.rels").async("text");
const recipeSlideRelId = /<Relationship\b[^>]*Id="([^"]+)"[^>]*Target="slides\/slideReview\.xml"/.exec(recipeSlideRels)?.[1];
assert.ok(recipeSlideRelId);
assert.match(await recipeSlideZip.file("ppt/presentation.xml").async("text"), new RegExp(`<p:sldId id="512" r:id="${recipeSlideRelId}"/>`));
assert.equal((await PresentationFile.importPptx(recipeSlidePptx)).slides.items.length, 3);
const removedRecipeSlidePptx = await PresentationFile.patchPptx(recipeSlidePptx, [{ path: "ppt/slides/slideReview.xml", remove: true, recipe: { kind: "slide", source: "ppt/presentation.xml", sourceReference: true } }]);
assert.equal((await PresentationFile.inspectPptx(removedRecipeSlidePptx)).ok, true);
assert.equal((await PresentationFile.importPptx(removedRecipeSlidePptx)).slides.items.length, 2);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await removedRecipeSlidePptx.arrayBuffer()))).file("ppt/presentation.xml").async("text"), /id="512"/);
const patchedPptxZip = await JSZip.loadAsync(new Uint8Array(await patchedPptx.arrayBuffer()));
assert.match(await patchedPptxZip.file("ppt/_rels/presentation.xml.rels").async("text"), /Id="rIdReview"[^>]*Target="\.\.\/customXml\/review\.json"/);
const removedReviewPptx = await PresentationFile.patchPptx(patchedPptx, [{ path: "customXml/review.json", remove: true }]);
assert.equal(removedReviewPptx.metadata.contentTypesUpdated, 1);
assert.equal(removedReviewPptx.metadata.relationshipsUpdated, 1);
assert.equal((await PresentationFile.inspectPptx(removedReviewPptx)).ok, true);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "customXml/unmanaged.json", json: {} }], { syncContentTypes: false }), /invalid OOXML package.*missingContentType/);
const unsyncedPptx = await PresentationFile.patchPptx(pptx, [{ path: "customXml/unmanaged.json", json: {} }], { syncContentTypes: false, validateResult: false });
const unsyncedPptxInspect = await PresentationFile.inspectPptx(unsyncedPptx);
assert.equal(unsyncedPptxInspect.ok, false);
assert.ok(unsyncedPptxInspect.issues.some((issue) => issue.type === "missingContentType" && issue.path === "customXml/unmanaged.json"));
const replacementPptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/presentation.xml", xml: "<p:presentation xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"/>" }]);
assert.ok((await PresentationFile.inspectPptx(replacementPptx)).parts.some((part) => part.path === "ppt/presentation.xml" && part.contentType.includes("presentationml.presentation.main+xml")));
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "customXml/no-type.json", json: {}, relationship: { source: "ppt/presentation.xml" } }]), /requires type/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "customXml/review.json", json: {}, relationship: { source: "ppt/missing.xml", type: "urn:review" } }]), /source part not found/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "../evil.xml", text: "bad" }]), /Unsafe PPTX part path/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "customXml/large.txt", text: "12345" }], { maxPatchBytes: 4 }), /exceeds maxPatchBytes/);
await assert.rejects(() => PresentationFile.inspectPptx(pptx, { maxParts: 1 }), /maxParts/);
const zip = await JSZip.loadAsync(new Uint8Array(await pptx.arrayBuffer()));
const themeXml = await zip.file("ppt/theme/theme1.xml").async("text");
assert.match(themeXml, /0ea5e9/i);
assert.match(themeXml, /Aptos Display/);
const masterXml = await zip.file("ppt/slideMasters/slideMaster1.xml").async("text");
assert.match(masterXml, /sldLayoutId/);
const layoutXml = await zip.file("ppt/slideLayouts/slideLayout1.xml").async("text");
assert.match(layoutXml, /Title and Content/);
assert.match(layoutXml, /<p:ph type="title"/);
const slideXml = await zip.file("ppt/slides/slide1.xml").async("text");
assert.match(slideXml, /<a:tbl>/);
assert.match(slideXml, /native-import-table/);
assert.match(slideXml, /<c:chart[^>]*r:id=/);
assert.match(slideXml, /<p:pic>/);
assert.match(slideXml, /<p:cxnSp>/);
assert.match(slideXml, /shape-to-table/);
assert.match(slideXml, /<a:rPr lang="en-US" sz="2400" b="1">/);
assert.match(slideXml, /<a:solidFill><a:srgbClr val="FFFFFF"\/><\/a:solidFill>/);
assert.match(slideXml, /<a:tcPr[^>]*><a:solidFill><a:srgbClr val="EDEDED"\/>/);
const notesXml = await zip.file("ppt/notesSlides/notesSlide1.xml").async("text");
assert.match(notesXml, /Speaker note/);
const commentsXml = await zip.file("ppt/comments/comment1.xml").async("text");
assert.match(commentsXml, /Tighten this headline/);
assert.match(commentsXml, /Updated wording/);
assert.match(commentsXml, /ooa:resolved="1"/);
const slideRelsXml = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(slideRelsXml, /Target="\.\.\/media\/image1\.png"/);
assert.match(slideRelsXml, /Target="\.\.\/charts\/chart1\.xml"/);
assert.match(slideRelsXml, /relationships\/slideLayout/);
assert.match(slideRelsXml, /relationships\/notesSlide/);
assert.match(slideRelsXml, /relationships\/comments/);
const chartXml = await zip.file("ppt/charts/chart1.xml").async("text");
assert.match(chartXml, /Native Import Chart/);
assert.match(chartXml, /<c:v>14<\/c:v>/);
assert.match(chartXml, /Quarter/);
assert.match(chartXml, /Revenue/);
assert.match(chartXml, /Forecast/);
assert.match(chartXml, /<c:showVal val="1"\/>/);
assert.match(chartXml, /<c:legendPos val="r"\/>/);
assert.match(chartXml, /0ea5e9/i);
const pieChartXml = await zip.file("ppt/charts/chart2.xml").async("text");
assert.match(pieChartXml, /<c:pieChart>/);
assert.match(pieChartXml, /Market Share/);
assert.match(pieChartXml, /Product A/);
assert.match(pieChartXml, /<c:v>45<\/c:v>/);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pptx`);
await pptx.save(out);
const loaded = await PresentationFile.importPptx(await FileBlob.load(out));
const loadedAll = loaded.inspect({ kind: "theme,layout,textbox,table,chart,image,notes,comment,connector", maxChars: 12000 }).ndjson;
assert.match(loadedAll, /Revenue plan/);
assert.match(loadedAll, /Imported Theme|Open Office Clean Room/);
assert.match(loadedAll, /Title and Content/);
assert.match(loadedAll, /Quarterly plan template/);
assert.match(loadedAll, /native-import-table/);
assert.match(loadedAll, /ARR/);
assert.match(loadedAll, /Native Import Chart/);
assert.match(loadedAll, /Market Share/);
assert.match(loadedAll, /Product A/);
assert.match(loadedAll, /Quarter/);
assert.match(loadedAll, /Forecast/);
assert.match(loadedAll, /native-import-image/);
assert.match(loadedAll, /Native import logo/);
assert.match(loadedAll, /Speaker note/);
assert.match(loadedAll, /Tighten this headline/);
assert.match(loadedAll, /shape-to-table/);
const loadedComment = loaded.slides.items[0].comments.items[0];
assert.ok(loaded.slides.items[0].resolve(loadedComment.targetId));
const loadedSummarySurface = loaded.slides.items[0].shapes.items.find((item) => item.name === "summary-surface");
assert.equal(loadedSummarySurface.text.style.fontSize, 32);
assert.equal(loadedSummarySurface.text.style.bold, true);
assert.equal(loadedSummarySurface.fill.toLowerCase(), "#ffffff");
console.log("presentation smoke ok");
