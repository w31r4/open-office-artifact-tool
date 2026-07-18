import assert from "node:assert/strict";
import JSZip from "jszip";

import {
  column,
  paragraph,
  Presentation,
  PresentationFile,
  row,
  run,
  shape as composeShape,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";
import {
  effectivePresentationImageCrop,
  presentationImageDataUrlDimensions,
} from "../src/presentation/image-crop.mjs";

const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const JPEG = "data:image/jpeg;base64,/9j/2Q==";
const WIDE_SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="200" viewBox="0 0 400 200"><rect width="200" height="200" fill="#2563eb"/><rect x="200" width="200" height="200" fill="#f97316"/></svg>').toString("base64")}`;
const TALL_SVG = `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 200 400"><rect width="200" height="200" fill="#2563eb"/><rect y="200" width="200" height="200" fill="#f97316"/></svg>').toString("base64")}`;

function itemByName(items, name) {
  const item = items.find((candidate) => candidate.name === name);
  assert.ok(item, "Missing presentation object " + name);
  return item;
}

assert.deepEqual(presentationImageDataUrlDimensions(WIDE_SVG), { width: 400, height: 200 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "cover", dataUrl: WIDE_SVG, frame: { width: 200, height: 200 } }), { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "contain", dataUrl: WIDE_SVG, frame: { width: 200, height: 200 } }), { left: 0, top: -0.5, right: 0, bottom: -0.5 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "cover", dataUrl: TALL_SVG, frame: { width: 200, height: 200 } }), { left: 0, top: 0.25, right: 0, bottom: 0.25 });
assert.deepEqual(effectivePresentationImageCrop({ fit: "contain", dataUrl: TALL_SVG, frame: { width: 200, height: 200 } }), { left: -0.5, top: 0, right: -0.5, bottom: 0 });
assert.deepEqual(effectivePresentationImageCrop({
  fit: "cover",
  crop: { left: 0.1, right: 0.1 },
  dataUrl: WIDE_SVG,
  frame: { width: 200, height: 200 },
}), { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.throws(() => effectivePresentationImageCrop({ fit: "cover", dataUrl: "data:image/png;base64,AA==", frame: { width: 100, height: 100 } }), /intrinsic dimensions/);

// The JavaScript layer remains the object model, Compose, inspect, resolve,
// semantic verification, and rendering surface.
const modelPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
assert.equal(modelPresentation.view.gridlinesVisible, false);
assert.equal(modelPresentation.view.guidesVisible, false);
assert.equal(modelPresentation.view.gridSpacingCxEmu, undefined);
assert.equal(modelPresentation.view.gridSpacingCyEmu, undefined);
assert.equal(modelPresentation.view.toProto(), undefined);
assert.equal(modelPresentation.view.showGridlines(), undefined);
assert.equal(modelPresentation.view.gridlinesVisible, true);
assert.equal(modelPresentation.view.toggleGridlines(), false);
assert.equal(modelPresentation.view.hideGridlines(), undefined);
assert.equal(modelPresentation.view.showGuides(), undefined);
assert.equal(modelPresentation.view.guidesVisible, true);
assert.deepEqual(modelPresentation.view.toProto(), { slideViewShowGuides: false, slideGuides: [] });
assert.equal(modelPresentation.view.toggleGuides(), false);
assert.equal(modelPresentation.view.hideGuides(), undefined);
assert.deepEqual(modelPresentation.toProto().viewProperties, { slideViewShowGuides: false, slideGuides: [] });
assert.deepEqual(modelPresentation.master.slideGuides, []);
assert.throws(() => modelPresentation.master.slideGuides.push({ orientation: "horizontal", position: 1 }), TypeError);
assert.throws(() => Presentation.create({ master: { slideGuides: [{ orientation: "diagonal", position: 1 }] } }), /horizontal or vertical/);
assert.throws(() => Presentation.create({ master: { slideGuides: [{ orientation: "horizontal", position: 1.5 }] } }), /signed 32-bit integer/);
const modelSlide = modelPresentation.slides.add({ name: "Compose model" });
const composed = modelSlide.compose(
  column({ name: "compose-root", width: "fill", height: "fill", gap: 18, padding: { x: 24, y: 20 } }, [
    paragraph({ id: "compose/headline", name: "compose-headline", className: "text-slate-950 text-4xl font-bold" }, [
      "Canonical ",
      run({ textStyle: { bold: true, color: "#2563EB" } }, ["Office"]),
      " model",
    ]),
    row({ name: "compose-row", width: "fill", height: 120, gap: 16 }, [
      paragraph({ name: "compose-card-a-copy", width: "fill", height: "fill", className: "text-slate-700 text-lg" }, ["Inspect"]),
      paragraph({ name: "compose-card-b-copy", width: "fill", height: "fill", className: "text-slate-700 text-lg" }, ["Verify"]),
    ]),
    composeShape({ name: "compose-pill", width: 220, height: 48, geometry: "roundRect", fill: "#DBEAFE" }, ["Agent-ready"]),
  ]),
  { frame: { left: 80, top: 80, width: 760, height: 420 } },
);
assert.ok(composed.length >= 4);
assert.equal(modelPresentation.resolve("compose/headline").text.value, "Canonical Office model");
assert.match(modelPresentation.inspect({ kind: "deck,slide,textbox,shape", maxChars: 10_000 }).ndjson, /compose-card-b-copy/);
assert.match(modelPresentation.inspect({ kind: "textbox", target: "compose\/headline", maxChars: 4000 }).ndjson, /Canonical Office model/);
assert.equal(modelPresentation.verify().ok, true);
assert.equal(modelPresentation.validateLayout().ok, true);
const unsupportedThemePresentation = Presentation.create({ theme: { colors: { accent1: "#FF0000" } } });
unsupportedThemePresentation.slides.add().shapes.add({ text: "Theme model only" });
await assert.rejects(
  () => PresentationFile.exportPptx(unsupportedThemePresentation),
  /presentation theme customization/i,
);

// A source-free layout is intentionally a small reusable authoring profile:
// one canonical master, direct-frame title/body text placeholders, and an
// explicit slide binding. It is native PresentationML, not preview-only model
// metadata or a reconstructed imported template graph.
const authoredLayoutPresentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  master: {
    name: "Authoring master",
    placeholders: [{
      type: "title",
      index: 0,
      name: "Title",
      text: "Master title prompt",
      position: { left: 60, top: 44, width: 1160, height: 82 },
      style: { fontSize: 30, bold: true, color: "#0F172A" },
    }],
  },
});
const authoredLayout = authoredLayoutPresentation.layouts.add({ name: "Title and body", type: "titleAndContent" });
authoredLayout.placeholders.add({
  type: "body",
  index: 1,
  name: "Body",
  text: "Master body prompt",
  position: { left: 72, top: 154, width: 1136, height: 490 },
  style: { fontSize: 18, color: "#334155" },
});
assert.equal(authoredLayoutPresentation.layouts.getById(authoredLayout.id), authoredLayout);
assert.equal(authoredLayout.placeholders.count, 1);
const authoredLayoutPlaceholderSummary = authoredLayout.placeholders.summary();
assert.deepEqual(authoredLayoutPlaceholderSummary, {
  ownerId: authoredLayout.id,
  count: 1,
  requiredCount: 0,
  types: ["body"],
  items: [{
    id: authoredLayout.placeholders.getItem("body").id,
    name: "Body",
    type: "body",
    idx: 1,
    index: 1,
    required: false,
    hasDirectPosition: true,
    position: { left: 72, top: 154, width: 1136, height: 490 },
  }],
});
authoredLayoutPlaceholderSummary.items[0].position.left = -1;
assert.equal(authoredLayout.placeholders.getItem("body").position.left, 72);
const authoredLayoutSlide = authoredLayoutPresentation.slides.add({ name: "Reusable layout", layout: "Title and body" });
assert.equal(authoredLayoutSlide.layoutId, authoredLayout.id);
assert.equal(authoredLayoutSlide.placeholders.count, 2);
const materializedPlaceholderCount = authoredLayoutSlide.shapes.items.length;
assert.equal(authoredLayoutSlide.setLayout(authoredLayout), authoredLayoutSlide);
assert.equal(authoredLayoutSlide.layoutId, authoredLayout.id);
assert.equal(authoredLayoutSlide.placeholders.count, 2);
assert.equal(authoredLayoutSlide.shapes.items.length, materializedPlaceholderCount);
const authoredTitle = authoredLayoutSlide.placeholders.getItem("title");
const authoredBody = authoredLayoutSlide.placeholders.getItem(1);
assert.ok(authoredTitle);
assert.ok(authoredBody);
authoredTitle.text.set("OpenChestnut layout title");
authoredBody.text.set("A direct-frame body placeholder survives native export and import.");
const authoredLayoutExport = await PresentationFile.exportPptx(authoredLayoutPresentation);
const authoredLayoutZip = await JSZip.loadAsync(new Uint8Array(await authoredLayoutExport.arrayBuffer()));
const authoredMasterXml = await authoredLayoutZip.file("ppt/slideMasters/slideMaster1.xml").async("text");
const authoredLayoutXml = await authoredLayoutZip.file("ppt/slideLayouts/slideLayout1.xml").async("text");
const authoredSlideXml = await authoredLayoutZip.file("ppt/slides/slide1.xml").async("text");
assert.match(authoredMasterXml, /<p:ph[^>]*type="title"[^>]*idx="0"/);
assert.match(authoredLayoutXml, /<p:sldLayout[^>]*type="obj"/);
assert.match(authoredLayoutXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="title"[^>]*idx="0"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
const authoredLayoutImported = await PresentationFile.importPptx(authoredLayoutExport);
assert.equal(authoredLayoutImported.master.placeholders.length, 1);
assert.equal(authoredLayoutImported.layouts.items.length, 1);
assert.equal(authoredLayoutImported.layouts.items[0].type, "obj");
const importedLayoutSlide = authoredLayoutImported.slides.getItem(0);
assert.equal(importedLayoutSlide.layoutId, authoredLayoutImported.layouts.items[0].id);
assert.equal(importedLayoutSlide.placeholders.getItem("title").text.value, "OpenChestnut layout title");
assert.equal(importedLayoutSlide.placeholders.getItem("body").text.value, "A direct-frame body placeholder survives native export and import.");
const authoredLayoutRoundTrip = await PresentationFile.exportPptx(authoredLayoutImported);
assert.equal((await PresentationFile.inspectPptx(authoredLayoutRoundTrip)).ok, true);

const guardedLayoutPresentation = Presentation.create();
const firstGuardedLayout = guardedLayoutPresentation.layouts.add({
  name: "First guarded layout",
  type: "title",
  placeholders: [{ type: "title", index: 0, position: { left: 80, top: 72, width: 960, height: 88 } }],
});
const secondGuardedLayout = guardedLayoutPresentation.layouts.add({ name: "Second guarded layout", type: "blank" });
const guardedLayoutSlide = guardedLayoutPresentation.slides.add({ layout: firstGuardedLayout });
assert.throws(
  () => guardedLayoutSlide.setLayout(secondGuardedLayout),
  /already has materialized placeholders.*changing layouts/i,
);
const slideCountBeforeUnknownLayout = guardedLayoutPresentation.slides.count;
assert.throws(
  () => guardedLayoutPresentation.slides.add({ layout: "Missing layout" }),
  /Unknown presentation layout: Missing layout/,
);
assert.equal(guardedLayoutPresentation.slides.count, slideCountBeforeUnknownLayout);

const insertionPresentation = Presentation.create();
const insertionLayout = insertionPresentation.layouts.add({
  name: "Inserted title",
  type: "title",
  placeholders: [{ type: "title", index: 0, position: { left: 88, top: 68, width: 920, height: 92 } }],
});
const insertionFirst = insertionPresentation.slides.add({ name: "First" });
const insertionThird = insertionPresentation.slides.add({ name: "Third" });
const insertionFront = insertionPresentation.slides.insert({ after: null, name: "Front" });
const insertionSecond = insertionPresentation.slides.insert({ after: 0, name: "Second" });
const insertionAfterFirst = insertionPresentation.slides.insert({ after: insertionFirst, name: "After first", layout: insertionLayout });
assert.deepEqual(insertionPresentation.slides.items.map((slide) => slide.name), ["Front", "Second", "First", "After first", "Third"]);
assert.equal(insertionAfterFirst.placeholders.count, 1);
assert.equal(insertionAfterFirst.placeholders.getItem("title").placeholder.layoutId, insertionLayout.id);
const insertionCountBeforeRejectedTarget = insertionPresentation.slides.count;
const foreignSlide = Presentation.create().slides.add({ name: "Foreign" });
assert.throws(
  () => insertionPresentation.slides.insert({ after: foreignSlide, name: "Rejected" }),
  /insertion target must belong to this presentation/i,
);
assert.throws(
  () => insertionPresentation.slides.insert({ after: 99, name: "Rejected" }),
  /after must be an existing Slide, a 0-based slide index, or null/i,
);
assert.equal(insertionPresentation.slides.count, insertionCountBeforeRejectedTarget);
const insertionRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(insertionPresentation));
assert.deepEqual(insertionRoundTrip.slides.items.map((slide) => slide.name), ["Front", "Second", "First", "After first", "Third"]);

const invalidSourceFreeLayout = Presentation.create();
const invalidSourceFreeSlide = invalidSourceFreeLayout.slides.add();
const invalidLayout = invalidSourceFreeLayout.layouts.add({
  name: "Missing direct placeholder frame",
  type: "title",
  placeholders: [{ type: "title", index: 0 }],
});
invalidSourceFreeSlide.setLayout(invalidLayout);
await assert.rejects(
  () => PresentationFile.exportPptx(invalidSourceFreeLayout),
  /requires a direct position/i,
);

const customGeometryPresentation = Presentation.create({ slideSize: { width: 400, height: 240 } });
const customGeometrySlide = customGeometryPresentation.slides.add({ name: "Custom geometry" });
const customGeometryShape = customGeometrySlide.shapes.add({
  name: "literal-custom-path",
  geometry: "custom",
  position: { left: 20, top: 20, width: 180, height: 120 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  customPaths: [{
    width: 21_600,
    height: 21_600,
    commands: [
      { moveTo: { x: 1_000, y: 2_000 } },
      { lineTo: { x: 20_000, y: 2_000 } },
      { cubicBezTo: { x1: 21_000, y1: 6_000, x2: 18_000, y2: 19_000, x: 10_800, y: 20_000 } },
      { close: {} },
    ],
  }],
});
assert.equal(customGeometryShape.customPaths[0].commands.length, 4);
assert.match(await (await customGeometrySlide.export()).text(), /<path d="M 1000 2000 L 20000 2000 C /);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 100, height: 100, commands: [{ arcTo: {} }] }] }),
  /unsupported command arcTo/,
);
assert.throws(
  () => customGeometrySlide.shapes.add({ geometry: "custom", customPaths: [{ width: 0, height: 100, commands: [{ close: true }] }] }),
  /width and height must be positive/,
);

// Groups are a recursive DrawingML ownership boundary, not flattened children
// with synthetic parent IDs. The public model keeps child coordinates local and
// OpenChestnut authors/imports native p:grpSp trees with fixed-topology edits.
const groupedPresentation = Presentation.create({ slideSize: { width: 960, height: 540 } });
const groupedSlide = groupedPresentation.slides.add({ name: "Native group tree" });
const authoredGroup = groupedSlide.groups.add({
  name: "Agent evidence group",
  position: { left: 100, top: 80, width: 600, height: 320 },
  childFrame: { left: -100, top: 50, width: 1200, height: 640 },
});
const groupedBefore = authoredGroup.shapes.add({
  name: "grouped-before",
  geometry: "roundRect",
  position: { left: 0, top: 100, width: 300, height: 120 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  text: "Before",
});
const groupedTarget = authoredGroup.shapes.add({
  name: "grouped-target",
  geometry: "rect",
  position: { left: 450, top: 100, width: 300, height: 120 },
  fill: "#DCFCE7",
  line: { fill: "#16A34A", width: 2 },
  text: "Target",
});
authoredGroup.connectors.add({
  name: "grouped-connector",
  connectorType: "straight",
  from: groupedBefore,
  to: groupedTarget,
  start: { x: 300, y: 160 },
  end: { x: 450, y: 160 },
  line: { fill: "#334155", width: 2, endArrow: "triangle" },
});
authoredGroup.images.add({
  name: "grouped-image",
  alt: "Grouped image evidence",
  position: { left: 800, top: 100, width: 120, height: 120 },
  fit: "stretch",
  dataUrl: PNG,
});
authoredGroup.tables.add({
  name: "grouped-table",
  position: { left: 0, top: 300, width: 400, height: 180 },
  values: [["Gate", "State"], ["Import", "Before"]],
  styleOptions: { headerRow: true, bandedRows: true },
});
authoredGroup.charts.add("combo", {
  name: "grouped-chart",
  title: "Grouped readiness",
  position: { left: 450, top: 300, width: 350, height: 200 },
  categories: ["Create", "Edit"],
  series: [
    { name: "Score", chartType: "bar", values: [7, 9], color: "#7C3AED" },
    {
      name: "Review", chartType: "line", axisGroup: "secondary", values: [5, 8],
      line: { fill: "#0F766E", width: 2 },
      marker: { symbol: "circle", size: 6, fill: "#0F766E" },
    },
  ],
  axes: {
    category: { title: "Stage" },
    value: { title: "Score" },
    secondary: { category: { title: "Stage" }, value: { title: "Review", min: 0, max: 10, majorUnit: 2 } },
  },
  legend: false,
});
const nestedGroup = authoredGroup.groups.add({
  name: "nested-group",
  position: { left: 850, top: 300, width: 250, height: 220 },
  childFrame: { left: 0, top: 0, width: 250, height: 220 },
});
nestedGroup.shapes.add({
  name: "nested-shape",
  geometry: "rect",
  position: { left: 20, top: 30, width: 200, height: 120 },
  fill: "#FCE7F3",
  line: { fill: "#BE185D", width: 1 },
  text: "Nested",
});
assert.equal(groupedPresentation.resolve(groupedBefore.id), groupedBefore);
assert.match(groupedPresentation.inspect({ kind: "groupShape,shape,connector,table,chart,image", maxChars: 20_000 }).ndjson, /Agent evidence group/);
assert.match(authoredGroup.toSvg(), /translate\(100 80\) scale\(0\.5 0\.5\) translate\(100 -50\)/);
assert.equal(groupedPresentation.verify().ok, true);
assert.equal(groupedPresentation.validateLayout().ok, true);

const groupedFirstExport = await PresentationFile.exportPptx(groupedPresentation);
const groupedFirstZip = await JSZip.loadAsync(new Uint8Array(await groupedFirstExport.arrayBuffer()));
const groupedFirstXml = await groupedFirstZip.file("ppt/slides/slide1.xml").async("text");
assert.equal((groupedFirstXml.match(/<p:grpSp>/g) || []).length, 2);
assert.match(groupedFirstXml, /<a:chOff x="-952500" y="476250"\s*\/>/);
assert.match(groupedFirstXml, /<a:chExt cx="11430000" cy="6096000"\s*\/>/);

const groupedImported = await PresentationFile.importPptx(groupedFirstExport);
const importedGroup = itemByName(groupedImported.slides.getItem(0).groups.items, "Agent evidence group");
assert.deepEqual(importedGroup.children.map((child) => child.layoutJson().kind), ["textbox", "textbox", "connector", "image", "table", "chart", "groupShape"]);
assert.deepEqual(importedGroup.childFrame, { left: -100, top: 50, width: 1200, height: 640 });
assert.equal(groupedImported.resolve(itemByName(importedGroup.shapes.items, "grouped-before").id).text.value, "Before");

importedGroup.name = "Edited agent evidence group";
importedGroup.position.left = 120;
importedGroup.childFrame.left = -50;
itemByName(importedGroup.shapes.items, "grouped-before").text.set("After");
delete itemByName(importedGroup.connectors.items, "grouped-connector").line.endArrow;
itemByName(importedGroup.images.items, "grouped-image").alt = "Edited grouped image evidence";
itemByName(importedGroup.tables.items, "grouped-table").cells.set(1, 1, "After");
const importedGroupedChart = itemByName(importedGroup.charts.items, "grouped-chart");
assert.equal(importedGroupedChart.chartType, "combo");
assert.deepEqual(importedGroupedChart.series.map((series) => [series.chartType, series.axisGroup || "primary"]), [["bar", "primary"], ["line", "secondary"]]);
assert.equal(importedGroupedChart.axes.secondary.value.max, 10);
importedGroupedChart.title = "Edited grouped readiness";
importedGroupedChart.series[0].values = [8, 10];
importedGroupedChart.series[1].values = [6, 9];
importedGroupedChart.axes.secondary.value.max = 12;
const importedNestedGroup = itemByName(importedGroup.groups.items, "nested-group");
importedNestedGroup.position.top = 320;
itemByName(importedNestedGroup.shapes.items, "nested-shape").fill = "#FDE68A";

const groupedSecondExport = await PresentationFile.exportPptx(groupedImported);
const groupedRoundTrip = await PresentationFile.importPptx(groupedSecondExport);
const roundTripGroup = itemByName(groupedRoundTrip.slides.getItem(0).groups.items, "Edited agent evidence group");
assert.equal(roundTripGroup.position.left, 120);
assert.equal(roundTripGroup.childFrame.left, -50);
assert.equal(itemByName(roundTripGroup.shapes.items, "grouped-before").text.value, "After");
assert.equal(itemByName(roundTripGroup.connectors.items, "grouped-connector").line.endArrow, undefined);
assert.equal(itemByName(roundTripGroup.images.items, "grouped-image").alt, "Edited grouped image evidence");
assert.equal(itemByName(roundTripGroup.tables.items, "grouped-table").values[1][1], "After");
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").chartType, "combo");
assert.deepEqual(itemByName(roundTripGroup.charts.items, "grouped-chart").series[0].values, [8, 10]);
assert.deepEqual(itemByName(roundTripGroup.charts.items, "grouped-chart").series[1].values, [6, 9]);
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").series[1].axisGroup, "secondary");
assert.equal(itemByName(roundTripGroup.charts.items, "grouped-chart").axes.secondary.value.max, 12);
assert.equal(itemByName(itemByName(roundTripGroup.groups.items, "nested-group").shapes.items, "nested-shape").fill, "#FDE68A");

const removedGroupedChild = roundTripGroup.children.pop();
await assert.rejects(
  () => PresentationFile.exportPptx(groupedRoundTrip),
  (error) => error?.code === "presentation_group_topology_changed",
);
roundTripGroup.children.push(removedGroupedChild);

const irregularGroupXml = groupedFirstXml.replace(
  /(<p:grpSp><p:nvGrpSpPr><p:cNvPr\b[^>]*name="Agent evidence group"[^>]*\/>[\s\S]*?<p:grpSpPr)(>)/,
  "$1 bwMode=\"gray\"$2",
);
assert.notEqual(irregularGroupXml, groupedFirstXml);
const irregularGroupFile = await PresentationFile.patchPptx(groupedFirstExport, [{ path: "ppt/slides/slide1.xml", xml: irregularGroupXml }]);
const irregularGroupZip = await JSZip.loadAsync(new Uint8Array(await irregularGroupFile.arrayBuffer()));
assert.match(await irregularGroupZip.file("ppt/slides/slide1.xml").async("text"), /<p:grpSpPr bwMode="gray">/);
const irregularGroupPresentation = await PresentationFile.importPptx(irregularGroupFile);
const irregularGroupSlide = irregularGroupPresentation.slides.getItem(0);
assert.equal(irregularGroupSlide.groups.items.length, 0);
const opaqueGroup = itemByName(irregularGroupSlide.nativeObjects.items, "Agent evidence group");
assert.equal(opaqueGroup.editable, false);
opaqueGroup.name = "Unsafe group edit";
await assert.rejects(
  () => PresentationFile.exportPptx(irregularGroupPresentation),
  (error) => error?.code === "unsupported_presentation_edit",
);

// The canonical file facade always crosses the OpenChestnut C# WASM layer.
const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const coreSlide = deck.slides.add({ name: "Core objects", background: { fill: "#F1F5F9", mode: "solid" } });
coreSlide.addNotes("Lead with the customer outcome.\nThen explain the operating model.");
coreSlide.shapes.add({
  name: "core-title",
  geometry: "textbox",
  position: { left: 50, top: 28, width: 1180, height: 70 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Presentation 0.2 core",
  textStyle: { fontFamily: "Arial", fontSize: 38, bold: true, color: "#0F172A" },
});
const rounded = coreSlide.shapes.add({
  name: "rounded-card",
  geometry: "roundRect",
  position: { left: 60, top: 140, width: 260, height: 100 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  shadow: { color: "#000000", blurRadius: 8, distance: 4, direction: 45, opacity: 0.25 },
  text: "Before edit",
  textStyle: { fontFamily: "Arial", fontSize: 25, bold: true, color: "#1E3A8A" },
});
const target = coreSlide.shapes.add({
  name: "target-textbox",
  geometry: "textbox",
  position: { left: 400, top: 140, width: 260, height: 100 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Target",
  textStyle: { fontFamily: "Arial", fontSize: 25, bold: true, color: "#14532D" },
});
coreSlide.shapes.add({
  name: "rich-copy",
  geometry: "textbox",
  position: { left: 840, top: 270, width: 380, height: 300 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [
    {
      runs: [
        { text: "Structured ", style: { fontSize: 24, bold: true, color: "#0F172A" } },
        { text: "text", style: { fontSize: 24, italic: true, color: "#2563EB" }, link: { uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/" } },
      ],
    },
    {
      bulletCharacter: "•",
      marginLeft: 28,
      indent: -14,
      runs: [{ text: "Character list", style: { fontSize: 19, color: "#334155" } }],
    },
    {
      autoNumber: { type: "arabicPeriod", startAt: 1 },
      marginLeft: 28,
      indent: -14,
      runs: [{ text: "Numbered list", style: { fontSize: 19, color: "#334155" } }],
    },
  ],
});
coreSlide.tables.add({
  name: "fixed-table",
  position: { left: 60, top: 300, width: 360, height: 190 },
  values: [["Layer", "State"], ["Office", "Before"], ["QA", "Ready"]],
  styleOptions: { headerRow: true, bandedRows: true },
});
coreSlide.images.add({
  name: "png-image",
  alt: "PNG evidence",
  position: { left: 470, top: 300, width: 140, height: 140 },
  fit: "stretch",
  dataUrl: PNG,
});
coreSlide.images.add({
  name: "jpeg-image",
  alt: "JPEG evidence",
  position: { left: 650, top: 300, width: 140, height: 140 },
  fit: "stretch",
  dataUrl: JPEG,
});
const coverImage = coreSlide.images.add({
  name: "cover-image",
  alt: "Wide image cropped to a square",
  position: { left: 650, top: 480, width: 140, height: 140 },
  fit: "cover",
  dataUrl: WIDE_SVG,
});
assert.match(coverImage.toSvg(), /viewBox="100 0 200 200"/);
assert.throws(() => { coverImage.crop = { left: 0.8, right: 0.3 }; }, /opposing sums/);
coreSlide.connectors.add({
  name: "straight-connector",
  connectorType: "straight",
  from: rounded,
  to: target,
  start: { x: 320, y: 180 },
  end: { x: 400, y: 180 },
  line: { fill: "#334155", width: 2, endArrow: "triangle" },
});
coreSlide.connectors.add({
  name: "elbow-polyline-connector",
  connectorType: "elbow",
  from: rounded,
  to: target,
  start: { x: 320, y: 210 },
  end: { x: 400, y: 225 },
  line: { fill: "#7C3AED", width: 2, startArrow: "triangle", endArrow: "triangle" },
});

const chartSlide = deck.slides.add({ name: "Literal charts" });
chartSlide.shapes.add({
  name: "chart-title",
  geometry: "textbox",
  position: { left: 50, top: 28, width: 1180, height: 70 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Source-free bar, line, and pie",
  textStyle: { fontFamily: "Arial", fontSize: 38, bold: true, color: "#0F172A" },
});
chartSlide.charts.add("bar", {
  name: "bar-chart",
  title: "Readiness",
  position: { left: 30, top: 130, width: 380, height: 320 },
  categories: ["Create", "Inspect", "Render"],
  series: [{ name: "Score", values: [78, 92, 85], color: "#2563EB" }],
  legend: false,
  axes: { category: { title: "Gate" }, value: { title: "Score", min: 0, max: 100, majorUnit: 20 } },
  dataLabels: { showValue: true, position: "outsideEnd" },
});
chartSlide.charts.add("line", {
  name: "line-chart",
  title: "Trend",
  position: { left: 450, top: 130, width: 380, height: 320 },
  categories: ["W1", "W2", "W3"],
  series: [{
    name: "Passes",
    values: [6, 9, 12],
    color: "#16A34A",
    line: { fill: "#16A34A", width: 2, style: "dash" },
    marker: { symbol: "circle", size: 7, fill: "#16A34A" },
  }],
  legend: false,
});
chartSlide.charts.add("pie", {
  name: "pie-chart",
  title: "Coverage",
  position: { left: 870, top: 130, width: 380, height: 320 },
  categories: ["Modeled", "Opaque"],
  series: [{ name: "Share", values: [80, 20], color: "#7C3AED" }],
  legend: true,
  dataLabels: { showCategoryName: true, showValue: true },
});

const comboSlide = deck.slides.add({ name: "Literal combo chart" });
comboSlide.charts.add("combo", {
  name: "revenue-margin-combo",
  title: "Revenue and margin",
  position: { left: 90, top: 120, width: 1080, height: 480 },
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48, 57], color: "#2563EB" },
    {
      name: "Margin",
      chartType: "line",
      values: [12, 15, 18],
      color: "#16A34A",
      line: { fill: "#16A34A", width: 2 },
      marker: { symbol: "circle", size: 7, fill: "#16A34A" },
    },
  ],
  legend: true,
  axes: { category: { title: "Quarter" }, value: { title: "Percent" } },
  dataLabels: { showValue: true, position: "top" },
});
const secondaryAxisCombo = Presentation.create({ slideSize: { width: 640, height: 360 } });
const secondaryAxisSlide = secondaryAxisCombo.slides.add({ name: "Secondary-axis combo" });
secondaryAxisSlide.charts.add("combo", {
  name: "secondary-axis-combo",
  title: "Revenue and gross margin",
  position: { left: 48, top: 60, width: 540, height: 250 },
  categories: ["Q1", "Q2"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48], color: "#2563EB" },
    { name: "Gross margin", chartType: "line", axisGroup: "secondary", values: [45, 50], line: { fill: "#16A34A", width: 2 }, marker: { symbol: "circle", size: 6, fill: "#16A34A" } },
  ],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Revenue ($M)" },
    secondary: { category: { title: "Quarter" }, value: { title: "Gross margin (%)", min: 0, max: 100, majorUnit: 10 } },
  },
  legend: true,
});
const secondaryAxisExport = await PresentationFile.exportPptx(secondaryAxisCombo);
const secondaryAxisZip = await JSZip.loadAsync(new Uint8Array(await secondaryAxisExport.arrayBuffer()));
const secondaryAxisChartXml = await Promise.all(Object.keys(secondaryAxisZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => secondaryAxisZip.file(name).async("text")))
  .then((items) => items.find((xml) => xml.includes("Revenue and gross margin")));
assert.ok(secondaryAxisChartXml);
assert.match(secondaryAxisChartXml, /<c:barChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:barChart>/);
assert.match(secondaryAxisChartXml, /<c:lineChart>[\s\S]*?<c:axId val="3"\s*\/><c:axId val="4"\s*\/><\/c:lineChart>/);
assert.match(secondaryAxisChartXml, /<c:catAx><c:axId val="3"\s*\/>[\s\S]*?<c:axPos val="t"\s*\/>/);
assert.match(secondaryAxisChartXml, /<c:valAx><c:axId val="4"\s*\/>[\s\S]*?<c:axPos val="r"\s*\/>/);
const importedSecondaryAxis = await PresentationFile.importPptx(secondaryAxisExport);
const importedSecondaryAxisChart = itemByName(importedSecondaryAxis.slides.getItem(0).charts.items, "secondary-axis-combo");
assert.equal(importedSecondaryAxisChart.chartType, "combo");
assert.deepEqual(importedSecondaryAxisChart.series.map((series) => [series.chartType, series.axisGroup || "primary"]), [["bar", "primary"], ["line", "secondary"]]);
assert.equal(importedSecondaryAxisChart.axes.secondary.category.title, "Quarter");
assert.equal(importedSecondaryAxisChart.axes.secondary.value.title, "Gross margin (%)");
assert.equal(importedSecondaryAxisChart.axes.secondary.value.max, 100);
importedSecondaryAxisChart.series[1].values = [47, 53];
importedSecondaryAxisChart.axes.secondary.value.max = 80;
const editedSecondaryAxis = await PresentationFile.exportPptx(importedSecondaryAxis);
const roundTripSecondaryAxis = await PresentationFile.importPptx(editedSecondaryAxis);
const roundTripSecondaryAxisChart = itemByName(roundTripSecondaryAxis.slides.getItem(0).charts.items, "secondary-axis-combo");
assert.deepEqual(roundTripSecondaryAxisChart.series[1].values, [47, 53]);
assert.equal(roundTripSecondaryAxisChart.series[1].axisGroup, "secondary");
assert.equal(roundTripSecondaryAxisChart.axes.secondary.value.max, 80);

const mixedAxisCombo = Presentation.create({ slideSize: { width: 640, height: 360 } });
mixedAxisCombo.slides.add({ name: "Rejected mixed combo" }).charts.add("combo", {
  name: "mixed-axis-combo",
  categories: ["Q1", "Q2"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48] },
    { name: "Primary line", chartType: "line", values: [12, 15] },
    { name: "Secondary line", chartType: "line", axisGroup: "secondary", values: [45, 50] },
  ],
});
await assert.rejects(PresentationFile.exportPptx(mixedAxisCombo), /cannot mix primary and secondary line plots/i);

assert.equal(deck.verify().ok, true);
assert.equal(deck.validateLayout().ok, true);
assert.equal(deck.resolve(rounded.id), rounded);
assert.equal(deck.resolve(rounded.id + "/text").text, "Before edit");
const deckInspect = deck.inspect({ kind: "deck,slide,textbox,shape,table,chart,image,connector,textRange,notes", maxChars: 24_000 }).ndjson;
assert.match(deckInspect, /Lead with the customer outcome/);
assert.match(deckInspect, /"background":\{"fill":"#F1F5F9","mode":"solid"\}/);
assert.equal(deck.resolve(coreSlide.speakerNotes.id), coreSlide.speakerNotes);
coreSlide.speakerNotes.textFrame.setText("Lead with the customer outcome.\nThen explain the operating model.");
assert.equal(coreSlide.speakerNotes.append("").text, "Lead with the customer outcome.\nThen explain the operating model.");

const firstExport = await PresentationFile.exportPptx(deck);
assert.equal(firstExport.metadata.codec, "open-chestnut");
assert.equal((await PresentationFile.inspectPptx(firstExport)).ok, true);

// Imported deck reordering is intentionally a shallow package operation: it
// preserves every original SlidePart exactly once and changes only the
// p:sldIdLst display order. It is separate from the even narrower, isolated
// layout-only imported-slide deletion profile below.
const reorderedImportedDeck = await PresentationFile.importPptx(firstExport);
const originalImportedSlideNames = reorderedImportedDeck.slides.items.map((slide) => slide.name);
const importedFirstSlide = reorderedImportedDeck.slides.getItem(0);
assert.equal(importedFirstSlide.moveTo(2), importedFirstSlide);
assert.deepEqual(reorderedImportedDeck.slides.items.map((slide) => slide.name), [...originalImportedSlideNames.slice(1), originalImportedSlideNames[0]]);
assert.throws(
  () => importedFirstSlide.moveTo(3),
  /destination must be an existing 0-based slide index/i,
);
const reorderedImportedPptx = await PresentationFile.exportPptx(reorderedImportedDeck);
const originalImportedZip = await JSZip.loadAsync(firstExport.bytes);
const reorderedImportedZip = await JSZip.loadAsync(reorderedImportedPptx.bytes);
for (const path of Object.keys(originalImportedZip.files).filter((path) => /^ppt\/slides\/slide\d+\.xml$/.test(path))) {
  assert.deepEqual(
    await reorderedImportedZip.file(path).async("uint8array"),
    await originalImportedZip.file(path).async("uint8array"),
    `reordering must preserve ${path} byte-for-byte`,
  );
}
const reorderedImportedRoundTrip = await PresentationFile.importPptx(reorderedImportedPptx);
assert.deepEqual(reorderedImportedRoundTrip.slides.items.map((slide) => slide.name), [...originalImportedSlideNames.slice(1), originalImportedSlideNames[0]]);
const reorderedEditedDeck = await PresentationFile.importPptx(firstExport);
const reorderedEditedSlide = reorderedEditedDeck.slides.getItem(0);
reorderedEditedSlide.moveTo(2);
itemByName(reorderedEditedSlide.shapes.items, "rounded-card").text.set("Edited after reorder");
const reorderedEditedPptx = await PresentationFile.exportPptx(reorderedEditedDeck);
const reorderedEditedRoundTrip = await PresentationFile.importPptx(reorderedEditedPptx);
assert.equal(itemByName(reorderedEditedRoundTrip.slides.getItem(2).shapes.items, "rounded-card").text.value, "Edited after reorder");
const importedTopologyChange = await PresentationFile.importPptx(firstExport);
importedTopologyChange.slides.add({ name: "Not source-bound" });
await assert.rejects(
  () => PresentationFile.exportPptx(importedTopologyChange),
  (error) => error?.code === "presentation_topology_changed",
);

// Imported deletion is a real OPC delete, not hiding a slide from p:sldIdLst.
// The bounded profile accepts an otherwise isolated SlidePart with only its
// layout relation, preserves every retained source part byte-for-byte, and
// refuses to remove the final remaining slide before a package write begins.
const deletionFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const deletionKeep = deletionFixture.slides.add({ name: "Keep" });
deletionKeep.shapes.add({ name: "keep-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Keep" });
const deletionRemove = deletionFixture.slides.add({ name: "Remove" });
deletionRemove.shapes.add({ name: "remove-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Remove" });
const deletionSourcePptx = await PresentationFile.exportPptx(deletionFixture);
const deletionSourceZip = await JSZip.loadAsync(deletionSourcePptx.bytes);
const deletionImportedDeck = await PresentationFile.importPptx(deletionSourcePptx);
const deletionImportedSlide = deletionImportedDeck.slides.getItem(1);
assert.equal(deletionImportedSlide.delete(), undefined);
assert.throws(() => deletionImportedDeck.slides.getItem(0).delete(), /retain at least one slide/i);
const deletionPptx = await PresentationFile.exportPptx(deletionImportedDeck);
const deletionZip = await JSZip.loadAsync(deletionPptx.bytes);
assert.equal(deletionZip.file("ppt/slides/slide2.xml"), null, "the deleted slide part must not remain in the package");
assert.equal(deletionZip.file("ppt/slides/_rels/slide2.xml.rels"), null, "the deleted slide relationship part must not remain in the package");
assert.deepEqual(
  await deletionZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await deletionSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "deleting another imported slide must retain the survivor byte-for-byte",
);
const deletionRoundTrip = await PresentationFile.importPptx(deletionPptx);
assert.deepEqual(deletionRoundTrip.slides.items.map((slide) => slide.name), ["Keep"]);

// Imported duplication is a distinct OPC graph operation, not another
// p:sldId reference to the same SlidePart. The first profile deliberately
// stays small: an unchanged shape-only slide with its layout as the only
// relationship becomes a fresh part; after export/reimport it is an ordinary
// source-bound slide and can use the normal supported edit path.
assert.throws(
  () => deletionFixture.slides.getItem(0).duplicate(),
  /available only for a supported imported PPTX source slide/i,
);
const cloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const cloneOriginal = cloneFixture.slides.add({ name: "Original" });
cloneOriginal.shapes.add({ name: "clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Original" });
cloneFixture.slides.add({ name: "Companion" }).shapes.add({ name: "companion-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Companion" });
const cloneSourcePptx = await PresentationFile.exportPptx(cloneFixture);
const cloneSourceZip = await JSZip.loadAsync(cloneSourcePptx.bytes);
const cloneImportedDeck = await PresentationFile.importPptx(cloneSourcePptx);
const importedCloneSource = cloneImportedDeck.slides.getItem(0);
const importedClone = importedCloneSource.duplicate();
assert.equal(importedClone.index, 1);
assert.notEqual(importedClone.id, importedCloneSource.id);
assert.notEqual(importedClone.shapes.items[0].id, importedCloneSource.shapes.items[0].id);
assert.equal(importedClone.shapes.items[0].text.value, "Original");
const clonePptx = await PresentationFile.exportPptx(cloneImportedDeck);
const cloneZip = await JSZip.loadAsync(clonePptx.bytes);
assert.deepEqual(
  await cloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await cloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning an imported slide must leave its origin SlidePart byte-for-byte intact",
);
assert.ok(cloneZip.file("ppt/slides/slide3.xml"), "the clone must be a new SlidePart rather than another reference to slide1");
const cloneRoundTrip = await PresentationFile.importPptx(clonePptx);
assert.deepEqual(cloneRoundTrip.slides.items.map((slide) => slide.name), ["Original", "Original", "Companion"]);
itemByName(cloneRoundTrip.slides.getItem(1).shapes.items, "clone-copy").text.set("Edited after reimport");
const cloneEditedPptx = await PresentationFile.exportPptx(cloneRoundTrip);
const cloneEditedRoundTrip = await PresentationFile.importPptx(cloneEditedPptx);
assert.equal(itemByName(cloneEditedRoundTrip.slides.getItem(1).shapes.items, "clone-copy").text.value, "Edited after reimport");

const immediateCloneEdit = await PresentationFile.importPptx(cloneSourcePptx);
immediateCloneEdit.slides.getItem(0).duplicate().shapes.items[0].text.set("Too soon");
await assert.rejects(
  () => PresentationFile.exportPptx(immediateCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// The ordinary deck carries media/shape relationships, so delete must stop at
// the C# OPC preflight instead of silently reconstructing a lossy template.
const complexImportedDeletion = await PresentationFile.importPptx(firstExport);
complexImportedDeletion.slides.getItem(0).delete();
await assert.rejects(
  () => PresentationFile.exportPptx(complexImportedDeletion),
  (error) => error?.code === "unsupported_presentation_slide_delete",
);

// Turn the canonical package into a source-bound template marker without
// creating a second writer, then prove its Master/Layout parts survive a
// modeled slide edit byte-for-byte.
const firstZip = await JSZip.loadAsync(new Uint8Array(await firstExport.arrayBuffer()));
const firstSlideXml = await firstZip.file("ppt/slides/slide1.xml").async("text");
assert.match(firstSlideXml, /<a:srcRect[^>]*l="25000"/);
assert.match(firstSlideXml, /<a:srcRect[^>]*r="25000"/);
const authoredChartXml = await Promise.all(Object.keys(firstZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => firstZip.file(name).async("text")));
const comboChartXml = authoredChartXml.find((xml) => xml.includes("Revenue and margin"));
assert.ok(comboChartXml);
assert.match(comboChartXml, /<c:barChart>/);
assert.match(comboChartXml, /<c:lineChart>/);
assert.match(comboChartXml, /<c:barChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:barChart>/);
assert.match(comboChartXml, /<c:lineChart>[\s\S]*?<c:axId val="1"\s*\/><c:axId val="2"\s*\/><\/c:lineChart>/);

// Reference 2.8.24 exposes imported PowerPoint grid spacing, snap settings,
// and guides through presentation.view plus read-only master/layout projections.
const viewPropertiesXml = '<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" lastView="sldView"><p:slideViewPr><p:cSldViewPr snapToGrid="1" snapToObjects="0" showGuides="1"><p:cViewPr varScale="1"><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide orient="vert" pos="2880"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="72008" cy="91440"/></p:viewPr>';
const presentationRelationships = await firstZip.file("ppt/_rels/presentation.xml.rels").async("text");
const viewSource = await PresentationFile.patchPptx(firstExport, [
  {
    path: "ppt/_rels/presentation.xml.rels",
    xml: presentationRelationships.replace("</Relationships>", '<Relationship Id="rIdViewProperties" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/></Relationships>'),
  },
  {
    path: "ppt/viewProps.xml",
    xml: viewPropertiesXml,
    contentType: "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
  },
]);
const importedViewPresentation = await PresentationFile.importPptx(viewSource);
assert.equal(importedViewPresentation.view.gridSpacingCxEmu, 72_008);
assert.equal(importedViewPresentation.view.gridSpacingCyEmu, 91_440);
assert.equal(importedViewPresentation.view.gridlinesVisible, false);
assert.equal(importedViewPresentation.view.guidesVisible, false);
assert.deepEqual(importedViewPresentation.view.toProto(), {
  gridSpacingCxEmu: 72_008,
  gridSpacingCyEmu: 91_440,
  slideViewSnapToGrid: true,
  slideViewSnapToObjects: false,
  slideViewShowGuides: false,
  slideGuides: [
    { orientation: "horizontal", position: 2160 },
    { orientation: "vertical", position: 2880 },
  ],
});
assert.deepEqual(importedViewPresentation.master.slideGuides, importedViewPresentation.view.toProto().slideGuides);
assert.deepEqual(importedViewPresentation.layouts.items[0].slideGuides, importedViewPresentation.view.toProto().slideGuides);
assert.throws(() => importedViewPresentation.layouts.items[0].slideGuides[0].position = 0, TypeError);
assert.equal(importedViewPresentation.view.showGridlines(), undefined);
assert.equal(importedViewPresentation.view.showGuides(), undefined);
assert.equal(importedViewPresentation.view.gridlinesVisible, true);
assert.equal(importedViewPresentation.view.guidesVisible, true);
const viewRoundTripFile = await PresentationFile.exportPptx(importedViewPresentation);
const viewRoundTripZip = await JSZip.loadAsync(viewRoundTripFile.bytes);
assert.equal(await viewRoundTripZip.file("ppt/viewProps.xml").async("text"), viewPropertiesXml);
const viewRoundTrip = await PresentationFile.importPptx(viewRoundTripFile);
assert.deepEqual(viewRoundTrip.view.toProto().slideGuides, importedViewPresentation.view.toProto().slideGuides);
const viewState = viewRoundTrip[Symbol.for("open-office-artifact-tool.open-chestnut-presentation-state")];
viewState.viewProperties.gridSpacingCxEmu = 72_009n;
await assert.rejects(
  () => PresentationFile.exportPptx(viewRoundTrip),
  (error) => error?.code === "unsupported_presentation_view_edit",
);

// Eligible imported top-level OLE objects expose one deliberately narrow edit:
// replacing the uniquely bound XLSX payload. The OLE shell, preview image,
// relationships, source package, and every other native part stay source-owned.
const embeddedSourceWorkbook = Workbook.create();
embeddedSourceWorkbook.worksheets.add("Embedded").getRange("A1").values = [["Original embedded workbook"]];
const embeddedSourceXlsx = await SpreadsheetFile.exportXlsx(embeddedSourceWorkbook);
const embeddedPreviewBytes = Buffer.from(PNG.split(",")[1], "base64");
const oleFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="100" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdEmbeddedWorkbook" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdEmbeddedPreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
const firstSlideRelationships = await firstZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleSource = await PresentationFile.patchPptx(firstExport, [
  { path: "ppt/slides/slide1.xml", xml: firstSlideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: firstSlideRelationships.replace("</Relationships>", '<Relationship Id="rIdEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-workbook.xlsx"/><Relationship Id="rIdEmbeddedPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/agent-workbook-preview.png"/></Relationships>') },
  { path: "ppt/embeddings/agent-workbook.xlsx", bytes: embeddedSourceXlsx.bytes, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { path: "ppt/media/agent-workbook-preview.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const oleSourceSnapshot = Uint8Array.from(oleSource.bytes);
const olePresentation = await PresentationFile.importPptx(oleSource);
const oleObject = itemByName(olePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(oleObject.nativeKind, "oleObject");
assert.equal(oleObject.editable, false);
assert.deepEqual(oleObject.inspectRecord().editableFields, ["embeddedWorkbook"]);
assert.throws(() => { oleObject.oleWorkbook = undefined; }, TypeError);
assert.throws(() => oleObject.setName("Unsafe shell rename"), /read-only/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook("not bytes"), /FileBlob, Uint8Array, ArrayBuffer/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook(new Uint8Array()), /1 through 16777216 bytes/);
assert.throws(() => oleObject.replaceEmbeddedWorkbook(new Uint8Array(16 * 1024 * 1024 + 1)), /1 through 16777216 bytes/);
const extractedSourceWorkbook = await SpreadsheetFile.importXlsx(oleObject.getEmbeddedWorkbook());
assert.equal(extractedSourceWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original embedded workbook");

const embeddedReplacementWorkbook = Workbook.create();
embeddedReplacementWorkbook.worksheets.add("Embedded").getRange("A1:B2").values = [["Replacement workbook", 42], ["Verified", true]];
const embeddedReplacementXlsx = await SpreadsheetFile.exportXlsx(embeddedReplacementWorkbook);
const mutableReplacement = Uint8Array.from(embeddedReplacementXlsx.bytes);
assert.equal(oleObject.replaceEmbeddedWorkbook(mutableReplacement), oleObject);
mutableReplacement.fill(0);
const pendingWorkbookFile = oleObject.getEmbeddedWorkbook();
assert.equal(pendingWorkbookFile.metadata.pendingReplacement, true);
pendingWorkbookFile.bytes.fill(0);
const pendingWorkbook = await SpreadsheetFile.importXlsx(oleObject.getEmbeddedWorkbook());
assert.deepEqual(pendingWorkbook.worksheets.getItem("Embedded").getRange("A1:B2").values, [["Replacement workbook", 42], ["Verified", true]]);
const replacementView = new DataView(embeddedReplacementXlsx.bytes.buffer, embeddedReplacementXlsx.bytes.byteOffset, embeddedReplacementXlsx.bytes.byteLength);
assert.equal(oleObject.replaceEmbeddedWorkbook(replacementView), oleObject);
assert.match(olePresentation.inspect({ kind: "nativeObject", target: oleObject.id, maxChars: 4000 }).ndjson, /"replacementPending":true/);

const oleExport = await PresentationFile.exportPptx(olePresentation);
assert.deepEqual(oleSource.bytes, oleSourceSnapshot);
const oleSourceZipForComparison = await JSZip.loadAsync(oleSource.bytes);
const oleOutputZip = await JSZip.loadAsync(oleExport.bytes);
assert.deepEqual(Object.keys(oleOutputZip.files).sort(), Object.keys(oleSourceZipForComparison.files).sort());
for (const partPath of Object.keys(oleSourceZipForComparison.files)) {
  if (oleSourceZipForComparison.files[partPath].dir || partPath === "ppt/embeddings/agent-workbook.xlsx") continue;
  assert.deepEqual(
    await oleOutputZip.file(partPath).async("uint8array"),
    await oleSourceZipForComparison.file(partPath).async("uint8array"),
    `OLE payload replacement must preserve ${partPath} byte-for-byte`,
  );
}
assert.deepEqual(await oleOutputZip.file("ppt/embeddings/agent-workbook.xlsx").async("uint8array"), embeddedReplacementXlsx.bytes);
assert.deepEqual(await oleOutputZip.file("ppt/media/agent-workbook-preview.png").async("uint8array"), Uint8Array.from(embeddedPreviewBytes));
assert.match(await oleOutputZip.file("ppt/slides/slide1.xml").async("text"), /r:id="rIdEmbeddedWorkbook"/);
const oleRoundTrip = await PresentationFile.importPptx(oleExport);
const reboundOleObject = itemByName(oleRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(reboundOleObject.inspectRecord().embeddedWorkbook.replacementPending, false);
assert.notEqual(reboundOleObject.oleWorkbook.sourceSha256, oleObject.oleWorkbook.sourceSha256);
const reboundWorkbook = await SpreadsheetFile.importXlsx(reboundOleObject.getEmbeddedWorkbook());
assert.deepEqual(reboundWorkbook.worksheets.getItem("Embedded").getRange("A1:B2").values, [["Replacement workbook", 42], ["Verified", true]]);

const invalidOlePresentation = await PresentationFile.importPptx(oleSource);
const invalidOleObject = itemByName(invalidOlePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
invalidOleObject.replaceEmbeddedWorkbook(Uint8Array.from([0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4]));
await assert.rejects(
  () => PresentationFile.exportPptx(invalidOlePresentation),
  (error) => new Set(["invalid_opc_package", "invalid_presentation_ole_workbook"]).has(error?.code),
);

const oleSourceZip = await JSZip.loadAsync(oleSource.bytes);
const oleSourceRelationships = await oleSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const sharedOleSource = await PresentationFile.patchPptx(oleSource, [{
  path: "ppt/slides/_rels/slide1.xml.rels",
  xml: oleSourceRelationships.replace("</Relationships>", '<Relationship Id="rIdSharedEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/agent-workbook.xlsx"/></Relationships>'),
}]);
const sharedOlePresentation = await PresentationFile.importPptx(sharedOleSource);
const sharedOleObject = itemByName(sharedOlePresentation.slides.getItem(0).nativeObjects.items, "Embedded workbook");
assert.equal(sharedOleObject.oleWorkbook, undefined);
assert.deepEqual(sharedOleObject.inspectRecord().editableFields, []);
assert.throws(() => sharedOleObject.getEmbeddedWorkbook(), /has no embedded XLSX workbook/);

const masterPath = "ppt/slideMasters/slideMaster1.xml";
const layoutPath = "ppt/slideLayouts/slideLayout1.xml";
const masterXml = await firstZip.file(masterPath).async("text");
const layoutXml = await firstZip.file(layoutPath).async("text");
const sourceMasterXml = masterXml.replace(/(<p:cSld\b[^>]*\bname=")[^"]*(")/, "$1Source Master Marker$2");
const sourceLayoutXml = layoutXml.replace(/(<p:cSld\b[^>]*\bname=")[^"]*(")/, "$1Source Layout Marker$2");
assert.notEqual(sourceMasterXml, masterXml);
assert.notEqual(sourceLayoutXml, layoutXml);
const sourceBound = await PresentationFile.patchPptx(firstExport, [
  { path: masterPath, xml: sourceMasterXml },
  { path: layoutPath, xml: sourceLayoutXml },
]);

const imported = await PresentationFile.importPptx(sourceBound);
assert.equal(imported.master.name, "Source Master Marker");
assert.equal(imported.layouts.items[0].name, "Source Layout Marker");
imported.master.name = "Unsupported master edit";
await assert.rejects(
  () => PresentationFile.exportPptx(imported),
  /master .*source-bound and read-only/i,
);
imported.master.name = "Source Master Marker";
imported.layouts.items[0].name = "Unsupported layout edit";
await assert.rejects(
  () => PresentationFile.exportPptx(imported),
  /layout .*source-bound and read-only/i,
);
imported.layouts.items[0].name = "Source Layout Marker";
assert.equal(imported.slides.getItem(0).speakerNotes.text, "Lead with the customer outcome.\nThen explain the operating model.");
imported.slides.getItem(0).addNotes("Lead with evidence.\nClose with the decision.");
imported.slides.getItem(1).addNotes("Cannot add a new notes part source-bound");
await assert.rejects(
  () => PresentationFile.exportPptx(imported),
  /cannot add speaker notes to slide 2.*no notes part/i,
);
imported.slides.getItem(1).addNotes("");
const importedCore = imported.slides.getItem(0);
assert.deepEqual(importedCore.background, { fill: "#f1f5f9", mode: "solid" });
assert.equal(itemByName(importedCore.shapes.items, "rounded-card").geometry, "roundRect");
assert.equal(itemByName(importedCore.shapes.items, "target-textbox").geometry, "textbox");
assert.deepEqual(itemByName(importedCore.shapes.items, "rounded-card").shadow, {
  color: "#000000",
  blurRadius: 8,
  distance: 4,
  direction: 45,
  opacity: 0.25,
});
assert.equal(itemByName(importedCore.images.items, "png-image").dataUrl, PNG);
assert.equal(itemByName(importedCore.images.items, "jpeg-image").dataUrl, JPEG);
const importedCover = itemByName(importedCore.images.items, "cover-image");
assert.equal(importedCover.fit, "stretch");
assert.deepEqual(importedCover.crop, { left: 0.25, top: 0, right: 0.25, bottom: 0 });
assert.equal(itemByName(importedCore.tables.items, "fixed-table").values[1][1], "Before");
const importedStraight = itemByName(importedCore.connectors.items, "straight-connector");
const importedElbow = itemByName(importedCore.connectors.items, "elbow-polyline-connector");
assert.equal(importedStraight.line.endArrow, "triangle");
assert.equal(importedElbow.connectorType, "elbow");
assert.equal(importedElbow.line.startArrow, "triangle");
assert.equal(importedElbow.line.endArrow, "triangle");
assert.ok(importedElbow.startTargetId && importedElbow.endTargetId);
const importedRich = itemByName(importedCore.shapes.items, "rich-copy");
assert.equal(importedRich.text.paragraphs[1].bulletCharacter, "•");
assert.deepEqual(importedRich.text.paragraphs[0].runs[1].link, {
  uri: "https://www.ecma-international.org/publications-and-standards/standards/ecma-376/",
});
const importedCharts = imported.slides.getItem(1).charts.items;
assert.deepEqual(importedCharts.map((chart) => chart.chartType), ["bar", "line", "pie"]);
assert.equal(importedCharts[1].series[0].marker.symbol, "circle");
assert.equal(importedCharts[2].dataLabels.showCategoryName, true);
const importedCombo = itemByName(imported.slides.getItem(2).charts.items, "revenue-margin-combo");
assert.equal(importedCombo.chartType, "combo");
assert.deepEqual(importedCombo.series.map((series) => series.chartType), ["bar", "line"]);
assert.equal(importedCombo.series[1].marker.symbol, "circle");
assert.equal(importedCombo.dataLabels.showValue, true);
assert.equal(importedCombo.dataLabels.position, "t");

const importedCard = itemByName(importedCore.shapes.items, "rounded-card");
importedCard.text.set("After edit");
importedCard.shadow.opacity = 0.35;
assert.equal(importedCore.setBackground({ fill: "accent2", mode: "reference", index: 1002 }), importedCore);
assert.equal(imported.slides.getItem(1).setBackground({ fill: "#FFF7ED", mode: "solid" }), imported.slides.getItem(1));
itemByName(importedCore.tables.items, "fixed-table").cells.set(1, 1, "After");
itemByName(importedCore.images.items, "png-image").alt = "Updated PNG evidence";
importedCover.fit = "contain";
importedCover.crop = undefined;
delete importedElbow.line.endArrow;
const editedParagraphs = importedRich.text.paragraphs;
editedParagraphs[0].runs[0].text = "Updated ";
importedRich.text.paragraphs = editedParagraphs;
const importedBar = itemByName(importedCharts, "bar-chart");
importedBar.title = "Updated readiness";
importedBar.series[0].values = [80, 94, 88];
importedCombo.title = "Updated revenue and margin";
importedCombo.series[1].values = [12, 16, 18];

const secondExport = await PresentationFile.exportPptx(imported);
assert.equal(secondExport.metadata.codec, "open-chestnut");
assert.equal((await PresentationFile.inspectPptx(secondExport)).ok, true);
const secondZip = await JSZip.loadAsync(new Uint8Array(await secondExport.arrayBuffer()));
assert.equal(await secondZip.file(masterPath).async("text"), sourceMasterXml);
assert.equal(await secondZip.file(layoutPath).async("text"), sourceLayoutXml);
assert.match(await secondZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /relationships\/slideLayout/);
assert.match(await secondZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /relationships\/slideMaster/);
const secondSlideXml = await secondZip.file("ppt/slides/slide1.xml").async("text");
const secondChartSlideXml = await secondZip.file("ppt/slides/slide2.xml").async("text");
assert.match(secondSlideXml, /<p:bgRef idx="1002">/);
assert.match(secondSlideXml, /<a:schemeClr val="accent2"/);
assert.match(secondChartSlideXml, /<a:srgbClr val="FFF7ED"/);
assert.match(secondSlideXml, /<a:srcRect[^>]*t="-50000"/);
assert.match(secondSlideXml, /<a:srcRect[^>]*b="-50000"/);
assert.match(secondSlideXml, /prst="roundRect"/);
assert.match(secondSlideXml, /txBox="1"/);
assert.match(secondSlideXml, /prst="line"/);
assert.match(secondSlideXml, /prst="bentConnector3"/);
assert.match(secondSlideXml, /<a:headEnd type="triangle"/);
assert.match(secondSlideXml, /<a:tailEnd type="triangle"/);
assert.ok(Object.keys(secondZip.files).some((name) => /\/media\/.+\.png$/.test(name)));
assert.ok(Object.keys(secondZip.files).some((name) => /\/media\/.+\.jpe?g$/.test(name)));
assert.equal(Object.keys(secondZip.files).filter((name) => /\/charts\/chart\d+\.xml$/.test(name)).length, 4);
assert.match(await secondZip.file("ppt/notesSlides/notesSlide1.xml").async("text"), /Lead with evidence/);

const roundTrip = await PresentationFile.importPptx(secondExport);
assert.equal(roundTrip.master.name, "Source Master Marker");
assert.equal(roundTrip.layouts.items[0].name, "Source Layout Marker");
const roundTripCore = roundTrip.slides.getItem(0);
assert.equal(roundTripCore.speakerNotes.text, "Lead with evidence.\nClose with the decision.");
assert.deepEqual(roundTripCore.background, { fill: "accent2", mode: "reference", index: 1002 });
assert.deepEqual(roundTrip.slides.getItem(1).background, { fill: "#fff7ed", mode: "solid" });
assert.equal(itemByName(roundTripCore.shapes.items, "rounded-card").text.value, "After edit");
assert.equal(itemByName(roundTripCore.shapes.items, "rounded-card").shadow.opacity, 0.35);
assert.equal(itemByName(roundTripCore.tables.items, "fixed-table").values[1][1], "After");
assert.equal(itemByName(roundTripCore.images.items, "png-image").alt, "Updated PNG evidence");
const roundTripCover = itemByName(roundTripCore.images.items, "cover-image");
assert.equal(roundTripCover.fit, "stretch");
assert.deepEqual(roundTripCover.crop, { left: 0, top: -0.5, right: 0, bottom: -0.5 });
assert.equal(itemByName(roundTripCore.connectors.items, "elbow-polyline-connector").line.endArrow, undefined);
assert.equal(itemByName(roundTripCore.shapes.items, "rich-copy").text.paragraphs[0].runs[0].text, "Updated ");
const roundTripBar = itemByName(roundTrip.slides.getItem(1).charts.items, "bar-chart");
assert.equal(roundTripBar.title, "Updated readiness");
assert.deepEqual(roundTripBar.series[0].values, [80, 94, 88]);
const roundTripCombo = itemByName(roundTrip.slides.getItem(2).charts.items, "revenue-margin-combo");
assert.equal(roundTripCombo.title, "Updated revenue and margin");
assert.deepEqual(roundTripCombo.series.map((series) => series.chartType), ["bar", "line"]);
assert.deepEqual(roundTripCombo.series[1].values, [12, 16, 18]);
assert.equal(roundTrip.verify().ok, true);

assert.equal(roundTripCore.clearBackground(), roundTripCore);
roundTripCover.fit = "stretch";
roundTripCover.crop = undefined;
const clearedBackgroundExport = await PresentationFile.exportPptx(roundTrip);
const clearedBackgroundZip = await JSZip.loadAsync(new Uint8Array(await clearedBackgroundExport.arrayBuffer()));
assert.doesNotMatch(await clearedBackgroundZip.file("ppt/slides/slide1.xml").async("text"), /<p:bg(?:Pr|Ref)\b/);
assert.doesNotMatch(await clearedBackgroundZip.file("ppt/slides/slide1.xml").async("text"), /<a:srcRect\b/);
const clearedBackgroundRoundTrip = await PresentationFile.importPptx(clearedBackgroundExport);
assert.deepEqual(clearedBackgroundRoundTrip.slides.getItem(0).background, {});
assert.equal(itemByName(clearedBackgroundRoundTrip.slides.getItem(0).images.items, "cover-image").crop, undefined);

const importedWithoutSourceSnapshot = await PresentationFile.importPptx(firstExport);
const presentationState = importedWithoutSourceSnapshot[Symbol.for("open-office-artifact-tool.open-chestnut-presentation-state")];
presentationState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => PresentationFile.exportPptx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

// OpenChestnut owns a deliberately narrow legacy PPTX comment profile: one
// slide-level text item at an explicit coordinate. It never turns the richer
// JS thread facade into a fake element anchor, reply graph, or resolved state.
const legacyCommentDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const legacyCommentSlide = legacyCommentDeck.slides.add({ name: "Legacy comments" });
const legacyCommentThread = legacyCommentSlide.comments.addThread(undefined, "Confirm the source before delivery.", {
  author: "Review Owner",
  created: "2026-07-18T03:05:00Z",
  position: { x: 360, y: 240 },
});
assert.match(legacyCommentDeck.inspect({ kind: "comment" }).ndjson, /Confirm the source before delivery/);
const legacyCommentExport = await PresentationFile.exportPptx(legacyCommentDeck);
const legacyCommentZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentExport.arrayBuffer()));
assert.ok(legacyCommentZip.file("ppt/comments/comment1.xml"));
assert.ok(legacyCommentZip.file("ppt/commentAuthors.xml"));
assert.match(await legacyCommentZip.file("ppt/comments/comment1.xml").async("text"), /Confirm the source before delivery/);
const legacyCommentImported = await PresentationFile.importPptx(legacyCommentExport);
assert.equal(legacyCommentImported.slides.getItem(0).comments.items.length, 1);
const importedLegacyThread = legacyCommentImported.slides.getItem(0).comments.items[0];
assert.equal(importedLegacyThread.nativeFormat, "legacy");
assert.equal(importedLegacyThread.targetId, undefined);
assert.equal(importedLegacyThread.comments.length, 1);
assert.equal(importedLegacyThread.comments[0].author, "Review Owner");
assert.equal(importedLegacyThread.comments[0].text, "Confirm the source before delivery.");
assert.deepEqual(importedLegacyThread.position, { x: 360, y: 240, unit: "px" });
const legacyCommentRoundTrip = await PresentationFile.exportPptx(legacyCommentImported);
const legacyCommentRoundTripZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentRoundTrip.arrayBuffer()));
assert.equal(
  await legacyCommentRoundTripZip.file("ppt/comments/comment1.xml").async("text"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("text"),
);
importedLegacyThread.addReply("Replies are not part of the legacy profile.");
await assert.rejects(
  () => PresentationFile.exportPptx(legacyCommentImported),
  (error) => error?.code === "unsupported_presentation_edit",
);

const invalidLegacyCommentDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const invalidLegacyCommentSlide = invalidLegacyCommentDeck.slides.add();
const invalidLegacyTarget = invalidLegacyCommentSlide.shapes.add({
  geometry: "rect",
  position: { left: 40, top: 40, width: 160, height: 80 },
});
invalidLegacyCommentSlide.comments.addThread(invalidLegacyTarget, "An element anchor is not a legacy comment.", {
  author: "Reviewer",
  position: { x: 120, y: 80 },
});
await assert.rejects(
  () => PresentationFile.exportPptx(invalidLegacyCommentDeck),
  (error) => error?.code === "unsupported_presentation_features",
);
assert.equal(legacyCommentThread.id.startsWith("pc"), true);

console.log("presentation smoke ok");
