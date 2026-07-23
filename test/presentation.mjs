import assert from "node:assert/strict";
import path from "node:path";
import JSZip from "jszip";

import {
  column,
  FileBlob,
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

// Custom shows are a real inline PresentationML graph. Source-free decks own
// the complete list; canonical imports may edit only names and ordered slide
// membership while show topology/native identity remain source-bound.
const customShowDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const customShowOverview = customShowDeck.slides.add({ name: "Overview" });
customShowOverview.shapes.add({
  name: "overview-title",
  position: { left: 80, top: 80, width: 800, height: 80 },
  text: [{ runs: [{ text: "Overview", link: { customShow: "Board route", returnToSlide: true, tooltip: "Open board route" } }] }],
});
const customShowEvidence = customShowDeck.slides.add({ name: "Evidence" });
customShowEvidence.shapes.add({ name: "evidence-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Evidence" });
const customShowAppendix = customShowDeck.slides.add({ name: "Appendix" });
customShowAppendix.shapes.add({ name: "appendix-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Appendix" });
const boardShow = customShowDeck.customShows.add({
  id: "custom-show/board",
  name: "Board route",
  nativeId: 7,
  slides: [customShowOverview, customShowAppendix],
});
customShowDeck.customShows.add({
  id: "custom-show/review",
  name: "Review route",
  nativeId: 11,
  slides: [customShowEvidence],
});
assert.equal(customShowDeck.resolve(boardShow.id), boardShow);
assert.match(customShowDeck.inspect({ kind: "customShow", maxChars: 4000 }).ndjson, /Board route/);
const customShowFirstExport = await PresentationFile.exportPptx(customShowDeck);
const customShowFirstZip = await JSZip.loadAsync(customShowFirstExport.bytes);
const customShowPresentationXml = await customShowFirstZip.file("ppt/presentation.xml").async("string");
assert.match(customShowPresentationXml, /<p:custShowLst>/);
assert.match(customShowPresentationXml, /<p:custShow name="Board route" id="7"><p:sldLst><p:sld r:id="rIdSlide1"[^>]*\/><p:sld r:id="rIdSlide3"[^>]*\/><\/p:sldLst><\/p:custShow>/);
assert.ok(customShowPresentationXml.indexOf("<p:custShowLst>") < customShowPresentationXml.indexOf("<p:defaultTextStyle"));
const customShowSlideXml = await customShowFirstZip.file("ppt/slides/slide1.xml").async("string");
assert.match(customShowSlideXml, /<a:hlinkClick r:id=""[^>]*action="ppaction:\/\/customshow\?id=7&amp;return=true"[^>]*tooltip="Open board route"/);

const customShowImported = await PresentationFile.importPptx(customShowFirstExport);
assert.equal(customShowImported.customShows.count, 2);
assert.deepEqual(customShowImported.customShows.getItem("Board route").slideIds, [customShowImported.slides.items[0].id, customShowImported.slides.items[2].id]);
assert.deepEqual(itemByName(customShowImported.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, {
  customShow: "Board route",
  returnToSlide: true,
  tooltip: "Open board route",
});
const editableBoardShow = customShowImported.customShows.getItem("Board route");
editableBoardShow.name = "Executive route";
editableBoardShow.setSlides([customShowImported.slides.items[2], customShowImported.slides.items[0], customShowImported.slides.items[2]]);
const customShowEditedExport = await PresentationFile.exportPptx(customShowImported);
const customShowEditedRoundTrip = await PresentationFile.importPptx(customShowEditedExport);
assert.equal(customShowEditedRoundTrip.customShows.count, 2);
assert.deepEqual(customShowEditedRoundTrip.customShows.getItem("Executive route").slideIds, [
  customShowEditedRoundTrip.slides.items[2].id,
  customShowEditedRoundTrip.slides.items[0].id,
  customShowEditedRoundTrip.slides.items[2].id,
]);
assert.equal(
  itemByName(customShowEditedRoundTrip.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link.customShow,
  "Executive route",
  "renaming a custom show must retain the native link identity and refresh its public display name",
);
const customShowEditedZip = await JSZip.loadAsync(customShowEditedExport.bytes);
for (const partPath of Object.keys(customShowFirstZip.files).filter((entry) => !customShowFirstZip.files[entry].dir && entry !== "ppt/presentation.xml")) {
  assert.deepEqual(
    await customShowEditedZip.file(partPath).async("uint8array"),
    await customShowFirstZip.file(partPath).async("uint8array"),
    `custom-show edit changed non-presentation part ${partPath}`,
  );
}
const customShowRetargeted = await PresentationFile.importPptx(customShowFirstExport);
const customShowRetargetedShape = itemByName(customShowRetargeted.slides.items[0].shapes.items, "overview-title");
const customShowRetargetedParagraph = customShowRetargetedShape.text.paragraphs[0];
customShowRetargetedParagraph.runs[0].link = {
  customShow: "Review route",
  returnToSlide: false,
};
customShowRetargetedShape.text.paragraphs = [customShowRetargetedParagraph];
const customShowRetargetedExport = await PresentationFile.exportPptx(customShowRetargeted);
const customShowRetargetedZip = await JSZip.loadAsync(customShowRetargetedExport.bytes);
assert.match(await customShowRetargetedZip.file("ppt/slides/slide1.xml").async("string"), /action="ppaction:\/\/customshow\?id=11&amp;return=false"/);
const customShowRetargetedRoundTrip = await PresentationFile.importPptx(customShowRetargetedExport);
assert.deepEqual(itemByName(customShowRetargetedRoundTrip.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, {
  customShow: "Review route",
  returnToSlide: false,
});
const missingCustomShowLink = Presentation.create({ slideSize: { width: 640, height: 360 } });
missingCustomShowLink.slides.add().shapes.add({
  position: { left: 40, top: 40, width: 360, height: 80 },
  text: [{ runs: [{ text: "Missing", link: { customShow: "Not present" } }] }],
});
await assert.rejects(
  () => PresentationFile.exportPptx(missingCustomShowLink),
  (error) => error?.code === "invalid_presentation_hyperlink",
);
const customShowCloneImport = await PresentationFile.importPptx(customShowFirstExport);
customShowCloneImport.slides.items[0].duplicate();
customShowCloneImport.customShows.getItem("Board route").name = "Cloned executive route";
const customShowCloneExport = await PresentationFile.exportPptx(customShowCloneImport);
const customShowCloneZip = await JSZip.loadAsync(customShowCloneExport.bytes);
assert.deepEqual(
  await customShowCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await customShowFirstZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a custom-show run link must preserve the retained source SlidePart byte-for-byte",
);
assert.deepEqual(
  await customShowCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await customShowFirstZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "a relationship-free custom-show action must not rewrite the retained source relationship graph",
);
assert.match(
  await customShowCloneZip.file("ppt/slides/slide4.xml").async("string"),
  /<a:hlinkClick\b[^>]*r:id=""[^>]*action="ppaction:\/\/customshow\?id=7&amp;return=true"[^>]*tooltip="Open board route"/,
);
assert.doesNotMatch(
  await customShowCloneZip.file("ppt/slides/_rels/slide4.xml.rels").async("string"),
  /relationships\/(?:hyperlink|slide)"/,
  "custom-show run links must remain relationship-free on the cloned SlidePart",
);
const customShowCloneRoundTrip = await PresentationFile.importPptx(customShowCloneExport);
assert.deepEqual(customShowCloneRoundTrip.slides.items.map((slide) => slide.name), ["Overview", "Overview", "Evidence", "Appendix"]);
assert.deepEqual(customShowCloneRoundTrip.customShows.getItem("Cloned executive route").slideIds, [
  customShowCloneRoundTrip.slides.items[0].id,
  customShowCloneRoundTrip.slides.items[3].id,
]);
assert.ok(!customShowCloneRoundTrip.customShows.getItem("Cloned executive route").slideIds.includes(customShowCloneRoundTrip.slides.items[1].id));
for (const slideIndex of [0, 1]) {
  assert.deepEqual(
    itemByName(customShowCloneRoundTrip.slides.items[slideIndex].shapes.items, "overview-title").text.paragraphs[0].runs[0].link,
    { customShow: "Cloned executive route", returnToSlide: true, tooltip: "Open board route" },
  );
}
const customShowCompoundCloneImport = await PresentationFile.importPptx(customShowFirstExport);
const customShowCompoundClone = customShowCompoundCloneImport.slides.items[0].duplicate();
customShowCompoundCloneImport.customShows.getItem("Board route").setSlides([customShowCompoundClone]);
await assert.rejects(
  () => PresentationFile.exportPptx(customShowCompoundCloneImport),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "custom-show membership changes must cross a separate export/reimport boundary from slide cloning",
);
const customShowMovedImport = await PresentationFile.importPptx(customShowFirstExport);
customShowMovedImport.slides.items[2].moveTo(0);
const customShowMovedRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(customShowMovedImport));
assert.deepEqual(customShowMovedRoundTrip.slides.items.map((slide) => slide.name), ["Appendix", "Overview", "Evidence"]);
assert.deepEqual(customShowMovedRoundTrip.customShows.getItem("Board route").slides.map((slide) => slide.name), ["Overview", "Appendix"]);

const customShowAddedImport = await PresentationFile.importPptx(customShowFirstExport);
customShowAddedImport.customShows.add("Added route", [customShowAddedImport.slides.items[0]]);
await assert.rejects(
  () => PresentationFile.exportPptx(customShowAddedImport),
  (error) => error?.code === "presentation_custom_show_topology_changed",
);
const customShowIdentityImport = await PresentationFile.importPptx(customShowFirstExport);
customShowIdentityImport.customShows.items[0].nativeId = 99;
await assert.rejects(
  () => PresentationFile.exportPptx(customShowIdentityImport),
  (error) => error?.code === "presentation_custom_show_topology_changed",
);

const irregularCustomShowZip = await JSZip.loadAsync(customShowFirstExport.bytes);
const irregularCustomShowXml = (await irregularCustomShowZip.file("ppt/presentation.xml").async("string"))
  .replace("</p:custShow>", "<p:extLst/></p:custShow>");
irregularCustomShowZip.file("ppt/presentation.xml", irregularCustomShowXml);
const irregularCustomShowBytes = await irregularCustomShowZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const irregularCustomShowFile = new FileBlob(irregularCustomShowBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const irregularCustomShowImport = await PresentationFile.importPptx(irregularCustomShowFile);
assert.equal(irregularCustomShowImport.customShows.count, 0);
assert.equal(itemByName(irregularCustomShowImport.slides.items[0].shapes.items, "overview-title").text.paragraphs[0].runs[0].link, undefined);
const irregularLinkMutation = await PresentationFile.importPptx(irregularCustomShowFile);
const irregularLinkShape = itemByName(irregularLinkMutation.slides.items[0].shapes.items, "overview-title");
const irregularLinkParagraph = irregularLinkShape.text.paragraphs[0];
irregularLinkParagraph.runs[0].link = { uri: "https://example.com/replacement" };
irregularLinkShape.text.paragraphs = [irregularLinkParagraph];
await assert.rejects(
  () => PresentationFile.exportPptx(irregularLinkMutation),
  (error) => error?.code === "unsupported_presentation_edit",
);
irregularCustomShowImport.customShows.add("Unsafe replacement", [irregularCustomShowImport.slides.items[0]]);
await assert.rejects(
  () => PresentationFile.exportPptx(irregularCustomShowImport),
  (error) => error?.code === "unsupported_presentation_custom_show_edit",
);

// PowerPoint sections are a p14 extension graph, not a custom-show route:
// they partition the complete ordered slide sequence through native p:sldId
// values and keep their GUID identity fixed after an import.
const sectionDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const sectionOpening = sectionDeck.slides.add({ name: "Opening" });
sectionOpening.shapes.add({ name: "section-opening", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Opening" });
const sectionEvidence = sectionDeck.slides.add({ name: "Evidence" });
sectionEvidence.shapes.add({ name: "section-evidence", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Evidence" });
const sectionAppendix = sectionDeck.slides.add({ name: "Appendix" });
sectionAppendix.shapes.add({ name: "section-appendix", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Appendix" });
const openingSection = sectionDeck.sections.add({
  id: "section/opening",
  name: "Opening",
  nativeId: "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}",
  slides: [sectionOpening],
});
sectionDeck.sections.add({
  id: "section/content",
  name: "Content",
  nativeId: "{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274}",
  slides: [sectionEvidence, sectionAppendix],
});
assert.equal(sectionDeck.resolve(openingSection.id), openingSection);
assert.match(sectionDeck.inspect({ kind: "section", maxChars: 4000 }).ndjson, /Opening/);
assert.equal(sectionDeck.verify().ok, true);
const sectionFirstExport = await PresentationFile.exportPptx(sectionDeck);
const sectionFirstZip = await JSZip.loadAsync(sectionFirstExport.bytes);
const sectionPresentationXml = await sectionFirstZip.file("ppt/presentation.xml").async("string");
assert.match(sectionPresentationXml, /<p:ext uri="\{521415D9-36F7-43E2-AB2F-B90AF26B5E84\}"><p14:sectionLst/);
assert.match(sectionPresentationXml, /<p14:section name="Opening" id="\{01F07B81-39E6-4BBB-9B89-66EA253FBD29\}"><p14:sldIdLst><p14:sldId id="256" \/><\/p14:sldIdLst><\/p14:section>/);
assert.match(sectionPresentationXml, /<p14:section name="Content" id="\{1FEF2C88-0CF2-4176-BA81-0DE6FD9D1274\}"><p14:sldIdLst><p14:sldId id="257" \/><p14:sldId id="258" \/><\/p14:sldIdLst><\/p14:section>/);
const sectionImported = await PresentationFile.importPptx(sectionFirstExport);
assert.equal(sectionImported.sections.count, 2);
assert.deepEqual(sectionImported.sections.getItem("Opening").slideIds, [sectionImported.slides.items[0].id]);
assert.deepEqual(sectionImported.sections.getItem("Content").slideIds, [sectionImported.slides.items[1].id, sectionImported.slides.items[2].id]);
const editableOpeningSection = sectionImported.sections.getItem("Opening");
editableOpeningSection.name = "Introduction";
editableOpeningSection.setSlides([sectionImported.slides.items[0], sectionImported.slides.items[1]]);
sectionImported.sections.getItem("Content").setSlides([sectionImported.slides.items[2]]);
const sectionEditedExport = await PresentationFile.exportPptx(sectionImported);
const sectionEditedRoundTrip = await PresentationFile.importPptx(sectionEditedExport);
assert.equal(sectionEditedRoundTrip.sections.getItem("Introduction").nativeId, "{01F07B81-39E6-4BBB-9B89-66EA253FBD29}");
assert.deepEqual(sectionEditedRoundTrip.sections.getItem("Introduction").slides.map((slide) => slide.name), ["Opening", "Evidence"]);
assert.deepEqual(sectionEditedRoundTrip.sections.getItem("Content").slides.map((slide) => slide.name), ["Appendix"]);
const sectionEditedZip = await JSZip.loadAsync(sectionEditedExport.bytes);
for (const partPath of Object.keys(sectionFirstZip.files).filter((entry) => !sectionFirstZip.files[entry].dir && entry !== "ppt/presentation.xml")) {
  assert.deepEqual(
    await sectionEditedZip.file(partPath).async("uint8array"),
    await sectionFirstZip.file(partPath).async("uint8array"),
    `section edit changed non-presentation part ${partPath}`,
  );
}
const sectionAddedImport = await PresentationFile.importPptx(sectionFirstExport);
sectionAddedImport.sections.add("Unsafe", [sectionAddedImport.slides.items[0], sectionAddedImport.slides.items[1], sectionAddedImport.slides.items[2]]);
await assert.rejects(
  () => PresentationFile.exportPptx(sectionAddedImport),
  (error) => error?.code === "presentation_section_topology_changed",
);
const sectionCloneImport = await PresentationFile.importPptx(sectionFirstExport);
sectionCloneImport.slides.items[0].duplicate();
await assert.rejects(
  () => PresentationFile.exportPptx(sectionCloneImport),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
const irregularSectionZip = await JSZip.loadAsync(sectionFirstExport.bytes);
const irregularSectionXml = (await irregularSectionZip.file("ppt/presentation.xml").async("string"))
  .replace("</p14:section>", "<p14:extLst/></p14:section>");
irregularSectionZip.file("ppt/presentation.xml", irregularSectionXml);
const irregularSectionBytes = await irregularSectionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const irregularSectionFile = new FileBlob(irregularSectionBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const irregularSectionImport = await PresentationFile.importPptx(irregularSectionFile);
assert.equal(irregularSectionImport.sections.count, 0);
irregularSectionImport.sections.add("Unsafe", irregularSectionImport.slides.items);
await assert.rejects(
  () => PresentationFile.exportPptx(irregularSectionImport),
  (error) => error?.code === "unsupported_presentation_section_edit",
);

// Slide transitions are a direct p:transition leaf, deliberately distinct
// from animation/timing graphs. The bounded profile owns fade plus four-way
// push, explicit speed/click behavior, and an optional auto-advance timer.
const transitionDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const transitionFade = transitionDeck.slides.add({ name: "Fade" });
transitionFade.shapes.add({ name: "transition-fade-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Fade" });
transitionFade.setTransition({ effect: "fade", speed: "medium", advanceOnClick: true, advanceAfterMs: 1250 });
const transitionPush = transitionDeck.slides.add({ name: "Push" });
transitionPush.shapes.add({ name: "transition-push-title", position: { left: 80, top: 80, width: 800, height: 80 }, text: "Push" });
transitionPush.setTransition({ effect: "push", direction: "right", speed: "fast", advanceOnClick: false, advanceAfterMs: 0 });
assert.equal(transitionDeck.resolve(`${transitionFade.id}/transition`), transitionFade.transition);
assert.deepEqual(transitionFade.transition.toJSON(), { effect: "fade", speed: "medium", advanceOnClick: true, advanceAfterMs: 1250 });
assert.match(transitionDeck.inspect({ kind: "transition", maxChars: 4000 }).ndjson, /"effect":"push"/);
assert.throws(() => transitionFade.setTransition({ effect: "fade", direction: "left" }), /does not accept direction/);
assert.throws(() => transitionPush.setTransition({ effect: "push", direction: "diagonal" }), /left, up, right, or down/);
assert.throws(() => transitionPush.setTransition({ effect: "wipe" }), /fade or push/);
const transitionFirstExport = await PresentationFile.exportPptx(transitionDeck);
const transitionFirstZip = await JSZip.loadAsync(transitionFirstExport.bytes);
const transitionFadeXml = await transitionFirstZip.file("ppt/slides/slide1.xml").async("string");
const transitionPushXml = await transitionFirstZip.file("ppt/slides/slide2.xml").async("string");
assert.match(transitionFadeXml, /<p:transition spd="med" advClick="1" advTm="1250"><p:fade \/><\/p:transition>/);
assert.match(transitionPushXml, /<p:transition spd="fast" advClick="0" advTm="0"><p:push dir="r" \/><\/p:transition>/);
const transitionImported = await PresentationFile.importPptx(transitionFirstExport);
assert.deepEqual(transitionImported.slides.items[0].transition.toJSON(), { effect: "fade", speed: "medium", advanceOnClick: true, advanceAfterMs: 1250 });
assert.deepEqual(transitionImported.slides.items[1].transition.toJSON(), { effect: "push", direction: "right", speed: "fast", advanceOnClick: false, advanceAfterMs: 0 });
assert.deepEqual(transitionImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: true, addable: false });
transitionImported.slides.items[0].setTransition({ effect: "push", direction: "down", speed: "slow", advanceOnClick: true });
const transitionEditedExport = await PresentationFile.exportPptx(transitionImported);
const transitionEdited = await PresentationFile.importPptx(transitionEditedExport);
assert.deepEqual(transitionEdited.slides.items[0].transition.toJSON(), { effect: "push", direction: "down", speed: "slow", advanceOnClick: true });
transitionEdited.slides.items[1].clearTransition();
const transitionClearedExport = await PresentationFile.exportPptx(transitionEdited);
const transitionCleared = await PresentationFile.importPptx(transitionClearedExport);
assert.equal(transitionCleared.slides.items[1].transition.configured, false);

// The clone is an exact new SlidePart on first export, so a modeled direct
// transition travels with the clone but cannot be changed before reimport.
const transitionCloneSource = await PresentationFile.importPptx(transitionFirstExport);
const transitionClone = transitionCloneSource.slides.items[0].duplicate();
assert.deepEqual(transitionClone.transition.toJSON(), transitionCloneSource.slides.items[0].transition.toJSON());
const transitionCloneExport = await PresentationFile.exportPptx(transitionCloneSource);
const transitionCloneRoundTrip = await PresentationFile.importPptx(transitionCloneExport);
assert.deepEqual(transitionCloneRoundTrip.slides.items[0].transition.toJSON(), transitionCloneRoundTrip.slides.items[1].transition.toJSON());

const transitionAbsentDeck = Presentation.create();
transitionAbsentDeck.slides.add({ name: "No transition" }).shapes.add({ text: "No transition" });
const transitionAbsentImported = await PresentationFile.importPptx(await PresentationFile.exportPptx(transitionAbsentDeck));
assert.deepEqual(transitionAbsentImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: false, editable: false, addable: false });
assert.throws(
  () => transitionAbsentImported.slides.items[0].setTransition({ effect: "fade" }),
  /source-bound/,
);

const opaqueTransitionZip = await JSZip.loadAsync(transitionFirstExport.bytes);
opaqueTransitionZip.file("ppt/slides/slide1.xml", transitionFadeXml.replace(/<p:fade\s*\/>/, '<p:wipe dir="l"/>'));
const opaqueTransitionBytes = await opaqueTransitionZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const opaqueTransitionFile = new FileBlob(opaqueTransitionBytes, { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const opaqueTransitionImported = await PresentationFile.importPptx(opaqueTransitionFile);
assert.equal(opaqueTransitionImported.slides.items[0].transition.configured, false);
assert.deepEqual(opaqueTransitionImported.slides.items[0].transition.capability, { sourceBound: true, partPresent: true, editable: false, addable: false });
assert.throws(
  () => opaqueTransitionImported.slides.items[0].setTransition({ effect: "fade" }),
  /source-bound/,
);
const opaqueTransitionPreserved = await PresentationFile.exportPptx(opaqueTransitionImported);
const opaqueTransitionPreservedZip = await JSZip.loadAsync(opaqueTransitionPreserved.bytes);
assert.match(await opaqueTransitionPreservedZip.file("ppt/slides/slide1.xml").async("string"), /<p:wipe dir="l"\/>/);
// Negative DrawingML offsets are retained only for an imported opaque,
// source-bound element. New authoring still rejects them instead of widening
// the public source-free layout profile.
const negativeSourceFreeFrame = Presentation.create();
negativeSourceFreeFrame.slides.add().shapes.add({
  name: "negative-source-free-frame",
  position: { left: -1, top: 0, width: 100, height: 40 },
  text: "must fail before export",
});
await assert.rejects(
  () => PresentationFile.exportPptx(negativeSourceFreeFrame),
  (error) => error?.code === "invalid_presentation_frame",
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
      text: [{ runs: [{ text: "Master title prompt", link: { customShow: "Layout route", returnToSlide: true } }] }],
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
  text: [{ runs: [{ text: "Master body prompt", link: { customShow: "Layout route" } }] }],
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
authoredLayoutPresentation.customShows.add("Layout route", [authoredLayoutSlide]);
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
assert.match(authoredMasterXml, /action="ppaction:\/\/customshow\?id=0&amp;return=true"/);
assert.match(authoredLayoutXml, /<p:sldLayout[^>]*type="obj"/);
assert.match(authoredLayoutXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
assert.match(authoredLayoutXml, /action="ppaction:\/\/customshow\?id=0"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="title"[^>]*idx="0"/);
assert.match(authoredSlideXml, /<p:ph[^>]*type="body"[^>]*idx="1"/);
const authoredLayoutImported = await PresentationFile.importPptx(authoredLayoutExport);
assert.equal(authoredLayoutImported.master.placeholders.length, 1);
assert.equal(authoredLayoutImported.layouts.items.length, 1);
assert.equal(authoredLayoutImported.layouts.items[0].type, "obj");
assert.equal(authoredLayoutImported.master.placeholders[0].text[0].runs[0].link.customShow, "Layout route");
assert.equal(authoredLayoutImported.layouts.items[0].placeholders[0].text[0].runs[0].link.customShow, "Layout route");
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

const chartFamilyDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const chartFamilySlide = chartFamilyDeck.slides.add({ name: "Native chart families" });
chartFamilySlide.charts.add("area", {
  name: "area-family",
  title: "Regional trajectory",
  position: { left: 40, top: 35, width: 570, height: 300 },
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Revenue", values: [42, 53, 68], fill: "#0EA5E9", line: { fill: "#0369A1", width: 1.5 } }],
  xAxis: { title: "Quarter" },
  yAxis: { title: "Revenue", min: 0, max: 80, majorUnit: 20 },
  legend: false,
});
chartFamilySlide.charts.add("doughnut", {
  name: "doughnut-family",
  title: "Regional mix",
  position: { left: 660, top: 35, width: 570, height: 300 },
  categories: ["North", "Central", "South"],
  series: [{ name: "Share", values: [52, 31, 17] }],
  dataLabels: { showCategoryName: true, showPercent: true, position: "outsideEnd" },
  legend: true,
});
chartFamilySlide.charts.add("scatter", {
  name: "scatter-family",
  title: "Reach relationship",
  position: { left: 40, top: 370, width: 570, height: 300 },
  series: [{ name: "Portfolio", xValues: [10, 20, 34], values: [35, 68, 84], marker: { symbol: "diamond", size: 8, fill: "#8B5CF6", line: { fill: "#6D28D9", width: 1 } } }],
  xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
  yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
  legend: false,
});
chartFamilySlide.charts.add("bubble", {
  name: "bubble-family",
  title: "Opportunity map",
  position: { left: 660, top: 370, width: 570, height: 300 },
  series: [{ name: "Opportunity", xValues: [10, 20, 34], values: [35, 68, 84], bubbleSizes: [4, 9, 16], fill: "#F97316", line: { fill: "#C2410C", width: 1 } }],
  xAxis: { title: "Reach", min: 0, max: 40, majorUnit: 10 },
  yAxis: { title: "Return", min: 0, max: 100, majorUnit: 20 },
  legend: false,
});
assert.equal(chartFamilyDeck.verify().ok, true);
const chartFamilySvg = chartFamilySlide.toSvg();
assert.match(chartFamilySvg, /Regional trajectory/);
assert.match(chartFamilySvg, /52%/);
assert.match(chartFamilySvg, /<circle[^>]+fill-opacity="0\.72"/);
assert.match(chartFamilySvg, /<path[^>]+fill-opacity="0\.45"/);
const chartFamilyExport = await PresentationFile.exportPptx(chartFamilyDeck);
const chartFamilyZip = await JSZip.loadAsync(new Uint8Array(await chartFamilyExport.arrayBuffer()));
const chartFamilyXml = await Promise.all(Object.keys(chartFamilyZip.files)
  .filter((name) => /\/charts\/chart\d+\.xml$/.test(name))
  .map((name) => chartFamilyZip.file(name).async("text")));
assert.equal(chartFamilyXml.filter((xml) => /<c:areaChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:doughnutChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:scatterChart>/.test(xml)).length, 1);
assert.equal(chartFamilyXml.filter((xml) => /<c:bubbleChart>/.test(xml)).length, 1);
assert.match(chartFamilyXml.find((xml) => /<c:doughnutChart>/.test(xml)), /<c:showPercent val="1"\s*\/>/);
assert.match(chartFamilyXml.find((xml) => /<c:scatterChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>/);
assert.match(chartFamilyXml.find((xml) => /<c:bubbleChart>/.test(xml)), /<c:xVal>[\s\S]*<c:yVal>[\s\S]*<c:bubbleSize>/);
const importedChartFamilyDeck = await PresentationFile.importPptx(chartFamilyExport);
const importedFamilies = importedChartFamilyDeck.slides.getItem(0).charts.items;
assert.deepEqual(importedFamilies.map((chart) => chart.chartType), ["area", "doughnut", "scatter", "bubble"]);
assert.equal(importedFamilies[1].dataLabels.showPercent, true);
assert.deepEqual(importedFamilies[2].series[0].xValues, [10, 20, 34]);
assert.deepEqual(importedFamilies[3].series[0].bubbleSizes, [4, 9, 16]);
importedFamilies[0].series[0].values[1] = 57;
importedFamilies[1].dataLabels.showPercent = false;
importedFamilies[2].series[0].xValues[1] = 22;
importedFamilies[3].series[0].bubbleSizes[1] = 12;
const editedChartFamilyExport = await PresentationFile.exportPptx(importedChartFamilyDeck);
const roundTripChartFamilies = (await PresentationFile.importPptx(editedChartFamilyExport)).slides.getItem(0).charts.items;
assert.equal(roundTripChartFamilies[0].series[0].values[1], 57);
assert.equal(roundTripChartFamilies[1].dataLabels.showPercent, false);
assert.equal(roundTripChartFamilies[2].series[0].xValues[1], 22);
assert.equal(roundTripChartFamilies[3].series[0].bubbleSizes[1], 12);
assert.throws(() => chartFamilySlide.charts.add("scatter", { categories: ["A"], series: [{ name: "Invalid", xValues: [1], values: [2] }] }), /per-series xValues/i);
assert.throws(() => chartFamilySlide.charts.add("bubble", { series: [{ name: "Invalid", xValues: [1], values: [2], bubbleSizes: [0] }] }), /positive bubbleSize/i);
assert.throws(() => chartFamilySlide.charts.add("doughnut", { categories: ["A"], series: [{ name: "Invalid", values: [1] }], xAxis: { title: "Invalid" } }), /cannot carry axes/i);
assert.throws(() => chartFamilySlide.charts.add("area", { categories: ["A"], series: [{ name: "Invalid marker", values: [1], marker: { symbol: "circle" } }] }), /area series 1 cannot carry a marker/i);
const scatterLineDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
scatterLineDeck.slides.add().charts.add("scatter", { series: [{ name: "Invalid line", xValues: [1, 2], values: [2, 3], line: { fill: "#000000", width: 1 } }] });
await assert.rejects(PresentationFile.exportPptx(scatterLineDeck), /marker-scatter.*cannot carry a series line/i);

const singleAxisChartDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
singleAxisChartDeck.slides.add().charts.add("bar", {
  categories: ["A", "B"],
  series: [{ name: "Values", values: [1, 2] }],
  yAxis: { title: "Configured value axis" },
});
const singleAxisChartRoundTrip = await PresentationFile.importPptx(await PresentationFile.exportPptx(singleAxisChartDeck));
assert.equal(singleAxisChartRoundTrip.slides.getItem(0).charts.items[0].axes.category.title, "");
assert.equal(singleAxisChartRoundTrip.slides.getItem(0).charts.items[0].axes.value.title, "Configured value axis");

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

// Speaker notes use the same paragraph/run model as visible slide text, but
// retain a deliberately narrower relationship-free contract. This proves the
// public facade can author, reimport, and edit a multi-run talk track without
// flattening it through the legacy `.text` convenience field.
const richNotesDeck = Presentation.create({ slideSize: { width: 640, height: 360 } });
const richNotesSlide = richNotesDeck.slides.add({
  name: "Rich speaker notes",
  notes: [
    {
      bulletCharacter: "•",
      runs: [
        { text: "Open with ", style: { bold: true, fontSize: 18, fontFamily: "Aptos", color: "#0F172A" } },
        { text: "the customer outcome.", style: { italic: true, fontSize: 18 } },
      ],
    },
    { autoNumber: { type: "arabicPeriod", startAt: 2 }, runs: [{ text: "Then explain the operating model.", style: { fontSize: 16 } }] },
  ],
});
richNotesSlide.shapes.add({ name: "rich-notes-title", text: "Visible slide", position: { left: 48, top: 48, width: 300, height: 72 } });
const richNotesPptx = await PresentationFile.exportPptx(richNotesDeck);
const richNotesZip = await JSZip.loadAsync(richNotesPptx.bytes);
const richNotesXml = await richNotesZip.file("ppt/notesSlides/notesSlide1.xml").async("text");
assert.match(richNotesXml, /<a:buChar\b[^>]*char="•"/);
assert.match(richNotesXml, /<a:rPr\b[^>]*\bb="1"/);
assert.match(richNotesXml, /<a:rPr\b[^>]*\bi="1"/);
assert.match(richNotesXml, /<a:buAutoNum\b[^>]*type="arabicPeriod"[^>]*startAt="2"/);
const importedRichNotesDeck = await PresentationFile.importPptx(richNotesPptx);
const importedRichNotes = importedRichNotesDeck.slides.getItem(0).speakerNotes;
assert.equal(importedRichNotes.text, "Open with the customer outcome.\nThen explain the operating model.");
assert.equal(importedRichNotes.capability.editable, true);
const importedRichParagraphs = importedRichNotes.textFrame.paragraphs;
assert.equal(importedRichParagraphs.length, 2);
assert.equal(importedRichParagraphs[0].runs.length, 2);
assert.equal(importedRichParagraphs[0].runs[0].style.bold, true);
assert.equal(importedRichParagraphs[0].runs[1].style.italic, true);
assert.deepEqual(importedRichParagraphs[1].autoNumber, { type: "arabicPeriod", startAt: 2 });
const richNotesNoOpPptx = await PresentationFile.exportPptx(importedRichNotesDeck);
const richNotesNoOpZip = await JSZip.loadAsync(richNotesNoOpPptx.bytes);
assert.deepEqual(
  await richNotesNoOpZip.file("ppt/notesSlides/notesSlide1.xml").async("uint8array"),
  await richNotesZip.file("ppt/notesSlides/notesSlide1.xml").async("uint8array"),
  "unchanged imported rich notes must retain their source NotesSlide bytes",
);
importedRichParagraphs[0].runs[1].text = "the operating decision.";
importedRichParagraphs[0].runs[1].style = { ...importedRichParagraphs[0].runs[1].style, bold: true, italic: false };
importedRichNotes.textFrame.paragraphs = importedRichParagraphs;
const richNotesEditedPptx = await PresentationFile.exportPptx(importedRichNotesDeck);
const richNotesEditedDeck = await PresentationFile.importPptx(richNotesEditedPptx);
const editedRichNotes = richNotesEditedDeck.slides.getItem(0).speakerNotes;
assert.equal(editedRichNotes.text, "Open with the operating decision.\nThen explain the operating model.");
assert.equal(editedRichNotes.textFrame.paragraphs[0].runs[1].style.bold, true);
assert.equal(editedRichNotes.textFrame.paragraphs[0].runs[1].style.italic, false);
const richNotesFlattenAttempt = await PresentationFile.importPptx(richNotesPptx);
richNotesFlattenAttempt.slides.getItem(0).speakerNotes.text = "Do not flatten this multi-run talk track.";
await assert.rejects(
  () => PresentationFile.exportPptx(richNotesFlattenAttempt),
  (error) => error?.code === "presentation_text_topology_changed",
);

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

// A retained imported SlidePart may change only its native p:cSld/@name. The
// transaction must leave every other decoded package part byte-for-byte intact.
const renamedImportedDeck = await PresentationFile.importPptx(firstExport);
const renamedImportedSlide = renamedImportedDeck.slides.getItem(0);
renamedImportedSlide.name = "Renamed imported overview";
const renamedImportedPptx = await PresentationFile.exportPptx(renamedImportedDeck);
const renamedImportedZip = await JSZip.loadAsync(renamedImportedPptx.bytes);
assert.deepEqual(Object.keys(renamedImportedZip.files).sort(), Object.keys(originalImportedZip.files).sort());
for (const [path, entry] of Object.entries(originalImportedZip.files)) {
  if (entry.dir || path === "ppt/slides/slide1.xml") continue;
  assert.deepEqual(
    await renamedImportedZip.file(path).async("uint8array"),
    await originalImportedZip.file(path).async("uint8array"),
    `renaming an imported slide must preserve ${path} byte-for-byte`,
  );
}
assert.match(await renamedImportedZip.file("ppt/slides/slide1.xml").async("text"), /<p:cSld\b[^>]*\bname="Renamed imported overview"/);
const renamedImportedRoundTrip = await PresentationFile.importPptx(renamedImportedPptx);
assert.equal(renamedImportedRoundTrip.slides.getItem(0).name, "Renamed imported overview");
assert.equal(itemByName(renamedImportedRoundTrip.slides.getItem(0).shapes.items, "rounded-card").text.value, "Before edit");

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
assert.throws(
  () => importedCloneSource.duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
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

// Canonical run hyperlinks are a modeled part of the same bounded clone leaf.
// The clone keeps the source XML r:ids, creates equivalent external/internal
// relationships on its fresh SlidePart, and retains action-only links inline.
const hyperlinkCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const hyperlinkCloneOrigin = hyperlinkCloneFixture.slides.add({ name: "Linked origin" });
const hyperlinkCloneTarget = hyperlinkCloneFixture.slides.add({ name: "Linked target" });
hyperlinkCloneTarget.shapes.add({ name: "target-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Target" });
hyperlinkCloneFixture.slides.add({ name: "Linked appendix" });
hyperlinkCloneOrigin.shapes.add({
  name: "linked-copy",
  geometry: "textbox",
  position: { left: 48, top: 48, width: 520, height: 96 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [{
    runs: [
      { text: "Guide ", link: { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank" } },
      { text: "Target ", link: { slideId: hyperlinkCloneTarget.id, tooltip: "Open target" } },
      { text: "Next", link: { action: "nextSlide" } },
    ],
  }],
});
const hyperlinkCloneSourcePptx = await PresentationFile.exportPptx(hyperlinkCloneFixture);
const hyperlinkCloneSourceZip = await JSZip.loadAsync(hyperlinkCloneSourcePptx.bytes);
const hyperlinkCloneImported = await PresentationFile.importPptx(hyperlinkCloneSourcePptx);
const hyperlinkCloneImportedOrigin = hyperlinkCloneImported.slides.getItem(0);
const hyperlinkClonePending = hyperlinkCloneImportedOrigin.duplicate();
const pendingRuns = itemByName(hyperlinkClonePending.shapes.items, "linked-copy").text.paragraphs[0].runs;
assert.equal(pendingRuns[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(pendingRuns[1].link.slideId, hyperlinkCloneImported.slides.getItem(2).id);
assert.equal(pendingRuns[2].link.action, "nextSlide");
const hyperlinkClonePptx = await PresentationFile.exportPptx(hyperlinkCloneImported);
const hyperlinkCloneZip = await JSZip.loadAsync(hyperlinkClonePptx.bytes);
assert.deepEqual(
  await hyperlinkCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await hyperlinkCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning canonical run hyperlinks must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await hyperlinkCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await hyperlinkCloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "cloning canonical run hyperlinks must retain the origin relationship part byte-for-byte",
);
const modeledRunLinkRelationships = (relationships) => [...relationships.matchAll(/<Relationship\b[^>]*>/g)]
  .map(([tag]) => ({
    id: /\bId="([^"]+)"/.exec(tag)?.[1],
    type: /\bType="([^"]+)"/.exec(tag)?.[1],
    target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
    targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
  }))
  .filter((relationship) => /\/(?:hyperlink|slide)$/.test(relationship.type || ""))
  .sort((left, right) => left.id.localeCompare(right.id));
assert.deepEqual(
  modeledRunLinkRelationships(await hyperlinkCloneZip.file("ppt/slides/_rels/slide4.xml.rels").async("text")),
  modeledRunLinkRelationships(await hyperlinkCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text")),
  "the clone must own the same canonical external and internal run-link graph",
);
const hyperlinkCloneRoundTrip = await PresentationFile.importPptx(hyperlinkClonePptx);
const hyperlinkCloneShape = itemByName(hyperlinkCloneRoundTrip.slides.getItem(1).shapes.items, "linked-copy");
const roundTripRuns = hyperlinkCloneShape.text.paragraphs[0].runs;
assert.equal(roundTripRuns[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(roundTripRuns[1].link.slideId, hyperlinkCloneRoundTrip.slides.getItem(2).id);
assert.equal(roundTripRuns[2].link.action, "nextSlide");
roundTripRuns[0].link = { uri: "https://example.com/clone-updated" };
hyperlinkCloneShape.text.paragraphs = [{ runs: roundTripRuns }];
const hyperlinkCloneEdited = await PresentationFile.exportPptx(hyperlinkCloneRoundTrip);
const hyperlinkCloneEditedRoundTrip = await PresentationFile.importPptx(hyperlinkCloneEdited);
assert.equal(itemByName(hyperlinkCloneEditedRoundTrip.slides.getItem(0).shapes.items, "linked-copy").text.paragraphs[0].runs[0].link.uri, "https://example.com/guide?x=1&y=2");
assert.equal(itemByName(hyperlinkCloneEditedRoundTrip.slides.getItem(1).shapes.items, "linked-copy").text.paragraphs[0].runs[0].link.uri, "https://example.com/clone-updated");

const immediateHyperlinkCloneEdit = await PresentationFile.importPptx(hyperlinkCloneSourcePptx);
const immediateHyperlinkCloneShape = itemByName(immediateHyperlinkCloneEdit.slides.getItem(0).duplicate().shapes.items, "linked-copy");
const immediateHyperlinkRuns = immediateHyperlinkCloneShape.text.paragraphs[0].runs;
immediateHyperlinkRuns[0].link = { uri: "https://example.com/too-soon" };
immediateHyperlinkCloneShape.text.paragraphs = [{ runs: immediateHyperlinkRuns }];
await assert.rejects(
  () => PresentationFile.exportPptx(immediateHyperlinkCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const immediateCloneEdit = await PresentationFile.importPptx(cloneSourcePptx);
immediateCloneEdit.slides.getItem(0).duplicate().shapes.items[0].text.set("Too soon");
await assert.rejects(
  () => PresentationFile.exportPptx(immediateCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const immediateCloneRename = await PresentationFile.importPptx(cloneSourcePptx);
immediateCloneRename.slides.getItem(0).duplicate().name = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateCloneRename),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const cloneWithoutOrigin = await PresentationFile.importPptx(cloneSourcePptx);
const cloneOrigin = cloneWithoutOrigin.slides.getItem(0);
cloneOrigin.duplicate();
cloneOrigin.delete();
await assert.rejects(
  () => PresentationFile.exportPptx(cloneWithoutOrigin),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// The next clone leaf includes canonical embedded images. The fresh SlidePart
// owns a fresh relationship, but both slides deliberately point at the same
// immutable media part; no source slide XML or media bytes are rewritten.
const imageCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const imageCloneOriginal = imageCloneFixture.slides.add({ name: "Image original" });
imageCloneOriginal.shapes.add({ name: "image-clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Image original" });
imageCloneOriginal.images.add({
  name: "image-clone-asset",
  alt: "Shared immutable clone asset",
  position: { left: 48, top: 150, width: 120, height: 120 },
  dataUrl: PNG,
  fit: "stretch",
});
const imageCloneSourcePptx = await PresentationFile.exportPptx(imageCloneFixture);
const imageCloneSourceZip = await JSZip.loadAsync(imageCloneSourcePptx.bytes);
const imageCloneSourceMediaPaths = Object.keys(imageCloneSourceZip.files)
  .filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.equal(imageCloneSourceMediaPaths.length, 1, "the source fixture must contain exactly one embedded image part");
const [imageCloneSourceMediaPath] = imageCloneSourceMediaPaths;
const imageCloneImportedDeck = await PresentationFile.importPptx(imageCloneSourcePptx);
const imageCloneImportedSource = imageCloneImportedDeck.slides.getItem(0);
const imageClone = imageCloneImportedSource.duplicate();
assert.equal(imageClone.images.items.length, 1);
assert.notEqual(imageClone.images.items[0].id, imageCloneImportedSource.images.items[0].id);
assert.equal(imageClone.images.items[0].alt, "Shared immutable clone asset");
assert.equal(imageClone.images.items[0].dataUrl, PNG);
const imageClonePptx = await PresentationFile.exportPptx(imageCloneImportedDeck);
const imageCloneZip = await JSZip.loadAsync(imageClonePptx.bytes);
assert.deepEqual(
  await imageCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await imageCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a slide with an embedded image must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await imageCloneZip.file(imageCloneSourceMediaPath).async("uint8array"),
  await imageCloneSourceZip.file(imageCloneSourceMediaPath).async("uint8array"),
  "cloning a slide with an embedded image must retain the shared media bytes",
);
assert.ok(imageCloneZip.file("ppt/slides/slide2.xml"), "the image clone must own a new SlidePart");
const imageParts = Object.keys(imageCloneZip.files).filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.deepEqual(imageParts, [imageCloneSourceMediaPath], "the clone must reuse the source ImagePart instead of duplicating media bytes");
const imageRelationshipTargets = (relationships) => [...relationships.matchAll(/<Relationship\b[^>]*>/g)]
  .filter(([tag]) => /\bType="[^\"]*\/image"/.test(tag))
  .map(([tag]) => /\bTarget="([^\"]+)"/.exec(tag)?.[1]);
const sourceImageRelationships = await imageCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const cloneImageRelationships = await imageCloneZip.file("ppt/slides/_rels/slide2.xml.rels").async("text");
assert.deepEqual(
  imageRelationshipTargets(cloneImageRelationships),
  imageRelationshipTargets(sourceImageRelationships),
  "the clone must own equivalent image relationships to the same media targets",
);
const imageCloneRoundTrip = await PresentationFile.importPptx(imageClonePptx);
assert.deepEqual(imageCloneRoundTrip.slides.items.map((slide) => slide.name), ["Image original", "Image original"]);
assert.equal(imageCloneRoundTrip.slides.getItem(1).images.items[0].alt, "Shared immutable clone asset");
imageCloneRoundTrip.slides.getItem(1).images.items[0].alt = "Edited after image clone reimport";
const imageCloneEditedPptx = await PresentationFile.exportPptx(imageCloneRoundTrip);
const imageCloneEditedRoundTrip = await PresentationFile.importPptx(imageCloneEditedPptx);
assert.equal(imageCloneEditedRoundTrip.slides.getItem(1).images.items[0].alt, "Edited after image clone reimport");

const immediateImageCloneEdit = await PresentationFile.importPptx(imageCloneSourcePptx);
immediateImageCloneEdit.slides.getItem(0).duplicate().images.items[0].alt = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateImageCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

// Canonical tables are an accepted GraphicFrame leaf. They are inline in slide
// XML, so duplicating them must create fresh model identity and exactly no
// table-specific OPC relationship. Closed literal-data charts are exercised
// separately below because each clone must own a distinct ChartPart.
const tableCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const tableCloneOriginal = tableCloneFixture.slides.add({ name: "Table original" });
const tableCloneSourceTable = tableCloneOriginal.tables.add({
  name: "decision-grid",
  position: { left: 48, top: 48, width: 450, height: 210 },
  values: [["Release evidence", "discarded", "discarded"], ["Native QA", "discarded", "Pass"], ["discarded", "discarded", "Release"]],
  rows: 3,
  columns: 3,
  styleOptions: { headerRow: true, bandedRows: true },
});
assert.equal(tableCloneSourceTable.merge({ startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 }), tableCloneSourceTable);
tableCloneSourceTable.merge({ startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 });
assert.deepEqual(tableCloneSourceTable.mergeRanges, [
  { startRow: 0, endRow: 0, startColumn: 0, endColumn: 2 },
  { startRow: 1, endRow: 2, startColumn: 0, endColumn: 1 },
]);
assert.deepEqual(tableCloneSourceTable.values, [["Release evidence", "", ""], ["Native QA", "", "Pass"], ["", "", "Release"]]);
assert.equal(tableCloneSourceTable.getCell(0, 0).columnSpan, 3);
assert.equal(tableCloneSourceTable.getCell(1, 0).rowSpan, 2);
assert.deepEqual(tableCloneSourceTable.getCell(2, 1).mergeOrigin, { row: 1, column: 0 });
assert.equal(tableCloneSourceTable.getCell(2, 1).editable, false);
assert.throws(() => tableCloneSourceTable.cells.set(2, 1, "hidden"), /covered by merge origin 1,0.*read-only/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 0, endRow: 1, startColumn: 2, endColumn: 2 }), /overlap at cell 0,2/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 3, endRow: 3, startColumn: 0, endColumn: 1 }), /outside the 3x3 grid/i);
assert.throws(() => tableCloneSourceTable.merge({ startRow: 2, endRow: 2, startColumn: 2, endColumn: 2 }), /at least two cells/i);
assert.match(tableCloneSourceTable.toSvg(), /width="450" height="70"/);
assert.match(tableCloneSourceTable.toSvg(), /width="300" height="140"/);
assert.match(tableCloneFixture.inspect({ kind: "table" }).ndjson, /"mergeRanges"/);
assert.deepEqual(tableCloneSourceTable.layoutJson().mergeRanges, tableCloneSourceTable.mergeRanges);
assert.equal(tableCloneFixture.verify().ok, true);
const tableCloneSourcePptx = await PresentationFile.exportPptx(tableCloneFixture);
const tableCloneSourceZip = await JSZip.loadAsync(tableCloneSourcePptx.bytes);
const tableCloneSourceXml = await tableCloneSourceZip.file("ppt/slides/slide1.xml").async("text");
assert.match(tableCloneSourceXml, /<a:tc gridSpan="3">/);
assert.match(tableCloneSourceXml, /<a:tc hMerge="1">/);
assert.match(tableCloneSourceXml, /<a:tc rowSpan="2" gridSpan="2">/);
assert.match(tableCloneSourceXml, /<a:tc hMerge="1" vMerge="1">/);
const tableCloneImportedDeck = await PresentationFile.importPptx(tableCloneSourcePptx);
const tableCloneImportedSource = tableCloneImportedDeck.slides.getItem(0);
assert.deepEqual(tableCloneImportedSource.tables.items[0].mergeRanges, tableCloneSourceTable.mergeRanges);
const tableClone = tableCloneImportedSource.duplicate();
assert.equal(tableClone.tables.items.length, 1);
assert.notEqual(tableClone.tables.items[0], tableCloneImportedSource.tables.items[0]);
assert.notEqual(tableClone.tables.items[0].id, tableCloneImportedSource.tables.items[0].id);
assert.deepEqual(tableClone.tables.items[0].values, tableCloneSourceTable.values);
assert.deepEqual(tableClone.tables.items[0].mergeRanges, tableCloneSourceTable.mergeRanges);
const tableClonePptx = await PresentationFile.exportPptx(tableCloneImportedDeck);
const tableCloneZip = await JSZip.loadAsync(tableClonePptx.bytes);
assert.deepEqual(
  await tableCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await tableCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning an inline table must retain the origin SlidePart byte-for-byte",
);
assert.ok(tableCloneZip.file("ppt/slides/slide2.xml"), "the table clone must own a new SlidePart");
const tableCloneRelationships = await tableCloneZip.file("ppt/slides/_rels/slide2.xml.rels").async("text");
assert.match(tableCloneRelationships, /\/slideLayout/);
assert.doesNotMatch(tableCloneRelationships, /\/(?:image|chart|hyperlink|oleObject|package)"/i, "canonical table clones must not add a table-specific OPC edge");
const tableCloneRoundTrip = await PresentationFile.importPptx(tableClonePptx);
assert.deepEqual(tableCloneRoundTrip.slides.items.map((slide) => slide.tables.items[0].values), [
  tableCloneSourceTable.values,
  tableCloneSourceTable.values,
]);
assert.deepEqual(tableCloneRoundTrip.slides.items.map((slide) => slide.tables.items[0].mergeRanges), [tableCloneSourceTable.mergeRanges, tableCloneSourceTable.mergeRanges]);
tableCloneRoundTrip.slides.getItem(1).tables.items[0].cells.set(2, 2, "Edited after table clone reimport");
const tableCloneEditedPptx = await PresentationFile.exportPptx(tableCloneRoundTrip);
const tableCloneEditedRoundTrip = await PresentationFile.importPptx(tableCloneEditedPptx);
assert.equal(tableCloneEditedRoundTrip.slides.getItem(1).tables.items[0].values[2][2], "Edited after table clone reimport");
assert.equal(tableCloneEditedRoundTrip.slides.getItem(0).tables.items[0].values[2][2], "Release");

const immediateTableCloneEdit = await PresentationFile.importPptx(tableCloneSourcePptx);
immediateTableCloneEdit.slides.getItem(0).duplicate().tables.items[0].values[2][2] = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateTableCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const importedTableMergeChange = await PresentationFile.importPptx(tableCloneSourcePptx);
importedTableMergeChange.slides.getItem(0).tables.items[0].merge({ startRow: 1, endRow: 2, startColumn: 2, endColumn: 2 });
await assert.rejects(
  () => PresentationFile.exportPptx(importedTableMergeChange),
  (error) => error?.code === "unsupported_presentation_edit",
);

// A recognized literal-data chart may travel with the bounded slide clone
// only when its ChartPart has no child/external/data relationship graph. The
// clone copies the exact chart XML into a distinct part: sharing would make a
// later chart edit affect both slides.
const chartCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const chartCloneOriginal = chartCloneFixture.slides.add({ name: "Chart original" });
chartCloneOriginal.charts.add("bar", {
  name: "pipeline-chart",
  position: { left: 48, top: 42, width: 500, height: 250 },
  title: "Quarterly pipeline",
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "Pipeline", values: [42, 48, 57], fill: "#2563EB" }],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Value", min: 0, max: 80, majorUnit: 20 },
  },
  legend: { visible: true, position: "r" },
  dataLabels: { showValue: true, position: "top" },
});
const chartCloneSourcePptx = await PresentationFile.exportPptx(chartCloneFixture);
const chartCloneSourceZip = await JSZip.loadAsync(chartCloneSourcePptx.bytes);
const chartPartPaths = (zip) => Object.keys(zip.files)
  .filter((partPath) => /^ppt\/(?:slides\/)?charts\/chart\d+\.xml$/i.test(partPath))
  .sort();
const chartCloneSourceParts = chartPartPaths(chartCloneSourceZip);
assert.equal(chartCloneSourceParts.length, 1, "the source fixture must own exactly one ChartPart");
const [chartCloneSourcePart] = chartCloneSourceParts;
const chartCloneImportedDeck = await PresentationFile.importPptx(chartCloneSourcePptx);
const chartCloneImportedSource = chartCloneImportedDeck.slides.getItem(0);
const chartClonePending = chartCloneImportedSource.duplicate();
assert.equal(chartClonePending.charts.items.length, 1);
assert.notEqual(chartClonePending.charts.items[0], chartCloneImportedSource.charts.items[0]);
assert.notEqual(chartClonePending.charts.items[0].id, chartCloneImportedSource.charts.items[0].id);
assert.deepEqual(chartClonePending.charts.items[0].categories, ["Q1", "Q2", "Q3"]);
assert.deepEqual(chartClonePending.charts.items[0].series[0].values, [42, 48, 57]);
const chartClonePptx = await PresentationFile.exportPptx(chartCloneImportedDeck);
const chartCloneZip = await JSZip.loadAsync(chartClonePptx.bytes);
assert.deepEqual(
  await chartCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await chartCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a closed native chart must retain the origin SlidePart byte-for-byte",
);
const chartCloneOutputParts = chartPartPaths(chartCloneZip);
assert.equal(chartCloneOutputParts.length, 2, "the chart clone must allocate exactly one additional ChartPart");
const chartCloneNewParts = chartCloneOutputParts.filter((partPath) => !chartCloneSourceZip.file(partPath));
assert.equal(chartCloneNewParts.length, 1, "the chart clone must own one distinct new ChartPart path");
const [chartClonePart] = chartCloneNewParts;
assert.deepEqual(
  await chartCloneZip.file(chartClonePart).async("uint8array"),
  await chartCloneSourceZip.file(chartCloneSourcePart).async("uint8array"),
  "the first clone export must copy the accepted ChartPart byte-for-byte",
);
const modeledChartRelationship = (xml) => {
  const relationships = [...xml.matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => ({
      id: /\bId="([^"]+)"/.exec(tag)?.[1],
      type: /\bType="([^"]+)"/.exec(tag)?.[1],
      target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
      targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
    }))
    .filter((relationship) => /\/chart$/.test(relationship.type || ""));
  assert.equal(relationships.length, 1, "a bounded chart clone slide must own exactly one chart relationship");
  return relationships[0];
};
const sourceChartRelationship = modeledChartRelationship(await chartCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"));
const cloneChartRelationship = modeledChartRelationship(await chartCloneZip.file("ppt/slides/_rels/slide2.xml.rels").async("text"));
assert.equal(cloneChartRelationship.id, sourceChartRelationship.id, "the clone must retain the slide-local chart relationship ID");
assert.equal(cloneChartRelationship.type, sourceChartRelationship.type);
assert.equal(cloneChartRelationship.targetMode, undefined);
const resolvedCloneChartTarget = cloneChartRelationship.target.startsWith("/")
  ? cloneChartRelationship.target.replace(/^\/+/, "")
  : path.posix.normalize(path.posix.join("ppt/slides", cloneChartRelationship.target));
assert.equal(resolvedCloneChartTarget, chartClonePart, "the clone chart relationship must target its newly allocated ChartPart");

const chartCloneRoundTrip = await PresentationFile.importPptx(chartClonePptx);
assert.deepEqual(chartCloneRoundTrip.slides.items.map((slide) => slide.charts.items[0].title), ["Quarterly pipeline", "Quarterly pipeline"]);
chartCloneRoundTrip.slides.getItem(1).charts.items[0].title = "Updated clone pipeline";
chartCloneRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1] = 63;
const chartCloneEditedPptx = await PresentationFile.exportPptx(chartCloneRoundTrip);
const chartCloneEditedZip = await JSZip.loadAsync(chartCloneEditedPptx.bytes);
assert.deepEqual(
  await chartCloneEditedZip.file(chartCloneSourcePart).async("uint8array"),
  await chartCloneSourceZip.file(chartCloneSourcePart).async("uint8array"),
  "editing the reimported clone chart must leave the origin ChartPart byte-for-byte intact",
);
const chartCloneEditedRoundTrip = await PresentationFile.importPptx(chartCloneEditedPptx);
assert.equal(chartCloneEditedRoundTrip.slides.getItem(0).charts.items[0].title, "Quarterly pipeline");
assert.equal(chartCloneEditedRoundTrip.slides.getItem(0).charts.items[0].series[0].values[1], 48);
assert.equal(chartCloneEditedRoundTrip.slides.getItem(1).charts.items[0].title, "Updated clone pipeline");
assert.equal(chartCloneEditedRoundTrip.slides.getItem(1).charts.items[0].series[0].values[1], 63);

const immediateChartCloneEdit = await PresentationFile.importPptx(chartCloneSourcePptx);
immediateChartCloneEdit.slides.getItem(0).duplicate().charts.items[0].title = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateChartCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const connectedChartCloneZip = await JSZip.loadAsync(chartCloneSourcePptx.bytes);
const chartRelationshipPart = path.posix.join(
  path.posix.dirname(chartCloneSourcePart),
  "_rels",
  path.posix.basename(chartCloneSourcePart) + ".rels",
);
connectedChartCloneZip.file(
  chartRelationshipPart,
  '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeChartChild" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/chart-child" TargetMode="External"/></Relationships>',
);
const connectedChartCloneSource = await connectedChartCloneZip.generateAsync({ type: "uint8array" });
const connectedChartCloneDeck = await PresentationFile.importPptx(connectedChartCloneSource);
connectedChartCloneDeck.slides.getItem(0).duplicate();
await assert.rejects(
  () => PresentationFile.exportPptx(connectedChartCloneDeck),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a chart with any child relationship graph must fail closed",
);

// A group is clone-safe only when every descendant is already in the narrow
// shape/table/closed-chart/image leaf. The group and every child receive fresh
// JS identity, while nested relationship-owning leaves still require the same
// native preflight as their top-level counterparts.
const groupCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const groupCloneOriginal = groupCloneFixture.slides.add({ name: "Recursive group original" });
const groupCloneRoot = groupCloneOriginal.addGroup({
  name: "evidence-cluster",
  position: { left: 48, top: 40, width: 420, height: 220 },
  childFrame: { left: 0, top: 0, width: 420, height: 220 },
});
groupCloneRoot.shapes.add({
  name: "group-copy",
  position: { left: 0, top: 0, width: 220, height: 48 },
  text: "Clone-safe grouped evidence",
  fill: "#FFFFFF",
  line: { fill: "#2563EB", width: 1 },
});
groupCloneRoot.tables.add({
  name: "group-decision-grid",
  position: { left: 0, top: 66, width: 220, height: 110 },
  values: [["Gate", "State"], ["QA", "Pass"]],
  rows: 2,
  columns: 2,
  styleOptions: { headerRow: true, bandedRows: true },
});
groupCloneRoot.images.add({
  name: "group-immutable-asset",
  alt: "Shared nested clone asset",
  position: { left: 246, top: 0, width: 96, height: 96 },
  dataUrl: PNG,
  fit: "stretch",
});
const groupCloneNested = groupCloneRoot.groups.add({
  name: "nested-evidence",
  position: { left: 246, top: 118, width: 150, height: 48 },
  childFrame: { left: 0, top: 0, width: 150, height: 48 },
});
groupCloneNested.shapes.add({
  name: "nested-copy",
  position: { left: 0, top: 0, width: 150, height: 48 },
  text: "Nested",
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 1 },
});
const groupCloneSourcePptx = await PresentationFile.exportPptx(groupCloneFixture);
const groupCloneSourceZip = await JSZip.loadAsync(groupCloneSourcePptx.bytes);
const groupCloneSourceMediaPaths = Object.keys(groupCloneSourceZip.files)
  .filter((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.equal(groupCloneSourceMediaPaths.length, 1, "the recursive group fixture must contain one nested image part");
const [groupCloneSourceMediaPath] = groupCloneSourceMediaPaths;
const groupCloneImportedDeck = await PresentationFile.importPptx(groupCloneSourcePptx);
const groupCloneImportedSource = groupCloneImportedDeck.slides.getItem(0);
const groupClone = groupCloneImportedSource.duplicate();
const groupCloneCopy = groupClone.groups.items[0];
assert.equal(groupClone.groups.items.length, 1);
assert.notEqual(groupCloneCopy, groupCloneImportedSource.groups.items[0]);
assert.notEqual(groupCloneCopy.id, groupCloneImportedSource.groups.items[0].id);
assert.notEqual(groupCloneCopy.shapes.items[0].id, groupCloneImportedSource.groups.items[0].shapes.items[0].id);
assert.notEqual(groupCloneCopy.tables.items[0].id, groupCloneImportedSource.groups.items[0].tables.items[0].id);
assert.notEqual(groupCloneCopy.images.items[0].id, groupCloneImportedSource.groups.items[0].images.items[0].id);
assert.notEqual(groupCloneCopy.groups.items[0].id, groupCloneImportedSource.groups.items[0].groups.items[0].id);
assert.deepEqual(groupCloneCopy.tables.items[0].values, [["Gate", "State"], ["QA", "Pass"]]);
assert.equal(groupCloneCopy.groups.items[0].shapes.items[0].text.value, "Nested");
const groupClonePptx = await PresentationFile.exportPptx(groupCloneImportedDeck);
const groupCloneZip = await JSZip.loadAsync(groupClonePptx.bytes);
assert.deepEqual(
  await groupCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await groupCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a recursively canonical group must retain its origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await groupCloneZip.file(groupCloneSourceMediaPath).async("uint8array"),
  await groupCloneSourceZip.file(groupCloneSourceMediaPath).async("uint8array"),
  "cloning a recursively canonical group must retain its shared media bytes",
);
assert.deepEqual(
  imageRelationshipTargets(await groupCloneZip.file("ppt/slides/_rels/slide2.xml.rels").async("text")),
  imageRelationshipTargets(await groupCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text")),
  "nested clone images must receive the same verified relationship target as the origin",
);
const groupCloneRoundTrip = await PresentationFile.importPptx(groupClonePptx);
assert.equal(groupCloneRoundTrip.slides.items.length, 2);
assert.equal(groupCloneRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.value, "Nested");
groupCloneRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.set("Edited after group clone reimport");
const groupCloneEditedPptx = await PresentationFile.exportPptx(groupCloneRoundTrip);
const groupCloneEditedRoundTrip = await PresentationFile.importPptx(groupCloneEditedPptx);
assert.equal(groupCloneEditedRoundTrip.slides.getItem(1).groups.items[0].groups.items[0].shapes.items[0].text.value, "Edited after group clone reimport");

const immediateGroupCloneEdit = await PresentationFile.importPptx(groupCloneSourcePptx);
immediateGroupCloneEdit.slides.getItem(0).duplicate().groups.items[0].groups.items[0].shapes.items[0].text.set("Too soon");
await assert.rejects(
  () => PresentationFile.exportPptx(immediateGroupCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const connectedGroupCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const connectedGroupCloneOriginal = connectedGroupCloneFixture.slides.add({ name: "Connected group original" });
const connectedGroupCloneRoot = connectedGroupCloneOriginal.addGroup({
  name: "connected-cluster",
  position: { left: 48, top: 40, width: 320, height: 120 },
  childFrame: { left: 0, top: 0, width: 320, height: 120 },
});
const connectedGroupCloneLeft = connectedGroupCloneRoot.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
const connectedGroupCloneRight = connectedGroupCloneRoot.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
connectedGroupCloneRoot.connectors.add({
  name: "join",
  from: connectedGroupCloneLeft,
  to: connectedGroupCloneRight,
  start: { x: 90, y: 41 },
  end: { x: 210, y: 41 },
  line: { fill: "#64748B", width: 1 },
});
const connectedGroupCloneSourcePptx = await PresentationFile.exportPptx(connectedGroupCloneFixture);
const connectedGroupCloneSourceZip = await JSZip.loadAsync(connectedGroupCloneSourcePptx.bytes);
const connectedGroupCloneImported = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
const connectedGroupCloneImportedSource = connectedGroupCloneImported.slides.getItem(0);
const connectedGroupCloneSourceGroup = connectedGroupCloneImportedSource.groups.items[0];
const connectedGroupCloneSourceConnector = connectedGroupCloneSourceGroup.connectors.items[0];
assert.equal(connectedGroupCloneSourceConnector.startTargetId, connectedGroupCloneSourceGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneSourceConnector.endTargetId, connectedGroupCloneSourceGroup.shapes.items[1].id);
const connectedGroupClone = connectedGroupCloneImportedSource.duplicate();
const connectedGroupCloneCopy = connectedGroupClone.groups.items[0];
const connectedGroupCloneConnector = connectedGroupCloneCopy.connectors.items[0];
assert.notEqual(connectedGroupCloneConnector.id, connectedGroupCloneSourceConnector.id);
assert.notEqual(connectedGroupCloneCopy.shapes.items[0].id, connectedGroupCloneSourceGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneConnector.startTargetId, connectedGroupCloneCopy.shapes.items[0].id);
assert.equal(connectedGroupCloneConnector.endTargetId, connectedGroupCloneCopy.shapes.items[1].id);
const connectedGroupClonePptx = await PresentationFile.exportPptx(connectedGroupCloneImported);
const connectedGroupCloneZip = await JSZip.loadAsync(connectedGroupClonePptx.bytes);
assert.deepEqual(
  await connectedGroupCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await connectedGroupCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a group with bounded connectors must retain its origin SlidePart byte-for-byte",
);
assert.ok(connectedGroupCloneZip.file("ppt/slides/slide2.xml"), "the connector clone must own a new SlidePart");
const connectedGroupCloneRoundTrip = await PresentationFile.importPptx(connectedGroupClonePptx);
const connectedGroupCloneRoundTripGroup = connectedGroupCloneRoundTrip.slides.getItem(1).groups.items[0];
const connectedGroupCloneRoundTripConnector = connectedGroupCloneRoundTripGroup.connectors.items[0];
assert.equal(connectedGroupCloneRoundTripConnector.startTargetId, connectedGroupCloneRoundTripGroup.shapes.items[0].id);
assert.equal(connectedGroupCloneRoundTripConnector.endTargetId, connectedGroupCloneRoundTripGroup.shapes.items[1].id);

const immediateConnectedGroupCloneEdit = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
immediateConnectedGroupCloneEdit.slides.getItem(0).duplicate().groups.items[0].connectors.items[0].line.width = 2;
await assert.rejects(
  () => PresentationFile.exportPptx(immediateConnectedGroupCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);

const unresolvedConnectedGroupClone = await PresentationFile.importPptx(connectedGroupCloneSourcePptx);
unresolvedConnectedGroupClone.slides.getItem(0).groups.items[0].connectors.items[0].startTargetId = "missing-source-target";
assert.throws(
  () => unresolvedConnectedGroupClone.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
);
assert.equal(unresolvedConnectedGroupClone.slides.items.length, 1, "connector-target preflight must not leave a partial clone behind");

// Speaker notes add one deliberately closed relationship leaf to the same
// clone profile. The NotesSlide itself is new and points at the clone, while
// its NotesMaster stays immutable and shared. This is raw part preservation,
// not permission to edit notes before the export/reimport boundary.
const notesCloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
const notesCloneOriginal = notesCloneFixture.slides.add({
  name: "Notes image original",
  notes: "Open with the customer outcome.\nClose with the operating decision.",
});
notesCloneOriginal.shapes.add({ name: "notes-clone-copy", position: { left: 48, top: 48, width: 300, height: 72 }, text: "Notes image original" });
notesCloneOriginal.images.add({
  name: "notes-clone-asset",
  alt: "Notes clone immutable asset",
  position: { left: 48, top: 150, width: 120, height: 120 },
  dataUrl: PNG,
  fit: "stretch",
});
const notesCloneSourcePptx = await PresentationFile.exportPptx(notesCloneFixture);
const notesCloneSourceZip = await JSZip.loadAsync(notesCloneSourcePptx.bytes);
const notesCloneSourceNotesPaths = Object.keys(notesCloneSourceZip.files)
  .filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path));
assert.deepEqual(notesCloneSourceNotesPaths, ["ppt/notesSlides/notesSlide1.xml"]);
const [notesCloneSourceNotesPath] = notesCloneSourceNotesPaths;
const notesCloneSourceMediaPath = Object.keys(notesCloneSourceZip.files)
  .find((path) => /^ppt\/media\/[^/]+\.(?:png|jpe?g|gif|svg)$/i.test(path));
assert.ok(notesCloneSourceMediaPath, "the notes clone fixture must contain one embedded image part");
const notesCloneImportedDeck = await PresentationFile.importPptx(notesCloneSourcePptx);
const notesCloneImportedSource = notesCloneImportedDeck.slides.getItem(0);
const notesClone = notesCloneImportedSource.duplicate();
assert.equal(notesClone.speakerNotes.text, "Open with the customer outcome.\nClose with the operating decision.");
const notesClonePptx = await PresentationFile.exportPptx(notesCloneImportedDeck);
const notesCloneZip = await JSZip.loadAsync(notesClonePptx.bytes);
assert.deepEqual(
  await notesCloneZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await notesCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "cloning a slide with notes must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await notesCloneZip.file(notesCloneSourceNotesPath).async("uint8array"),
  await notesCloneSourceZip.file(notesCloneSourceNotesPath).async("uint8array"),
  "cloning a slide with notes must retain the origin NotesSlide byte-for-byte",
);
assert.deepEqual(
  await notesCloneZip.file(notesCloneSourceMediaPath).async("uint8array"),
  await notesCloneSourceZip.file(notesCloneSourceMediaPath).async("uint8array"),
  "cloning a slide with notes must retain the shared media bytes",
);
const notesClonePaths = Object.keys(notesCloneZip.files)
  .filter((path) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(path));
assert.deepEqual(notesClonePaths, ["ppt/notesSlides/notesSlide1.xml", "ppt/notesSlides/notesSlide2.xml"]);
assert.equal(
  await notesCloneZip.file("ppt/notesSlides/notesSlide2.xml").async("text"),
  await notesCloneZip.file(notesCloneSourceNotesPath).async("text"),
  "the clone NotesSlide XML must be a verbatim copy of the source notes XML",
);
const relationshipTagForType = (relationships, suffix) => [...relationships.matchAll(/<Relationship\b[^>]*>/gi)]
  .find(([tag]) => new RegExp(`\\bType="[^"]*\\/${suffix}"`, "i").test(tag))?.[0];
const relationshipAttributeForType = (relationships, suffix, attribute) => {
  const tag = relationshipTagForType(relationships, suffix);
  return tag && new RegExp(`\\b${attribute}="([^"]+)"`, "i").exec(tag)?.[1];
};
const sourceSlideRelationships = await notesCloneZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const cloneSlideRelationships = await notesCloneZip.file("ppt/slides/_rels/slide2.xml.rels").async("text");
const sourceNotesRelationships = await notesCloneZip.file("ppt/notesSlides/_rels/notesSlide1.xml.rels").async("text");
const cloneNotesRelationships = await notesCloneZip.file("ppt/notesSlides/_rels/notesSlide2.xml.rels").async("text");
assert.equal(relationshipAttributeForType(cloneSlideRelationships, "notesSlide", "Id"), relationshipAttributeForType(sourceSlideRelationships, "notesSlide", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "notesMaster", "Id"), relationshipAttributeForType(sourceNotesRelationships, "notesMaster", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "notesMaster", "Target"), relationshipAttributeForType(sourceNotesRelationships, "notesMaster", "Target"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "slide", "Id"), relationshipAttributeForType(sourceNotesRelationships, "slide", "Id"));
assert.equal(relationshipAttributeForType(cloneNotesRelationships, "slide", "Target"), "/ppt/slides/slide2.xml");
const notesCloneRoundTrip = await PresentationFile.importPptx(notesClonePptx);
assert.deepEqual(notesCloneRoundTrip.slides.items.map((slide) => slide.speakerNotes.text), [
  "Open with the customer outcome.\nClose with the operating decision.",
  "Open with the customer outcome.\nClose with the operating decision.",
]);
notesCloneRoundTrip.slides.getItem(1).speakerNotes.text = "Edited after notes clone reimport.";
const notesCloneEditedPptx = await PresentationFile.exportPptx(notesCloneRoundTrip);
const notesCloneEditedRoundTrip = await PresentationFile.importPptx(notesCloneEditedPptx);
assert.equal(notesCloneEditedRoundTrip.slides.getItem(1).speakerNotes.text, "Edited after notes clone reimport.");

const immediateNotesCloneEdit = await PresentationFile.importPptx(notesCloneSourcePptx);
immediateNotesCloneEdit.slides.getItem(0).duplicate().speakerNotes.text = "Too soon";
await assert.rejects(
  () => PresentationFile.exportPptx(immediateNotesCloneEdit),
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
const sharedOleSlideCount = sharedOlePresentation.slides.items.length;
assert.throws(
  () => sharedOlePresentation.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "an OLE package with more than one inbound relationship must fail clone preflight before mutating the model",
);
assert.equal(sharedOlePresentation.slides.items.length, sharedOleSlideCount);

// The bounded imported-slide clone may carry the same uniquely bound,
// top-level embedded-XLSX OLE frame. The mutable workbook package is copied
// into a distinct part, while the immutable preview ImagePart is shared.
const oleCloneBaseSlideXml = await cloneSourceZip.file("ppt/slides/slide1.xml").async("text");
const oleCloneBaseRelationships = await cloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleCloneSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", '<Relationship Id="rIdEmbeddedWorkbook" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/clone-agent-workbook.xlsx"/><Relationship Id="rIdEmbeddedPreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/clone-agent-workbook-preview.png"/></Relationships>') },
  { path: "ppt/embeddings/clone-agent-workbook.xlsx", bytes: embeddedSourceXlsx.bytes, contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
  { path: "ppt/media/clone-agent-workbook-preview.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const oleCloneSourceSnapshot = Uint8Array.from(oleCloneSource.bytes);
const oleCloneImported = await PresentationFile.importPptx(oleCloneSource);
const oleCloneOrigin = oleCloneImported.slides.getItem(0);
const oleCloneOriginObject = itemByName(oleCloneOrigin.nativeObjects.items, "Embedded workbook");
const oleClonePending = oleCloneOrigin.duplicate();
const oleClonePendingObject = itemByName(oleClonePending.nativeObjects.items, "Embedded workbook");
assert.notEqual(oleClonePendingObject, oleCloneOriginObject);
assert.notEqual(oleClonePendingObject.id, oleCloneOriginObject.id);
assert.equal(oleClonePendingObject.oleWorkbook.partPath, oleCloneOriginObject.oleWorkbook.partPath);
assert.equal(oleClonePendingObject.oleWorkbook.sourceSha256, oleCloneOriginObject.oleWorkbook.sourceSha256);

const oleCloneExport = await PresentationFile.exportPptx(oleCloneImported);
assert.deepEqual(oleCloneSource.bytes, oleCloneSourceSnapshot);
const oleCloneSourceZip = await JSZip.loadAsync(oleCloneSource.bytes);
const oleCloneOutputZip = await JSZip.loadAsync(oleCloneExport.bytes);
assert.deepEqual(
  await oleCloneOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await oleCloneSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "OLE cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await oleCloneOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await oleCloneSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "OLE cloning must retain the origin relationship part byte-for-byte",
);
const oleCloneWorkbookPaths = Object.keys(oleCloneOutputZip.files)
  .filter((partPath) => /^ppt\/(?:slides\/)?embeddings\/[^/]+\.xlsx$/i.test(partPath))
  .sort();
assert.equal(oleCloneWorkbookPaths.length, 2, "the clone must allocate exactly one additional XLSX package part");
const oleCloneWorkbookPart = oleCloneWorkbookPaths.find((partPath) => partPath !== "ppt/embeddings/clone-agent-workbook.xlsx");
assert.ok(oleCloneWorkbookPart);
assert.deepEqual(
  await oleCloneOutputZip.file(oleCloneWorkbookPart).async("uint8array"),
  await oleCloneSourceZip.file("ppt/embeddings/clone-agent-workbook.xlsx").async("uint8array"),
  "the first clone export must copy the embedded XLSX bytes exactly",
);
const relationshipForType = (xml, typeSuffix) => {
  const matches = [...xml.matchAll(/<Relationship\b[^>]*>/g)]
    .map(([tag]) => ({
      id: /\bId="([^"]+)"/.exec(tag)?.[1],
      type: /\bType="([^"]+)"/.exec(tag)?.[1],
      target: /\bTarget="([^"]+)"/.exec(tag)?.[1],
      targetMode: /\bTargetMode="([^"]+)"/.exec(tag)?.[1],
    }))
    .filter((relationship) => relationship.type?.endsWith(`/${typeSuffix}`));
  assert.equal(matches.length, 1, `expected one ${typeSuffix} relationship in ${xml}`);
  return matches[0];
};
const resolveSlideRelationshipTarget = (target) => target.startsWith("/")
  ? target.replace(/^\/+/, "")
  : path.posix.normalize(path.posix.join("ppt/slides", target));
const oleCloneSourceRelationships = await oleCloneOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const oleCloneCopyRelationships = await oleCloneOutputZip.file("ppt/slides/_rels/slide3.xml.rels").async("text");
const oleCloneSourcePackageRelationship = relationshipForType(oleCloneSourceRelationships, "package");
const oleCloneCopyPackageRelationship = relationshipForType(oleCloneCopyRelationships, "package");
assert.equal(oleCloneCopyPackageRelationship.id, oleCloneSourcePackageRelationship.id);
assert.equal(oleCloneCopyPackageRelationship.targetMode, undefined);
assert.equal(resolveSlideRelationshipTarget(oleCloneSourcePackageRelationship.target), "ppt/embeddings/clone-agent-workbook.xlsx");
assert.equal(resolveSlideRelationshipTarget(oleCloneCopyPackageRelationship.target), oleCloneWorkbookPart);
const oleCloneSourcePreviewRelationship = relationshipForType(oleCloneSourceRelationships, "image");
const oleCloneCopyPreviewRelationship = relationshipForType(oleCloneCopyRelationships, "image");
assert.equal(oleCloneCopyPreviewRelationship.id, oleCloneSourcePreviewRelationship.id);
assert.equal(
  resolveSlideRelationshipTarget(oleCloneCopyPreviewRelationship.target),
  resolveSlideRelationshipTarget(oleCloneSourcePreviewRelationship.target),
  "both OLE frames must share the same immutable preview ImagePart",
);

const oleCloneRoundTrip = await PresentationFile.importPptx(oleCloneExport);
const oleCloneRoundTripOrigin = itemByName(oleCloneRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook");
const oleCloneRoundTripCopy = itemByName(oleCloneRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded workbook");
assert.notEqual(oleCloneRoundTripCopy.oleWorkbook.partPath, oleCloneRoundTripOrigin.oleWorkbook.partPath);
assert.equal(oleCloneRoundTripCopy.oleWorkbook.sourceSha256, oleCloneRoundTripOrigin.oleWorkbook.sourceSha256);
oleCloneRoundTripCopy.replaceEmbeddedWorkbook(embeddedReplacementXlsx);
const oleCloneEditedExport = await PresentationFile.exportPptx(oleCloneRoundTrip);
const oleCloneEditedZip = await JSZip.loadAsync(oleCloneEditedExport.bytes);
assert.deepEqual(
  await oleCloneEditedZip.file(oleCloneRoundTripOrigin.oleWorkbook.partPath).async("uint8array"),
  embeddedSourceXlsx.bytes,
  "editing the reimported clone workbook must leave the origin package byte-for-byte intact",
);
assert.deepEqual(
  await oleCloneEditedZip.file(oleCloneRoundTripCopy.oleWorkbook.partPath).async("uint8array"),
  embeddedReplacementXlsx.bytes,
);
const oleCloneEditedRoundTrip = await PresentationFile.importPptx(oleCloneEditedExport);
const oleCloneEditedOriginWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleCloneEditedRoundTrip.slides.getItem(0).nativeObjects.items, "Embedded workbook").getEmbeddedWorkbook());
const oleCloneEditedCopyWorkbook = await SpreadsheetFile.importXlsx(itemByName(oleCloneEditedRoundTrip.slides.getItem(1).nativeObjects.items, "Embedded workbook").getEmbeddedWorkbook());
assert.equal(oleCloneEditedOriginWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Original embedded workbook");
assert.equal(oleCloneEditedCopyWorkbook.worksheets.getItem("Embedded").getRange("A1").values[0][0], "Replacement workbook");

const immediateOleCloneEdit = await PresentationFile.importPptx(oleCloneSource);
itemByName(immediateOleCloneEdit.slides.getItem(0).duplicate().nativeObjects.items, "Embedded workbook")
  .replaceEmbeddedWorkbook(embeddedReplacementXlsx);
await assert.rejects(
  () => PresentationFile.exportPptx(immediateOleCloneEdit),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a cloned OLE payload may be edited only after export and reimport establish independent source identity",
);

// A canonical top-level SmartArt frame owns exactly the four standard
// relationship-free DrawingML diagram roots. The bounded clone copies all
// four into distinct parts so a later source-bound edit cannot couple the
// origin and clone through shared diagram state.
const smartArtFrame = '<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="120" name="Clone-safe SmartArt"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="5486400" cy="2743200"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdAgentDiagramData" r:lo="rIdAgentDiagramLayout" r:qs="rIdAgentDiagramStyle" r:cs="rIdAgentDiagramColors"/></a:graphicData></a:graphic></p:graphicFrame>';
const smartArtRelationships = '<Relationship Id="rIdAgentDiagramData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/agent-data.xml"/><Relationship Id="rIdAgentDiagramLayout" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/agent-layout.xml"/><Relationship Id="rIdAgentDiagramStyle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/agent-style.xml"/><Relationship Id="rIdAgentDiagramColors" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/agent-colors.xml"/>';
const smartArtParts = [
  ["ppt/diagrams/agent-data.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>'],
  ["ppt/diagrams/agent-layout.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml", '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:agent-layout"><dgm:title val="Agent"/><dgm:desc val="Agent layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>'],
  ["ppt/diagrams/agent-style.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml", '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:agent-style"><dgm:title val="Agent"/><dgm:desc val="Agent style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>'],
  ["ppt/diagrams/agent-colors.xml", "application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml", '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:agent-colors"><dgm:title val="Agent"/><dgm:desc val="Agent colors"/><dgm:catLst/></dgm:colorsDef>'],
];
const smartArtSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${smartArtFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${smartArtRelationships}</Relationships>`) },
  ...smartArtParts.map(([partPath, contentType, xml]) => ({ path: partPath, contentType, xml })),
]);
const smartArtSourceSnapshot = Uint8Array.from(smartArtSource.bytes);
const smartArtImported = await PresentationFile.importPptx(smartArtSource);
const smartArtOriginSlide = smartArtImported.slides.getItem(0);
const smartArtOrigin = itemByName(smartArtOriginSlide.nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtOrigin.nativeKind, "diagram");
assert.equal(smartArtOrigin.parts.length, 4);
assert.ok(smartArtOrigin.parts.every((part) => part.relationships.length === 0));
const smartArtPendingSlide = smartArtOriginSlide.duplicate();
const smartArtPending = itemByName(smartArtPendingSlide.nativeObjects.items, "Clone-safe SmartArt");
assert.notEqual(smartArtPending, smartArtOrigin);
assert.notEqual(smartArtPending.id, smartArtOrigin.id);
assert.deepEqual(smartArtPending.parts.map((part) => part.path), smartArtOrigin.parts.map((part) => part.path));

const smartArtExport = await PresentationFile.exportPptx(smartArtImported);
assert.deepEqual(smartArtSource.bytes, smartArtSourceSnapshot);
const smartArtSourceZip = await JSZip.loadAsync(smartArtSource.bytes);
const smartArtOutputZip = await JSZip.loadAsync(smartArtExport.bytes);
assert.deepEqual(
  await smartArtOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await smartArtSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "SmartArt cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await smartArtOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await smartArtSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "SmartArt cloning must retain the origin relationship part byte-for-byte",
);
const smartArtSourceRels = await smartArtOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const smartArtCloneRels = await smartArtOutputZip.file("ppt/slides/_rels/slide3.xml.rels").async("text");
const smartArtTypeSuffixes = ["diagramData", "diagramLayout", "diagramQuickStyle", "diagramColors"];
const smartArtClonePartPaths = [];
for (const typeSuffix of smartArtTypeSuffixes) {
  const sourceRelationship = relationshipForType(smartArtSourceRels, typeSuffix);
  const cloneRelationship = relationshipForType(smartArtCloneRels, typeSuffix);
  const sourcePath = resolveSlideRelationshipTarget(sourceRelationship.target);
  const clonePath = resolveSlideRelationshipTarget(cloneRelationship.target);
  assert.equal(cloneRelationship.id, sourceRelationship.id);
  assert.notEqual(clonePath, sourcePath);
  assert.deepEqual(
    await smartArtOutputZip.file(clonePath).async("uint8array"),
    await smartArtSourceZip.file(sourcePath).async("uint8array"),
    `SmartArt cloning must byte-copy the closed ${typeSuffix} part`,
  );
  smartArtClonePartPaths.push(clonePath);
}
assert.equal(new Set(smartArtClonePartPaths).size, 4);

const smartArtRoundTrip = await PresentationFile.importPptx(smartArtExport);
const smartArtRoundTripOrigin = itemByName(smartArtRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
const smartArtRoundTripClone = itemByName(smartArtRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtRoundTripOrigin.parts.length, 4);
assert.equal(smartArtRoundTripClone.parts.length, 4);
assert.equal(
  smartArtRoundTripOrigin.parts.some((part) => smartArtRoundTripClone.parts.some((clonePart) => clonePart.path === part.path)),
  false,
  "reimported SmartArt origin and clone must not share any mutable diagram part",
);
assert.deepEqual(
  smartArtRoundTripOrigin.parts.map((part) => part.sourceSha256).sort(),
  smartArtRoundTripClone.parts.map((part) => part.sourceSha256).sort(),
);

// A canonical closed SmartArt data model can expose only its direct plain
// document-node text. The bounded edit rewrites the one hash-bound data part;
// the frame, relationships, and layout/style/color leaves must stay intact.
const smartArtTextData = '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dgm:ptLst><dgm:pt modelId="{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Original node</a:t></a:r></a:p></dgm:t></dgm:pt><dgm:pt modelId="1" type="doc"><dgm:t><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Second node</a:t></a:r></a:p></dgm:t></dgm:pt></dgm:ptLst><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>';
const smartArtTextSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${smartArtFrame}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${smartArtRelationships}</Relationships>`) },
  ...smartArtParts.map(([partPath, contentType, xml]) => ({ path: partPath, contentType, xml: partPath === "ppt/diagrams/agent-data.xml" ? smartArtTextData : xml })),
]);
const smartArtTextInput = Uint8Array.from(smartArtTextSource.bytes);
const smartArtTextImported = await PresentationFile.importPptx(smartArtTextSource);
const smartArtTextObject = itemByName(smartArtTextImported.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
assert.equal(smartArtTextObject.editable, false);
assert.deepEqual(smartArtTextObject.diagramText?.nodes, [
  { id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", text: "Original node" },
  { id: "1", text: "Second node" },
]);
assert.deepEqual(smartArtTextObject.inspectRecord().editableFields, ["diagramText"]);
assert.throws(
  () => smartArtTextObject.setDiagramNodeText("missing", "nope"),
  /not part of the source-bound diagram profile/,
);
assert.throws(
  () => smartArtTextObject.setDiagramNodeText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", "x".repeat(32_768)),
  /32767 XML-safe characters/,
);
smartArtTextObject.setDiagramNodeText("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", " Revised node ");
const smartArtTextSlideCount = smartArtTextImported.slides.items.length;
assert.throws(
  () => smartArtTextImported.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a pending SmartArt text edit must cross an export/reimport boundary before cloning",
);
assert.equal(smartArtTextImported.slides.items.length, smartArtTextSlideCount);
const smartArtTextExport = await PresentationFile.exportPptx(smartArtTextImported);
assert.deepEqual(smartArtTextSource.bytes, smartArtTextInput, "SmartArt text edits must preserve the caller input bytes");
const smartArtTextSourceZip = await JSZip.loadAsync(smartArtTextSource.bytes);
const smartArtTextOutputZip = await JSZip.loadAsync(smartArtTextExport.bytes);
for (const path of [
  "ppt/slides/slide1.xml",
  "ppt/slides/_rels/slide1.xml.rels",
  "ppt/diagrams/agent-layout.xml",
  "ppt/diagrams/agent-style.xml",
  "ppt/diagrams/agent-colors.xml",
]) {
  assert.deepEqual(
    await smartArtTextOutputZip.file(path).async("uint8array"),
    await smartArtTextSourceZip.file(path).async("uint8array"),
    `SmartArt text edits must not alter ${path}`,
  );
}
const smartArtTextOutputData = await smartArtTextOutputZip.file("ppt/diagrams/agent-data.xml").async("text");
assert.match(smartArtTextOutputData, / Revised node /);
assert.match(smartArtTextOutputData, /xml:space="preserve"/);
assert.notDeepEqual(
  await smartArtTextOutputZip.file("ppt/diagrams/agent-data.xml").async("uint8array"),
  await smartArtTextSourceZip.file("ppt/diagrams/agent-data.xml").async("uint8array"),
);
const smartArtTextRoundTrip = await PresentationFile.importPptx(smartArtTextExport);
const smartArtTextRebound = itemByName(smartArtTextRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
assert.deepEqual(smartArtTextRebound.diagramText?.nodes, [
  { id: "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", text: " Revised node " },
  { id: "1", text: "Second node" },
]);
assert.notEqual(smartArtTextRebound.diagramText?.sourceSha256, smartArtTextObject.diagramText?.sourceSha256);

const richSmartArtTextSource = await PresentationFile.patchPptx(smartArtTextSource, [{
  path: "ppt/diagrams/agent-data.xml",
  xml: smartArtTextData.replace("<a:r><a:t>Original node</a:t></a:r>", "<a:r><a:t>Original</a:t></a:r><a:r><a:t> node</a:t></a:r>"),
}]);
const richSmartArtText = await PresentationFile.importPptx(richSmartArtTextSource);
assert.equal(itemByName(richSmartArtText.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt").diagramText, undefined,
  "multi-run SmartArt text must remain opaque rather than being flattened");

const invalidSmartArtModelIdSource = await PresentationFile.patchPptx(smartArtTextSource, [{
  path: "ppt/diagrams/agent-data.xml",
  xml: smartArtTextData.replace("{B31B1833-2B65-4D6B-B3D4-9B3988427B21}", "agent-node-1"),
}]);
const invalidSmartArtModelId = await PresentationFile.importPptx(invalidSmartArtModelIdSource);
assert.equal(itemByName(invalidSmartArtModelId.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt").diagramText, undefined,
  "an invalid ST_ModelId must not expose a SmartArt text-edit capability");

const connectedSmartArtSource = await PresentationFile.patchPptx(smartArtSource, [{
  path: "ppt/diagrams/_rels/agent-data.xml.rels",
  xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeSmartArtLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/smartart" TargetMode="External"/></Relationships>',
}]);
const connectedSmartArt = await PresentationFile.importPptx(connectedSmartArtSource);
const connectedSmartArtSlideCount = connectedSmartArt.slides.items.length;
assert.throws(
  () => connectedSmartArt.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a relationship-bearing SmartArt part must fail before mutating the model",
);
assert.equal(connectedSmartArt.slides.items.length, connectedSmartArtSlideCount);

const nestedSmartArt = await PresentationFile.importPptx(smartArtSource);
const nestedSmartArtObject = itemByName(nestedSmartArt.slides.getItem(0).nativeObjects.items, "Clone-safe SmartArt");
nestedSmartArtObject.rawXml = nestedSmartArtObject.rawXml.replace(/^<p:graphicFrame/, "<p:grpSp");
const nestedSmartArtSlideCount = nestedSmartArt.slides.items.length;
assert.throws(
  () => nestedSmartArt.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a SmartArt graph whose source binding is not a top-level graphicFrame must fail before mutating the model",
);
assert.equal(nestedSmartArt.slides.items.length, nestedSmartArtSlideCount);

const foreignRelationshipNamespaceSource = await PresentationFile.patchPptx(smartArtSource, [{
  path: "ppt/slides/slide1.xml",
  xml: (await smartArtSourceZip.file("ppt/slides/slide1.xml").async("text")).replace(
    smartArtFrame,
    smartArtFrame.replace(
      "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "https://example.invalid/relationships",
    ),
  ),
}]);
const foreignRelationshipNamespaceSmartArt = await PresentationFile.importPptx(foreignRelationshipNamespaceSource);
const foreignNamespaceSlideCount = foreignRelationshipNamespaceSmartArt.slides.items.length;
assert.throws(
  () => foreignRelationshipNamespaceSmartArt.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a relationship-like but non-OOXML SmartArt namespace must fail before mutating the model",
);
assert.equal(foreignRelationshipNamespaceSmartArt.slides.items.length, foreignNamespaceSlideCount);

// A canonical top-level p:contentPart is the PresentationML carrier for one
// standard InkML CustomXmlPart. The clone must allocate a new InkML part under
// the same slide-local r:id rather than sharing mutable ink XML.
const inkContentElement = '<p:contentPart xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdAgentInk"><p14:nvContentPartPr><p14:cNvPr id="121" name="Clone-safe ink"/><p14:cNvContentPartPr/><p14:nvPr/></p14:nvContentPartPr><p14:xfrm><a:off x="914400" y="1828800"/><a:ext cx="4572000" cy="2286000"/></p14:xfrm></p:contentPart>';
const inkRelationship = '<Relationship Id="rIdAgentInk" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/agent-ink.xml"/>';
const inkXml = '<ink xmlns="http://www.w3.org/2003/InkML"><trace>0 0, 100 100, 200 0</trace></ink>';
const inkSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${inkContentElement}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${inkRelationship}</Relationships>`) },
  { path: "ppt/customXml/agent-ink.xml", contentType: "application/inkml+xml", xml: inkXml },
]);
const inkSourceSnapshot = Uint8Array.from(inkSource.bytes);
const inkImported = await PresentationFile.importPptx(inkSource);
const inkOriginSlide = inkImported.slides.getItem(0);
const inkOrigin = itemByName(inkOriginSlide.nativeObjects.items, "Clone-safe ink");
assert.equal(inkOrigin.nativeKind, "contentPart");
assert.deepEqual(inkOrigin.position, { left: 96, top: 192, width: 480, height: 240 });
assert.equal(inkOrigin.parts.length, 1);
assert.equal(inkOrigin.parts[0].contentType, "application/inkml+xml");
assert.equal(inkOrigin.parts[0].relationships.length, 0);
const inkPendingSlide = inkOriginSlide.duplicate();
const inkPending = itemByName(inkPendingSlide.nativeObjects.items, "Clone-safe ink");
assert.notEqual(inkPending, inkOrigin);
assert.notEqual(inkPending.id, inkOrigin.id);
assert.equal(inkPending.parts[0].path, inkOrigin.parts[0].path);

const inkExport = await PresentationFile.exportPptx(inkImported);
assert.deepEqual(inkSource.bytes, inkSourceSnapshot);
const inkSourceZip = await JSZip.loadAsync(inkSource.bytes);
const inkOutputZip = await JSZip.loadAsync(inkExport.bytes);
assert.deepEqual(
  await inkOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await inkSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "InkML cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await inkOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await inkSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "InkML cloning must retain the origin relationship part byte-for-byte",
);
const inkSourceRelationship = relationshipForType(await inkOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), "customXml");
const inkCloneRelationship = relationshipForType(await inkOutputZip.file("ppt/slides/_rels/slide3.xml.rels").async("text"), "customXml");
const inkSourcePartPath = resolveSlideRelationshipTarget(inkSourceRelationship.target);
const inkClonePartPath = resolveSlideRelationshipTarget(inkCloneRelationship.target);
assert.equal(inkCloneRelationship.id, inkSourceRelationship.id);
assert.equal(inkSourcePartPath, "ppt/customXml/agent-ink.xml");
assert.match(inkClonePartPath, /^ppt\/customXml\/item\d+\.xml$/i);
assert.notEqual(inkClonePartPath, inkSourcePartPath);
assert.deepEqual(
  await inkOutputZip.file(inkClonePartPath).async("uint8array"),
  await inkSourceZip.file(inkSourcePartPath).async("uint8array"),
  "InkML cloning must byte-copy the closed content part",
);

const inkRoundTrip = await PresentationFile.importPptx(inkExport);
const inkRoundTripOrigin = itemByName(inkRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe ink");
const inkRoundTripClone = itemByName(inkRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe ink");
assert.notEqual(inkRoundTripOrigin.parts[0].path, inkRoundTripClone.parts[0].path);
assert.equal(inkRoundTripOrigin.parts[0].sourceSha256, inkRoundTripClone.parts[0].sourceSha256);

const connectedInkSource = await PresentationFile.patchPptx(inkSource, [{
  path: "ppt/customXml/_rels/agent-ink.xml.rels",
  xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdUnsafeInkLink" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://example.invalid/ink" TargetMode="External"/></Relationships>',
}]);
const connectedInk = await PresentationFile.importPptx(connectedInkSource);
const connectedInkSlideCount = connectedInk.slides.items.length;
assert.throws(
  () => connectedInk.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a relationship-bearing InkML part must fail before mutating the model",
);
assert.equal(connectedInk.slides.items.length, connectedInkSlideCount);

const nestedInk = await PresentationFile.importPptx(inkSource);
const nestedInkObject = itemByName(nestedInk.slides.getItem(0).nativeObjects.items, "Clone-safe ink");
nestedInkObject.rawXml = nestedInkObject.rawXml.replace(/^<p:contentPart/, "<p:grpSp");
const nestedInkSlideCount = nestedInk.slides.items.length;
assert.throws(
  () => nestedInk.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "an InkML graph whose source binding is not a top-level contentPart must fail before mutating the model",
);
assert.equal(nestedInk.slides.items.length, nestedInkSlideCount);

// PowerPoint represents an embedded video as one top-level picture with a
// poster ImagePart plus paired video/media data relationships to one MP4.
// The bounded clone shares the immutable poster but copies the MP4 into a new
// SDK-allocated MediaDataPart, preserving both slide-local relationship IDs.
const embeddedVideoBytes = Buffer.from("AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDEAAAMVbW9vdgAAAGxtdmhkAAAAAAAAAAAAAAAAAAAD6AAAACgAAQAAAQAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAAAj90cmFrAAAAXHRraGQAAAADAAAAAAAAAAAAAAABAAAAAAAAACgAAAAAAAAAAAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAABAAAAAABAAAAAQAAAAAAAkZWR0cwAAABxlbHN0AAAAAAAAAAEAAAAoAAAAAAABAAAAAAG3bWRpYQAAACBtZGhkAAAAAAAAAAAAAAAAAAAyAAAAAgBVxAAAAAAALWhkbHIAAAAAAAAAAHZpZGUAAAAAAAAAAAAAAABWaWRlb0hhbmRsZXIAAAABYm1pbmYAAAAUdm1oZAAAAAEAAAAAAAAAAAAAACRkaW5mAAAAHGRyZWYAAAAAAAAAAQAAAAx1cmwgAAAAAQAAASJzdGJsAAAAvnN0c2QAAAAAAAAAAQAAAK5hdmMxAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAABAAEABIAAAASAAAAAAAAAABFUxhdmM2Mi4yOC4xMDIgbGlieDI2NAAAAAAAAAAAAAAAGP//AAAANGF2Y0MBZAAK/+EAF2dkAAqs2V7ARAAAAwAEAAADAMg8SJZYAQAGaOvjyyLA/fj4AAAAABBwYXNwAAAAAQAAAAEAAAAUYnRydAAAAAAAAinoAAAAAAAAABhzdHRzAAAAAAAAAAEAAAABAAACAAAAABxzdHNjAAAAAAAAAAEAAAABAAAAAQAAAAEAAAAUc3RzegAAAAAAAALFAAAAAQAAABRzdGNvAAAAAAAAAAEAAANFAAAAYnVkdGEAAABabWV0YQAAAAAAAAAhaGRscgAAAAAAAAAAbWRpcmFwcGwAAAAAAAAAAAAAAAAtaWxzdAAAACWpdG9vAAAAHWRhdGEAAAABAAAAAExhdmY2Mi4xMi4xMDIAAAAIZnJlZQAAAs1tZGF0AAACrgYF//+q3EXpvebZSLeWLNgg2SPu73gyNjQgLSBjb3JlIDE2NSByMzIyMiBiMzU2MDVhIC0gSC4yNjQvTVBFRy00IEFWQyBjb2RlYyAtIENvcHlsZWZ0IDIwMDMtMjAyNSAtIGh0dHA6Ly93d3cudmlkZW9sYW4ub3JnL3gyNjQuaHRtbCAtIG9wdGlvbnM6IGNhYmFjPTEgcmVmPTMgZGVibG9jaz0xOjA6MCBhbmFseXNpPTB4MzoweDExMyBtZT1oZXggc3VibWU9NyBwc3k9MSBwc3lfcmQ9MS4wMDowLjAwIG1peGVkX3JlZj0xIG1lX3JhbmdlPTE2IGNocm9tYV9tZT0xIHRyZWxsaXM9MSA4eDhkY3Q9MSBjcW09MCBkZWFkem9uZT0yMSwxMSBmYXN0X3Bza2lwPTEgY2hyb21hX3FwX29mZnNldD0tMiB0aHJlYWRzPTEgbG9va2FoZWFkX3RocmVhZHM9MSBzbGljZWRfdGhyZWFkcz0wIG5yPTAgZGVjaW1hdGU9MSBpbnRlcmxhY2VkPTAgYmx1cmF5X2NvbXBhdD0wIGNvbnN0cmFpbmVkX2ludHJhPTAgYmZyYW1lcz0zIGJfcHlyYW1pZD0yIGJfYWRhcHQ9MSBiX2JpYXM9MCBkaXJlY3Q9MSB3ZWlnaHRiPTEgb3Blbl9nb3A9MCB3ZWlnaHRwPTIga2V5aW50PTI1MCBrZXlpbnRfbWluPTI1IHNjZW5lY3V0PTQwIGludHJhX3JlZnJlc2g9MCByY19sb29rYWhlYWQ9NDAgcmM9Y3JmIG1idHJlZT0xIGNyZj0yMy4wIHFjb21wPTAuNjAgcXBtaW49MCBxcG1heD02OSBxcHN0ZXA9NCBpcF9yYXRpbz0xLjQwIGFxPTE6MS4wMACAAAAAD2WIhAAr//72c3wKa22xgQ==", "base64");
const mediaPicture = '<p:pic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvPicPr><p:cNvPr id="122" name="Clone-safe video"><a:hlinkClick r:id="" action="ppaction://media"/></p:cNvPr><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr><a:videoFile r:link="rIdAgentVideo"/><p:extLst><p:ext uri="{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}"><p14:media r:embed="rIdAgentMedia"/></p:ext></p:extLst></p:nvPr></p:nvPicPr><p:blipFill><a:blip r:embed="rIdAgentVideoPoster"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="1828800"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';
const mediaRelationships = '<Relationship Id="rIdAgentVideo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="../media/agent-video.mp4"/><Relationship Id="rIdAgentMedia" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="../media/agent-video.mp4"/><Relationship Id="rIdAgentVideoPoster" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/agent-video-poster.png"/>';
const mediaSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${mediaPicture}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${mediaRelationships}</Relationships>`) },
  { path: "ppt/media/agent-video.mp4", bytes: embeddedVideoBytes, contentType: "video/mp4" },
  { path: "ppt/media/agent-video-poster.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const mediaSourceSnapshot = Uint8Array.from(mediaSource.bytes);
const mediaImported = await PresentationFile.importPptx(mediaSource);
const mediaOriginSlide = mediaImported.slides.getItem(0);
const mediaOrigin = itemByName(mediaOriginSlide.nativeObjects.items, "Clone-safe video");
assert.equal(mediaOrigin.nativeKind, "media");
assert.equal(mediaOrigin.relationshipReferences.length, 3);
assert.equal(mediaOrigin.parts.length, 2);
assert.equal(mediaOrigin.parts.filter((part) => part.contentType === "video/mp4").length, 1);
assert.equal(mediaOrigin.parts.filter((part) => part.contentType === "image/png").length, 1);
const mediaPendingSlide = mediaOriginSlide.duplicate();
const mediaPending = itemByName(mediaPendingSlide.nativeObjects.items, "Clone-safe video");
assert.notEqual(mediaPending, mediaOrigin);
assert.notEqual(mediaPending.id, mediaOrigin.id);
assert.deepEqual(mediaPending.parts.map((part) => part.path), mediaOrigin.parts.map((part) => part.path));

const mediaExport = await PresentationFile.exportPptx(mediaImported);
assert.deepEqual(mediaSource.bytes, mediaSourceSnapshot);
const mediaSourceZip = await JSZip.loadAsync(mediaSource.bytes);
const mediaOutputZip = await JSZip.loadAsync(mediaExport.bytes);
assert.deepEqual(
  await mediaOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await mediaSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
  "embedded-video cloning must retain the origin SlidePart byte-for-byte",
);
assert.deepEqual(
  await mediaOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  await mediaSourceZip.file("ppt/slides/_rels/slide1.xml.rels").async("uint8array"),
  "embedded-video cloning must retain the origin relationship part byte-for-byte",
);
const mediaSourceRels = await mediaOutputZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const mediaCloneRels = await mediaOutputZip.file("ppt/slides/_rels/slide3.xml.rels").async("text");
const sourceVideoRelationship = relationshipForType(mediaSourceRels, "video");
const sourceMediaRelationship = relationshipForType(mediaSourceRels, "media");
const sourcePosterRelationship = relationshipForType(mediaSourceRels, "image");
const cloneVideoRelationship = relationshipForType(mediaCloneRels, "video");
const cloneMediaRelationship = relationshipForType(mediaCloneRels, "media");
const clonePosterRelationship = relationshipForType(mediaCloneRels, "image");
assert.equal(cloneVideoRelationship.id, sourceVideoRelationship.id);
assert.equal(cloneMediaRelationship.id, sourceMediaRelationship.id);
assert.equal(clonePosterRelationship.id, sourcePosterRelationship.id);
const sourceVideoPartPath = resolveSlideRelationshipTarget(sourceVideoRelationship.target);
const sourceMediaPartPath = resolveSlideRelationshipTarget(sourceMediaRelationship.target);
const cloneVideoPartPath = resolveSlideRelationshipTarget(cloneVideoRelationship.target);
const cloneMediaPartPath = resolveSlideRelationshipTarget(cloneMediaRelationship.target);
assert.equal(sourceVideoPartPath, sourceMediaPartPath);
assert.equal(cloneVideoPartPath, cloneMediaPartPath);
assert.match(cloneVideoPartPath, /^(?:ppt\/)?media\/[^/]+\.mp4$/i);
assert.notEqual(cloneVideoPartPath, sourceVideoPartPath);
assert.deepEqual(
  await mediaOutputZip.file(cloneVideoPartPath).async("uint8array"),
  await mediaSourceZip.file(sourceVideoPartPath).async("uint8array"),
  "embedded-video cloning must byte-copy the accepted MP4 into an independent MediaDataPart",
);
assert.equal(
  resolveSlideRelationshipTarget(clonePosterRelationship.target),
  resolveSlideRelationshipTarget(sourcePosterRelationship.target),
  "embedded-video cloning must share the immutable poster ImagePart",
);

const mediaRoundTrip = await PresentationFile.importPptx(mediaExport);
const mediaRoundTripOrigin = itemByName(mediaRoundTrip.slides.getItem(0).nativeObjects.items, "Clone-safe video");
const mediaRoundTripClone = itemByName(mediaRoundTrip.slides.getItem(1).nativeObjects.items, "Clone-safe video");
const mediaRoundTripOriginVideo = mediaRoundTripOrigin.parts.find((part) => part.contentType === "video/mp4");
const mediaRoundTripCloneVideo = mediaRoundTripClone.parts.find((part) => part.contentType === "video/mp4");
const mediaRoundTripOriginPoster = mediaRoundTripOrigin.parts.find((part) => part.contentType.startsWith("image/"));
const mediaRoundTripClonePoster = mediaRoundTripClone.parts.find((part) => part.contentType.startsWith("image/"));
assert.notEqual(mediaRoundTripOriginVideo.path, mediaRoundTripCloneVideo.path);
assert.equal(mediaRoundTripOriginVideo.sourceSha256, mediaRoundTripCloneVideo.sourceSha256);
assert.equal(mediaRoundTripOriginPoster.path, mediaRoundTripClonePoster.path);

const malformedMedia = await PresentationFile.importPptx(mediaSource);
const malformedMediaObject = itemByName(malformedMedia.slides.getItem(0).nativeObjects.items, "Clone-safe video");
malformedMediaObject.rawXml = malformedMediaObject.rawXml.replace(
  "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}",
  "{00000000-0000-0000-0000-000000000000}",
);
const malformedMediaSlideCount = malformedMedia.slides.items.length;
assert.throws(
  () => malformedMedia.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a media picture with a non-canonical extension must fail before mutating the model",
);
assert.equal(malformedMedia.slides.items.length, malformedMediaSlideCount);

const extensionRichMedia = await PresentationFile.importPptx(mediaSource);
const extensionRichMediaObject = itemByName(extensionRichMedia.slides.getItem(0).nativeObjects.items, "Clone-safe video");
extensionRichMediaObject.rawXml = extensionRichMediaObject.rawXml.replace(
  "</p:extLst>",
  '<p:ext uri="{00000000-0000-0000-0000-000000000000}"><p14:placeholder/></p:ext></p:extLst>',
);
const extensionRichMediaSlideCount = extensionRichMedia.slides.items.length;
assert.throws(
  () => extensionRichMedia.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a media picture with an extra extension must fail before mutating the model",
);
assert.equal(extensionRichMedia.slides.items.length, extensionRichMediaSlideCount);

const wrongTypeMediaSource = await PresentationFile.patchPptx(cloneSourcePptx, [
  { path: "ppt/slides/slide1.xml", xml: oleCloneBaseSlideXml.replace("</p:spTree>", `${mediaPicture}</p:spTree>`) },
  { path: "ppt/slides/_rels/slide1.xml.rels", xml: oleCloneBaseRelationships.replace("</Relationships>", `${mediaRelationships}</Relationships>`) },
  { path: "ppt/media/agent-video.mp4", bytes: embeddedVideoBytes, contentType: "video/quicktime" },
  { path: "ppt/media/agent-video-poster.png", bytes: embeddedPreviewBytes, contentType: "image/png" },
]);
const wrongTypeMedia = await PresentationFile.importPptx(wrongTypeMediaSource);
const wrongTypeMediaSlideCount = wrongTypeMedia.slides.items.length;
assert.throws(
  () => wrongTypeMedia.slides.getItem(0).duplicate(),
  (error) => error?.code === "unsupported_presentation_slide_clone",
  "a non-MP4 media payload must fail before mutating the model",
);
assert.equal(wrongTypeMedia.slides.items.length, wrongTypeMediaSlideCount);

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
assert.deepEqual(imported.slides.getItem(0).speakerNotes.capability, {
  sourceBound: true,
  partPresent: true,
  editable: true,
  addable: false,
});
assert.deepEqual(imported.slides.getItem(1).speakerNotes.capability, {
  sourceBound: true,
  partPresent: false,
  editable: false,
  addable: true,
});
imported.slides.getItem(0).addNotes("Lead with evidence.\nClose with the decision.");
imported.slides.getItem(1).addNotes("Explain the chart assumptions.\nInvite questions on the forecast.");
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
assert.equal(Object.keys(secondZip.files).filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name)).length, 2);
assert.ok((await Promise.all(
  Object.keys(secondZip.files)
    .filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/.test(name))
    .map(async (name) => (await secondZip.file(name).async("text")).includes("Explain the chart assumptions")),
)).includes(true));

const roundTrip = await PresentationFile.importPptx(secondExport);
assert.equal(roundTrip.master.name, "Source Master Marker");
assert.equal(roundTrip.layouts.items[0].name, "Source Layout Marker");
const roundTripCore = roundTrip.slides.getItem(0);
assert.equal(roundTripCore.speakerNotes.text, "Lead with evidence.\nClose with the decision.");
assert.equal(roundTrip.slides.getItem(1).speakerNotes.text, "Explain the chart assumptions.\nInvite questions on the forecast.");
assert.deepEqual(roundTrip.slides.getItem(1).speakerNotes.capability, {
  sourceBound: true,
  partPresent: true,
  editable: true,
  addable: false,
});
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
const legacyAdditionSourceDeck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const legacyAdditionTarget = legacyAdditionSourceDeck.slides.add({ name: "Imported review target" });
legacyAdditionTarget.shapes.add({
  name: "visible-review-title",
  geometry: "textbox",
  text: "Visible content must not change",
  position: { left: 96, top: 96, width: 900, height: 88 },
});
const legacyAdditionControl = legacyAdditionSourceDeck.slides.add({ name: "Imported review control" });
legacyAdditionControl.shapes.add({
  name: "visible-control-title",
  geometry: "textbox",
  text: "Control slide",
  position: { left: 96, top: 96, width: 900, height: 88 },
});
const legacyAdditionSource = await PresentationFile.exportPptx(legacyAdditionSourceDeck);
const legacyAdditionSourceBytes = new Uint8Array(await legacyAdditionSource.arrayBuffer());
const legacyAdditionImported = await PresentationFile.importPptx(legacyAdditionSource);
assert.deepEqual(legacyAdditionImported.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: false,
  addable: true,
});
assert.match(legacyAdditionImported.inspect({ kind: "slide" }).ndjson, /"commentsCapability":\{"sourceBound":true,"format":"legacy","partPresent":false,"addable":true\}/);
legacyAdditionImported.slides.getItem(0).comments.addThread(undefined, "Confirm the imported evidence.", {
  author: "Review Owner",
  created: "2026-07-20T03:04:05Z",
  position: { x: 360, y: 240 },
});
const legacyAdditionExport = await PresentationFile.exportPptx(legacyAdditionImported);
const legacyAdditionOutputBytes = new Uint8Array(await legacyAdditionExport.arrayBuffer());
const legacyAdditionSourceZip = await JSZip.loadAsync(legacyAdditionSourceBytes);
const legacyAdditionOutputZip = await JSZip.loadAsync(legacyAdditionOutputBytes);
assert.deepEqual(
  await legacyAdditionOutputZip.file("ppt/slides/slide1.xml").async("uint8array"),
  await legacyAdditionSourceZip.file("ppt/slides/slide1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyAdditionOutputZip.file("ppt/slides/slide2.xml").async("uint8array"),
  await legacyAdditionSourceZip.file("ppt/slides/slide2.xml").async("uint8array"),
);
assert.ok(legacyAdditionOutputZip.file("ppt/commentAuthors.xml"));
assert.ok(legacyAdditionOutputZip.file("ppt/comments/comment1.xml"));
const legacyAdditionRoundTrip = await PresentationFile.importPptx(legacyAdditionExport);
assert.equal(legacyAdditionRoundTrip.slides.getItem(0).comments.items[0].comments[0].text, "Confirm the imported evidence.");
assert.deepEqual(legacyAdditionRoundTrip.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: true,
  addable: false,
});
assert.deepEqual(legacyAdditionRoundTrip.slides.getItem(1).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: false,
  addable: false,
});

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
assert.deepEqual(legacyCommentImported.slides.getItem(0).comments.capability, {
  sourceBound: true,
  format: "legacy",
  partPresent: true,
  addable: false,
});
const importedLegacyThread = legacyCommentImported.slides.getItem(0).comments.items[0];
assert.equal(importedLegacyThread.nativeFormat, "legacy");
assert.equal(importedLegacyThread.targetId, undefined);
assert.equal(importedLegacyThread.comments.length, 1);
assert.equal(importedLegacyThread.comments[0].author, "Review Owner");
assert.equal(importedLegacyThread.comments[0].text, "Confirm the source before delivery.");
assert.deepEqual(importedLegacyThread.position, { x: 360, y: 240, unit: "px" });

// The bounded imported-slide clone profile may carry a closed legacy-comments
// leaf. The clone gets a distinct model/thread object and comments part, while
// both p:cm records keep their IDs against the one immutable author catalog.
const legacyCommentCloneDeck = await PresentationFile.importPptx(legacyCommentExport);
const legacyCommentCloneSource = legacyCommentCloneDeck.slides.getItem(0);
const legacyCommentClone = legacyCommentCloneSource.duplicate();
assert.equal(legacyCommentClone.comments.items.length, 1);
assert.notEqual(legacyCommentClone.comments.items[0], legacyCommentCloneSource.comments.items[0]);
assert.equal(legacyCommentClone.comments.items[0].comments[0].text, "Confirm the source before delivery.");
const legacyCommentCloneExport = await PresentationFile.exportPptx(legacyCommentCloneDeck);
const legacyCommentCloneZip = await JSZip.loadAsync(new Uint8Array(await legacyCommentCloneExport.arrayBuffer()));
assert.ok(legacyCommentCloneZip.file("ppt/comments/comment2.xml"));
assert.deepEqual(
  await legacyCommentCloneZip.file("ppt/comments/comment2.xml").async("uint8array"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyCommentCloneZip.file("ppt/comments/comment1.xml").async("uint8array"),
  await legacyCommentZip.file("ppt/comments/comment1.xml").async("uint8array"),
);
assert.deepEqual(
  await legacyCommentCloneZip.file("ppt/commentAuthors.xml").async("uint8array"),
  await legacyCommentZip.file("ppt/commentAuthors.xml").async("uint8array"),
);
const legacyCommentCloneRoundTrip = await PresentationFile.importPptx(legacyCommentCloneExport);
assert.equal(legacyCommentCloneRoundTrip.slides.items.length, 2);
assert.deepEqual(
  legacyCommentCloneRoundTrip.slides.items.map((slide) => slide.comments.items[0].comments[0].text),
  ["Confirm the source before delivery.", "Confirm the source before delivery."],
);

const editedLegacyCommentCloneDeck = await PresentationFile.importPptx(legacyCommentExport);
const editedLegacyCommentClone = editedLegacyCommentCloneDeck.slides.getItem(0).duplicate();
editedLegacyCommentClone.comments.items[0].comments[0].text = "This comment cannot change before the clone boundary.";
await assert.rejects(
  () => PresentationFile.exportPptx(editedLegacyCommentCloneDeck),
  (error) => error?.code === "unsupported_presentation_edit",
);

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

// Office 2021 modern comments use their native author/comments graph instead
// of the legacy slide annotation part. OpenChestnut owns a bounded root +
// direct replies profile with top-level drawing and shape-text-range anchors.
const modernCommentDeck = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  commentFormat: "modern",
});
const modernCommentSlide = modernCommentDeck.slides.add({ name: "Modern comments" });
const modernCommentTarget = modernCommentSlide.shapes.add({
  id: "modern-comment-target",
  name: "Decision evidence",
  geometry: "rect",
  position: { left: 80, top: 80, width: 520, height: 120 },
  text: "Customer evidence is ready",
});
const modernCommentThread = modernCommentSlide.comments.addThread({
  textMatch: { element: modernCommentTarget, query: "Customer evidence is ready", occurrence: 0 },
}, "Confirm the customer evidence.", {
  id: "{11111111-1111-4111-8111-111111111111}",
  author: "Review Owner",
  created: "2026-07-19T02:55:00Z",
  nativeFormat: "modern",
  position: { x: 1_234_500, y: 2_345_600, unit: "emu" },
  comments: [{
    nativeId: "{11111111-1111-4111-8111-111111111111}",
    author: "Review Owner",
    person: {
      id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      name: "Review Owner",
      initials: "RO",
      userId: "review.owner@example.test",
      providerId: "None",
    },
    text: "Confirm the customer evidence.",
    created: "2026-07-19T02:55:00Z",
    status: "active",
  }],
});
modernCommentThread.addReply("Evidence is attached.", {
  nativeId: "{22222222-2222-4222-8222-222222222222}",
  author: "Evidence Owner",
  person: {
    id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    name: "Evidence Owner",
    initials: "EO",
    userId: "evidence.owner@example.test",
    providerId: "None",
  },
  created: "2026-07-19T03:05:00Z",
  status: "active",
});
assert.equal(modernCommentDeck.verify().ok, true);
const modernCommentExport = await PresentationFile.exportPptx(modernCommentDeck);
const modernCommentZip = await JSZip.loadAsync(new Uint8Array(await modernCommentExport.arrayBuffer()));
const modernCommentPartPath = Object.keys(modernCommentZip.files).find((name) => /^ppt\/comments\/(?:modernComment|comment)\d*\.xml$/.test(name));
const modernAuthorsPartPath = Object.keys(modernCommentZip.files).find((name) => /^ppt\/(?:authors|authors\/author\d+)\.xml$/.test(name));
assert.ok(modernCommentPartPath);
assert.ok(modernAuthorsPartPath);
const modernCommentXml = await modernCommentZip.file(modernCommentPartPath).async("text");
const modernAuthorsXml = await modernCommentZip.file(modernAuthorsPartPath).async("text");
assert.match(modernCommentXml, /<p188:replyLst>/);
assert.match(modernCommentXml, /<oac:txMkLst>/);
assert.match(modernCommentXml, /<oac:txMk cp="0" len="26"/);
assert.match(modernAuthorsXml, /Review Owner/);
assert.match(modernAuthorsXml, /Evidence Owner/);

const modernCommentImported = await PresentationFile.importPptx(modernCommentExport);
assert.equal(modernCommentImported.commentFormat, "modern");
const importedModernSlide = modernCommentImported.slides.getItem(0);
const importedModernThread = importedModernSlide.comments.items[0];
assert.equal(importedModernThread.nativeFormat, "modern");
assert.equal(importedModernThread.comments.length, 2);
assert.equal(importedModernThread.comments[0].author, "Review Owner");
assert.equal(importedModernThread.comments[1].author, "Evidence Owner");
assert.equal(importedModernThread.nativeAnchor.type, "textRange");
assert.equal(importedModernThread.nativeAnchor.textLength, 26);
assert.equal(importedModernSlide.resolve(importedModernThread.targetId).kind, "textRange");

const modernUnchanged = await PresentationFile.exportPptx(modernCommentImported);
const modernUnchangedZip = await JSZip.loadAsync(new Uint8Array(await modernUnchanged.arrayBuffer()));
assert.deepEqual(
  await modernUnchangedZip.file(modernCommentPartPath).async("uint8array"),
  await modernCommentZip.file(modernCommentPartPath).async("uint8array"),
);
assert.deepEqual(
  await modernUnchangedZip.file(modernAuthorsPartPath).async("uint8array"),
  await modernCommentZip.file(modernAuthorsPartPath).async("uint8array"),
);

importedModernThread.comments[0].text = "Customer evidence confirmed.";
importedModernThread.comments[1].text = "Recorded in the decision log.";
importedModernThread.resolve();
const modernEdited = await PresentationFile.exportPptx(modernCommentImported);
const modernEditedRoundTrip = await PresentationFile.importPptx(modernEdited);
const editedModernThread = modernEditedRoundTrip.slides.getItem(0).comments.items[0];
assert.equal(editedModernThread.comments[0].text, "Customer evidence confirmed.");
assert.equal(editedModernThread.comments[1].text, "Recorded in the decision log.");
assert.equal(editedModernThread.resolved, true);
assert.equal(editedModernThread.comments[0].status, "resolved");

editedModernThread.comments[0].author = "Changed identity";
await assert.rejects(
  () => PresentationFile.exportPptx(modernEditedRoundTrip),
  (error) => error?.code === "presentation_comment_topology_changed",
);

const invalidModernCommentDeck = Presentation.create({ commentFormat: "modern" });
const invalidModernCommentSlide = invalidModernCommentDeck.slides.add();
const invalidModernTarget = invalidModernCommentSlide.shapes.add({ text: "Short" });
invalidModernCommentSlide.comments.addThread(`${invalidModernTarget.id}/text`, "Out of bounds.", {
  nativeFormat: "modern",
  nativeAnchor: { type: "textRange", cp: 3, length: 99 },
  author: "Reviewer",
  created: "2026-07-19T04:00:00Z",
  position: { x: 100, y: 100, unit: "emu" },
});
await assert.rejects(
  () => PresentationFile.exportPptx(invalidModernCommentDeck),
  (error) => error?.code === "invalid_presentation_modern_comment",
);

const missingModernCommentPositionDeck = Presentation.create({ commentFormat: "modern" });
const missingModernCommentPositionSlide = missingModernCommentPositionDeck.slides.add();
const missingModernCommentPositionTarget = missingModernCommentPositionSlide.shapes.add({ text: "Coordinate required" });
missingModernCommentPositionSlide.comments.addThread(missingModernCommentPositionTarget, "No implicit origin.", {
  nativeFormat: "modern",
  author: "Reviewer",
  created: "2026-07-19T04:10:00Z",
});
await assert.rejects(
  () => PresentationFile.exportPptx(missingModernCommentPositionDeck),
  (error) => error?.code === "invalid_presentation_modern_comment" && /explicit.*position/i.test(error.message),
);

console.log("presentation smoke ok");
