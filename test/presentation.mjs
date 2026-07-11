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
  series: [{ name: "Revenue", values: [10, 14] }],
});
const nativeImage = slide.images.add({
  name: "native-import-image",
  alt: "Native import logo",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  position: { left: 840, top: 450, width: 120, height: 90 },
});
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
assert.equal(presentation.resolve(nativeImage.id).alt, "Native import logo");

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
assert.match(qaBroken.verify({ maxChars: 8000 }).ndjson, /danglingComment/);
assert.match(qaBroken.verify({ maxChars: 8000 }).ndjson, /placeholderMissingContent/);
assert.match(qaBroken.help("presentation.validateLayout").ndjson, /off-canvas/);

const layout = await slide.export({ format: "layout" });
assert.equal(layout.type, "application/vnd.open-office-artifact.layout+json");
assert.match(await layout.text(), /summary-surface/);
const preview = await presentation.export({ slide, format: "svg" });
assert.equal(preview.type, "image/svg+xml");

const pptx = await PresentationFile.exportPptx(presentation);
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
assert.match(loadedAll, /native-import-image/);
assert.match(loadedAll, /Native import logo/);
assert.match(loadedAll, /Speaker note/);
assert.match(loadedAll, /Tighten this headline/);
assert.match(loadedAll, /shape-to-table/);
console.log("presentation smoke ok");
