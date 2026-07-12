import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { box, column, FileBlob, paragraph, Presentation, PresentationFile, row, run, rule, shape as composeShape } from "../src/index.mjs";

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const aliasedThemePresentation = Presentation.create({ theme: { colors: { dk1: "#112233", lt1: "#fefefe" } } });
assert.equal(aliasedThemePresentation.theme.colors.tx1, "#112233");
assert.equal(aliasedThemePresentation.theme.colors.bg1, "#fefefe");
assert.equal("dk1" in aliasedThemePresentation.theme.colors, false);
const masterOnlyPresentation = Presentation.create({ master: { name: "Master Only", background: "#123456" } });
masterOnlyPresentation.slides.add();
const masterOnlyLoaded = await PresentationFile.importPptx(await PresentationFile.exportPptx(masterOnlyPresentation));
assert.equal(masterOnlyLoaded.layouts.items.length, 1);
assert.equal(masterOnlyLoaded.master.name, "Master Only");
assert.deepEqual(masterOnlyLoaded.slides.items[0].effectiveBackground(), { fill: "#123456", mode: "solid" });
const multiMasterPresentation = Presentation.create({
  theme: { colors: { accent1: "#abcdef", accent6: "#123456" }, fonts: { minor: "Inter" } },
  masters: [
    { id: "master/brand", name: "Brand Master", background: "#112233", placeholders: [{ type: "title", idx: 1, name: "Brand Title", style: { bold: true, color: "accent1" } }] },
    { id: "master/data", name: "Data Master", background: { fill: "accent1", mode: "reference", index: 1001 }, theme: { name: "Data Theme", colors: { accent1: "#ff0066", accent2: "#663399" }, fonts: { major: "Georgia" }, textStyles: { title: { fontSize: 44, color: "accent2" } }, colorMap: { accent1: "accent2" } }, placeholders: [{ type: "title", idx: 1, name: "Data Title", style: { bold: true, color: "accent2" } }] },
  ],
  layouts: [
    { id: "layout/brand", name: "Brand Layout", masterId: "master/brand", placeholders: [] },
    { id: "layout/data", name: "Data Layout", masterId: "master/data", placeholders: [] },
  ],
});
multiMasterPresentation.slides.add({ layoutId: "layout/brand" });
multiMasterPresentation.slides.add({ layoutId: "layout/data" });
assert.equal(multiMasterPresentation.master, multiMasterPresentation.masters.items[0]);
assert.equal(multiMasterPresentation.masters.count, 2);
assert.equal(multiMasterPresentation.layouts.getItem("layout/data").effectivePlaceholders()[0].name, "Data Title");
assert.deepEqual(multiMasterPresentation.slides.items[1].effectiveBackground(), { fill: "accent1", mode: "reference", index: 1001 });
assert.equal(multiMasterPresentation.slides.items[1].effectiveTheme().colors.accent1, "#ff0066");
assert.equal(multiMasterPresentation.slides.items[1].effectiveTheme().colors.accent6, "#123456");
assert.equal(multiMasterPresentation.slides.items[1].effectiveTheme().fonts.minor, "Inter");
assert.match(multiMasterPresentation.slides.items[1].toSvg(), /fill="#663399"/);
assert.match(multiMasterPresentation.inspect({ kind: "slideMaster" }).ndjson, /Brand Master/);
assert.match(multiMasterPresentation.inspect({ kind: "slideMaster" }).ndjson, /Data Master/);
assert.equal(multiMasterPresentation.resolve("master/data").name, "Data Master");
const multiMasterPptx = await PresentationFile.exportPptx(multiMasterPresentation);
assert.equal((await PresentationFile.inspectPptx(multiMasterPptx)).ok, true);
const multiMasterZip = await JSZip.loadAsync(new Uint8Array(await multiMasterPptx.arrayBuffer()));
assert.ok(multiMasterZip.file("ppt/slideMasters/slideMaster1.xml"));
assert.ok(multiMasterZip.file("ppt/slideMasters/slideMaster2.xml"));
assert.ok(multiMasterZip.file("ppt/theme/theme2.xml"));
assert.match(await multiMasterZip.file("ppt/presentation.xml").async("text"), /<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId4"\/><p:sldMasterId id="2147483649" r:id="rId5"\/><\/p:sldMasterIdLst>/);
assert.match(await multiMasterZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /Brand Master/);
assert.match(await multiMasterZip.file("ppt/slideMasters/slideMaster2.xml").async("text"), /Data Master/);
assert.match(await multiMasterZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /slideMaster1\.xml/);
assert.match(await multiMasterZip.file("ppt/slideLayouts/_rels/slideLayout2.xml.rels").async("text"), /slideMaster2\.xml/);
assert.match(await multiMasterZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels").async("text"), /relationships\/theme" Target="\.\.\/theme\/theme1\.xml"/);
assert.match(await multiMasterZip.file("ppt/slideMasters/_rels/slideMaster2.xml.rels").async("text"), /relationships\/theme" Target="\.\.\/theme\/theme2\.xml"/);
assert.match(await multiMasterZip.file("ppt/theme/theme2.xml").async("text"), /name="Data Theme"/);
assert.match(await multiMasterZip.file("ppt/theme/theme2.xml").async("text"), /<a:accent1><a:srgbClr val="FF0066"\/>/);
assert.match(await multiMasterZip.file("ppt/slideMasters/slideMaster2.xml").async("text"), /<p:clrMap[^>]*accent1="accent2"/);
assert.match(await multiMasterZip.file("ppt/slideMasters/slideMaster2.xml").async("text"), /<p:titleStyle>[\s\S]*?<a:defRPr sz="4400"/);
assert.match(await multiMasterZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /slideLayout1\.xml/);
assert.match(await multiMasterZip.file("ppt/slides/_rels/slide2.xml.rels").async("text"), /slideLayout2\.xml/);
const multiMasterLoaded = await PresentationFile.importPptx(multiMasterPptx);
assert.equal(multiMasterLoaded.masters.count, 2);
assert.equal(multiMasterLoaded.masters.items[0].name, "Brand Master");
assert.equal(multiMasterLoaded.masters.items[1].name, "Data Master");
assert.equal(multiMasterLoaded.layouts.items[0].masterId, multiMasterLoaded.masters.items[0].id);
assert.equal(multiMasterLoaded.layouts.items[1].masterId, multiMasterLoaded.masters.items[1].id);
assert.deepEqual(multiMasterLoaded.slides.items[1].effectiveBackground(), { fill: "accent1", mode: "reference", index: 1001 });
assert.equal(multiMasterLoaded.theme.colors.accent1, "#abcdef");
assert.equal(multiMasterLoaded.masters.items[1].theme?.name, "Data Theme");
assert.equal(multiMasterLoaded.masters.items[1].effectiveTheme().colors.accent1, "#ff0066");
assert.equal(multiMasterLoaded.masters.items[1].effectiveTheme().colors.accent6, "#123456");
assert.equal(multiMasterLoaded.masters.items[1].effectiveTheme().fonts.major, "Georgia");
assert.equal(multiMasterLoaded.masters.items[1].effectiveTheme().textStyles.title.fontSize, 44);
assert.equal(multiMasterLoaded.masters.items[1].effectiveTheme().colorMap.accent1, "accent2");
assert.match(multiMasterLoaded.slides.items[1].toSvg(), /fill="#663399"/);
const multiMasterRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(multiMasterLoaded)).arrayBuffer()));
assert.ok(multiMasterRoundtripZip.file("ppt/slideMasters/slideMaster2.xml"));
assert.match(await multiMasterRoundtripZip.file("ppt/slideLayouts/_rels/slideLayout2.xml.rels").async("text"), /slideMaster2\.xml/);
assert.match(await multiMasterRoundtripZip.file("ppt/theme/theme2.xml").async("text"), /<a:accent1><a:srgbClr val="FF0066"\/>/);
const masterThemeFallbackZip = await JSZip.loadAsync(new Uint8Array(await multiMasterPptx.arrayBuffer()));
const multiMasterPresentationRels = await masterThemeFallbackZip.file("ppt/_rels/presentation.xml.rels").async("text");
masterThemeFallbackZip.file("ppt/_rels/presentation.xml.rels", multiMasterPresentationRels.replace(/<Relationship\b(?=[^>]*Type="[^"]*\/theme")[^>]*\/>/, ""));
const alternateMultiMasterPresentationXml = (await masterThemeFallbackZip.file("ppt/presentation.xml").async("text")).replaceAll("<p:", "<deck:").replaceAll("</p:", "</deck:").replace("xmlns:p=", "xmlns:deck=");
masterThemeFallbackZip.file("ppt/presentation.xml", alternateMultiMasterPresentationXml);
const masterThemeFallbackPptx = new FileBlob(await masterThemeFallbackZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: multiMasterPptx.type });
const masterThemeFallbackLoaded = await PresentationFile.importPptx(masterThemeFallbackPptx);
assert.equal(masterThemeFallbackLoaded.theme.colors.accent1, "#abcdef");
assert.equal(masterThemeFallbackLoaded.masters.count, 2);
assert.equal(masterThemeFallbackLoaded.masters.items[1].effectiveTheme().colors.accent1, "#ff0066");
assert.throws(() => multiMasterPresentation.masters.items[1].setTheme({ colors: { accent7: "#000000" } }), /Unsupported presentation theme color/);
multiMasterPresentation.masters.items[1].setTheme(null);
assert.equal(multiMasterPresentation.masters.items[1].theme, undefined);
assert.equal(multiMasterPresentation.masters.items[1].effectiveTheme().colors.accent1, "#abcdef");
const missingMasterPresentation = Presentation.create();
missingMasterPresentation.layouts.add({ id: "layout/missing-master", masterId: "master/missing" });
missingMasterPresentation.slides.add({ layoutId: "layout/missing-master" });
assert.ok(missingMasterPresentation.verify().issues.some((issue) => issue.type === "missingMaster"));
await assert.rejects(() => PresentationFile.exportPptx(missingMasterPresentation), /references missing master/);
const missingLayoutPresentation = Presentation.create({ master: { id: "master/valid" } });
missingLayoutPresentation.slides.add({ layoutId: "layout/missing" });
assert.ok(missingLayoutPresentation.verify().issues.some((issue) => issue.type === "missingLayout"));
await assert.rejects(() => PresentationFile.exportPptx(missingLayoutPresentation), /references missing layout/);
assert.throws(() => Presentation.create({ masters: [{ id: "master/duplicate" }, { id: "master/duplicate" }] }), /Duplicate presentation master ID/);
presentation.theme
  .setColors({ accent1: "#0ea5e9", accent2: "#f97316", accent4: "#7c3aed", accent5: "#16a34a", accent6: "#dc2626", tx2: "#334155", bg2: "#f8fafc", hlink: "#2563eb", folHlink: "#9333ea" })
  .setFonts({ major: "Aptos Display", minor: "Aptos", majorEastAsia: "PingFang SC", majorComplexScript: "Arial", minorEastAsia: "Noto Sans CJK SC", minorComplexScript: "Arial" })
  .setTextStyles({ title: { fontSize: 38, color: "accent1", alignment: "center" }, body: { fontSize: 22, color: "tx2", fontFamily: "+mn-lt" }, other: { fontSize: 16, italic: true, color: "#475569" } })
  .setColorMap({ accent1: "accent2", accent2: "accent1" });
presentation.master.update({
  id: "master/clean-room",
  name: "Clean Room Master",
  background: { fill: "accent1", mode: "reference", index: 1001 },
  placeholders: [
    { id: "master/title", type: "title", idx: 1, name: "Master Title", position: { left: 48, top: 24, width: 900, height: 72 }, required: true, style: { fontSize: 32, bold: true, color: "accent2", fontFamily: "Aptos Display" } },
    { id: "master/body", type: "body", idx: 2, name: "Master Body", position: { left: 72, top: 140, width: 760, height: 360 }, style: { fontSize: 20, color: "tx2" } },
  ],
});
assert.throws(() => presentation.theme.setColors({ accent7: "#000000" }), /Unsupported presentation theme color accent7/);
assert.throws(() => presentation.theme.setColors({ accent3: "transparent" }), /six-digit RGB color/);
assert.throws(() => presentation.theme.setFonts({ major: "" }), /must not be empty/);
assert.throws(() => presentation.theme.setTextStyles({ title: { fontSize: 0 } }), /fontSize must be greater than 0/);
assert.throws(() => presentation.theme.setTextStyles({ body: { alignment: "middle" } }), /alignment must be/);
assert.throws(() => presentation.theme.setColorMap({ accent1: "unknown" }), /unsupported target/);
const titleLayout = presentation.layouts.add({
  id: "layout/title-content",
  name: "Title and Content",
  type: "titleAndContent",
  background: "#fff7ed",
  placeholders: [
    { id: "ph/title", type: "title", idx: 1, name: "Title Placeholder", position: { left: 72, top: 36, width: 720, height: 52 }, required: true, style: { fontSize: 28 } },
    { id: "ph/body", type: "body", idx: 2, name: "Body Placeholder", position: { left: 72, top: 520, width: 560, height: 80 } },
  ],
});
const slide = presentation.slides.add();
const [titlePlaceholder] = slide.applyLayout(titleLayout);
titlePlaceholder.text = "Quarterly plan template";
assert.deepEqual(slide.effectiveBackground(), { fill: "#fff7ed", mode: "solid" });
assert.equal(titlePlaceholder.text.style.fontSize, 28);
assert.equal(titlePlaceholder.text.style.bold, true);
assert.equal(titlePlaceholder.text.style.color, "accent2");
assert.equal(titleLayout.effectivePlaceholders()[1].style.fontSize, 20);
assert.throws(() => presentation.layouts.add({ name: "Invalid Placeholder", placeholders: [{ type: "title", idx: 0 }] }), /unsigned positive 32-bit/);
assert.throws(() => presentation.master.setBackground("transparent"), /scheme color, six-digit RGB/);
const themeLayoutSnapshot = presentation.inspect({ kind: "theme,layout,textbox", maxChars: 8000 }).ndjson;
assert.match(themeLayoutSnapshot, /Open Office Clean Room/);
assert.match(themeLayoutSnapshot, /Title and Content/);
assert.match(themeLayoutSnapshot, /Quarterly plan template/);
assert.equal(presentation.resolve("layout/title-content").name, "Title and Content");
assert.match(presentation.help("presentation.theme").ndjson, /theme colors/);
assert.match(presentation.help("presentation.masters.add").ndjson, /Slide Master/);
assert.match(presentation.help("presentation.masters.getItem").ndjson, /master ID/);
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
pieSlide.background = { fill: "#ecfeff", mode: "solid" };
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
thread.addReply("Updated wording to plan.", { author: "Editor" }).resolve();
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
assert.match(themeXml, /<a:dk2><a:srgbClr val="334155"\/><\/a:dk2>/);
assert.match(themeXml, /<a:lt2><a:srgbClr val="F8FAFC"\/><\/a:lt2>/);
assert.match(themeXml, /<a:accent4><a:srgbClr val="7C3AED"\/><\/a:accent4>/);
assert.match(themeXml, /<a:accent5><a:srgbClr val="16A34A"\/><\/a:accent5>/);
assert.match(themeXml, /<a:accent6><a:srgbClr val="DC2626"\/><\/a:accent6>/);
assert.match(themeXml, /<a:majorFont><a:latin typeface="Aptos Display"\/><a:ea typeface="PingFang SC"\/><a:cs typeface="Arial"\/><\/a:majorFont>/);
assert.equal((themeXml.match(/<a:fillStyleLst>[\s\S]*?<\/a:fillStyleLst>/g) || []).length, 1);
assert.equal((/<a:fillStyleLst>([\s\S]*?)<\/a:fillStyleLst>/.exec(themeXml)?.[1].match(/<a:solidFill>/g) || []).length, 3);
assert.equal((/<a:lnStyleLst>([\s\S]*?)<\/a:lnStyleLst>/.exec(themeXml)?.[1].match(/<a:ln\b/g) || []).length, 3);
assert.equal((/<a:effectStyleLst>([\s\S]*?)<\/a:effectStyleLst>/.exec(themeXml)?.[1].match(/<a:effectStyle>/g) || []).length, 3);
assert.equal((/<a:bgFillStyleLst>([\s\S]*?)<\/a:bgFillStyleLst>/.exec(themeXml)?.[1].match(/<a:solidFill>/g) || []).length, 3);
const masterXml = await zip.file("ppt/slideMasters/slideMaster1.xml").async("text");
assert.match(masterXml, /sldLayoutId/);
assert.match(masterXml, /<p:cSld name="Clean Room Master"><p:bg><p:bgRef idx="1001"><a:schemeClr val="accent1"\/><\/p:bgRef><\/p:bg>/);
assert.match(masterXml, /<p:ph type="title" idx="1"\/>/);
assert.match(masterXml, /<a:rPr lang="en-US" sz="2400" b="1">/);
assert.match(masterXml, /<a:schemeClr val="accent2"\/>/);
assert.match(masterXml, /<p:clrMap[^>]*accent1="accent2"[^>]*accent2="accent1"/);
assert.match(masterXml, /<p:titleStyle>[\s\S]*?<a:defRPr sz="3800" b="1" i="0">[\s\S]*?<a:schemeClr val="accent1"/);
assert.match(masterXml, /<p:bodyStyle>[\s\S]*?<a:defRPr sz="2200" b="0" i="0">[\s\S]*?<a:schemeClr val="tx2"/);
assert.match(masterXml, /<p:otherStyle>[\s\S]*?<a:defRPr sz="1600" b="0" i="1">[\s\S]*?<a:srgbClr val="475569"/);
assert.equal((masterXml.match(/<a:lvl[1-9]pPr\b/g) || []).length, 27);
const alternateThemePrefixXml = themeXml.replaceAll("<a:", "<d:").replaceAll("</a:", "</d:").replace("xmlns:a=", "xmlns:d=").replace('<d:dk1><d:srgbClr val="0F172A"/></d:dk1>', '<d:dk1><d:sysClr val="windowText" lastClr="0F172A"/></d:dk1>');
const alternateMasterPrefixXml = masterXml.replaceAll("<p:", "<m:").replaceAll("</p:", "</m:").replaceAll("<a:", "<d:").replaceAll("</a:", "</d:").replace("xmlns:p=", "xmlns:m=").replace("xmlns:a=", "xmlns:d=");
const alternatePrefixPptx = await PresentationFile.patchPptx(pptx, [
  { path: "ppt/theme/theme1.xml", xml: alternateThemePrefixXml },
  { path: "ppt/slideMasters/slideMaster1.xml", xml: alternateMasterPrefixXml },
]);
const alternatePrefixLoaded = await PresentationFile.importPptx(alternatePrefixPptx);
assert.equal(alternatePrefixLoaded.master.name, "Clean Room Master");
assert.deepEqual(alternatePrefixLoaded.master.background, { fill: "accent1", mode: "reference", index: 1001 });
assert.equal(alternatePrefixLoaded.master.placeholders[0].style.bold, true);
assert.equal(alternatePrefixLoaded.theme.colors.accent6, "#dc2626");
assert.equal(alternatePrefixLoaded.theme.colors.tx1, "#0f172a");
assert.equal(alternatePrefixLoaded.theme.fonts.majorEastAsia, "PingFang SC");
assert.equal(alternatePrefixLoaded.theme.textStyles.title.fontSize, 38);
assert.equal(alternatePrefixLoaded.theme.colorMap.accent1, "accent2");
const layoutXml = await zip.file("ppt/slideLayouts/slideLayout1.xml").async("text");
assert.match(layoutXml, /Title and Content/);
assert.match(layoutXml, /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFF7ED"\/><\/a:solidFill>/);
assert.match(layoutXml, /<p:ph type="title"/);
assert.match(await zip.file("ppt/presentation.xml").async("text"), /<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId\d+"\/><\/p:sldMasterIdLst>/);
assert.match(await zip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /relationships\/slideMaster" Target="\.\.\/slideMasters\/slideMaster1\.xml"/);
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
const pieSlideXml = await zip.file("ppt/slides/slide2.xml").async("text");
assert.match(pieSlideXml, /<p:bg><p:bgPr><a:solidFill><a:srgbClr val="ECFEFF"\/><\/a:solidFill>/);
const notesXml = await zip.file("ppt/notesSlides/notesSlide1.xml").async("text");
assert.match(notesXml, /Speaker note/);
const commentsXml = await zip.file("ppt/comments/comment1.xml").async("text");
assert.match(commentsXml, /Tighten this headline/);
assert.match(commentsXml, /Updated wording/);
assert.match(commentsXml, /ooa:resolved="1"/);
assert.match(commentsXml, /authorId="0"/);
assert.match(commentsXml, /authorId="1"/);
assert.equal((commentsXml.match(/idx="1"/g) || []).length, 2);
const commentAuthorsXml = await zip.file("ppt/commentAuthors.xml").async("text");
assert.match(commentAuthorsXml, /<p:cmAuthor id="0" name="Reviewer" initials="RE" lastIdx="1" clrIdx="0"\/>/);
assert.match(commentAuthorsXml, /<p:cmAuthor id="1" name="Editor" initials="ED" lastIdx="1" clrIdx="1"\/>/);
const presentationRelsXml = await zip.file("ppt/_rels/presentation.xml.rels").async("text");
assert.match(presentationRelsXml, /relationships\/commentAuthors" Target="commentAuthors\.xml"/);
const slideRelsXml = await zip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(slideRelsXml, /Target="\.\.\/media\/image1\.png"/);
assert.match(slideRelsXml, /Target="\.\.\/charts\/chart1\.xml"/);
assert.match(slideRelsXml, /relationships\/slideLayout/);
assert.match(slideRelsXml, /relationships\/notesSlide/);
assert.match(slideRelsXml, /relationships\/comments/);
assert.match(await zip.file("[Content_Types].xml").async("text"), /PartName="\/ppt\/commentAuthors\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.presentationml\.commentAuthors\+xml"/);
const relocatedNotesComments = await PresentationFile.patchPptx(pptx, [
  { path: "ppt/notesSlides/notesSlide1.xml", remove: true },
  { path: "ppt/comments/comment1.xml", remove: true },
  { path: "ppt/commentAuthors.xml", remove: true },
  { path: "ppt/custom/notes/review.xml", xml: notesXml, recipe: { kind: "notesSlide", source: "ppt/slides/slide1.xml" } },
  { path: "ppt/custom/comments/review.xml", xml: commentsXml, recipe: { kind: "comments", source: "ppt/slides/slide1.xml" } },
  { path: "ppt/custom/comments/authors.xml", xml: commentAuthorsXml, recipe: { kind: "commentAuthors", source: "ppt/presentation.xml" } },
]);
assert.equal((await PresentationFile.inspectPptx(relocatedNotesComments)).ok, true);
const relocatedZip = await JSZip.loadAsync(new Uint8Array(await relocatedNotesComments.arrayBuffer()));
const relocatedSlideRels = await relocatedZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(relocatedSlideRels, /Target="\.\.\/custom\/notes\/review\.xml"/);
assert.match(relocatedSlideRels, /Target="\.\.\/custom\/comments\/review\.xml"/);
assert.match(await relocatedZip.file("ppt/_rels/presentation.xml.rels").async("text"), /Target="custom\/comments\/authors\.xml"/);
const relocatedLoaded = await PresentationFile.importPptx(relocatedNotesComments);
assert.match(relocatedLoaded.slides.items[0].speakerNotes.text, /Speaker note/);
assert.deepEqual(relocatedLoaded.slides.items[0].comments.items[0].comments.map((comment) => comment.author), ["Reviewer", "Editor"]);
const invalidCommentRootXml = commentsXml.replace(/<p:cmLst\b/, "<p:notes").replace(/<\/p:cmLst>/, "</p:notes>");
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: invalidCommentRootXml }]), /invalid OOXML package.*pptxSemanticPartRootInvalid/);
const invalidCommentRootPptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: invalidCommentRootXml }], { validateResult: false });
const invalidCommentRootInspect = await PresentationFile.inspectPptx(invalidCommentRootPptx);
assert.equal(invalidCommentRootInspect.ok, false);
assert.ok(invalidCommentRootInspect.issues.some((issue) => issue.type === "pptxSemanticPartRootInvalid" && issue.path === "ppt/comments/comment1.xml"));
const strictPresentationNamespace = "http://purl.oclc.org/ooxml/presentationml/main";
const strictReviewPptx = await PresentationFile.patchPptx(pptx, [
  { path: "ppt/notesSlides/notesSlide1.xml", xml: notesXml.replaceAll("http://schemas.openxmlformats.org/presentationml/2006/main", strictPresentationNamespace) },
  { path: "ppt/comments/comment1.xml", xml: commentsXml.replaceAll("http://schemas.openxmlformats.org/presentationml/2006/main", strictPresentationNamespace) },
  { path: "ppt/commentAuthors.xml", xml: commentAuthorsXml.replaceAll("http://schemas.openxmlformats.org/presentationml/2006/main", strictPresentationNamespace) },
]);
assert.equal((await PresentationFile.inspectPptx(strictReviewPptx)).ok, true);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: commentsXml.replace('authorId="0"', 'authorId="99"') }]), /invalid OOXML package.*pptxCommentAuthorReferenceNotFound/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: commentsXml.replace('authorId="1"', 'authorId="0"') }]), /invalid OOXML package.*pptxCommentIndexDuplicate/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: commentsXml.replace(/ dt="[^"]+"/, ' dt="not-a-date"') }]), /invalid OOXML package.*pptxCommentDateInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/comments/comment1.xml", xml: commentsXml.replace(/<p:pos\b[^>]*\/>/, "") }]), /invalid OOXML package.*pptxCommentPositionInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/commentAuthors.xml", xml: commentAuthorsXml.replace('lastIdx="1"', 'lastIdx="0"') }]), /invalid OOXML package.*pptxCommentLastIndexTooSmall/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/commentAuthors.xml", xml: commentAuthorsXml.replace('id="1"', 'id="0"') }]), /invalid OOXML package.*pptxCommentAuthorDuplicateId/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/notesSlides/notesSlide1.xml", xml: notesXml.replace(/<p:cSld>[\s\S]*?<\/p:cSld>/, "") }]), /invalid OOXML package.*pptxNotesCommonSlideDataMissing/);
const orphanedCommentsRels = slideRelsXml.replace(/<Relationship\b[^>]*relationships\/comments[^>]*\/>/, "");
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/_rels/slide1.xml.rels", xml: orphanedCommentsRels }]), /invalid OOXML package.*pptxSemanticPartOrphaned/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/commentAuthors.xml", remove: true, recipe: { kind: "commentAuthors", source: "ppt/presentation.xml" } }]), /invalid OOXML package.*pptxCommentAuthorsMissing/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/custom/notes/wrong-source.xml", xml: notesXml, recipe: { kind: "notesSlide", source: "ppt/presentation.xml" } }]), /invalid OOXML package.*pptxSemanticRelationshipSourceInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(pptx, [{ path: "ppt/custom/comments/authors-duplicate.xml", xml: commentAuthorsXml, recipe: { kind: "commentAuthors", source: "ppt/presentation.xml" } }]), /invalid OOXML package.*pptxCommentAuthorsMultiplicity/);
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
assert.equal(loaded.master.id, "pptx-master-2147483648");
assert.equal(loaded.master.name, "Clean Room Master");
assert.deepEqual(loaded.master.background, { fill: "accent1", mode: "reference", index: 1001 });
assert.equal(loaded.master.placeholders.length, 2);
assert.equal(loaded.master.placeholders[0].style.bold, true);
assert.equal(loaded.master.placeholders[0].style.color, "accent2");
assert.equal(loaded.layouts.items[0].id, "pptx-layout-2147483649");
assert.equal(loaded.layouts.items[0].masterId, "pptx-master-2147483648");
assert.deepEqual(loaded.layouts.items[0].background, { fill: "#fff7ed", mode: "solid" });
assert.equal(loaded.layouts.items[0].effectivePlaceholders()[0].style.bold, true);
assert.equal(loaded.layouts.items[0].effectivePlaceholders()[0].style.fontSize, 28);
assert.deepEqual(loaded.slides.items[0].effectiveBackground(), { fill: "#fff7ed", mode: "solid" });
assert.deepEqual(loaded.slides.items[1].effectiveBackground(), { fill: "#ecfeff", mode: "solid" });
assert.equal(loaded.theme.colors.accent6, "#dc2626");
assert.equal(loaded.theme.colors.tx2, "#334155");
assert.equal(loaded.theme.fonts.majorEastAsia, "PingFang SC");
assert.equal(loaded.theme.fonts.minorComplexScript, "Arial");
assert.deepEqual(loaded.theme.textStyles.title, { fontSize: 38, bold: true, italic: false, color: "accent1", fontFamily: "+mj-lt", alignment: "center" });
assert.deepEqual(loaded.theme.textStyles.body, { fontSize: 22, bold: false, italic: false, color: "tx2", fontFamily: "+mn-lt", alignment: "left" });
assert.equal(loaded.theme.textStyles.other.italic, true);
assert.equal(loaded.theme.textStyles.other.color, "#475569");
assert.equal(loaded.theme.colorMap.accent1, "accent2");
assert.equal(loaded.theme.colorMap.accent2, "accent1");
const loadedThemeRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(loaded)).arrayBuffer()));
assert.match(await loadedThemeRoundtripZip.file("ppt/theme/theme1.xml").async("text"), /<a:accent6><a:srgbClr val="DC2626"\/><\/a:accent6>/);
assert.match(await loadedThemeRoundtripZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /<p:clrMap[^>]*accent1="accent2"[^>]*accent2="accent1"/);
assert.match(await loadedThemeRoundtripZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /<p:titleStyle>[\s\S]*?<a:defRPr sz="3800"/);
assert.match(await loadedThemeRoundtripZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /<p:bg><p:bgRef idx="1001">/);
assert.match(await loadedThemeRoundtripZip.file("ppt/slideLayouts/slideLayout1.xml").async("text"), /<a:srgbClr val="FFF7ED"/);
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
assert.match(loadedAll, /"textStyles"/);
assert.match(loadedAll, /"colorMap"/);
const loadedComment = loaded.slides.items[0].comments.items[0];
assert.ok(loaded.slides.items[0].resolve(loadedComment.targetId));
assert.equal(loadedComment.author, "Reviewer");
assert.deepEqual(loadedComment.comments.map((comment) => comment.author), ["Reviewer", "Editor"]);
const loadedSummarySurface = loaded.slides.items[0].shapes.items.find((item) => item.name === "summary-surface");
assert.equal(loadedSummarySurface.text.style.fontSize, 32);
assert.equal(loadedSummarySurface.text.style.bold, true);
assert.equal(loadedSummarySurface.fill.toLowerCase(), "#ffffff");

const modernPresentation = Presentation.create({ commentFormat: "modern" });
const modernSlide = modernPresentation.slides.add();
const modernTarget = modernSlide.shapes.add({ name: "modern-comment-target", text: "Office 2021 review target" });
modernSlide.comments.addThread(modernTarget, "Review the modern comment contract.", { author: "Modern Reviewer", created: "2026-07-13T01:02:03Z" })
  .addReply("The native reply is preserved.", { author: "Modern Editor", created: "2026-07-13T02:03:04Z", userId: "editor@example.test", providerId: "None" })
  .resolve();
const modernPptx = await PresentationFile.exportPptx(modernPresentation);
assert.equal((await PresentationFile.inspectPptx(modernPptx)).ok, true);
const modernZip = await JSZip.loadAsync(new Uint8Array(await modernPptx.arrayBuffer()));
const modernCommentsXml = await modernZip.file("ppt/comments/comment1.xml").async("text");
const modernAuthorsXml = await modernZip.file("ppt/authors.xml").async("text");
assert.match(modernCommentsXml, /p188:cmLst/);
assert.match(modernCommentsXml, /oac:deMkLst/);
assert.match(modernCommentsXml, /pc:sldMk sldId="256"/);
assert.match(modernCommentsXml, /oac:spMk id="2" creationId="\{[0-9A-F-]+\}"/);
assert.match(modernCommentsXml, /p188:replyLst/);
assert.match(modernCommentsXml, /status="resolved"/);
assert.equal((modernCommentsXml.match(/id="\{[0-9A-F-]+\}"/g) || []).length, 2);
assert.match(modernAuthorsXml, /p188:authorLst/);
assert.match(modernAuthorsXml, /name="Modern Reviewer"/);
assert.match(modernAuthorsXml, /name="Modern Editor"/);
assert.match(await modernZip.file("ppt/_rels/presentation.xml.rels").async("text"), /office\/2018\/10\/relationships\/authors/);
assert.match(await modernZip.file("ppt/presentation.xml").async("text"), /<p:notesSz cx="6858000" cy="9144000"\/>/);
assert.match(await modernZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /office\/2018\/10\/relationships\/comments/);
assert.match(await modernZip.file("[Content_Types].xml").async("text"), /application\/vnd\.ms-powerpoint\.comments\+xml/);
const modernSlideXml = await modernZip.file("ppt/slides/slide1.xml").async("text");
assert.match(modernSlideXml, /a16:creationId[^>]*id="\{[0-9A-F-]+\}"/);
const modernLoaded = await PresentationFile.importPptx(modernPptx);
assert.equal(modernLoaded.commentFormat, "modern");
const modernLoadedThread = modernLoaded.slides.items[0].comments.items[0];
assert.equal(modernLoadedThread.nativeFormat, "modern");
assert.equal(modernLoadedThread.resolved, true);
assert.equal(modernLoadedThread.targetId, modernLoaded.slides.items[0].shapes.items[0].id);
assert.equal(modernLoadedThread.nativeAnchor.creationId, modernLoaded.slides.items[0].shapes.items[0].creationId);
assert.deepEqual(modernLoadedThread.comments.map((comment) => comment.author), ["Modern Reviewer", "Modern Editor"]);
assert.ok(modernLoadedThread.comments.every((comment) => /^\{[0-9A-F-]+\}$/.test(comment.nativeId)));
const modernSecondExport = await PresentationFile.exportPptx(modernLoaded);
const modernSecondZip = await JSZip.loadAsync(new Uint8Array(await modernSecondExport.arrayBuffer()));
const modernSecondXml = await modernSecondZip.file("ppt/comments/comment1.xml").async("text");
assert.match(modernSecondXml, new RegExp(modernLoadedThread.comments[0].nativeId.replace(/[{}]/g, "\\$&")));
assert.match(modernSecondXml, new RegExp(modernLoadedThread.comments[1].nativeId.replace(/[{}]/g, "\\$&")));

const relocatedModern = await PresentationFile.patchPptx(modernPptx, [
  { path: "ppt/comments/comment1.xml", remove: true },
  { path: "ppt/authors.xml", remove: true },
  { path: "ppt/custom/reviews/modern.xml", xml: modernCommentsXml.replaceAll("<p188:", "<review:").replaceAll("</p188:", "</review:").replace("xmlns:p188=", "xmlns:review="), recipe: { kind: "modernComments", source: "ppt/slides/slide1.xml" } },
  { path: "ppt/custom/reviews/people.xml", xml: modernAuthorsXml.replaceAll("<p188:", "<people:").replaceAll("</p188:", "</people:").replace("xmlns:p188=", "xmlns:people="), recipe: { kind: "modernAuthors", source: "ppt/presentation.xml" } },
]);
assert.equal((await PresentationFile.inspectPptx(relocatedModern)).ok, true);
const relocatedModernLoaded = await PresentationFile.importPptx(relocatedModern);
assert.equal(relocatedModernLoaded.commentFormat, "modern");
assert.deepEqual(relocatedModernLoaded.slides.items[0].comments.items[0].comments.map((comment) => comment.text), ["Review the modern comment contract.", "The native reply is preserved."]);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(/id="\{[0-9A-F-]+\}"/, 'id="not-a-guid"') }]), /invalid OOXML package.*pptxModernCommentIdInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/authors.xml", xml: modernAuthorsXml.replace(/<p188:author\b[^>]*\/>/, "") }]), /invalid OOXML package.*pptxModernCommentAuthorNotFound/);
const modernAnchorXml = /<oac:deMkLst>[\s\S]*?<\/oac:deMkLst>/.exec(modernCommentsXml)[0];
const modernAnchorCreationId = /creationId="(\{[0-9A-F-]+\})"/.exec(modernAnchorXml)[1];
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(modernAnchorXml, "") }]), /invalid OOXML package.*pptxModernCommentAnchorMissing/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(modernAnchorXml, modernAnchorXml + modernAnchorXml) }]), /invalid OOXML package.*pptxModernCommentAnchorMissing/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace('sldId="256"', 'sldId="257"') }]), /invalid OOXML package.*pptxModernCommentSlideMonikerMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace("oac:spMk", "oac:picMk") }]), /invalid OOXML package.*pptxModernCommentDrawingTargetNotFound/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace('oac:spMk id="2"', 'oac:spMk id="3"') }]), /invalid OOXML package.*pptxModernCommentDrawingIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(modernAnchorCreationId, "{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}") }]), /invalid OOXML package.*pptxModernCommentDrawingTargetNotFound/);
const idOnlyModernPptx = await PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(/ creationId="\{[0-9A-F-]+\}"/, "") }]);
assert.equal((await PresentationFile.inspectPptx(idOnlyModernPptx)).ok, true);
assert.ok((await PresentationFile.importPptx(idOnlyModernPptx)).slides.items[0].comments.items[0].targetId);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/comment1.xml", xml: modernCommentsXml.replace(/<p188:pos x="[^"]+"/, '<p188:pos x="NaN"') }]), /invalid OOXML package.*pptxModernCommentPositionInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(modernPptx, [{ path: "ppt/comments/_rels/comment1.xml.rels", xml: '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide1.xml"/></Relationships>' }]), /invalid OOXML package.*pptxModernCommentPartRelationshipsForbidden/);

const textAnchorPresentation = Presentation.create({ commentFormat: "modern" });
const textAnchorSlide = textAnchorPresentation.slides.add();
const textAnchorShape = textAnchorSlide.shapes.add({ name: "text-anchor-shape", text: "Review this exact phrase" });
const textAnchorTarget = textAnchorSlide.resolve(`${textAnchorShape.id}/text`);
textAnchorSlide.comments.addThread(textAnchorTarget, "Review the text range.", { author: "Text Reviewer", created: "2026-07-13T04:05:06Z" });
const textAnchorPptx = await PresentationFile.exportPptx(textAnchorPresentation);
assert.equal((await PresentationFile.inspectPptx(textAnchorPptx)).ok, true);
const textAnchorZip = await JSZip.loadAsync(new Uint8Array(await textAnchorPptx.arrayBuffer()));
const textAnchorCommentsXml = await textAnchorZip.file("ppt/comments/comment1.xml").async("text");
assert.match(textAnchorCommentsXml, /<oac:txMkLst><pc:sldMkLst><pc:docMk\/><pc:sldMk sldId="256"\/><\/pc:sldMkLst><oac:spMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:txMk cp="0" len="24"\/><\/oac:txMkLst>/);
assert.doesNotMatch(textAnchorCommentsXml, /oac:deMkLst/);
const textAnchorLoaded = await PresentationFile.importPptx(textAnchorPptx);
const textAnchorLoadedSlide = textAnchorLoaded.slides.items[0];
const textAnchorLoadedThread = textAnchorLoadedSlide.comments.items[0];
assert.equal(textAnchorLoadedThread.targetId, `${textAnchorLoadedSlide.shapes.items[0].id}/text`);
assert.equal(textAnchorLoadedSlide.resolve(textAnchorLoadedThread.targetId)?.kind, "textRange");
assert.deepEqual({ type: textAnchorLoadedThread.nativeAnchor.type, cp: textAnchorLoadedThread.nativeAnchor.cp, length: textAnchorLoadedThread.nativeAnchor.length }, { type: "textRange", cp: 0, length: 24 });
assert.equal(textAnchorLoadedThread.nativeAnchor.creationId, textAnchorLoadedSlide.shapes.items[0].creationId);
const subsetTextAnchorPptx = await PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace('cp="0" len="24"', 'cp="7" len="4"') }]);
const subsetTextAnchorLoaded = await PresentationFile.importPptx(subsetTextAnchorPptx);
const subsetTextThread = subsetTextAnchorLoaded.slides.items[0].comments.items[0];
assert.deepEqual({ cp: subsetTextThread.nativeAnchor.cp, length: subsetTextThread.nativeAnchor.length }, { cp: 7, length: 4 });
assert.equal(subsetTextAnchorLoaded.slides.items[0].resolve(subsetTextThread.targetId)?.text, "Review this exact phrase");
const subsetTextSecondExport = await PresentationFile.exportPptx(subsetTextAnchorLoaded);
const subsetTextSecondZip = await JSZip.loadAsync(new Uint8Array(await subsetTextSecondExport.arrayBuffer()));
assert.match(await subsetTextSecondZip.file("ppt/comments/comment1.xml").async("text"), /<oac:txMk cp="7" len="4"\/>/);
const contextualTextAnchorPptx = await PresentationFile.patchPptx(textAnchorPptx, [{
  path: "ppt/comments/comment1.xml",
  xml: textAnchorCommentsXml.replace('<oac:txMk cp="0" len="24"/>', '<oac:txMk cp="0" len="24"><oac:context len="24" hash="123456789"/></oac:txMk>'),
}]);
const contextualTextLoaded = await PresentationFile.importPptx(contextualTextAnchorPptx);
assert.deepEqual(
  { contextLength: contextualTextLoaded.slides.items[0].comments.items[0].nativeAnchor.contextLength, contextHash: contextualTextLoaded.slides.items[0].comments.items[0].nativeAnchor.contextHash },
  { contextLength: 24, contextHash: 123456789 },
);
const contextualTextSecondExport = await PresentationFile.exportPptx(contextualTextLoaded);
const contextualTextSecondZip = await JSZip.loadAsync(new Uint8Array(await contextualTextSecondExport.arrayBuffer()));
assert.match(await contextualTextSecondZip.file("ppt/comments/comment1.xml").async("text"), /<oac:txMk cp="0" len="24"><oac:context len="24" hash="123456789"\/><\/oac:txMk>/);
const alternatePrefixTextAnchorPptx = await PresentationFile.patchPptx(textAnchorPptx, [{
  path: "ppt/comments/comment1.xml",
  xml: textAnchorCommentsXml
    .replaceAll("<oac:", "<anchor:").replaceAll("</oac:", "</anchor:").replace("xmlns:oac=", "xmlns:anchor=")
    .replaceAll("<pc:", "<location:").replaceAll("</pc:", "</location:").replace("xmlns:pc=", "xmlns:location="),
}]);
assert.equal((await PresentationFile.inspectPptx(alternatePrefixTextAnchorPptx)).ok, true);
assert.equal((await PresentationFile.importPptx(alternatePrefixTextAnchorPptx)).slides.items[0].comments.items[0].nativeAnchor.type, "textRange");
const textAnchorXml = /<oac:txMkLst>[\s\S]*?<\/oac:txMkLst>/.exec(textAnchorCommentsXml)[0];
const textAnchorCreationId = /creationId="(\{[0-9A-F-]+\})"/.exec(textAnchorXml)[1];
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace(/<oac:txMk\b[^>]*\/>/, "") }]), /invalid OOXML package.*pptxModernCommentTextRangeMonikerInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace('cp="0" len="24"', 'cp="22" len="3"') }]), /invalid OOXML package.*pptxModernCommentTextRangeOutOfBounds/);
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace('<oac:txMk cp="0" len="24"/>', '<oac:txMk cp="0" len="24"><oac:context len="24"/></oac:txMk>') }]), /invalid OOXML package.*pptxModernCommentTextRangeContextInvalid/);
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace('oac:spMk id="2"', 'oac:spMk id="3"') }]), /invalid OOXML package.*pptxModernCommentTextIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace(textAnchorCreationId, "{AAAAAAAA-BBBB-4CCC-8DDD-EEEEEEEEEEEE}") }]), /invalid OOXML package.*pptxModernCommentTextTargetNotFound/);
await assert.rejects(() => PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace(textAnchorXml, textAnchorXml + textAnchorXml) }]), /invalid OOXML package.*pptxModernCommentAnchorMissing/);
const idOnlyTextAnchorPptx = await PresentationFile.patchPptx(textAnchorPptx, [{ path: "ppt/comments/comment1.xml", xml: textAnchorCommentsXml.replace(/ creationId="\{[0-9A-F-]+\}"/, "") }]);
assert.equal((await PresentationFile.inspectPptx(idOnlyTextAnchorPptx)).ok, true);
assert.equal((await PresentationFile.importPptx(idOnlyTextAnchorPptx)).slides.items[0].comments.items[0].nativeAnchor.type, "textRange");

const typedAnchorPresentation = Presentation.create({ commentFormat: "modern" });
const typedAnchorSlide = typedAnchorPresentation.slides.add();
const typedShape = typedAnchorSlide.shapes.add({ name: "typed-shape", text: "Shape" });
const typedTable = typedAnchorSlide.tables.add({ name: "typed-table", values: [["A"]] });
const typedChart = typedAnchorSlide.charts.add("bar", { name: "typed-chart", categories: ["A"], series: [{ name: "Value", values: [1] }] });
const typedImage = typedAnchorSlide.images.add({ name: "typed-image", dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=" });
const typedConnector = typedAnchorSlide.connectors.add({ name: "typed-connector", start: { x: 10, y: 10 }, end: { x: 100, y: 100 } });
for (const target of [typedShape, typedTable, typedChart, typedImage, typedConnector]) typedAnchorSlide.comments.addThread(target, `Review ${target.name}`, { author: "Anchor Reviewer", created: "2026-07-13T03:04:05Z" });
const typedAnchorPptx = await PresentationFile.exportPptx(typedAnchorPresentation);
assert.equal((await PresentationFile.inspectPptx(typedAnchorPptx)).ok, true);
const typedAnchorZip = await JSZip.loadAsync(new Uint8Array(await typedAnchorPptx.arrayBuffer()));
const typedAnchorCommentsXml = await typedAnchorZip.file("ppt/comments/comment1.xml").async("text");
const typedAnchorSlideXml = await typedAnchorZip.file("ppt/slides/slide1.xml").async("text");
assert.equal((typedAnchorCommentsXml.match(/<oac:spMk\b/g) || []).length, 1);
assert.equal((typedAnchorCommentsXml.match(/<oac:graphicFrameMk\b/g) || []).length, 2);
assert.equal((typedAnchorCommentsXml.match(/<oac:picMk\b/g) || []).length, 1);
assert.equal((typedAnchorCommentsXml.match(/<oac:cxnSpMk\b/g) || []).length, 1);
const typedCreationIds = [...typedAnchorSlideXml.matchAll(/a16:creationId[^>]*id="(\{[0-9A-F-]+\})"/g)].map((match) => match[1]);
assert.equal(new Set(typedCreationIds).size, 5);
await assert.rejects(() => PresentationFile.patchPptx(typedAnchorPptx, [{ path: "ppt/slides/slide1.xml", xml: typedAnchorSlideXml.replace(typedCreationIds[1], typedCreationIds[0]) }]), /invalid OOXML package.*pptxModernDrawingCreationIdDuplicate/);
const typedAnchorLoaded = await PresentationFile.importPptx(typedAnchorPptx);
assert.deepEqual(typedAnchorLoaded.slides.items[0].comments.items.map((thread) => typedAnchorLoaded.slides.items[0].resolve(thread.targetId)?.name), ["typed-shape", "typed-table", "typed-chart", "typed-image", "typed-connector"]);

const groupedPresentation = Presentation.create({ commentFormat: "modern", layouts: [{ id: "group-blank", name: "Group Blank", type: "blank", masterId: "master/default" }] });
const groupedSlide = groupedPresentation.slides.add({ layoutId: "group-blank" });
const reviewGroup = groupedSlide.groups.add({ name: "review-group", position: { left: 100, top: 80, width: 400, height: 200 }, childFrame: { left: 0, top: 0, width: 200, height: 100 } });
const groupedChild = reviewGroup.shapes.add({ name: "grouped-child", text: "Grouped child text", position: { left: 20, top: 10, width: 100, height: 30 } });
reviewGroup.connectors.add({ name: "grouped-connector", start: { x: 20, y: 55 }, end: { x: 170, y: 55 } });
const nestedGroup = reviewGroup.groups.add({ name: "nested-group", position: { left: 125, top: 10, width: 60, height: 35 }, childFrame: { left: 0, top: 0, width: 60, height: 35 } });
nestedGroup.shapes.add({ name: "nested-child", text: "Nested", position: { left: 5, top: 5, width: 45, height: 20 } });
groupedSlide.comments.addThread(reviewGroup, "Review the grouped surface.", { author: "Group Reviewer", created: "2026-07-13T05:06:07Z" });
groupedSlide.comments.addThread(groupedChild, "Review the grouped child.", { author: "Group Reviewer", created: "2026-07-13T05:07:08Z" });
groupedSlide.comments.addThread(groupedSlide.resolve(`${groupedChild.id}/text`), "Review the grouped text.", { author: "Group Reviewer", created: "2026-07-13T05:08:09Z" });
groupedSlide.comments.addThread(nestedGroup, "Review the nested group.", { author: "Group Reviewer", created: "2026-07-13T05:09:10Z" });
const groupedPptx = await PresentationFile.exportPptx(groupedPresentation);
assert.equal((await PresentationFile.inspectPptx(groupedPptx)).ok, true);
const groupedZip = await JSZip.loadAsync(new Uint8Array(await groupedPptx.arrayBuffer()));
const groupedSlideXml = await groupedZip.file("ppt/slides/slide1.xml").async("text");
const groupedCommentsXml = await groupedZip.file("ppt/comments/comment1.xml").async("text");
assert.equal((groupedSlideXml.match(/<p:grpSp>/g) || []).length, 2);
assert.match(groupedSlideXml, /<a:xfrm><a:off x="952500" y="762000"\/><a:ext cx="3810000" cy="1905000"\/><a:chOff x="0" y="0"\/><a:chExt cx="1905000" cy="952500"\/><\/a:xfrm>/);
assert.match(groupedCommentsXml, /<oac:deMkLst>[\s\S]*?<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><\/oac:deMkLst>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/>/);
assert.match(groupedCommentsXml, /<oac:txMkLst>[\s\S]*?<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/><oac:txMk cp="0" len="18"\/>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:grpSpMk id="5" creationId="\{[0-9A-F-]+\}"\/>/);
const groupedLoaded = await PresentationFile.importPptx(groupedPptx);
const groupedLoadedSlide = groupedLoaded.slides.items[0];
const groupedLoadedGroup = groupedLoadedSlide.groups.items[0];
assert.equal(groupedLoadedGroup.name, "review-group");
assert.equal(groupedLoadedGroup.shapes.items[0].text.value, "Grouped child text");
assert.equal(groupedLoadedGroup.groups.items[0].shapes.items[0].text.value, "Nested");
assert.deepEqual(groupedLoadedGroup.absoluteChildFrame(groupedLoadedGroup.shapes.items[0]), { left: 140, top: 100, width: 200, height: 60 });
assert.match(groupedLoaded.inspect({ kind: "groupShape,shape,textRange", maxChars: 20_000 }).ndjson, /review-group[\s\S]*grouped-child[\s\S]*nested-group/);
const groupedLayout = groupedLoadedGroup.layoutJson();
assert.deepEqual(groupedLayout.children[0].frame, { left: 140, top: 100, width: 200, height: 60 });
assert.equal(groupedLoadedSlide.resolve(groupedLoadedGroup.groups.items[0].shapes.items[0].id)?.text.value, "Nested");
assert.deepEqual(groupedLoadedSlide.comments.items.map((thread) => groupedLoadedSlide.resolve(thread.targetId)?.name), ["review-group", "grouped-child", undefined, "nested-group"]);
assert.equal(groupedLoadedSlide.resolve(groupedLoadedSlide.comments.items[2].targetId)?.kind, "textRange");
assert.deepEqual(groupedLoadedSlide.comments.items.map((thread) => thread.nativeAnchor.monikers.length), [1, 2, 2, 2]);
const groupedSecondPptx = await PresentationFile.exportPptx(groupedLoaded);
const groupedSecondZip = await JSZip.loadAsync(new Uint8Array(await groupedSecondPptx.arrayBuffer()));
assert.match(await groupedSecondZip.file("ppt/comments/comment1.xml").async("text"), /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/>/);
const alternatePrefixGrouped = await PresentationFile.patchPptx(groupedPptx, [{
  path: "ppt/slides/slide1.xml",
  xml: groupedSlideXml.replaceAll("<p:", "<slide:").replaceAll("</p:", "</slide:").replace("xmlns:p=", "xmlns:slide=").replaceAll("<a:", "<draw:").replaceAll("</a:", "</draw:").replace("xmlns:a=", "xmlns:draw="),
}]);
const alternatePrefixGroupedLoaded = await PresentationFile.importPptx(alternatePrefixGrouped);
assert.equal(alternatePrefixGroupedLoaded.slides.items[0].groups.items[0].shapes.items[0].text.value, "Grouped child text");
const groupCreationIds = [...groupedSlideXml.matchAll(/a16:creationId[^>]*id="(\{[0-9A-F-]+\})"/g)].map((match) => match[1]);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/comments/comment1.xml", xml: groupedCommentsXml.replace('<oac:grpSpMk id="2"', '<oac:grpSpMk id="99"') }]), /invalid OOXML package.*pptxModernCommentDrawingIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/comments/comment1.xml", xml: groupedCommentsXml.replace(/<oac:grpSpMk id="2"[^>]*\/><oac:spMk id="3"/, '<oac:spMk id="3"') }]), /invalid OOXML package.*pptxModernCommentDrawingIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/slides/slide1.xml", xml: groupedSlideXml.replace(groupCreationIds[3], groupCreationIds[0]) }]), /invalid OOXML package.*pptxModernDrawingCreationIdDuplicate/);
const invalidGroupLayout = Presentation.create();
const invalidGroupSlide = invalidGroupLayout.slides.add();
invalidGroupSlide.groups.add({ name: "bounded-group", position: { left: 0, top: 0, width: 100, height: 100 }, childFrame: { left: 0, top: 0, width: 100, height: 100 }, shapes: [{ name: "outside-child", position: { left: 90, top: 10, width: 20, height: 20 } }] });
assert.ok(invalidGroupLayout.verify().issues.some((issue) => issue.type === "groupChildOutOfBounds"));

const invalidModernModel = Presentation.create({ commentFormat: "modern" });
const invalidModernSlide = invalidModernModel.slides.add();
invalidModernSlide.comments.addThread(undefined, "Invalid identity", { comments: [{ nativeId: "not-a-guid", author: "Reviewer", text: "Invalid identity", created: "2026-07-13T01:02:03Z" }] });
assert.ok(invalidModernModel.verify().issues.some((issue) => issue.type === "invalidModernCommentMetadata"));
await assert.rejects(() => PresentationFile.exportPptx(invalidModernModel), /must be a brace-delimited GUID/);
console.log("presentation smoke ok");
