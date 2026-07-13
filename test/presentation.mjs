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
  styleId: 10,
  varyColors: true,
  barOptions: { direction: "bar", grouping: "stacked", gapWidth: 80, overlap: -20 },
  series: [
    { name: "Revenue", values: [10, 14], color: "#0ea5e9", line: { fill: "#0369a1", width: 1.5, style: "dash" }, points: [{ idx: 1, fill: "#facc15", line: { fill: "#dc2626", width: 2, style: "dot" } }] },
    { name: "Forecast", values: [12, 16], fill: "accent1" },
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
  series: [{ name: "Share", values: [45, 35, 20], points: [{ idx: 1, fill: "#facc15" }] }],
});
const lineChart = pieSlide.charts.add("line", {
  name: "trend-line",
  title: "Quality Trend",
  position: { left: 540, top: 120, width: 420, height: 240 },
  categories: ["Plan", "Build", "Verify"],
  styleId: 13,
  lineOptions: { grouping: "stacked", smooth: true, marker: { symbol: "diamond", size: 9 } },
  series: [
    { name: "Model", values: [4, 7, 9], color: "#22c55e", line: { fill: "#15803d", width: 3, style: "dashDot" }, points: [{ idx: 1, fill: "#eab308" }] },
    { name: "Native", values: [3, 6, 10], color: "#a855f7", marker: { symbol: "triangle", size: 8 }, smooth: false },
  ],
});
const comboChart = pieSlide.charts.add("combo", {
  name: "revenue-margin-combo",
  title: "Revenue and Margin",
  position: { left: 120, top: 400, width: 840, height: 250 },
  categories: ["Q1", "Q2", "Q3"],
  styleId: 8,
  barOptions: { direction: "column", grouping: "clustered", gapWidth: 90, overlap: 0 },
  lineOptions: { grouping: "standard", marker: { symbol: "circle", size: 7 }, smooth: false },
  axes: { category: { title: "Quarter" }, value: { title: "Amount" } },
  legend: { visible: true, position: "b" },
  dataLabels: { showValue: true, position: "outsideEnd" },
  series: [
    { chartType: "bar", name: "Revenue", values: [10, 14, 18], color: "#3d8dff", dataLabels: { showValue: true, showCategoryName: true, position: "insideEnd" } },
    { chartType: "line", name: "Margin", values: [4, 6, 7], color: "#dc2626", line: { fill: "#dc2626", width: 2, style: "dash" }, marker: { symbol: "diamond", size: 8 }, dataLabels: { showValue: true, position: "top" }, trendline: { type: "linear", name: "Margin trend", forward: 0.5, backward: 0.5, intercept: 1, displayEquation: true, displayRSquared: true, line: { fill: "#111827", width: 1.25, style: "dashDot" } } },
    { chartType: "bar", name: "Forecast", values: [12, 16, 20], color: "#6dcbf4", dataLabels: false },
  ],
});
assert.throws(() => pieSlide.charts.add("bar", { styleId: 49 }), /styleId must be an integer from 1 to 48/);
assert.throws(() => pieSlide.charts.add("bar", { barOptions: { grouping: "standard" } }), /chart bar grouping must be one of/);
assert.throws(() => pieSlide.charts.add("line", { lineOptions: { marker: { symbol: "hexagon" } } }), /chart marker symbol must be one of/);
assert.throws(() => pieSlide.charts.add("bar", { series: [{ values: [1], dataLabels: { position: "floating" } }] }), /data-label position must be one of/);
assert.throws(() => pieSlide.charts.add("pie", { series: [{ values: [1, 2, 3], trendline: { type: "linear" } }] }), /supported only for bar and line series/);
assert.throws(() => pieSlide.charts.add("line", { series: [{ values: [1, 2], trendline: { type: "movingAverage", period: 2 } }] }), /require at least three series values/);
assert.throws(() => pieSlide.charts.add("line", { series: [{ values: [1, 2, 3], trendline: { type: "linear", forward: 0.25 } }] }), /must use 0.5 increments/);
assert.throws(() => pieSlide.charts.add("line", { series: [{ values: [1, 2, 3], trendline: { type: "linear", order: 2 } }] }), /order is supported only for polynomial/);
assert.throws(() => pieSlide.charts.add("line", { series: [{ values: [1, 2, 3], trendline: { type: "polynomial", order: 7 } }] }), /order must be an integer from 2 to 6/);
assert.throws(() => pieSlide.charts.add("combo", { series: [{ name: "Missing type", values: [1] }] }), /series chartType must be bar or line/);
assert.throws(() => pieSlide.charts.add("combo", { series: [{ chartType: "bar", name: "Only bars", values: [1] }] }), /requires at least one bar series and one line series/);
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
assert.equal(presentation.resolve(nativeChart.id).styleId, 10);
assert.deepEqual(presentation.resolve(nativeChart.id).barOptions, { direction: "bar", grouping: "stacked", gapWidth: 80, overlap: -20 });
assert.deepEqual(presentation.resolve(lineChart.id).lineOptions, { grouping: "stacked", marker: { symbol: "diamond", size: 9 }, smooth: true });
assert.deepEqual(presentation.resolve(comboChart.id).series.map((series) => series.chartType), ["bar", "line", "bar"]);
assert.deepEqual(presentation.resolve(comboChart.id).barOptions, { direction: "column", grouping: "clustered", gapWidth: 90, overlap: 0 });
assert.deepEqual(presentation.resolve(comboChart.id).lineOptions, { grouping: "standard", marker: { symbol: "circle", size: 7 }, smooth: false });
assert.deepEqual(presentation.resolve(comboChart.id).series.map((series) => series.dataLabels), [
  { showValue: true, showCategoryName: true, position: "inEnd" },
  { showValue: true, showCategoryName: false, position: "t" },
  { showValue: false, showCategoryName: false, position: "bestFit" },
]);
assert.deepEqual(presentation.resolve(comboChart.id).series[1].trendlines, [{
  type: "linear",
  name: "Margin trend",
  forward: 0.5,
  backward: 0.5,
  intercept: 1,
  displayEquation: true,
  displayRSquared: true,
  line: { fill: "#111827", width: 1.25, style: "dashDot" },
}]);
const trendlineCatalog = Presentation.create().slides.add().charts.add("line", {
  categories: ["A", "B", "C", "D"],
  series: [{
    values: [1, 2, 4, 8],
    trendlines: [
      { type: "exponential" }, { type: "linear" }, { type: "logarithmic" },
      { type: "movingAverage", period: 2 }, { type: "polynomial", order: 3 }, { type: "power" },
    ],
  }],
});
assert.deepEqual(trendlineCatalog.series[0].trendlines.map((trendline) => trendline.type), ["exp", "linear", "log", "movingAvg", "poly", "power"]);
const comboChartSvg = comboChart.toSvg();
assert.match(comboChartSvg, /<rect [^>]*fill="#3d8dff"/i);
assert.match(comboChartSvg, /<polyline [^>]*stroke="#dc2626"/i);
assert.match(comboChartSvg, />6<\/text>/);
assert.match(comboChartSvg, />Q2: 14<\/text>/);
assert.doesNotMatch(comboChartSvg, />16<\/text>/);
assert.match(comboChartSvg, /<line [^>]*stroke="#111827"[^>]*stroke-width="1\.25"[^>]*stroke-dasharray="6 3 1 3"/i);
assert.ok(comboChartSvg.indexOf('fill="#3d8dff"') < comboChartSvg.indexOf("<polyline"));
const lineChartSvg = lineChart.toSvg();
assert.match(lineChartSvg, /<path d="M/);
assert.match(lineChartSvg, /930,162/);
assert.match(lineChartSvg, /M 930 158 L 934 166 L 926 166/);
assert.throws(() => slide.charts.add("bar", { styleId: 49 }), /styleId must be an integer from 1 to 48/);
assert.throws(() => slide.charts.add("bar", { barOptions: { gapWidth: 501 } }), /gapWidth must be an integer from 0 to 500/);
assert.throws(() => slide.charts.add("line", { lineOptions: { marker: { symbol: "picture" } } }), /chart marker symbol must be one of/);
assert.throws(() => slide.charts.add("bar", { series: [{ values: [1], points: [{ idx: 1, fill: "#ffffff" }] }] }), /outside the series value range/);
assert.throws(() => slide.charts.add("bar", { series: [{ values: [1], points: [{ idx: 0 }, { idx: 0 }] }] }), /point idx 0 is duplicated/);
assert.throws(() => slide.charts.add("line", { series: [{ values: [1], line: { style: "scribble" } }] }), /chart line style must be one of/);
assert.equal(Presentation.create().slides.add().charts.add("bar", { styleIndex: 12 }).styleId, 12);
assert.match(presentation.inspect({ kind: "chart", target: nativeChart.id, maxChars: 8000 }).ndjson, /"axes"/);
assert.match(presentation.inspect({ kind: "chart", target: nativeChart.id, maxChars: 8000 }).ndjson, /Forecast/);
assert.match(presentation.help("slide.charts.add").ndjson, /bar\/line\/pie.*combo/);
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
assert.match(slideXml, /<a:tcPr[^>]*><a:lnL[\s\S]*?<a:solidFill><a:srgbClr val="EDEDED"\/><\/a:solidFill><\/a:tcPr>/);
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
assert.match(chartXml, /<a:solidFill><a:schemeClr val="accent1"\/><\/a:solidFill>/);
assert.match(chartXml, /<c:style val="10"\/>/);
assert.match(chartXml, /<c:barDir val="bar"\/>/);
assert.match(chartXml, /<c:grouping val="stacked"\/>/);
assert.match(chartXml, /<c:varyColors val="1"\/>/);
assert.match(chartXml, /<c:gapWidth val="80"\/>/);
assert.match(chartXml, /<c:overlap val="-20"\/>/);
assert.match(chartXml, /<a:ln w="19050"><a:solidFill><a:srgbClr val="0369A1"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln>/);
assert.match(chartXml, /<c:dPt><c:idx val="1"\/><c:spPr><a:solidFill><a:srgbClr val="FACC15"\/><\/a:solidFill><a:ln w="25400">/);
assert.match(chartXml, /<a:prstDash val="dot"\/><\/a:ln><\/c:spPr><\/c:dPt>/);
const pieChartXml = await zip.file("ppt/charts/chart2.xml").async("text");
assert.match(pieChartXml, /<c:pieChart>/);
assert.match(pieChartXml, /Market Share/);
assert.match(pieChartXml, /Product A/);
assert.match(pieChartXml, /<c:v>45<\/c:v>/);
assert.match(pieChartXml, /<c:dPt><c:idx val="1"\/><c:spPr><a:solidFill><a:srgbClr val="FACC15"\/><\/a:solidFill><\/c:spPr><\/c:dPt>/);
const lineChartXml = await zip.file("ppt/charts/chart3.xml").async("text");
assert.match(lineChartXml, /<c:style val="13"\/>/);
assert.match(lineChartXml, /<c:lineChart><c:grouping val="stacked"\/>/);
assert.match(lineChartXml, /<c:marker><c:symbol val="diamond"\/><c:size val="9"\/><\/c:marker>/);
assert.match(lineChartXml, /<c:marker><c:symbol val="triangle"\/><c:size val="8"\/><\/c:marker>/);
assert.equal((lineChartXml.match(/<c:smooth val="1"\/>/g) || []).length, 1);
assert.equal((lineChartXml.match(/<c:smooth val="0"\/>/g) || []).length, 1);
assert.match(lineChartXml, /<a:ln w="38100"><a:solidFill><a:srgbClr val="15803D"\/><\/a:solidFill><a:prstDash val="dashDot"\/><\/a:ln>/);
assert.match(lineChartXml, /<c:dPt><c:idx val="1"\/><c:spPr><a:solidFill><a:srgbClr val="EAB308"\/><\/a:solidFill><\/c:spPr><\/c:dPt>/);
const comboChartXml = await zip.file("ppt/charts/chart4.xml").async("text");
assert.match(comboChartXml, /<c:style val="8"\/>/);
assert.match(comboChartXml, /<c:barChart><c:barDir val="col"\/><c:grouping val="clustered"\/>/);
assert.match(comboChartXml, /<c:lineChart><c:grouping val="standard"\/>/);
assert.match(comboChartXml, /<c:gapWidth val="90"\/><c:overlap val="0"\/>/);
assert.match(comboChartXml, /<c:idx val="0"\/><c:order val="0"\/>[\s\S]*?<c:v>Revenue<\/c:v>/);
assert.match(comboChartXml, /<c:idx val="1"\/><c:order val="1"\/>[\s\S]*?<c:v>Margin<\/c:v>/);
assert.match(comboChartXml, /<c:idx val="2"\/><c:order val="2"\/>[\s\S]*?<c:v>Forecast<\/c:v>/);
assert.equal((comboChartXml.match(/<c:catAx>/g) || []).length, 1);
assert.equal((comboChartXml.match(/<c:valAx>/g) || []).length, 1);
assert.equal((comboChartXml.match(/<c:axId val="1"\/>/g) || []).length, 3);
assert.equal((comboChartXml.match(/<c:axId val="2"\/>/g) || []).length, 3);
assert.match(comboChartXml, /<c:marker><c:symbol val="diamond"\/><c:size val="8"\/><\/c:marker>/);
assert.match(comboChartXml, /<a:ln w="25400"><a:solidFill><a:srgbClr val="DC2626"\/><\/a:solidFill><a:prstDash val="dash"\/><\/a:ln>/);
const comboSeriesBlocks = [...comboChartXml.matchAll(/<c:ser>[\s\S]*?<\/c:ser>/g)].map((match) => match[0]);
const revenueSeriesXml = comboSeriesBlocks.find((block) => /<c:v>Revenue<\/c:v>/.test(block));
const marginSeriesXml = comboSeriesBlocks.find((block) => /<c:v>Margin<\/c:v>/.test(block));
const forecastSeriesXml = comboSeriesBlocks.find((block) => /<c:v>Forecast<\/c:v>/.test(block));
assert.match(revenueSeriesXml, /<c:dLbls><c:dLblPos val="inEnd"\/>[\s\S]*?<c:showVal val="1"\/>[\s\S]*?<c:showCatName val="1"\/>/);
assert.match(marginSeriesXml, /<c:dLbls><c:dLblPos val="t"\/>[\s\S]*?<c:showVal val="1"\/>[\s\S]*?<c:showCatName val="0"\/>/);
assert.match(marginSeriesXml, /<c:trendline><c:name>Margin trend<\/c:name><c:spPr><a:ln w="15875">[\s\S]*?<a:srgbClr val="111827"\/>[\s\S]*?<a:prstDash val="dashDot"\/>[\s\S]*?<c:trendlineType val="linear"\/><c:forward val="0\.5"\/><c:backward val="0\.5"\/><c:intercept val="1"\/><c:dispRSqr val="1"\/><c:dispEq val="1"\/><\/c:trendline>/);
assert.match(forecastSeriesXml, /<c:dLbls><c:showLegendKey val="0"\/><c:showVal val="0"\/><c:showCatName val="0"\/>/);
assert.equal((comboChartXml.match(/<c:dLblPos val="outEnd"\/>/g) || []).length, 2);
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
assert.match(loadedAll, /Quality Trend/);
assert.match(loadedAll, /Revenue and Margin/);
assert.match(loadedAll, /native-import-image/);
assert.match(loadedAll, /Native import logo/);
assert.match(loadedAll, /Speaker note/);
assert.match(loadedAll, /Tighten this headline/);
assert.match(loadedAll, /shape-to-table/);
assert.match(loadedAll, /"textStyles"/);
const loadedBarChart = loaded.slides.items[0].charts.items.find((chart) => chart.name === "native-import-chart");
assert.equal(loadedBarChart.styleId, 10);
assert.deepEqual(loadedBarChart.barOptions, { direction: "bar", grouping: "stacked", gapWidth: 80, overlap: -20 });
assert.deepEqual(loadedBarChart.series[0].line, { fill: "#0369A1", width: 1.5, style: "dash" });
assert.deepEqual(loadedBarChart.series[0].points, [{ idx: 1, fill: "#FACC15", line: { fill: "#DC2626", width: 2, style: "dot" } }]);
assert.equal(loadedBarChart.series[1].color, "accent1");
const loadedLineChart = loaded.slides.items[1].charts.items.find((chart) => chart.name === "trend-line");
assert.equal(loadedLineChart.styleId, 13);
assert.equal(loadedLineChart.lineOptions.grouping, "stacked");
assert.deepEqual(loadedLineChart.series.map((series) => series.marker), [{ symbol: "diamond", size: 9 }, { symbol: "triangle", size: 8 }]);
assert.deepEqual(loadedLineChart.series.map((series) => series.smooth), [true, false]);
assert.deepEqual(loadedLineChart.series[0].line, { fill: "#15803D", width: 3, style: "dashDot" });
assert.deepEqual(loadedLineChart.series[0].points, [{ idx: 1, fill: "#EAB308" }]);
const loadedComboChart = loaded.slides.items[1].charts.items.find((chart) => chart.name === "revenue-margin-combo");
assert.equal(loadedComboChart.chartType, "combo");
assert.deepEqual(loadedComboChart.series.map((series) => series.name), ["Revenue", "Margin", "Forecast"]);
assert.deepEqual(loadedComboChart.series.map((series) => series.chartType), ["bar", "line", "bar"]);
assert.deepEqual(loadedComboChart.barOptions, { direction: "column", grouping: "clustered", gapWidth: 90, overlap: 0 });
assert.deepEqual(loadedComboChart.lineOptions, { grouping: "standard", marker: { symbol: "diamond", size: 8 }, smooth: false });
assert.deepEqual(loadedComboChart.series[1].line, { fill: "#DC2626", width: 2, style: "dash" });
assert.deepEqual(loadedComboChart.series[1].trendlines, [{
  type: "linear",
  name: "Margin trend",
  forward: 0.5,
  backward: 0.5,
  intercept: 1,
  displayEquation: true,
  displayRSquared: true,
  line: { fill: "#111827", width: 1.25, style: "dashDot" },
}]);
assert.deepEqual(loadedComboChart.dataLabels, { showValue: true, showCategoryName: false, position: "outEnd" });
assert.deepEqual(loadedComboChart.series.map((series) => series.dataLabels), [
  { showValue: true, showCategoryName: true, position: "inEnd" },
  { showValue: true, showCategoryName: false, position: "t" },
  { showValue: false, showCategoryName: false, position: "bestFit" },
]);
const alternateChartPrefixXml = lineChartXml.replace('<c:smooth val="1"/>', "<c:smooth/>")
  .replaceAll("xmlns:c=", "xmlns:cx=").replaceAll("<c:", "<cx:").replaceAll("</c:", "</cx:")
  .replaceAll("xmlns:a=", "xmlns:ax=").replaceAll("<a:", "<ax:").replaceAll("</a:", "</ax:");
const alternateChartPrefixPptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/charts/chart3.xml", xml: alternateChartPrefixXml }]);
const alternateChartPrefixLoaded = await PresentationFile.importPptx(alternateChartPrefixPptx);
const alternateLineChart = alternateChartPrefixLoaded.slides.items[1].charts.items.find((chart) => chart.name === "trend-line");
assert.equal(alternateLineChart.styleId, 13);
assert.equal(alternateLineChart.title, "Quality Trend");
assert.deepEqual(alternateLineChart.series[0].marker, { symbol: "diamond", size: 9 });
const alternateChartSecondZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(alternateChartPrefixLoaded)).arrayBuffer()));
const alternateChartSecondXml = await alternateChartSecondZip.file("ppt/charts/chart3.xml").async("text");
assert.match(alternateChartSecondXml, /<c:style val="13"\/>/);
assert.match(alternateChartSecondXml, /<a:prstDash val="dashDot"\/>/);
assert.match(alternateChartSecondXml, /<c:dPt><c:idx val="1"\/>[\s\S]*?<a:srgbClr val="EAB308"\/>/);
const alternateComboPrefixXml = comboChartXml
  .replaceAll("xmlns:c=", "xmlns:cx=").replaceAll("<c:", "<cx:").replaceAll("</c:", "</cx:")
  .replaceAll("xmlns:a=", "xmlns:ax=").replaceAll("<a:", "<ax:").replaceAll("</a:", "</ax:");
const alternateComboPrefixPptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/charts/chart4.xml", xml: alternateComboPrefixXml }]);
const alternateComboPrefixLoaded = await PresentationFile.importPptx(alternateComboPrefixPptx);
const alternateComboChart = alternateComboPrefixLoaded.slides.items[1].charts.items.find((chart) => chart.name === "revenue-margin-combo");
assert.equal(alternateComboChart.chartType, "combo");
assert.deepEqual(alternateComboChart.series.map((series) => `${series.chartType}:${series.name}`), ["bar:Revenue", "line:Margin", "bar:Forecast"]);
assert.deepEqual(alternateComboChart.series.map((series) => series.dataLabels?.position), ["inEnd", "t", "bestFit"]);
assert.equal(alternateComboChart.series[1].trendlines[0].name, "Margin trend");
const alternateComboSecondZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(alternateComboPrefixLoaded)).arrayBuffer()));
const alternateComboSecondXml = await alternateComboSecondZip.file("ppt/charts/chart4.xml").async("text");
assert.match(alternateComboSecondXml, /<c:barChart>[\s\S]*?<c:lineChart>/);
assert.equal((alternateComboSecondXml.match(/<c:catAx>/g) || []).length, 1);
assert.deepEqual([...alternateComboSecondXml.matchAll(/<c:order val="(\d+)"\/>/g)].map((match) => Number(match[1])).sort((left, right) => left - right), [0, 1, 2]);
assert.match(alternateComboSecondXml, /<c:dLblPos val="inEnd"\/>/);
assert.match(alternateComboSecondXml, /<c:dLblPos val="t"\/>/);
assert.match(alternateComboSecondXml, /<c:trendlineType val="linear"\/>/);
assert.match(alternateComboSecondXml, /<c:dispRSqr val="1"\/><c:dispEq val="1"\/>/);
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
const groupedTable = reviewGroup.tables.add({ name: "grouped-table", position: { left: 10, top: 65, width: 70, height: 25 }, values: [["A", "B"], [1, 2]], rows: 2, columns: 2, styleOptions: { headerRow: true } });
const groupedChart = reviewGroup.charts.add("bar", { name: "grouped-chart", title: "Grouped chart", position: { left: 85, top: 60, width: 60, height: 35 }, categories: ["A", "B"], series: [{ name: "Value", values: [1, 2] }] });
const groupedImage = reviewGroup.images.add({ name: "grouped-image", alt: "Grouped pixel", position: { left: 150, top: 60, width: 30, height: 30 }, dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=" });
groupedSlide.comments.addThread(reviewGroup, "Review the grouped surface.", { author: "Group Reviewer", created: "2026-07-13T05:06:07Z" });
groupedSlide.comments.addThread(groupedChild, "Review the grouped child.", { author: "Group Reviewer", created: "2026-07-13T05:07:08Z" });
groupedSlide.comments.addThread(groupedSlide.resolve(`${groupedChild.id}/text`), "Review the grouped text.", { author: "Group Reviewer", created: "2026-07-13T05:08:09Z" });
groupedSlide.comments.addThread(nestedGroup, "Review the nested group.", { author: "Group Reviewer", created: "2026-07-13T05:09:10Z" });
groupedSlide.comments.addThread(groupedTable, "Review the grouped table.", { author: "Group Reviewer", created: "2026-07-13T05:10:11Z" });
groupedSlide.comments.addThread(groupedChart, "Review the grouped chart.", { author: "Group Reviewer", created: "2026-07-13T05:11:12Z" });
groupedSlide.comments.addThread(groupedImage, "Review the grouped image.", { author: "Group Reviewer", created: "2026-07-13T05:12:13Z" });
const groupedPptx = await PresentationFile.exportPptx(groupedPresentation);
assert.equal((await PresentationFile.inspectPptx(groupedPptx)).ok, true);
const groupedZip = await JSZip.loadAsync(new Uint8Array(await groupedPptx.arrayBuffer()));
const groupedSlideXml = await groupedZip.file("ppt/slides/slide1.xml").async("text");
const groupedCommentsXml = await groupedZip.file("ppt/comments/comment1.xml").async("text");
assert.equal((groupedSlideXml.match(/<p:grpSp>/g) || []).length, 2);
assert.equal((groupedSlideXml.match(/<p:graphicFrame>/g) || []).length, 2);
assert.equal((groupedSlideXml.match(/<p:pic>/g) || []).length, 1);
assert.match(groupedSlideXml, /<a:xfrm><a:off x="952500" y="762000"\/><a:ext cx="3810000" cy="1905000"\/><a:chOff x="0" y="0"\/><a:chExt cx="1905000" cy="952500"\/><\/a:xfrm>/);
assert.match(groupedCommentsXml, /<oac:deMkLst>[\s\S]*?<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><\/oac:deMkLst>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/>/);
assert.match(groupedCommentsXml, /<oac:txMkLst>[\s\S]*?<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/><oac:txMk cp="0" len="18"\/>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:grpSpMk id="5" creationId="\{[0-9A-F-]+\}"\/>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:graphicFrameMk id="7" creationId="\{[0-9A-F-]+\}"\/>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:graphicFrameMk id="8" creationId="\{[0-9A-F-]+\}"\/>/);
assert.match(groupedCommentsXml, /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:picMk id="9" creationId="\{[0-9A-F-]+\}"\/>/);
assert.ok(groupedZip.file("ppt/charts/chart1.xml"));
assert.ok(groupedZip.file("ppt/media/image1.png"));
const groupedLoaded = await PresentationFile.importPptx(groupedPptx);
const groupedLoadedSlide = groupedLoaded.slides.items[0];
const groupedLoadedGroup = groupedLoadedSlide.groups.items[0];
assert.equal(groupedLoadedGroup.name, "review-group");
assert.equal(groupedLoadedGroup.shapes.items[0].text.value, "Grouped child text");
assert.equal(groupedLoadedGroup.groups.items[0].shapes.items[0].text.value, "Nested");
assert.deepEqual(groupedLoadedGroup.tables.items[0].values, [["A", "B"], ["1", "2"]]);
assert.equal(groupedLoadedGroup.charts.items[0].title, "Grouped chart");
assert.equal(groupedLoadedGroup.images.items[0].alt, "Grouped pixel");
assert.deepEqual(groupedLoadedGroup.absoluteChildFrame(groupedLoadedGroup.shapes.items[0]), { left: 140, top: 100, width: 200, height: 60 });
assert.match(groupedLoaded.inspect({ kind: "groupShape,shape,textRange,table,chart,image", maxChars: 20_000 }).ndjson, /review-group[\s\S]*grouped-child[\s\S]*nested-group[\s\S]*grouped-table[\s\S]*grouped-chart[\s\S]*grouped-image/);
const groupedLayout = groupedLoadedGroup.layoutJson();
assert.deepEqual(groupedLayout.children[0].frame, { left: 140, top: 100, width: 200, height: 60 });
assert.equal(groupedLoadedSlide.resolve(groupedLoadedGroup.groups.items[0].shapes.items[0].id)?.text.value, "Nested");
assert.deepEqual(groupedLoadedSlide.comments.items.map((thread) => groupedLoadedSlide.resolve(thread.targetId)?.name), ["review-group", "grouped-child", undefined, "nested-group", "grouped-table", "grouped-chart", "grouped-image"]);
assert.equal(groupedLoadedSlide.resolve(groupedLoadedSlide.comments.items[2].targetId)?.kind, "textRange");
assert.deepEqual(groupedLoadedSlide.comments.items.map((thread) => thread.nativeAnchor.monikers.length), [1, 2, 2, 2, 2, 2, 2]);
const groupedSecondPptx = await PresentationFile.exportPptx(groupedLoaded);
const groupedSecondZip = await JSZip.loadAsync(new Uint8Array(await groupedSecondPptx.arrayBuffer()));
assert.match(await groupedSecondZip.file("ppt/comments/comment1.xml").async("text"), /<oac:grpSpMk id="2" creationId="\{[0-9A-F-]+\}"\/><oac:spMk id="3" creationId="\{[0-9A-F-]+\}"\/>/);
assert.ok(groupedSecondZip.file("ppt/charts/chart1.xml"));
assert.ok(groupedSecondZip.file("ppt/media/image1.png"));
const alternatePrefixGrouped = await PresentationFile.patchPptx(groupedPptx, [{
  path: "ppt/slides/slide1.xml",
  xml: groupedSlideXml.replaceAll("<p:", "<slide:").replaceAll("</p:", "</slide:").replace("xmlns:p=", "xmlns:slide=").replaceAll("<a:", "<draw:").replaceAll("</a:", "</draw:").replace("xmlns:a=", "xmlns:draw="),
}]);
const alternatePrefixGroupedLoaded = await PresentationFile.importPptx(alternatePrefixGrouped);
assert.equal(alternatePrefixGroupedLoaded.slides.items[0].groups.items[0].shapes.items[0].text.value, "Grouped child text");
assert.equal(alternatePrefixGroupedLoaded.slides.items[0].groups.items[0].tables.items[0].values[0][0], "A");
const groupCreationIds = [...groupedSlideXml.matchAll(/a16:creationId[^>]*id="(\{[0-9A-F-]+\})"/g)].map((match) => match[1]);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/comments/comment1.xml", xml: groupedCommentsXml.replace('<oac:grpSpMk id="2"', '<oac:grpSpMk id="99"') }]), /invalid OOXML package.*pptxModernCommentDrawingIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/comments/comment1.xml", xml: groupedCommentsXml.replace(/<oac:grpSpMk id="2"[^>]*\/><oac:spMk id="3"/, '<oac:spMk id="3"') }]), /invalid OOXML package.*pptxModernCommentDrawingIdentityMismatch/);
await assert.rejects(() => PresentationFile.patchPptx(groupedPptx, [{ path: "ppt/slides/slide1.xml", xml: groupedSlideXml.replace(groupCreationIds[3], groupCreationIds[0]) }]), /invalid OOXML package.*pptxModernDrawingCreationIdDuplicate/);
const invalidGroupLayout = Presentation.create();
const invalidGroupSlide = invalidGroupLayout.slides.add();
invalidGroupSlide.groups.add({ name: "bounded-group", position: { left: 0, top: 0, width: 100, height: 100 }, childFrame: { left: 0, top: 0, width: 100, height: 100 }, shapes: [{ name: "outside-child", position: { left: 90, top: 10, width: 20, height: 20 } }] });
assert.ok(invalidGroupLayout.verify().issues.some((issue) => issue.type === "groupChildOutOfBounds"));
const legacyGroupedPresentation = Presentation.create();
const legacyGroupedSlide = legacyGroupedPresentation.slides.add();
const legacyGroup = legacyGroupedSlide.groups.add({ name: "legacy-group", position: { left: 20, top: 20, width: 200, height: 100 }, images: [{ name: "legacy-image", alt: "Legacy grouped image", position: { left: 10, top: 10, width: 40, height: 40 }, dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=" }] });
legacyGroupedSlide.comments.addThread(legacyGroup.images.items[0], "Legacy grouped picture review.", { author: "Legacy Reviewer" });
const legacyGroupedPptx = await PresentationFile.exportPptx(legacyGroupedPresentation);
assert.equal((await PresentationFile.inspectPptx(legacyGroupedPptx)).ok, true);
const legacyGroupedLoaded = await PresentationFile.importPptx(legacyGroupedPptx);
assert.equal(legacyGroupedLoaded.slides.items[0].resolve(legacyGroupedLoaded.slides.items[0].comments.items[0].targetId)?.name, "legacy-image");

const pictureBulletPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAAHUlEQVR4nGNQOhr3nxLMMGrA/9EwiBsNg6PDIgwAUQdEH39xn2wAAAAASUVORK5CYII=";
const paragraphPresentation = Presentation.create({
  master: {
    id: "master/lists",
    name: "List Master",
    textParagraphStyles: {
      body: {
        0: { bulletCharacter: "○", bulletFont: "Arial", bulletColor: "accent2", bulletSizePercent: 1.1, marginLeft: 30, indent: -15, style: { fontSize: 22, color: "tx1" } },
        1: { bulletCharacter: "–", bulletFont: "Arial", bulletColor: "tx2", bulletSize: 18, marginLeft: 52, indent: -20, spaceBeforePercent: 0.2, style: { fontSize: 18, color: "tx2" } },
        2: { bulletImage: { dataUrl: pictureBulletPng, alt: "Master picture bullet" }, bulletSize: 16, marginLeft: 72, indent: -20, style: { fontSize: 18, color: "tx2" } },
      },
      other: {
        0: { bulletImage: { uri: "https://example.com/master-status.png", alt: "External master picture bullet" }, bulletSizePercent: 0.75 },
      },
    },
    placeholders: [{
      type: "body",
      idx: 1,
      name: "Inherited List",
      position: { left: 80, top: 80, width: 720, height: 360 },
      paragraphStyles: { 0: { bulletCharacter: "•", marginLeft: 28, indent: -14, spaceAfter: 6, style: { fontSize: 20, color: "accent1" } } },
    }],
  },
  layouts: [{ id: "layout/lists", name: "List Layout", type: "obj", masterId: "master/lists", placeholders: [{ type: "body", idx: 1, name: "Inherited List", paragraphStyles: { 1: { bulletImage: { dataUrl: pictureBulletPng, alt: "Layout picture bullet" }, bulletSizePercent: 0.9 } } }] }],
});
const paragraphSlide = paragraphPresentation.slides.add({ layoutId: "layout/lists" });
const [inheritedListShape] = paragraphPresentation.layouts.getItem("layout/lists").apply(paragraphSlide);
inheritedListShape.text.set([
  { runs: ["Inherited one"] },
  { level: 1, runs: ["Layout picture"] },
  { level: 2, runs: ["Master picture"] },
  { level: 2, bulletCharacter: "!", runs: ["Direct override"] },
]);
const richTextShape = paragraphSlide.shapes.add({ name: "rich-list", position: { left: 820, top: 80, width: 380, height: 420 }, text: "" });
richTextShape.text.set([
  [{ run: "Status", textStyle: { bold: true, color: "#0f172a" } }, " review"],
  { bulletCharacter: "•", bulletFont: "Georgia", bulletColor: "#dc2626", bulletSizePercent: 1.5, marginLeft: 24, indent: -12, spaceAfter: 5, runs: [{ run: "Quality:", textStyle: { bold: true } }, " defects down"] },
  { level: 1, autoNumber: { type: "arabicPeriod", startAt: 3 }, bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true, marginLeft: 48, indent: -14, runs: [{ run: "Ship", textStyle: { italic: true, underline: "sng" } }] },
]);
richTextShape.text.style = { fontFamily: "Arial", fontSize: 20, color: "#334155", lineSpacing: 1.15 };
const pictureBulletShape = paragraphSlide.shapes.add({ name: "picture-list", position: { left: 80, top: 470, width: 420, height: 90 }, text: [{ bulletImage: { dataUrl: pictureBulletPng, alt: "Green status" }, bulletSize: 18, marginLeft: 30, indent: -20, runs: ["Embedded picture bullet"] }] });
const externalPictureBulletShape = paragraphSlide.shapes.add({ name: "external-picture-list", position: { left: 520, top: 470, width: 420, height: 90 }, text: [{ pictureBullet: "https://example.com/status.png", bulletSizePercent: 0.8, runs: ["External picture bullet"] }] });
const pictureBulletGroup = paragraphSlide.groups.add({ name: "picture-bullet-group", position: { left: 940, top: 470, width: 250, height: 100 }, childFrame: { left: 0, top: 0, width: 250, height: 100 } });
const groupedPictureBulletShape = pictureBulletGroup.shapes.add({ name: "grouped-picture-list", position: { left: 0, top: 0, width: 250, height: 100 }, text: [{ bullet: { image: pictureBulletPng }, runs: ["Grouped bullet"] }] });
assert.equal(richTextShape.text.value, "Status review\nQuality: defects down\nShip");
assert.equal(richTextShape.text.paragraphs[1].bulletCharacter, "•");
assert.deepEqual(richTextShape.text.paragraphs[2].autoNumber, { type: "arabicPeriod", startAt: 3 });
assert.match(richTextShape.toSvg(), />•<\/text>/);
assert.match(richTextShape.toSvg(), />3\.<\/text>/);
assert.match(richTextShape.toSvg(), /font-family="Georgia" font-size="30" fill="#dc2626">•<\/text>/);
assert.equal(pictureBulletShape.text.paragraphs[0].bulletImage.alt, "Green status");
assert.match(pictureBulletShape.toSvg(), /<image[^>]*data:image\/png;base64/);
assert.match(externalPictureBulletShape.toSvg(), /data-picture-bullet="external"/);
assert.equal(groupedPictureBulletShape.text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "xx", runs: ["bad"] }]), /exactly one Unicode character/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletImage: pictureBulletPng, runs: ["bad"] }]), /exactly one of bulletCharacter, autoNumber, bulletImage, or bulletNone/);
assert.throws(() => richTextShape.text.set([{ bulletImage: "data:text\/plain;base64,SGVsbG8=", runs: ["bad"] }]), /base64 PNG, JPEG, GIF, or SVG/);
assert.throws(() => richTextShape.text.set([{ autoNumber: { type: "unsupported" }, runs: ["bad"] }]), /Unsupported Presentation auto-number type/);
assert.throws(() => richTextShape.text.set([{ spaceBefore: 4, spaceBeforePercent: 0.2, runs: ["bad"] }]), /either spaceBefore or spaceBeforePercent/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletFont: "Arial", bulletFontFollowText: true, runs: ["bad"] }]), /cannot combine bulletFont/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletSizePercent: 0.2, runs: ["bad"] }]), /bulletSizePercent must be between 0.25 and 4/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletSize: 1, runs: ["bad"] }]), /bulletSize must be between 1.333 and 1024/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletSize: 20, bulletSizePercent: 1.2, runs: ["bad"] }]), /exactly one of bulletSize/);
assert.throws(() => richTextShape.text.set([{ bulletCharacter: "•", bulletColor: "not-a-color", runs: ["bad"] }]), /scheme color.*RGB color.*color token/);
const followOnlyDeck = Presentation.create();
const followOnlyShape = followOnlyDeck.slides.add().shapes.add({ text: { text: "Follow marker", bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true } });
assert.deepEqual(followOnlyShape.text.paragraphs[0], { runs: [{ text: "Follow marker", style: {} }], level: 0, bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true, style: {} });
assert.throws(() => richTextShape.text.set([[{ run: "linked", link: { uri: "https:\/\/example.com", isExternal: true } }]]), /structured-run links are not supported yet/);
assert.throws(() => Presentation.create({ master: { textParagraphStyles: { body: { 9: { bulletCharacter: "•" } } } } }), /level must be an integer from 0 through 8/);
richTextShape.text.set([
  [{ run: "Status", textStyle: { bold: true, color: "#0f172a" } }, " review"],
  { bulletCharacter: "•", bulletFont: "Georgia", bulletColor: "#dc2626", bulletSizePercent: 1.5, marginLeft: 24, indent: -12, spaceAfter: 5, runs: [{ run: "Quality:", textStyle: { bold: true } }, " defects down"] },
  { level: 1, autoNumber: { type: "arabicPeriod", startAt: 3 }, bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true, marginLeft: 48, indent: -14, runs: [{ run: "Ship", textStyle: { italic: true, underline: "sng" } }] },
]);
const paragraphInspect = paragraphPresentation.inspect({ kind: "textbox", target: richTextShape.id, maxChars: 20_000 });
assert.match(paragraphInspect.ndjson, /"bulletCharacter":"•"/);
assert.match(paragraphInspect.ndjson, /"bulletFont":"Georgia","bulletColor":"#dc2626"/);
assert.match(paragraphInspect.ndjson, /"bulletSizePercent":1.5/);
assert.match(paragraphInspect.ndjson, /"type":"arabicPeriod","startAt":3/);
assert.match(paragraphPresentation.inspect({ kind: "textbox", target: pictureBulletShape.id, maxChars: 20_000 }).ndjson, /"bulletImage":\{"dataUrl":"data:image\/png;base64/);
assert.match(paragraphPresentation.help("shape.text.set").ndjson, /structured paragraphs/);
const paragraphPptx = await PresentationFile.exportPptx(paragraphPresentation);
assert.equal((await PresentationFile.inspectPptx(paragraphPptx)).ok, true);
const paragraphZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
const paragraphSlideXml = await paragraphZip.file("ppt/slides/slide1.xml").async("text");
const paragraphSlideRelsXml = await paragraphZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
const paragraphMasterXml = await paragraphZip.file("ppt/slideMasters/slideMaster1.xml").async("text");
const paragraphMasterRelsXml = await paragraphZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels").async("text");
const paragraphLayoutXml = await paragraphZip.file("ppt/slideLayouts/slideLayout1.xml").async("text");
const paragraphLayoutRelsXml = await paragraphZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text");
assert.match(paragraphSlideXml, /<a:buChar char="•"\/>/);
assert.match(paragraphSlideXml, /<a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
assert.match(paragraphSlideXml, /<a:buClr><a:srgbClr val="DC2626"\/><\/a:buClr><a:buSzPct val="150000"\/><a:buFont typeface="Georgia"\/><a:buChar char="•"\/>/);
assert.match(paragraphSlideXml, /<a:buClrTx\/><a:buSzTx\/><a:buFontTx\/><a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
assert.equal((paragraphSlideXml.match(/<a:buBlip><a:blip r:embed="rId\d+"\/><\/a:buBlip>/g) || []).length, 6);
assert.match(paragraphSlideXml, /<a:buBlip><a:blip r:link="rId\d+"\/><\/a:buBlip>/);
assert.match(paragraphSlideRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="\.\.\/media\/image1\.png"/);
assert.match(paragraphSlideRelsXml, /Target="https:\/\/example\.com\/status\.png" TargetMode="External"/);
assert.ok(paragraphZip.file("ppt/media/image1.png"));
assert.equal(Object.keys(paragraphZip.files).filter((file) => /^ppt\/media\/image\d+\.png$/.test(file)).length, 1);
assert.match(paragraphMasterXml, /<p:bodyStyle>[\s\S]*?<a:lvl3pPr[\s\S]*?<a:buBlip><a:blip r:embed="rId3"\/><\/a:buBlip>/);
assert.match(paragraphMasterRelsXml, /Id="rId3" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="\.\.\/media\/image1\.png"/);
assert.match(paragraphMasterRelsXml, /Id="rId4" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="https:\/\/example\.com\/master-status\.png" TargetMode="External"/);
assert.match(paragraphLayoutXml, /<a:lstStyle>[\s\S]*?<a:lvl2pPr[\s\S]*?<a:buBlip><a:blip r:embed="rId2"\/><\/a:buBlip>/);
assert.match(paragraphLayoutRelsXml, /Id="rId2" Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/image" Target="\.\.\/media\/image1\.png"/);
const missingPictureBulletRelationshipZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
missingPictureBulletRelationshipZip.file("ppt/slides/_rels/slide1.xml.rels", paragraphSlideRelsXml.replace(/<Relationship[^>]*Target="\.\.\/media\/image1\.png"[^>]*\/>/, ""));
const missingPictureBulletRelationshipPptx = new FileBlob(await missingPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
await assert.rejects(() => PresentationFile.importPptx(missingPictureBulletRelationshipPptx), /picture bullet references missing image relationship/);
const missingMasterPictureBulletRelationshipZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
missingMasterPictureBulletRelationshipZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", paragraphMasterRelsXml.replace(/<Relationship Id="rId3"[^>]*\/>/, ""));
const missingMasterPictureBulletRelationshipPptx = new FileBlob(await missingMasterPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
await assert.rejects(() => PresentationFile.importPptx(missingMasterPictureBulletRelationshipPptx), /picture bullet references missing image relationship rId3/);
const missingLayoutPictureBulletRelationshipZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
missingLayoutPictureBulletRelationshipZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels", paragraphLayoutRelsXml.replace(/<Relationship Id="rId2"[^>]*\/>/, ""));
const missingLayoutPictureBulletRelationshipPptx = new FileBlob(await missingLayoutPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
await assert.rejects(() => PresentationFile.importPptx(missingLayoutPictureBulletRelationshipPptx), /picture bullet references missing image relationship rId2/);
const wrongTypePictureBulletRelationshipZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
wrongTypePictureBulletRelationshipZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", paragraphMasterRelsXml.replace(/(Id="rId3" Type="[^"]+)\/image"/, "$1/chart\""));
const wrongTypePictureBulletRelationshipPptx = new FileBlob(await wrongTypePictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
await assert.rejects(() => PresentationFile.importPptx(wrongTypePictureBulletRelationshipPptx), /picture bullet references missing image relationship rId3/);
const missingPictureBulletPartZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
missingPictureBulletPartZip.remove("ppt/media/image1.png");
const missingPictureBulletPartPptx = new FileBlob(await missingPictureBulletPartZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
await assert.rejects(() => PresentationFile.importPptx(missingPictureBulletPartPptx), /picture bullet relationship rId3 targets missing part ppt\/media\/image1\.png/);
const strictPictureBulletRelationshipZip = await JSZip.loadAsync(new Uint8Array(await paragraphPptx.arrayBuffer()));
strictPictureBulletRelationshipZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels", paragraphMasterRelsXml.replace("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "http://purl.oclc.org/ooxml/officeDocument/relationships/image"));
const strictPictureBulletRelationshipPptx = new FileBlob(await strictPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
assert.equal((await PresentationFile.importPptx(strictPictureBulletRelationshipPptx)).master.textParagraphStyles.body[2].bulletImage.dataUrl, pictureBulletPng);
assert.match(paragraphSlideXml, /<a:rPr[^>]*b="1"/);
assert.match(paragraphSlideXml, /<a:rPr[^>]*i="1"[^>]*u="sng"/);
assert.match(paragraphMasterXml, /<a:lstStyle><a:lvl1pPr[^>]*marL="266700"[^>]*indent="-133350"/);
assert.match(paragraphMasterXml, /<a:buChar char="•"\/>/);
assert.match(paragraphMasterXml, /<p:bodyStyle>[\s\S]*?<a:lvl2pPr[^>]*marL="495300"[^>]*indent="-190500"[\s\S]*?<a:spcPct val="20000"\/>[\s\S]*?<a:buChar char="–"\/>/);
assert.match(paragraphMasterXml, /<p:bodyStyle>[\s\S]*?<a:buClr><a:schemeClr val="accent2"\/><\/a:buClr><a:buSzPct val="110000"\/><a:buFont typeface="Arial"\/><a:buChar char="○"\/>/);
const inheritedListShapeXml = /<p:sp>[\s\S]*?<p:cNvPr[^>]*name="Inherited List"[\s\S]*?<\/p:sp>/.exec(paragraphSlideXml)[0];
paragraphZip.file("ppt/slides/slide1.xml", paragraphSlideXml.replace(inheritedListShapeXml, inheritedListShapeXml.replaceAll(/<a:buChar char="(?:•|○|–)"\/>/g, "").replaceAll(/<a:buBlip>[\s\S]*?<\/a:buBlip>/g, "").replaceAll(/<a:buClr>[\s\S]*?<\/a:buClr>/g, "").replaceAll(/<a:buSz(?:Pct|Pts)\b[^>]*\/>/g, "").replaceAll(/<a:buFont\b[^>]*\/>/g, "").replaceAll(/ marL="-?\d+" indent="-?\d+"/g, "").replaceAll(/<a:spcBef>[\s\S]*?<\/a:spcBef>/g, "")));
const inheritedParagraphPptx = new FileBlob(await paragraphZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
const paragraphLoaded = await PresentationFile.importPptx(inheritedParagraphPptx);
const paragraphLoadedInherited = paragraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "Inherited List");
const paragraphLoadedRich = paragraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "rich-list");
const paragraphLoadedPicture = paragraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "picture-list");
const paragraphLoadedExternalPicture = paragraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "external-picture-list");
const paragraphLoadedGroupedPicture = paragraphLoaded.slides.items[0].groups.items.find((group) => group.name === "picture-bullet-group").shapes.items[0];
assert.deepEqual(paragraphLoadedInherited.text.effectiveParagraphs().map((paragraph) => paragraph.bulletCharacter), ["•", undefined, undefined, "!"]);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[1].bulletImage.dataUrl, pictureBulletPng);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[2].bulletImage.dataUrl, pictureBulletPng);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[3].bulletImage, undefined);
assert.deepEqual(paragraphLoadedInherited.position, { left: 80, top: 80, width: 720, height: 360 });
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[0].marginLeft, 28);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[0].style.color, "accent1");
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[1].marginLeft, 52);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[1].indent, -20);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[1].spaceBeforePercent, 0.2);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[0].bulletFont, "Arial");
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[0].bulletColor, "accent2");
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[0].bulletSizePercent, 1.1);
assert.equal(paragraphLoaded.master.textParagraphStyles.body[1].bulletSize, 18);
assert.equal(paragraphLoadedInherited.text.effectiveParagraphs()[1].bulletSizePercent, 0.9);
assert.equal(paragraphLoaded.master.textParagraphStyles.body[2].bulletImage.dataUrl, pictureBulletPng);
assert.deepEqual(paragraphLoaded.master.textParagraphStyles.other[0].bulletImage, { uri: "https://example.com/master-status.png", relationshipMode: "link" });
assert.equal(paragraphLoaded.layouts.getItem("List Layout").placeholders[0].paragraphStyles[1].bulletImage.dataUrl, pictureBulletPng);
assert.equal(paragraphLoadedRich.text.paragraphs[1].bulletCharacter, "•");
assert.equal(paragraphLoadedRich.text.paragraphs[1].bulletFont, "Georgia");
assert.equal(paragraphLoadedRich.text.paragraphs[1].bulletColor, "#dc2626");
assert.equal(paragraphLoadedRich.text.paragraphs[1].bulletSizePercent, 1.5);
assert.deepEqual(paragraphLoadedRich.text.paragraphs[2].autoNumber, { type: "arabicPeriod", startAt: 3 });
assert.equal(paragraphLoadedRich.text.paragraphs[2].bulletFontFollowText, true);
assert.equal(paragraphLoadedRich.text.paragraphs[2].bulletColorFollowText, true);
assert.equal(paragraphLoadedRich.text.paragraphs[2].bulletSizeFollowText, true);
assert.equal(paragraphLoadedRich.text.paragraphs[2].runs[0].style.italic, true);
assert.equal(paragraphLoadedRich.text.paragraphs[2].runs[0].style.underline, "sng");
assert.equal(paragraphLoadedPicture.text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);
assert.deepEqual(paragraphLoadedExternalPicture.text.paragraphs[0].bulletImage, { uri: "https://example.com/status.png", relationshipMode: "link" });
assert.equal(paragraphLoadedGroupedPicture.text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);
const paragraphSecondPptx = await PresentationFile.exportPptx(paragraphLoaded);
assert.equal((await PresentationFile.inspectPptx(paragraphSecondPptx)).ok, true);
const paragraphSecondZip = await JSZip.loadAsync(new Uint8Array(await paragraphSecondPptx.arrayBuffer()));
assert.match(await paragraphSecondZip.file("ppt/slides/slide1.xml").async("text"), /<a:buAutoNum type="arabicPeriod" startAt="3"\/>/);
assert.equal(((await paragraphSecondZip.file("ppt/slides/slide1.xml").async("text")).match(/<a:buBlip>/g) || []).length, 7);
assert.match(await paragraphSecondZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /<p:bodyStyle>[\s\S]*?<a:buChar char="–"\/>/);
assert.match(await paragraphSecondZip.file("ppt/slideMasters/slideMaster1.xml").async("text"), /<p:bodyStyle>[\s\S]*?<a:lvl3pPr[\s\S]*?<a:buBlip>/);
assert.match(await paragraphSecondZip.file("ppt/slideLayouts/slideLayout1.xml").async("text"), /<a:lvl2pPr[\s\S]*?<a:buBlip>/);
const alternatePrefixParagraphZip = await JSZip.loadAsync(new Uint8Array(await inheritedParagraphPptx.arrayBuffer()));
for (const file of ["ppt/slideMasters/slideMaster1.xml", "ppt/slideLayouts/slideLayout1.xml", "ppt/slides/slide1.xml"]) {
  const xml = await alternatePrefixParagraphZip.file(file).async("text");
  let remapped = xml.replaceAll("<p:", "<deck:").replaceAll("</p:", "</deck:").replace("xmlns:p=", "xmlns:deck=").replaceAll("<a:", "<draw:").replaceAll("</a:", "</draw:").replace("xmlns:a=", "xmlns:draw=");
  remapped = remapped.replaceAll("r:", "rel:").replace("xmlns:r=", "xmlns:rel=");
  alternatePrefixParagraphZip.file(file, remapped);
}
for (const file of ["ppt/slideMasters/_rels/slideMaster1.xml.rels", "ppt/slideLayouts/_rels/slideLayout1.xml.rels", "ppt/slides/_rels/slide1.xml.rels"]) {
  const xml = await alternatePrefixParagraphZip.file(file).async("text");
  alternatePrefixParagraphZip.file(file, xml.replaceAll("http://schemas.openxmlformats.org/officeDocument/2006/relationships/image", "http://purl.oclc.org/ooxml/officeDocument/relationships/image"));
}
const alternatePrefixParagraphPptx = new FileBlob(await alternatePrefixParagraphZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: paragraphPptx.type });
const alternatePrefixParagraphLoaded = await PresentationFile.importPptx(alternatePrefixParagraphPptx);
assert.equal(alternatePrefixParagraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "Inherited List").text.effectiveParagraphs()[0].bulletCharacter, "•");
assert.deepEqual(alternatePrefixParagraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "rich-list").text.paragraphs[2].autoNumber, { type: "arabicPeriod", startAt: 3 });
assert.equal(alternatePrefixParagraphLoaded.slides.items[0].shapes.items.find((shape) => shape.name === "picture-list").text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);
assert.equal(alternatePrefixParagraphLoaded.master.textParagraphStyles.body[2].bulletImage.dataUrl, pictureBulletPng);
assert.deepEqual(alternatePrefixParagraphLoaded.master.textParagraphStyles.other[0].bulletImage, { uri: "https://example.com/master-status.png", relationshipMode: "link" });
assert.equal(alternatePrefixParagraphLoaded.layouts.getItem("List Layout").placeholders[0].paragraphStyles[1].bulletImage.dataUrl, pictureBulletPng);

// Unsupported native drawing objects remain agent-visible and preserve their complete OPC relationship graph.
const nativeObjectSource = Presentation.create();
const nativeObjectSourceSlide = nativeObjectSource.slides.add();
nativeObjectSourceSlide.images.add({ name: "shared-native-preview", alt: "Shared native preview", position: { left: 20, top: 20, width: 40, height: 40 }, dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=" });
const nativeObjectZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(nativeObjectSource)).arrayBuffer()));
const nativeObjectSlideXml = await nativeObjectZip.file("ppt/slides/slide1.xml").async("text");
const nativeOleXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="10" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdNativeOle" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>';
const nativeDiagramXml = '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="11" name="Preserved diagram"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="3657600"/><a:ext cx="5486400" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdNativeDm" r:lo="rIdNativeLo" r:qs="rIdNativeQs" r:cs="rIdNativeCs"/></a:graphicData></a:graphic></p:graphicFrame>';
const nativeContentGroupXml = '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="12" name="Native content group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="7000000" y="5000000"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></p:grpSpPr><p:contentPart r:id="rIdNativeContent"/></p:grpSp>';
nativeObjectZip.file("ppt/slides/slide1.xml", nativeObjectSlideXml.replace("</p:spTree>", `${nativeOleXml}${nativeDiagramXml}${nativeContentGroupXml}</p:spTree>`));
const nativeObjectRelsXml = await nativeObjectZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
nativeObjectZip.file("ppt/slides/_rels/slide1.xml.rels", nativeObjectRelsXml.replace("</Relationships>", '<Relationship Id="rIdNativeOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/native-workbook.xlsx"/><Relationship Id="rIdNativeDm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/native-data.xml"/><Relationship Id="rIdNativeLo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/native-layout.xml"/><Relationship Id="rIdNativeQs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/native-style.xml"/><Relationship Id="rIdNativeCs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/native-colors.xml"/><Relationship Id="rIdNativeContent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/native-content.xml"/></Relationships>'));
const embeddedWorkbookBytes = Uint8Array.of(80, 75, 3, 4, 1, 2, 3, 4);
nativeObjectZip.file("ppt/embeddings/native-workbook.xlsx", embeddedWorkbookBytes);
nativeObjectZip.file("ppt/diagrams/native-data.xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>');
nativeObjectZip.file("ppt/diagrams/native-layout.xml", '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-layout"><dgm:title val="Native"/><dgm:desc val="Native layout"/><dgm:catLst/></dgm:layoutDef>');
nativeObjectZip.file("ppt/diagrams/native-style.xml", '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-style"><dgm:title val="Native"/><dgm:desc val="Native style"/><dgm:catLst/></dgm:styleDef>');
nativeObjectZip.file("ppt/diagrams/native-colors.xml", '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-colors"><dgm:title val="Native"/><dgm:desc val="Native colors"/><dgm:catLst/></dgm:colorsDef>');
nativeObjectZip.file("ppt/customXml/native-content.xml", '<native xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rIdPayload">preserve me</native>');
nativeObjectZip.file("ppt/customXml/_rels/native-content.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPayload" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="assets/payload.svg"/></Relationships>');
nativeObjectZip.file("ppt/customXml/assets/payload.svg", '<svg xmlns="http://www.w3.org/2000/svg" width="8" height="8"><rect width="8" height="8" fill="#0ea5e9"/></svg>');
const nativeObjectContentTypes = await nativeObjectZip.file("[Content_Types].xml").async("text");
nativeObjectZip.file("[Content_Types].xml", nativeObjectContentTypes.replace("</Types>", '<Override PartName="/ppt/embeddings/native-workbook.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/diagrams/native-data.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/><Override PartName="/ppt/diagrams/native-layout.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/><Override PartName="/ppt/diagrams/native-style.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/><Override PartName="/ppt/diagrams/native-colors.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/><Override PartName="/ppt/customXml/native-content.xml" ContentType="application/xml"/><Override PartName="/ppt/customXml/assets/payload.svg" ContentType="image/svg+xml"/></Types>'));
const nativeObjectFixture = new FileBlob(await nativeObjectZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" });
const missingNativeRelationshipZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectFixture.arrayBuffer()));
missingNativeRelationshipZip.file("ppt/slides/slide1.xml", (await missingNativeRelationshipZip.file("ppt/slides/slide1.xml").async("text")).replace('r:id="rIdNativeOle"', 'r:id="rIdMissingNativeOle"'));
const missingNativeRelationshipPptx = new FileBlob(await missingNativeRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: nativeObjectFixture.type });
await assert.rejects(() => PresentationFile.importPptx(missingNativeRelationshipPptx), /native object references missing slide relationship rIdMissingNativeOle/);
const missingDiagramRelationshipZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectFixture.arrayBuffer()));
missingDiagramRelationshipZip.file("ppt/slides/slide1.xml", (await missingDiagramRelationshipZip.file("ppt/slides/slide1.xml").async("text")).replace('r:dm="rIdNativeDm"', 'r:dm="rIdMissingNativeDm"'));
const missingDiagramRelationshipPptx = new FileBlob(await missingDiagramRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: nativeObjectFixture.type });
const missingDiagramRelationshipInspect = await PresentationFile.inspectPptx(missingDiagramRelationshipPptx);
assert.ok(missingDiagramRelationshipInspect.issues.some((issue) => issue.type === "relationshipReferenceIdNotFound" && issue.referenceAttribute === "r:dm"));
const alternateNativePrefixZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectFixture.arrayBuffer()));
alternateNativePrefixZip.file("ppt/slides/slide1.xml", (await alternateNativePrefixZip.file("ppt/slides/slide1.xml").async("text")).replaceAll("r:", "rel:").replace("xmlns:r=", "xmlns:rel="));
const alternateNativePrefixPptx = new FileBlob(await alternateNativePrefixZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: nativeObjectFixture.type });
assert.deepEqual((await PresentationFile.importPptx(alternateNativePrefixPptx)).slides.items[0].nativeObjects.items.map((object) => object.nativeKind), ["oleObject", "diagram"]);
const nativeObjectLoaded = await PresentationFile.importPptx(nativeObjectFixture);
const nativeObjectLoadedSlide = nativeObjectLoaded.slides.items[0];
const nativeObjectVerification = nativeObjectLoaded.verify();
assert.equal(nativeObjectVerification.ok, true, nativeObjectVerification.ndjson);
assert.deepEqual(nativeObjectLoadedSlide.nativeObjects.items.map((object) => object.nativeKind), ["oleObject", "diagram"]);
assert.equal(nativeObjectLoadedSlide.groups.items[0].nativeObjects.items[0].nativeKind, "contentPart");
assert.equal(nativeObjectLoadedSlide.nativeObjects.items[0].parts.length, 2);
assert.equal(nativeObjectLoadedSlide.nativeObjects.items[1].parts.length, 4);
assert.equal(nativeObjectLoadedSlide.groups.items[0].nativeObjects.items[0].parts.length, 2);
const nativeObjectInspect = nativeObjectLoaded.inspect({ kind: "nativeObject", maxChars: 20_000 }).ndjson;
assert.match(nativeObjectInspect, /"nativeKind":"oleObject"/);
assert.match(nativeObjectInspect, /"nativeKind":"diagram"/);
assert.match(nativeObjectInspect, /"nativeKind":"contentPart"/);
assert.equal(nativeObjectLoadedSlide.resolve(nativeObjectLoadedSlide.nativeObjects.items[1].id)?.name, "Preserved diagram");
assert.equal(nativeObjectLoadedSlide.resolve(nativeObjectLoadedSlide.groups.items[0].nativeObjects.items[0].id)?.nativeKind, "contentPart");
const nativeObjectSecondPptx = await PresentationFile.exportPptx(nativeObjectLoaded);
const nativeObjectSecondInspect = await PresentationFile.inspectPptx(nativeObjectSecondPptx);
assert.equal(nativeObjectSecondInspect.ok, true);
const nativeObjectSecondZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectSecondPptx.arrayBuffer()));
const nativeObjectSecondSlideXml = await nativeObjectSecondZip.file("ppt/slides/slide1.xml").async("text");
const nativeObjectSecondRelsXml = await nativeObjectSecondZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(nativeObjectSecondSlideXml, /<p:oleObj[^>]*r:id="rId\d+"[\s\S]*?<a:blip r:embed="rId\d+"/);
assert.match(nativeObjectSecondSlideXml, /<dgm:relIds[^>]*r:dm="rId\d+"[^>]*r:lo="rId\d+"[^>]*r:qs="rId\d+"[^>]*r:cs="rId\d+"/);
assert.match(nativeObjectSecondSlideXml, /<p:contentPart[^>]*r:id="rId\d+"/);
assert.doesNotMatch(nativeObjectSecondSlideXml, /\/\s+xmlns:/);
assert.match(nativeObjectSecondRelsXml, /relationships\/package/);
assert.match(nativeObjectSecondRelsXml, /relationships\/diagramData/);
assert.match(nativeObjectSecondRelsXml, /relationships\/customXml/);
const preservedPreviewPath = Object.keys(nativeObjectSecondZip.files).find((file) => /^ppt\/preserved\/native\d+\/image1\.png$/.test(file));
assert.ok(preservedPreviewPath, "OLE preview that collides with a modeled image is remapped into the preserved namespace");
assert.deepEqual([...await nativeObjectSecondZip.file("ppt/embeddings/native-workbook.xlsx").async("uint8array")], [...embeddedWorkbookBytes]);
assert.equal(await nativeObjectSecondZip.file("ppt/customXml/native-content.xml").async("text"), '<native xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rIdPayload">preserve me</native>');
assert.match(await nativeObjectSecondZip.file("ppt/customXml/_rels/native-content.xml.rels").async("text"), /Target="assets\/payload\.svg"/);
assert.match(await nativeObjectSecondZip.file("[Content_Types].xml").async("text"), /native-workbook\.xlsx" ContentType="application\/vnd\.openxmlformats-officedocument\.spreadsheetml\.sheet"/);
nativeObjectLoaded.commentFormat = "modern";
nativeObjectLoadedSlide.comments.addThread(nativeObjectLoadedSlide.nativeObjects.items[1], "Review preserved SmartArt.", { author: "Native Reviewer" });
const nativeObjectCommentPptx = await PresentationFile.exportPptx(nativeObjectLoaded);
assert.equal((await PresentationFile.inspectPptx(nativeObjectCommentPptx)).ok, true);
const nativeObjectCommentZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectCommentPptx.arrayBuffer()));
assert.match(await nativeObjectCommentZip.file("ppt/comments/comment1.xml").async("text"), /<p188:unknownAnchor\/>/);
const relocatedNativeZip = await JSZip.loadAsync(new Uint8Array(await nativeObjectFixture.arrayBuffer()));
const relocatedNativeSlideXml = await relocatedNativeZip.file("ppt/slides/slide1.xml").async("text");
const relocatedNativeRelsXml = (await relocatedNativeZip.file("ppt/slides/_rels/slide1.xml.rels").async("text")).replaceAll('Target="../', 'Target="../../');
relocatedNativeZip.file("ppt/custom/slides/native.xml", relocatedNativeSlideXml);
relocatedNativeZip.file("ppt/custom/slides/_rels/native.xml.rels", relocatedNativeRelsXml);
relocatedNativeZip.remove("ppt/slides/slide1.xml");
relocatedNativeZip.remove("ppt/slides/_rels/slide1.xml.rels");
relocatedNativeZip.file("ppt/_rels/presentation.xml.rels", (await relocatedNativeZip.file("ppt/_rels/presentation.xml.rels").async("text")).replace('Target="slides/slide1.xml"', 'Target="custom/slides/native.xml"'));
relocatedNativeZip.file("[Content_Types].xml", (await relocatedNativeZip.file("[Content_Types].xml").async("text")).replace('/ppt/slides/slide1.xml', '/ppt/custom/slides/native.xml'));
const relocatedNativeFixture = new FileBlob(await relocatedNativeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: nativeObjectFixture.type });
const relocatedNativeLoaded = await PresentationFile.importPptx(relocatedNativeFixture);
assert.ok(presentationSlideNativeSourcePaths(relocatedNativeLoaded).every((sourcePart) => sourcePart === "ppt/custom/slides/native.xml"));
const relocatedNativeSecondZip = await JSZip.loadAsync(new Uint8Array(await (await PresentationFile.exportPptx(relocatedNativeLoaded)).arrayBuffer()));
assert.deepEqual([...await relocatedNativeSecondZip.file("ppt/embeddings/native-workbook.xlsx").async("uint8array")], [...embeddedWorkbookBytes]);

function presentationSlideNativeSourcePaths(deck) {
  const paths = [];
  for (const slide of deck.slides.items) {
    paths.push(...slide.nativeObjects.items.map((object) => object.sourcePart));
    for (const group of slide.groups.items) paths.push(...group.allElements().filter((element) => element.kind === "nativeObject").map((object) => object.sourcePart));
  }
  return paths;
}

const invalidModernModel = Presentation.create({ commentFormat: "modern" });
const invalidModernSlide = invalidModernModel.slides.add();
invalidModernSlide.comments.addThread(undefined, "Invalid identity", { comments: [{ nativeId: "not-a-guid", author: "Reviewer", text: "Invalid identity", created: "2026-07-13T01:02:03Z" }] });
assert.ok(invalidModernModel.verify().issues.some((issue) => issue.type === "invalidModernCommentMetadata"));
await assert.rejects(() => PresentationFile.exportPptx(invalidModernModel), /must be a brace-delimited GUID/);
console.log("presentation smoke ok");
