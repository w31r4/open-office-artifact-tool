import assert from "node:assert/strict";
import JSZip from "jszip";
import {
  DocumentModel,
  Presentation,
  Workbook,
} from "../src/index.mjs";
import {
  exportDocxWithOpenChestnut,
  exportPptxWithOpenChestnut,
  exportXlsxWithOpenChestnut,
  importDocxWithOpenChestnut,
  importPptxWithOpenChestnut,
  importXlsxWithOpenChestnut,
  openChestnutStatus,
} from "../src/codecs/open-chestnut.mjs";

const status = await openChestnutStatus();
assert.equal(status.available, true);
assert.equal(status.protocolVersion, 2);
assert.equal(status.assemblyName, "OpenChestnut.Runtime.dll");

const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";

// XLSX: create, import, edit, and re-export the canonical 0.2 slice.
const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Core");
sheet.getRange("A1:D3").values = [
  ["Date", "Category", "Value", "Status"],
  [new Date("2026-07-16T00:00:00.000Z"), "A", 8, "Ready"],
  [new Date("2026-07-17T00:00:00.000Z"), "B", 11, "Review"],
];
sheet.getRange("A1:D1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
sheet.getRange("A1:A3").format.numberFormat = "yyyy-mm-dd";
sheet.getRange("B1:B3").format.columnWidthPx = 120;
sheet.getRange("A2:D2").format.rowHeightPx = 28;
sheet.mergeCells("A5:B5");
sheet.getRange("A5").values = [["Merged"]];
sheet.freezePanes.freezeRows(1).freezeColumns(1);
sheet.tables.add({ range: "A1:D3", name: "CoreTable", style: "TableStyleMedium4" });
sheet.getRange("D2:D3").dataValidation = { rule: { type: "list", values: ["Ready", "Review", "Done"] } };
sheet.getRange("C2:C3").conditionalFormats.add("cellIs", { operator: "greaterThan", formula: 9, format: { fill: "#DCFCE7" } });
workbook.comments.setSelf({ displayName: "Analyst" });
workbook.comments.addThread({ cell: sheet.getRange("D2") }, "Canonical threaded comment");
sheet.images.add({ name: "Logo", dataUrl: png, alt: "One pixel logo", anchor: { from: { row: 6, col: 0 }, extent: { widthPx: 32, heightPx: 32 } } });
sheet.charts.add("bar", {
  name: "Values",
  title: "Values by category",
  categories: ["A", "B"],
  series: [{ name: "Value", values: [8, 11], fill: "#2563EB" }],
  position: { left: 260, top: 180, width: 360, height: 220 },
});

const xlsx = await exportXlsxWithOpenChestnut(workbook);
assert.equal(xlsx.metadata.codec, "open-chestnut");
assert.deepEqual([...xlsx.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
const importedWorkbook = await importXlsxWithOpenChestnut(xlsx);
const importedSheet = importedWorkbook.worksheets.getItem("Core");
assert.equal(importedSheet.getRange("B3").values[0][0], "B");
assert.ok(importedSheet.getRange("A2").values[0][0] > 40_000, "Date must cross the wire as an Excel serial");
assert.equal(importedSheet.getRange("D2:D3").dataValidation.type, "list");
assert.equal(importedSheet.conditionalFormattings.items.length, 1);
assert.equal(importedWorkbook.comments.threads.length, 1);
assert.equal(importedSheet.images.items.length, 1);
assert.equal(importedSheet.charts.items[0].type, "bar");
assert.equal(importedSheet.freezePanes.frozen, true);
importedSheet.getRange("C3").values = [[12]];
const xlsx2 = await exportXlsxWithOpenChestnut(importedWorkbook, { recalculate: false });
assert.equal((await importXlsxWithOpenChestnut(xlsx2)).worksheets.getItem("Core").getRange("C3").values[0][0], 12);
await assert.rejects(exportXlsxWithOpenChestnut(workbook, { allowLossy: true }), /does not accept option/i);

// DOCX: styles, run/paragraph formatting, section/header/footer, fields, image,
// table, list, link, and classic comment.
const document = DocumentModel.create({
  name: "Core document",
  blocks: [],
  defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#111827" },
});
document.styles.add("CoreHeading", { name: "Core Heading", basedOn: "Normal", fontSize: 22, bold: true, color: "#1D4ED8" });
const heading = document.addParagraph("Quarterly brief", {
  styleId: "CoreHeading",
  paragraphFormat: { alignment: "center", spaceAfterPt: 8 },
  runs: [{ text: "Quarterly ", style: { bold: true } }, { text: "brief", style: { italic: true, color: "#DC2626" } }],
});
const bodyParagraph = document.addParagraph("Canonical OpenChestnut document.");
document.addParagraph("Editable paragraph.");
document.addHeader("Confidential", { referenceType: "default", sectionIndex: 0 });
document.addFooter("Page ", { referenceType: "default", sectionIndex: 0, fieldInstruction: "PAGE" });
document.addField("PAGE", "1");
document.addHyperlink("Evidence", "https://example.com/evidence");
document.addListItem("First action", { listType: "number", numberingId: 7 });
document.addTable({ values: [["Metric", "Value"], ["Revenue", "42"]], widthDxa: 9000, columnWidthsDxa: [3600, 5400], styleId: "TableGrid" });
document.addImage({ name: "Logo", dataUrl: png, alt: "One pixel logo", widthPx: 32, heightPx: 32 });
document.addSection({ breakType: "nextPage", orientation: "landscape", pageSize: { widthTwips: 15840, heightTwips: 12240 }, margins: { top: 720, right: 720, bottom: 720, left: 720 } });
document.addComment(bodyParagraph, "Review body paragraph", { author: "Reviewer", initials: "RV", date: "2026-07-16T08:00:00Z" });

const docx = await exportDocxWithOpenChestnut(document);
assert.equal(docx.metadata.codec, "open-chestnut");
assert.deepEqual([...docx.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
const docxZip = await JSZip.loadAsync(docx.bytes);
assert.ok(docxZip.file("word/document.xml"));
assert.ok(docxZip.file("word/styles.xml"));
assert.ok(Object.keys(docxZip.files).some((part) => /(?:^|\/)media\/[^/]+\.png$/.test(part)));
const importedDocument = await importDocxWithOpenChestnut(docx);
assert.equal(importedDocument.defaultRunStyle.fontFamily, "Aptos");
assert.ok(importedDocument.styles.get("CoreHeading"));
assert.equal(importedDocument.blocks[0].text, "Quarterly brief");
assert.equal(importedDocument.blocks[0].runs[1].style.italic, true);
assert.equal(importedDocument.headers[0].text, "Confidential");
assert.equal(importedDocument.footers[0].fieldInstruction, "PAGE");
assert.equal(importedDocument.comments.length, 1);
importedDocument.blocks[2].text = "Edited through OpenChestnut.";
importedDocument.blocks[2].runs = [{ text: importedDocument.blocks[2].text, style: {} }];
const docx2 = await exportDocxWithOpenChestnut(importedDocument);
assert.equal((await importDocxWithOpenChestnut(docx2)).blocks[2].text, "Edited through OpenChestnut.");
await assert.rejects(exportDocxWithOpenChestnut(document, { allowLossy: true }), /does not accept option/i);

// PPTX: source-free roundRect/textbox, basic effect styling, connector arrows,
// and bar/line/pie charts, followed by a bounded second edit.
const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = presentation.slides.add({ name: "Core presentation" });
const roundedCard = slide.shapes.add({
  name: "Rounded card",
  geometry: "roundRect",
  position: { left: 48, top: 48, width: 260, height: 100 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  shadow: { color: "#000000", blurRadius: 8, distance: 4, direction: 45, opacity: 0.25 },
  text: "Rounded",
});
const textBox = slide.shapes.add({
  name: "Text box",
  geometry: "textbox",
  position: { left: 380, top: 48, width: 260, height: 100 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [{ bulletCharacter: "•", runs: [{ text: "Linked text", style: { bold: true }, link: { uri: "https://example.com" } }] }],
});
slide.connectors.add({
  name: "Elbow connector",
  connectorType: "elbow",
  from: roundedCard,
  to: textBox,
  line: { fill: "#334155", width: 2, startArrow: "triangle", endArrow: "triangle" },
});
slide.charts.add("bar", { name: "Revenue bars", position: { left: 48, top: 200, width: 340, height: 210 }, title: "Revenue", categories: ["Q1", "Q2"], series: [{ name: "Actual", values: [8, 11], color: "#2563EB" }], legend: false, dataLabels: { showValue: true, position: "outsideEnd" } });
slide.charts.add("line", { name: "Trend line", position: { left: 420, top: 200, width: 340, height: 210 }, title: "Trend", categories: ["Q1", "Q2"], series: [{ name: "Actual", values: [8, 11], color: "#16A34A", line: { fill: "#16A34A", width: 2, style: "dash" }, marker: { symbol: "circle", size: 7, fill: "#16A34A" } }], legend: false });
slide.charts.add("pie", { name: "Mix", position: { left: 790, top: 200, width: 340, height: 210 }, title: "Mix", categories: ["Direct", "Partner"], series: [{ name: "Share", values: [60, 40], color: "#7C3AED" }], legend: true, dataLabels: { showCategoryName: true, showValue: true, position: "bestFit" } });

const pptx = await exportPptxWithOpenChestnut(presentation);
assert.equal(pptx.metadata.codec, "open-chestnut");
assert.deepEqual([...pptx.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
const importedPresentation = await importPptxWithOpenChestnut(pptx);
const importedSlide = importedPresentation.slides.getItem(0);
const importedRounded = importedSlide.shapes.items.find((shape) => shape.name === "Rounded card");
assert.equal(importedRounded.geometry, "roundRect");
assert.deepEqual(importedRounded.shadow, { color: "#000000", blurRadius: 8, distance: 4, direction: 45, opacity: 0.25 });
assert.equal(importedSlide.shapes.items.find((shape) => shape.name === "Text box").geometry, "textbox");
assert.equal(importedSlide.connectors.items[0].connectorType, "elbow");
assert.equal(importedSlide.connectors.items[0].line.startArrow, "triangle");
assert.deepEqual(importedSlide.charts.items.map((chart) => chart.chartType), ["bar", "line", "pie"]);
importedRounded.shadow.opacity = 0.35;
importedSlide.connectors.items[0].line.endArrow = undefined;
importedSlide.charts.items[0].title = "Updated revenue";
importedSlide.charts.items[0].series[0].values[1] = 12;
const pptx2 = await exportPptxWithOpenChestnut(importedPresentation);
const editedPresentation = await importPptxWithOpenChestnut(pptx2);
assert.equal(editedPresentation.slides.items[0].shapes.items.find((shape) => shape.name === "Rounded card").shadow.opacity, 0.35);
assert.equal(editedPresentation.slides.items[0].connectors.items[0].line.endArrow, undefined);
assert.equal(editedPresentation.slides.items[0].charts.items[0].title, "Updated revenue");
assert.deepEqual(editedPresentation.slides.items[0].charts.items[0].series[0].values, [8, 12]);
await assert.rejects(exportPptxWithOpenChestnut(presentation, { allowLossy: true }), /does not accept option/i);

console.log("OpenChestnut protocol 2 canonical core smoke ok");
