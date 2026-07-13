import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, Presentation, PresentationFile, Workbook, SpreadsheetFile } from "../src/index.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";
import {
  OpenXmlWasmCodecError,
  exportDocxWithOpenXmlWasm,
  exportPptxWithOpenXmlWasm,
  exportXlsxWithOpenXmlWasm,
  importDocxWithOpenXmlWasm,
  importPptxWithOpenXmlWasm,
  importXlsxWithOpenXmlWasm,
  openXmlWasmStatus,
} from "../src/codecs/openxml-wasm.mjs";

const workbook = Workbook.create({ dateSystem: "1904" });
const summary = workbook.worksheets.add("Summary");
summary.getRange("A1:B2").values = [["Quarter", 42.5], [true, null]];
summary.getRange("B2").formulas = [["=B1*2"]];
summary.freezePanes.freezeRows(1).freezeColumns(1);
summary.showGridLines = false;
summary.columnDimensions.set(0, { width: 18, bestFit: true });
summary.rowDimensions.set(0, { height: 24 });
summary.mergeCells("A3:B3");
const details = workbook.worksheets.add("Details");
details.getRange("A1:B1").values = [["Status", "ready"]];

const concurrentWorkbook = Workbook.create();
concurrentWorkbook.worksheets.add("Concurrent").getRange("A1").values = [["cached runtime"]];
const [status, exported, concurrentExport] = await Promise.all([
  openXmlWasmStatus(),
  exportXlsxWithOpenXmlWasm(workbook),
  exportXlsxWithOpenXmlWasm(concurrentWorkbook),
]);
assert.deepEqual([...concurrentExport.bytes.slice(0, 2)], [0x50, 0x4b]);
assert.deepEqual([...exported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(exported.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
assert.equal(exported.metadata.codec, "openxml-wasm");
assert.equal((await SpreadsheetFile.inspectXlsx(exported)).ok, true);

const imported = await importXlsxWithOpenXmlWasm(exported);
assert.equal(imported.dateSystem, "1904");
assert.equal(imported.worksheets.items.length, 2);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(imported.worksheets.getItem("Summary").freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
assert.equal(imported.worksheets.getItem("Summary").showGridLines, false);
assert.equal(imported.worksheets.getItem("Summary").columnDimensions.get(0).width, 18);
assert.deepEqual(imported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);
assert.match(imported.inspect({ kind: "workbook,sheet,formula" }).ndjson, /"dateSystem":"1904"/);
assert.equal(imported.verify().ok, true);
assert.equal(imported.resolve(imported.worksheets.getItem("Summary").id).name, "Summary");

// Open XML SDK serializes SpreadsheetML with a legal namespace prefix. Keep the
// JavaScript migration oracle able to read the same package while both codecs
// coexist, so cross-codec fixtures compare semantics instead of XML spelling.
const javascriptImported = await SpreadsheetFile.importXlsx(exported);
assert.equal(javascriptImported.dateSystem, "1904");
assert.equal(javascriptImported.worksheets.items.length, 2);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);

const secondExport = await exportXlsxWithOpenXmlWasm(imported, { recalculate: false });
assert.deepEqual([...secondExport.bytes.slice(0, 2)], [0x50, 0x4b]);

const externalZip = await JSZip.loadAsync(exported.bytes);
const relationshipPath = "xl/_rels/workbook.xml.rels";
const relationships = await externalZip.file(relationshipPath).async("text");
externalZip.file(relationshipPath, relationships.replace("</Relationships>", '<Relationship Id="rIdExternal" Type="urn:open-office-artifact-tool:test" Target="https://example.invalid/data" TargetMode="External"/></Relationships>'));
const externalBytes = await externalZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const opaqueImported = await importXlsxWithOpenXmlWasm(externalBytes);
opaqueImported.worksheets.getItem("Summary").getRange("B1").values = [[99]];
const preserved = await exportXlsxWithOpenXmlWasm(opaqueImported, { recalculate: false });
assert.equal(preserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const preservedZip = await JSZip.loadAsync(preserved.bytes);
assert.match(await preservedZip.file(relationshipPath).async("text"), /Id="rIdExternal"[^>]*Target="https:\/\/example\.invalid\/data"/);
const preservedImported = await importXlsxWithOpenXmlWasm(preserved);
assert.equal(preservedImported.worksheets.getItem("Summary").getRange("B1").values[0][0], 99);

await assert.rejects(
  importXlsxWithOpenXmlWasm(exported, { limits: { maxInputBytes: 16 } }),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "input_budget_exceeded",
);

const styled = Workbook.create();
const styledSheet = styled.worksheets.add("Sheet1");
styledSheet.getRange("A1:B2").values = [["Label", "Value"], ["styled", 1]];
styledSheet.getRange("A1:B1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
styledSheet.tables.add("A1:B2", true, "StyledTable").style = "TableStyleMedium4";
await assert.rejects(
  exportXlsxWithOpenXmlWasm(styled),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_workbook_features",
);
const styledSource = await SpreadsheetFile.exportXlsx(styled);
const styledImported = await importXlsxWithOpenXmlWasm(styledSource);
styledImported.worksheets.getItem("Sheet1").getRange("B2").values = [[2]];
const styledPreserved = await exportXlsxWithOpenXmlWasm(styledImported, { recalculate: false });
const styledRoundTrip = await SpreadsheetFile.importXlsx(styledPreserved);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("B2").values[0][0], 2);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("A1").format.font.bold, true);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").tables.items[0].name, "StyledTable");

const minimalDocument = DocumentModel.create({
  name: "OpenXML WASM brief",
  blocks: [
    { kind: "paragraph", text: "Quarterly brief", styleId: "Normal", runs: [{ text: "Quarterly ", style: { bold: true } }, { text: "brief", style: { italic: true } }] },
    { kind: "table", styleId: "TableGrid", values: [["Revenue", "42"], ["Status", "Ready"]] },
  ],
});
const docxExported = await exportDocxWithOpenXmlWasm(minimalDocument);
assert.deepEqual([...docxExported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(docxExported.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
assert.equal(docxExported.metadata.codec, "openxml-wasm");
assert.equal((await DocumentFile.inspectDocx(docxExported)).ok, true);
const docxImported = await importDocxWithOpenXmlWasm(docxExported);
assert.equal(docxImported.name, "Imported document");
assert.equal(docxImported.blocks[0].text, "Quarterly brief");
assert.equal(docxImported.blocks[0].runs[0].style.bold, true);
assert.deepEqual(docxImported.blocks[1].values, [["Revenue", "42"], ["Status", "Ready"]]);
assert.equal(docxImported.verify().ok, true);

const richDocument = DocumentModel.create({ name: "Source preservation", blocks: [] });
richDocument.addParagraph("Editable lead", { styleId: "Normal" });
richDocument.addHyperlink("Preserved link", "https://example.invalid/source");
richDocument.addHeader("Preserved header", { sectionIndex: 0, referenceType: "default" });
const richSource = await DocumentFile.exportDocx(richDocument);
const richImported = await importDocxWithOpenXmlWasm(richSource);
assert.equal(richImported.blocks[0].text, "Editable lead");
assert.equal(richImported.blocks[1].text, "Preserved link");
richImported.blocks[0].text = "Edited lead";
richImported.blocks[0].runs[0].text = "Edited lead";
const richPreserved = await exportDocxWithOpenXmlWasm(richImported);
assert.equal(richPreserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const richRoundTrip = await DocumentFile.importDocx(richPreserved, { preferNative: true });
assert.equal(richRoundTrip.blocks[0].text, "Edited lead");
assert.equal(richRoundTrip.blocks.some((block) => block.kind === "hyperlink" && block.text === "Preserved link"), true);
assert.equal(richRoundTrip.headers.some((header) => header.text === "Preserved header"), true);
richImported.blocks[1].text = "Unsafe hyperlink edit";
richImported.blocks[1].runs[0].text = "Unsafe hyperlink edit";
await assert.rejects(
  exportDocxWithOpenXmlWasm(richImported),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_document_edit",
);

const unsupportedDocument = DocumentModel.create({ blocks: [] });
unsupportedDocument.addHyperlink("Unsupported direct authoring", "https://example.invalid/direct");
await assert.rejects(
  exportDocxWithOpenXmlWasm(unsupportedDocument),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_document_features",
);

const minimalPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
minimalPresentation.slides.add({ name: "Overview" }).shapes.add({
  name: "Title",
  geometry: "rect",
  position: { left: 60, top: 40, width: 860, height: 70 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: "OpenXML WASM presentation",
});
const pptxExported = await exportPptxWithOpenXmlWasm(minimalPresentation);
assert.deepEqual([...pptxExported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(pptxExported.type, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
assert.equal(pptxExported.metadata.codec, "openxml-wasm");
assert.equal((await PresentationFile.inspectPptx(pptxExported)).ok, true);
const pptxImported = await importPptxWithOpenXmlWasm(pptxExported);
assert.equal(pptxImported.slides.count, 1);
assert.equal(pptxImported.slides.getItem(0).shapes.items[0].text.value, "OpenXML WASM presentation");
assert.equal(pptxImported.verify().ok, true);

const richPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const richShape = richPresentation.slides.add({ name: "Rich text" }).shapes.add({
  name: "Rich text",
  geometry: "rect",
  position: { left: 60, top: 40, width: 920, height: 180 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [
    {
      alignment: "center",
      bulletCharacter: "•",
      bulletFont: "Georgia",
      bulletColor: "#DC2626",
      bulletSizePercent: 1.5,
      runs: [
        { text: "Quarterly ", style: { bold: true, fontSize: 36, fontFamily: "Aptos Display", color: "#0F172A" } },
        { text: "brief", style: { italic: true, fontSize: 36 } },
      ],
    },
    { level: 1, autoNumber: { type: "romanLcPeriod", startAt: 3 }, bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true, runs: [{ text: "Source-bound detail", style: { fontSize: 20, color: "#475569" } }] },
    { bulletNone: true, bulletSize: 24, runs: [{ text: "Explicitly unbulleted", style: { fontSize: 20 } }] },
  ],
});
const richPptx = await exportPptxWithOpenXmlWasm(richPresentation);
const richPptxImported = await importPptxWithOpenXmlWasm(richPptx);
const richImportedShape = richPptxImported.slides.getItem(0).shapes.items[0];
assert.equal(richImportedShape.text.value, "Quarterly brief\nSource-bound detail\nExplicitly unbulleted");
assert.equal(richImportedShape.text.paragraphs.length, 3);
assert.equal(richImportedShape.text.paragraphs[0].alignment, "center");
assert.equal(richImportedShape.text.paragraphs[0].bulletCharacter, "•");
assert.equal(richImportedShape.text.paragraphs[0].bulletFont, "Georgia");
assert.equal(richImportedShape.text.paragraphs[0].bulletColor, "#DC2626");
assert.equal(richImportedShape.text.paragraphs[0].bulletSizePercent, 1.5);
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.bold, true);
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.fontSize, 36);
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.fontFamily, "Aptos Display");
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.color, "#0F172A");
assert.equal(richImportedShape.text.paragraphs[0].runs[1].style.italic, true);
assert.deepEqual(richImportedShape.text.paragraphs[1].autoNumber, { type: "romanLcPeriod", startAt: 3 });
assert.equal(richImportedShape.text.paragraphs[1].bulletFontFollowText, true);
assert.equal(richImportedShape.text.paragraphs[1].bulletColorFollowText, true);
assert.equal(richImportedShape.text.paragraphs[1].bulletSizeFollowText, true);
assert.equal(richImportedShape.text.paragraphs[2].bulletNone, true);
assert.equal(richImportedShape.text.paragraphs[2].bulletSize, 24);
richImportedShape.text.paragraphs = richImportedShape.text.paragraphs.map((paragraph, paragraphIndex) => ({
  ...paragraph,
  ...(paragraphIndex === 0 ? { bulletCharacter: "◆", bulletFont: undefined, bulletFontFollowText: true, bulletColor: "#2563EB", bulletSizePercent: undefined, bulletSize: 24 } : {}),
  ...(paragraphIndex === 1 ? { autoNumber: { type: "arabicPeriod", startAt: 5 }, bulletFontFollowText: undefined, bulletFont: "Aptos", bulletColorFollowText: undefined, bulletColor: "#16A34A", bulletSizeFollowText: undefined, bulletSizePercent: 1.25 } : {}),
  ...(paragraphIndex === 2 ? { bulletNone: undefined, bulletCharacter: "–", bulletSize: undefined, bulletSizeFollowText: true } : {}),
  runs: paragraph.runs.map((run, runIndex) => paragraphIndex === 0 && runIndex === 0
    ? { ...run, text: "Updated ", style: { ...run.style, bold: false, color: "#2563EB" } }
    : run),
}));
const richPptxEdited = await exportPptxWithOpenXmlWasm(richPptxImported);
const richPptxRoundTrip = await importPptxWithOpenXmlWasm(richPptxEdited);
const richRoundTripShape = richPptxRoundTrip.slides.getItem(0).shapes.items[0];
assert.equal(richRoundTripShape.text.value, "Updated brief\nSource-bound detail\nExplicitly unbulleted");
assert.equal(richRoundTripShape.text.paragraphs[0].runs[0].style.bold, false);
assert.equal(richRoundTripShape.text.paragraphs[0].runs[0].style.color, "#2563EB");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletCharacter, "◆");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletFontFollowText, true);
assert.equal(richRoundTripShape.text.paragraphs[0].bulletColor, "#2563EB");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletSize, 24);
assert.deepEqual(richRoundTripShape.text.paragraphs[1].autoNumber, { type: "arabicPeriod", startAt: 5 });
assert.equal(richRoundTripShape.text.paragraphs[1].bulletFont, "Aptos");
assert.equal(richRoundTripShape.text.paragraphs[1].bulletColor, "#16A34A");
assert.equal(richRoundTripShape.text.paragraphs[1].bulletSizePercent, 1.25);
assert.equal(richRoundTripShape.text.paragraphs[2].bulletCharacter, "–");
assert.equal(richRoundTripShape.text.paragraphs[2].bulletSizeFollowText, true);

const hyperlinkPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const hyperlinkSlide = hyperlinkPresentation.slides.add({ name: "Links" });
const hyperlinkTarget = hyperlinkPresentation.slides.add({ name: "Details" });
const hyperlinkAppendix = hyperlinkPresentation.slides.add({ name: "Appendix" });
hyperlinkSlide.shapes.add({
  name: "Run links",
  position: { left: 60, top: 40, width: 920, height: 120 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [{ runs: [
    { text: "Guide ", link: { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank", history: false, highlightClick: true } },
    { text: "Details ", link: { slideId: hyperlinkTarget.id } },
    { text: "Next ", link: { action: "nextSlide" } },
    { text: "End", link: { action: "endShow" } },
  ] }],
});
const hyperlinkPptx = await exportPptxWithOpenXmlWasm(hyperlinkPresentation);
const hyperlinkZip = await JSZip.loadAsync(hyperlinkPptx.bytes);
const hyperlinkSlideXml = await hyperlinkZip.file("ppt/slides/slide1.xml").async("text");
const hyperlinkRelsXml = await hyperlinkZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(hyperlinkSlideXml, /<a:hlinkClick[^>]*tgtFrame="_blank"[^>]*tooltip="Read the guide"[^>]*history="0"[^>]*highlightClick="1"/);
assert.match(hyperlinkSlideXml, /action="ppaction:\/\/hlinksldjump"/);
assert.match(hyperlinkSlideXml, /action="ppaction:\/\/hlinkshowjump\?jump=nextslide"/);
assert.match(hyperlinkRelsXml, /relationships\/hyperlink[^>]*Target="https:\/\/example\.com\/guide\?x=1&amp;y=2"[^>]*TargetMode="External"/);
assert.match(hyperlinkRelsXml, /relationships\/slide[^>]*Target="(?:\/ppt\/slides\/)?slide2\.xml"/);
const hyperlinkImported = await importPptxWithOpenXmlWasm(hyperlinkPptx);
const hyperlinkImportedShape = hyperlinkImported.slides.getItem(0).shapes.items[0];
const hyperlinkRuns = hyperlinkImportedShape.text.paragraphs[0].runs;
assert.deepEqual(hyperlinkRuns[0].link, { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank", history: false, highlightClick: true });
assert.deepEqual(hyperlinkRuns[1].link, { slideId: hyperlinkImported.slides.getItem(1).id });
assert.deepEqual(hyperlinkRuns[2].link, { action: "nextSlide" });
hyperlinkImportedShape.text.paragraphs = hyperlinkImportedShape.text.paragraphs.map((paragraph) => ({
  ...paragraph,
  runs: paragraph.runs.map((run, index) => ({
    ...run,
    link: index === 0
      ? { uri: "https://example.com/updated", targetFrame: "_self" }
      : index === 1
        ? { slideId: hyperlinkImported.slides.getItem(2).id, tooltip: "Appendix" }
        : index === 2
          ? { action: "lastSlide", highlightClick: false }
          : undefined,
  })),
}));
const hyperlinkEdited = await exportPptxWithOpenXmlWasm(hyperlinkImported);
const hyperlinkRoundTrip = await importPptxWithOpenXmlWasm(hyperlinkEdited);
const hyperlinkRoundTripRuns = hyperlinkRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].runs;
assert.deepEqual(hyperlinkRoundTripRuns[0].link, { uri: "https://example.com/updated", targetFrame: "_self" });
assert.deepEqual(hyperlinkRoundTripRuns[1].link, { slideId: hyperlinkRoundTrip.slides.getItem(2).id, tooltip: "Appendix" });
assert.deepEqual(hyperlinkRoundTripRuns[2].link, { action: "lastSlide", highlightClick: false });
assert.equal(hyperlinkRoundTripRuns[3].link, undefined);

const unsupportedHyperlinkPresentation = Presentation.create();
unsupportedHyperlinkPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "Tour", link: { customShow: "Evidence" } }] }] });
await assert.rejects(
  exportPptxWithOpenXmlWasm(unsupportedHyperlinkPresentation),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_features" && /custom-show hyperlink/.test(error.message),
);

richImportedShape.text.paragraphs = [
  ...richImportedShape.text.paragraphs.slice(0, 1),
  { ...richImportedShape.text.paragraphs[1], runs: [...richImportedShape.text.paragraphs[1].runs, { text: "unsafe", style: {} }] },
];
await assert.rejects(
  exportPptxWithOpenXmlWasm(richPptxImported),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "presentation_text_topology_changed",
);

const unsupportedRichPresentation = Presentation.create();
unsupportedRichPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "underline", style: { underline: "single" } }] }] });
await assert.rejects(
  exportPptxWithOpenXmlWasm(unsupportedRichPresentation),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_features",
);

const invalidRichPresentation = Presentation.create();
invalidRichPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "transparent", style: { color: "transparent" } }] }] });
await assert.rejects(
  exportPptxWithOpenXmlWasm(invalidRichPresentation),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_features",
);

const invalidBulletPresentation = Presentation.create();
const invalidBulletShape = invalidBulletPresentation.slides.add().shapes.add({ text: [{ autoNumber: "arabicPeriod", runs: ["invalid"] }] });
invalidBulletShape.text._paragraphs[0].autoNumber = { type: "not-a-scheme" };
await assert.rejects(
  exportPptxWithOpenXmlWasm(invalidBulletPresentation),
  (error) => error instanceof RangeError && /auto-number type/.test(error.message),
);

const unsupportedBulletColorPresentation = Presentation.create();
unsupportedBulletColorPresentation.slides.add().shapes.add({ text: [{ bulletCharacter: "•", bulletColor: "accent1", runs: ["styled"] }] });
await assert.rejects(
  exportPptxWithOpenXmlWasm(unsupportedBulletColorPresentation),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_features",
);

const preservedPresentation = Presentation.create({
  master: { id: "master/preservation", name: "Preservation master", background: "#FFFFFF", placeholders: [] },
});
const preservedSlide = preservedPresentation.slides.add({ name: "Opaque graph" });
preservedSlide.shapes.add({ name: "Editable title", position: { left: 40, top: 32, width: 720, height: 64 }, fill: "#FFFFFF", line: { fill: "#334155", width: 1 }, text: "Before WASM" });
preservedSlide.images.add({
  name: "Preserved image",
  alt: "Opaque image evidence",
  position: { left: 40, top: 140, width: 180, height: 120 },
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
});
preservedSlide.charts.add("bar", {
  name: "Preserved chart",
  position: { left: 300, top: 140, width: 560, height: 360 },
  categories: ["Model", "Package"],
  series: [{ name: "Evidence", values: [8, 12] }],
});
const presentationSource = await PresentationFile.exportPptx(preservedPresentation);
const presentationImported = await importPptxWithOpenXmlWasm(presentationSource);
assert.equal(presentationImported.slides.getItem(0).nativeObjects.items.length, 2);
presentationImported.slides.getItem(0).shapes.items[0].text.set("After WASM");
const presentationPreserved = await exportPptxWithOpenXmlWasm(presentationImported);
assert.equal(presentationPreserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const presentationRoundTrip = await PresentationFile.importPptx(presentationPreserved);
assert.equal(presentationRoundTrip.slides.getItem(0).shapes.items[0].text.value, "After WASM");
assert.equal(presentationRoundTrip.slides.getItem(0).images.items.length, 1);
assert.equal(presentationRoundTrip.slides.getItem(0).charts.items.length, 1);
presentationImported.slides.getItem(0).nativeObjects.items[0].name = "Unsafe native edit";
await assert.rejects(
  exportPptxWithOpenXmlWasm(presentationImported),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_edit",
);

const unsupportedPresentation = Presentation.create();
unsupportedPresentation.slides.add().images.add({ prompt: "Unsupported direct image authoring" });
await assert.rejects(
  exportPptxWithOpenXmlWasm(unsupportedPresentation),
  (error) => error instanceof OpenXmlWasmCodecError && error.code === "unsupported_presentation_features",
);

assert.equal(status.available, true);
assert.equal(status.protocolVersion, 1);
assert.equal(status.assemblyName, "OpenOffice.OpenXmlWasm.dll");
assert.ok(status.manifest.totalBytes > 1_000_000);
assert.equal(status.manifest.files.some((file) => file.path.endsWith(".map") || file.path.endsWith(".symbols")), false);

const sofficeStatus = spawnSync("soffice", ["--version"], { encoding: "utf8" });
const popplerStatus = spawnSync("pdftoppm", ["-v"], { encoding: "utf8" });
if (sofficeStatus.status === 0 && popplerStatus.status === 0) {
  const pdf = await createLibreOfficeRenderer({ timeoutMs: 60_000 })({
    input: exported,
    inputType: exported.type,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "workbook",
  });
  assert.equal(new TextDecoder().decode(pdf.bytes.slice(0, 4)), "%PDF");
  const png = await createPopplerRenderer({ dpi: 96, timeoutMs: 60_000 })({
    input: pdf,
    inputType: pdf.type,
    outputType: "image/png",
    format: "png",
    artifactKind: "workbook",
    pageIndex: 0,
  });
  assert.deepEqual([...png.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
}

const noDotnetProbe = `
  import { Workbook } from ${JSON.stringify(new URL("../src/index.mjs", import.meta.url).href)};
  import { exportXlsxWithOpenXmlWasm } from ${JSON.stringify(new URL("../src/codecs/openxml-wasm.mjs", import.meta.url).href)};
  const workbook = Workbook.create();
  workbook.worksheets.add("Sheet1").getRange("A1").values = [["no dotnet on PATH"]];
  const file = await exportXlsxWithOpenXmlWasm(workbook);
  if (file.bytes[0] !== 0x50 || file.bytes[1] !== 0x4b) process.exit(1);
`;
const child = spawnSync(process.execPath, ["--input-type=module", "-e", noDotnetProbe], {
  cwd: new URL("..", import.meta.url),
  encoding: "utf8",
  env: { ...process.env, PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin" },
});
assert.equal(child.status, 0, `bundled runtime failed without dotnet on PATH\nSTDOUT:\n${child.stdout}\nSTDERR:\n${child.stderr}`);

console.log("openxml wasm smoke ok");
