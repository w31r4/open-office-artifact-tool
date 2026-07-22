import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import JSZip from "jszip";

import { DocumentFile, DocumentModel } from "../src/index.mjs";
import { DocumentFile as DocumentFileModule, DocumentModel as DocumentModelModule } from "../src/document/index.mjs";

async function changedZipParts(left, right) {
  const hashes = async (value) => {
    const zip = await JSZip.loadAsync(await value.arrayBuffer());
    const result = new Map();
    for (const [name, entry] of Object.entries(zip.files)) {
      if (entry.dir) continue;
      result.set(name, createHash("sha256").update(await entry.async("uint8array")).digest("hex"));
    }
    return result;
  };
  const [before, after] = await Promise.all([hashes(left), hashes(right)]);
  return [...new Set([...before.keys(), ...after.keys()])]
    .filter((name) => before.get(name) !== after.get(name))
    .sort();
}

assert.strictEqual(DocumentModel, DocumentModelModule);
assert.strictEqual(DocumentFile, DocumentFileModule);

const textRangeFixture = DocumentModel.create({ name: "Text range formatting guard" });
const singleRunRangeTarget = textRangeFixture.addParagraph("Original", {
  runs: [{ text: "Original", style: { bold: true, color: "#315A83" } }],
});
const singleRunStyle = structuredClone(singleRunRangeTarget.runs[0].style);
textRangeFixture.resolve(`${singleRunRangeTarget.id}/text`).text = "Replacement";
assert.equal(singleRunRangeTarget.text, "Replacement");
assert.equal(singleRunRangeTarget.runs[0].text, "Replacement");
assert.deepEqual(singleRunRangeTarget.runs[0].style, singleRunStyle);
const multiRunRangeTarget = textRangeFixture.addParagraph("Two runs", {
  runs: [{ text: "Two ", style: { bold: true } }, { text: "runs", style: { italic: true } }],
});
assert.throws(
  () => { textRangeFixture.resolve(`${multiRunRangeTarget.id}/text`).text = "Flattened"; },
  /multiple source runs.*formatting boundaries/i,
);
const sourceBoundPatchTarget = textRangeFixture.addParagraph("Patch target", {
  textEditable: false,
  textPatchable: true,
});
assert.throws(
  () => sourceBoundPatchTarget.replaceText("target", "\ud800"),
  /XML-safe strings/i,
);

const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
const jpeg = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAACAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAABQf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCQAHTd/9k=";

const floatingPlacement = {
  type: "floating",
  horizontal: { relativeTo: "margin", offsetPx: 240 },
  vertical: { relativeTo: "margin", offsetPx: 24 },
  wrap: "square",
  wrapSide: "right",
  distanceFromTextPx: { top: 2, right: 12, bottom: 6, left: 12 },
};
const floatingDocument = DocumentModel.create({ name: "Floating image model", blocks: [] });
floatingDocument.addParagraph("Text before the floating figure.");
const floatingImage = floatingDocument.addImage({
  name: "floating-figure",
  dataUrl: png,
  alt: "Floating figure",
  widthPx: 120,
  heightPx: 80,
  placement: floatingPlacement,
});
floatingDocument.addParagraph("Text after the floating figure.");
assert.deepEqual(floatingImage.placement, floatingPlacement);
assert.deepEqual(floatingImage.toProto().placement, floatingPlacement);
const floatingLayout = floatingDocument.layoutJson();
const floatingLayoutElement = floatingLayout.elements.find((element) => element.id === floatingImage.id);
assert.deepEqual(floatingLayoutElement.bbox, [312, 96, 120, 80]);
assert.equal(floatingLayoutElement.placement.wrapSide, "right");
const floatingInspect = floatingDocument.inspect({ kind: "image,layout", maxChars: 12_000 }).ndjson
  .trim().split("\n").map((line) => JSON.parse(line));
assert.equal(floatingInspect.find((record) => record.kind === "image")?.placement.horizontal.relativeTo, "margin");
const floatingPreview = await floatingDocument.render();
assert.match(await floatingPreview.text(), /<image[^>]*x="312"[^>]*y="96"[^>]*width="120"[^>]*height="80"[^>]*aria-label="Floating figure"/);
assert.equal(floatingDocument.verify({ visualQa: true }).ok, true);

const validFloatingImage = { dataUrl: png, widthPx: 40, heightPx: 30, placement: floatingPlacement };
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { ...floatingPlacement, zIndex: 4 } }), /unsupported field zIndex/i);
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { ...floatingPlacement, horizontal: { relativeTo: "margin" } } }), /offsetPx is required/i);
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { ...floatingPlacement, horizontal: { relativeTo: "margin", offsetPx: "24" } } }), /must be a number/i);
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { ...floatingPlacement, horizontal: { relativeTo: "margin", offsetPx: 10_001 } } }), /finite pixel value/i);
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { ...floatingPlacement, wrap: "topAndBottom", wrapSide: "left" } }), /cannot specify wrapSide/i);
assert.throws(() => floatingDocument.addImage({ ...validFloatingImage, placement: { type: "inline", wrap: "square" } }), /cannot carry floating-image fields/i);
floatingImage.placement.horizontal.offsetPx = Number.POSITIVE_INFINITY;
assert.equal(floatingDocument.verify().issues.some((issue) => issue.type === "invalidImagePlacement"), true);
floatingImage.placement.horizontal.offsetPx = 240;

const watermarkModel = DocumentModel.create({ name: "Watermark model", blocks: [] });
watermarkModel.addParagraph("Draft review body");
const modelWatermark = watermarkModel.addWatermark("DRAFT", { id: "watermark/model", sectionIndex: 0 });
assert.equal(watermarkModel.resolve(modelWatermark.id), modelWatermark);
assert.deepEqual(modelWatermark.toProto(), {
  kind: "watermark",
  id: "watermark/model",
  text: "DRAFT",
  referenceType: "default",
  sectionIndex: 0,
  editable: true,
  sourceBound: false,
});
assert.match(watermarkModel.inspect({ kind: "document,watermark" }).ndjson, /"kind":"watermark".*"text":"DRAFT"/);
assert.match(await (await watermarkModel.render()).text(), /rotate\(-45 306 396\).*DRAFT/);
assert.equal(watermarkModel.verify().ok, true);
assert.throws(() => watermarkModel.addWatermark("SECOND", { sectionIndex: 0 }), /already has a default text watermark/i);
assert.throws(() => DocumentModel.create({ blocks: [], watermarks: [{ text: "   " }] }), /cannot be blank/i);
assert.throws(() => DocumentModel.create({ blocks: [], watermarks: [{ text: "DRAFT", referenceType: "odd" }] }), /must be default, first, or even/i);
const evenWatermarkModel = DocumentModel.create({ name: "Even watermark activation", blocks: [] });
evenWatermarkModel.addWatermark("EVEN REVIEW", { referenceType: "even", sectionIndex: 0 });
assert.equal(evenWatermarkModel.settings.evenAndOddHeaders, true);
modelWatermark.remove();
assert.equal(watermarkModel.watermarks.length, 0);
assert.throws(() => modelWatermark.remove(), /no longer attached/i);

const document = DocumentModel.create({
  name: "OpenChestnut document profile",
  defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#202020" },
  blocks: [
    { kind: "paragraph", name: "title", styleId: "Title", text: "OpenChestnut document profile" },
  ],
});
document.styles.add("TableGrid", { name: "Table Grid", type: "table" });
document.styles.add("BodyAccent", {
  name: "Body Accent",
  type: "paragraph",
  basedOn: "Normal",
  fontFamily: "Aptos",
  fontSize: 13,
  color: "#315A83",
  bold: true,
  alignment: "left",
  spaceAfterTwips: 240,
  keepNext: true,
});

const commentTarget = document.addParagraph("Review this paragraph before release.", {
  name: "comment-target",
  styleId: "Normal",
});
const formatted = document.addParagraph("Bold and colored", {
  name: "formatted-paragraph",
  styleId: "BodyAccent",
  paragraphFormat: {
    alignment: "center",
    leftIndentTwips: 360,
    spaceBeforeTwips: 120,
    spaceAfterTwips: 240,
    lineSpacingTwips: 300,
    lineSpacingRule: "auto",
    keepNext: true,
  },
  runs: [
    { text: "Bold ", style: { bold: true, fontFamily: "Aptos Display", fontSize: 15 } },
    { text: "and colored", style: { italic: true, underline: true, color: "#CC0000", characterSpacingTwips: 10 } },
  ],
});
const contentControlParagraph = document.addParagraph("", {
  name: "customer-template-field",
  styleId: "Normal",
  runs: [
    { text: "Customer: " },
    { text: "Ada Lovelace", contentControl: { id: "customer-name-control", tag: "CUSTOMER_NAME", alias: "Customer name" } },
    { text: "." },
  ],
});
const bullet = document.addListItem("Inspect the semantic model.", {
  name: "bullet-item",
  styleId: "Normal",
  listType: "bullet",
  numberFormat: "bullet",
  start: 1,
  levelText: "•",
  numberingId: 71,
  abstractNumberingId: 7,
});
const numbered = document.addListItem("Render and verify the DOCX.", {
  name: "numbered-item",
  styleId: "Normal",
  listType: "number",
  numberFormat: "decimal",
  start: 1,
  levelText: "%1.",
  numberingId: 72,
  abstractNumberingId: 8,
});
const table = document.addTable({
  name: "readiness-table",
  styleId: "TableGrid",
  widthDxa: 9000,
  indentDxa: 120,
  columnWidthsDxa: [4500, 4500],
  cellMarginsDxa: { top: 60, right: 100, bottom: 60, left: 100 },
  borderColor: "445566",
  borderSize: 8,
  headerFill: "E2E8F0",
  values: [["Gate", "Status"], ["Semantic", "Pending"], ["Visual", "Required"]],
});
const hyperlink = document.addHyperlink(
  "Open XML SDK documentation",
  "https://learn.microsoft.com/office/open-xml/open-xml-sdk",
  { name: "external-link", styleId: "Normal", tooltip: "Open documentation", history: true },
);
const field = document.addField("PAGE", "1", { name: "page-field", styleId: "Normal" });
const insertion = document.addInsertion("Added wording", { name: "tracked-insertion", styleId: "Normal", author: "Reviewer", date: "2026-07-17T08:00:00Z" });
const deletion = document.addDeletion("Removed wording", { name: "tracked-deletion", styleId: "Normal", author: "Reviewer", date: "2026-07-17T08:05:00Z" });
const footnoteTarget = document.addParagraph("Paragraph with a source-free footnote.", { name: "footnote-target", styleId: "Normal" });
const endnoteTarget = document.addParagraph("Paragraph with a source-free endnote.", { name: "endnote-target", styleId: "Normal" });
const footnote = document.addFootnote(footnoteTarget, "Source-free footnote", { name: "footnote-evidence" });
const endnote = document.addEndnote(endnoteTarget, "Source-free endnote", { name: "endnote-evidence" });
const pngImage = document.addImage({
  name: "png-mark",
  styleId: "Normal",
  dataUrl: png,
  alt: "PNG approval mark",
  widthPx: 48,
  heightPx: 48,
});
const jpegImage = document.addImage({
  name: "jpeg-mark",
  styleId: "Normal",
  dataUrl: jpeg,
  alt: "JPEG approval mark",
  widthPx: 40,
  heightPx: 32,
});
const section = document.addSection({
  name: "landscape-section",
  breakType: "continuous",
  orientation: "landscape",
  pageSize: { widthTwips: 15840, heightTwips: 12240 },
  margins: { top: 720, right: 900, bottom: 720, left: 900 },
});
const secondSection = document.addParagraph("Second-section evidence.", { name: "second-section", styleId: "Normal" });
const secondSectionBookmark = document.addBookmark(secondSection, "SecondSection");
const internalLink = document.addHyperlink("Jump to second-section evidence", "#SecondSection", { name: "internal-link", styleId: "Normal" });

document.setSectionSettings(0, { differentFirstPage: true });
const defaultHeader = document.addHeader("Default header", {
  name: "default-header",
  referenceType: "default",
  sectionIndex: 0,
  variantActive: true,
});
const firstHeader = document.addHeader("First-page header", {
  name: "first-header",
  referenceType: "first",
  sectionIndex: 0,
  variantActive: true,
});
const evenHeader = document.addHeader("Even-page header", {
  name: "even-header",
  referenceType: "even",
  sectionIndex: 0,
  variantActive: true,
});
const defaultFooter = document.addFooter("1", {
  name: "default-footer",
  referenceType: "default",
  sectionIndex: 0,
  variantActive: true,
  fieldInstruction: "PAGE",
});
const firstFooter = document.addFooter("First-page footer", {
  name: "first-footer",
  referenceType: "first",
  sectionIndex: 0,
  variantActive: true,
});
const evenFooter = document.addFooter("Even-page footer", {
  name: "even-footer",
  referenceType: "even",
  sectionIndex: 0,
  variantActive: true,
});
const comment = document.addComment(commentTarget, "Confirm the release evidence.", {
  author: "Reviewer",
  initials: "RV",
  date: "2026-07-16T08:00:00Z",
});

const inspect = document.inspect({
  kind: "document,paragraph,listItem,table,comment,bookmark,note,header,footer,hyperlink,field,change,image,section,style,layout",
  maxChars: 24_000,
}).ndjson;
for (const expected of [
  "OpenChestnut document profile",
  "BodyAccent",
  "Bold and colored",
  "CUSTOMER_NAME",
  "readiness-table",
  "Default header",
  "First-page header",
  "Even-page header",
  "PNG approval mark",
  "JPEG approval mark",
  "landscape-section",
  "Added wording",
  "Removed wording",
  "Confirm the release evidence",
  "SecondSection",
  "Source-free footnote",
  "Source-free endnote",
]) assert.match(inspect, new RegExp(expected));
const contentControl = document.contentControls[0];
assert.equal(contentControl.tag, "CUSTOMER_NAME");
assert.equal(contentControl.text, "Ada Lovelace");
assert.equal(document.resolve(contentControl.id).targetId, contentControlParagraph.id);
assert.throws(() => document.fillContentControls({ UNKNOWN_FIELD: "value" }), /Unknown document content-control tag/);
assert.equal(document.resolve(formatted.id), formatted);
assert.equal(document.resolve(table.id).getCell(1, 1).value, "Pending");
assert.equal(document.resolve(pngImage.id).alt, "PNG approval mark");
assert.equal(document.resolve(section.id).orientation, "landscape");
assert.equal(document.resolve(comment.id).targetId, commentTarget.id);
assert.equal(document.resolve(defaultHeader.id).referenceType, "default");
assert.equal(document.resolve(firstHeader.id).referenceType, "first");
assert.equal(document.resolve(evenHeader.id).referenceType, "even");
assert.equal(document.resolve(defaultFooter.id).fieldInstruction, "PAGE");
assert.equal(document.resolve(insertion.id).changeType, "insert");
assert.equal(document.resolve(deletion.id).changeType, "delete");
assert.equal(document.resolve(footnote.id).targetId, footnoteTarget.id);
assert.equal(document.resolve(endnote.id).targetId, endnoteTarget.id);
assert.equal(document.resolve(secondSectionBookmark.id).targetId, secondSection.id);
assert.equal(internalLink.anchor, "SecondSection");
assert.equal(document.resolve(firstFooter.id).referenceType, "first");
assert.equal(document.resolve(evenFooter.id).referenceType, "even");
const modelVerification = document.verify({ visualQa: true });
assert.equal(modelVerification.ok, true, JSON.stringify(modelVerification.issues));
document.setSettings({ trackRevisions: true });

const firstDocx = await DocumentFile.exportDocx(document);
assert.equal(firstDocx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const firstDocxBytes = Buffer.from(await firstDocx.arrayBuffer());
const firstDocxSha256 = createHash("sha256").update(firstDocxBytes).digest("hex");

const watermarkDocument = DocumentModel.create({ name: "Watermark OpenChestnut slice", blocks: [] });
watermarkDocument.addParagraph("Native watermark verification body.");
watermarkDocument.addWatermark("CONFIDENTIAL", { id: "watermark/confidential", sectionIndex: 0 });
const watermarkDocx = await DocumentFile.exportDocx(watermarkDocument);
const watermarkZip = await JSZip.loadAsync(await watermarkDocx.arrayBuffer());
const watermarkHeaderPath = Object.keys(watermarkZip.files).find((name) => /^word\/header\d+\.xml$/.test(name));
assert.ok(watermarkHeaderPath);
assert.match(await watermarkZip.file(watermarkHeaderPath).async("text"), /<v:textpath[^>]*string="CONFIDENTIAL"/);
const importedWatermarkDocument = await DocumentFile.importDocx(watermarkDocx);
assert.equal(importedWatermarkDocument.watermarks.length, 1);
assert.equal(importedWatermarkDocument.watermarks[0].text, "CONFIDENTIAL");
assert.equal(importedWatermarkDocument.watermarks[0].sourceBound, true);
assert.equal(importedWatermarkDocument.watermarks[0].editable, true);
assert.equal(importedWatermarkDocument.resolve(importedWatermarkDocument.watermarks[0].id), importedWatermarkDocument.watermarks[0]);
const importedWatermarkInspect = importedWatermarkDocument.inspect({ kind: "watermark" }).ndjson;
assert.match(importedWatermarkInspect, /"sourceBound":true/);
const watermarkNoOp = await DocumentFile.exportDocx(importedWatermarkDocument);
assert.deepEqual(Buffer.from(await watermarkNoOp.arrayBuffer()), Buffer.from(await watermarkDocx.arrayBuffer()), "unchanged imported watermark must preserve exact source bytes");
importedWatermarkDocument.watermarks[0].text = "INTERNAL REVIEW";
const editedWatermarkDocx = await DocumentFile.exportDocx(importedWatermarkDocument);
assert.deepEqual(await changedZipParts(watermarkDocx, editedWatermarkDocx), [watermarkHeaderPath]);
const editedWatermarkRoundTrip = await DocumentFile.importDocx(editedWatermarkDocx);
assert.equal(editedWatermarkRoundTrip.watermarks[0].text, "INTERNAL REVIEW");
editedWatermarkRoundTrip.watermarks[0].remove();
const removedWatermarkDocx = await DocumentFile.exportDocx(editedWatermarkRoundTrip);
assert.deepEqual(await changedZipParts(editedWatermarkDocx, removedWatermarkDocx), [watermarkHeaderPath]);
const removedWatermarkRoundTrip = await DocumentFile.importDocx(removedWatermarkDocx);
assert.equal(removedWatermarkRoundTrip.watermarks.length, 0);
removedWatermarkRoundTrip.addWatermark("UNSAFE ADD", { sectionIndex: 0 });
await assert.rejects(
  () => DocumentFile.exportDocx(removedWatermarkRoundTrip),
  (error) => error?.code === "document_watermark_topology_changed" && /cannot add a watermark/i.test(error.message),
);
const watermarkScopeTamper = await DocumentFile.importDocx(watermarkDocx);
watermarkScopeTamper.watermarks[0].referenceType = "first";
await assert.rejects(
  () => DocumentFile.exportDocx(watermarkScopeTamper),
  (error) => error?.code === "unsupported_document_watermark_edit" && /fixed after import/i.test(error.message),
);

const fragmentedPatchFixture = DocumentModel.create({ name: "Fragmented source-bound patch", blocks: [] });
fragmentedPatchFixture.addParagraph("Quarterly plan");
fragmentedPatchFixture.addTable({ values: [["Revenue", "42"]] });
const fragmentedPatchBase = await DocumentFile.exportDocx(fragmentedPatchFixture);
const fragmentedPatchBaseZip = await JSZip.loadAsync(await fragmentedPatchBase.arrayBuffer());
const fragmentedPatchBaseXml = await fragmentedPatchBaseZip.file("word/document.xml").async("text");
const fragmentedPatchXml = fragmentedPatchBaseXml
  .replace(
    '<w:pPr><w:pStyle w:val="Normal" /></w:pPr><w:r><w:t>Quarterly plan</w:t></w:r>',
    '<w:pPr><w:pStyle w:val="Normal" /><w:widowControl /></w:pPr><w:r><w:t>Quarter</w:t></w:r><w:r><w:t>ly plan</w:t></w:r>',
  )
  .replace(
    '<w:r><w:rPr><w:b /></w:rPr><w:t>Revenue</w:t></w:r>',
    '<w:r><w:rPr><w:b /></w:rPr><w:t>Rev</w:t></w:r><w:r><w:rPr><w:b /></w:rPr><w:t>enue</w:t></w:r>',
  );
assert.notEqual(fragmentedPatchXml, fragmentedPatchBaseXml);
const fragmentedPatchSource = await DocumentFile.patchDocx(fragmentedPatchBase, [
  { path: "word/document.xml", xml: fragmentedPatchXml },
]);
const fragmentedImported = await DocumentFile.importDocx(fragmentedPatchSource);
const fragmentedParagraph = fragmentedImported.blocks.find((block) => block.kind === "paragraph" && block.text === "Quarterly plan");
const fragmentedTable = fragmentedImported.blocks.find((block) => block.kind === "table");
const fragmentedCell = fragmentedTable.getCell(0, 0);
assert.equal(fragmentedParagraph.textEditable, false);
assert.equal(fragmentedParagraph.textPatchable, true);
assert.equal(fragmentedCell.editable, false);
assert.equal(fragmentedCell.textPatchable, true);
fragmentedImported.resolve(`${fragmentedParagraph.id}/text`).replace("Quarterly", "Annual");
fragmentedImported.resolve(`${fragmentedCell.id}/text`).replace("Revenue", "Net revenue");
const fragmentedPatchOutput = await DocumentFile.exportDocx(fragmentedImported);
assert.deepEqual(await changedZipParts(fragmentedPatchSource, fragmentedPatchOutput), ["word/document.xml"]);
const fragmentedPatchOutputZip = await JSZip.loadAsync(await fragmentedPatchOutput.arrayBuffer());
const fragmentedPatchOutputXml = await fragmentedPatchOutputZip.file("word/document.xml").async("text");
assert.match(fragmentedPatchOutputXml, /<w:widowControl\s*\/>[\s\S]*?<w:r><w:t>Annual<\/w:t><\/w:r><w:r><w:t xml:space="preserve"> plan<\/w:t><\/w:r>/);
assert.match(fragmentedPatchOutputXml, /<w:r><w:rPr><w:b\s*\/><\/w:rPr><w:t>Net revenue<\/w:t><\/w:r><w:r><w:rPr><w:b\s*\/><\/w:rPr><w:t\s*\/><\/w:r>/);
const fragmentedRoundTrip = await DocumentFile.importDocx(fragmentedPatchOutput);
assert.equal(fragmentedRoundTrip.blocks.find((block) => block.kind === "paragraph")?.text, "Annual plan");
assert.equal(fragmentedRoundTrip.blocks.find((block) => block.kind === "table")?.getCell(0, 0).value, "Net revenue");

const mixedFormattingXml = fragmentedPatchXml.replace(
  '<w:r><w:t>ly plan</w:t></w:r>',
  '<w:r><w:rPr><w:i /></w:rPr><w:t>ly plan</w:t></w:r>',
);
const mixedFormattingSource = await DocumentFile.patchDocx(fragmentedPatchBase, [
  { path: "word/document.xml", xml: mixedFormattingXml },
]);
const mixedFormattingDocument = await DocumentFile.importDocx(mixedFormattingSource);
const mixedFormattingParagraph = mixedFormattingDocument.blocks.find((block) => block.kind === "paragraph");
assert.equal(mixedFormattingParagraph.textPatchable, true);
mixedFormattingDocument.resolve(`${mixedFormattingParagraph.id}/text`).replace("Quarterly", "Annual");
await assert.rejects(
  () => DocumentFile.exportDocx(mixedFormattingDocument),
  (error) => error?.code === "unsupported_document_edit" && /different formatting/.test(error.message),
);

const trackedReplacement = await DocumentFile.addTrackedReplacement(firstDocx, {
  expectedSourceSha256: firstDocxSha256,
  targetBlockIndex: document.blocks.indexOf(formatted),
  expectedText: "Bold and colored",
  search: "colored",
  replacement: "reviewed",
  author: "Reviewer",
  date: "2026-07-21T09:30:00Z",
});
const trackedReplacementBytes = Buffer.from(await trackedReplacement.arrayBuffer());
const trackedReplacementSha256 = createHash("sha256").update(trackedReplacementBytes).digest("hex");
assert.deepEqual(trackedReplacement.metadata.trackedReplacement, {
  sourceSha256: firstDocxSha256,
  outputSha256: trackedReplacementSha256,
  target: { kind: "paragraph", blockIndex: document.blocks.indexOf(formatted) },
  targetBlockIndex: document.blocks.indexOf(formatted),
  targetBodyIndex: document.blocks.indexOf(formatted),
  sourceElementSha256: trackedReplacement.metadata.trackedReplacement.sourceElementSha256,
  outputElementSha256: trackedReplacement.metadata.trackedReplacement.outputElementSha256,
  deletedTextSha256: createHash("sha256").update("colored").digest("hex"),
  insertedTextSha256: createHash("sha256").update("reviewed").digest("hex"),
  deletedTextChars: "colored".length,
  insertedTextChars: "reviewed".length,
  matchedSourceRunCount: 1,
  deletionNativeRevisionId: trackedReplacement.metadata.trackedReplacement.deletionNativeRevisionId,
  insertionNativeRevisionId: trackedReplacement.metadata.trackedReplacement.insertionNativeRevisionId,
  changedParts: ["word/document.xml"],
});
assert.match(trackedReplacement.metadata.trackedReplacement.sourceElementSha256, /^[0-9a-f]{64}$/);
assert.match(trackedReplacement.metadata.trackedReplacement.outputElementSha256, /^[0-9a-f]{64}$/);
assert.notEqual(trackedReplacement.metadata.trackedReplacement.sourceElementSha256, trackedReplacement.metadata.trackedReplacement.outputElementSha256);
assert.notEqual(trackedReplacement.metadata.trackedReplacement.deletionNativeRevisionId, trackedReplacement.metadata.trackedReplacement.insertionNativeRevisionId);
assert.deepEqual(Buffer.from(await firstDocx.arrayBuffer()), firstDocxBytes, "tracked replacement must not mutate its source blob");
const trackedReplacementDocument = await DocumentFile.importDocx(trackedReplacement);
const trackedReplacementTarget = trackedReplacementDocument.blocks[document.blocks.indexOf(formatted)];
assert.equal(trackedReplacementTarget.kind, "paragraph");
assert.equal(trackedReplacementTarget.text, "Bold and reviewed");
assert.equal(trackedReplacementTarget.textEditable, false);

const acceptedTrackedReplacement = await DocumentFile.finalizeRevisions(trackedReplacement, {
  mode: "accept",
  expectedSourceSha256: trackedReplacementSha256,
});
assert.equal(acceptedTrackedReplacement.metadata.revisionFinalization.insertionCount, 2);
assert.equal(acceptedTrackedReplacement.metadata.revisionFinalization.deletionCount, 2);
const acceptedTrackedReplacementDocument = await DocumentFile.importDocx(acceptedTrackedReplacement);
assert.equal(acceptedTrackedReplacementDocument.blocks[document.blocks.indexOf(formatted)].text, "Bold and reviewed");
assert.equal(acceptedTrackedReplacementDocument.blocks.some((block) => block.text === "Added wording"), true);
assert.equal(acceptedTrackedReplacementDocument.blocks.some((block) => block.text === "Removed wording"), false);

const rejectedTrackedReplacement = await DocumentFile.finalizeRevisions(trackedReplacement, {
  mode: "reject",
  keepTracking: true,
  expectedSourceSha256: trackedReplacementSha256,
});
const rejectedTrackedReplacementDocument = await DocumentFile.importDocx(rejectedTrackedReplacement);
assert.equal(rejectedTrackedReplacementDocument.blocks[document.blocks.indexOf(formatted)].text, "Bold and colored");
assert.equal(rejectedTrackedReplacementDocument.blocks.some((block) => block.text === "Added wording"), false);
assert.equal(rejectedTrackedReplacementDocument.blocks.some((block) => block.text === "Removed wording"), true);

const tableBlockIndex = document.blocks.indexOf(table);
const tableTrackedReplacement = await DocumentFile.addTrackedReplacement(firstDocx, {
  expectedSourceSha256: firstDocxSha256,
  target: { kind: "tableCell", blockIndex: tableBlockIndex, row: 1, column: 1 },
  expectedText: "Pending",
  search: "Pending",
  replacement: "Approved",
  author: "Table reviewer",
  date: "2026-07-21T11:00:00Z",
});
const tableTrackedBytes = Buffer.from(await tableTrackedReplacement.arrayBuffer());
const tableTrackedSha256 = createHash("sha256").update(tableTrackedBytes).digest("hex");
assert.deepEqual(tableTrackedReplacement.metadata.trackedReplacement.target, {
  kind: "tableCell",
  blockIndex: tableBlockIndex,
  row: 1,
  column: 1,
});
assert.equal(tableTrackedReplacement.metadata.trackedReplacement.targetBlockIndex, tableBlockIndex);
assert.equal(tableTrackedReplacement.metadata.trackedReplacement.targetBodyIndex, tableBlockIndex);
assert.equal(tableTrackedReplacement.metadata.trackedReplacement.outputSha256, tableTrackedSha256);
assert.equal(tableTrackedReplacement.metadata.trackedReplacement.matchedSourceRunCount, 1);
assert.deepEqual(tableTrackedReplacement.metadata.trackedReplacement.changedParts, ["word/document.xml"]);
assert.deepEqual(Buffer.from(await firstDocx.arrayBuffer()), firstDocxBytes, "table tracked replacement must not mutate its source blob");
const tableTrackedDocument = await DocumentFile.importDocx(tableTrackedReplacement);
assert.equal(tableTrackedDocument.blocks[tableBlockIndex].kind, "table");
assert.equal(tableTrackedDocument.blocks[tableBlockIndex].getCell(1, 1).value, "Approved");
const tableTrackedZip = await JSZip.loadAsync(tableTrackedBytes);
const tableTrackedXml = await tableTrackedZip.file("word/document.xml").async("text");
assert.match(tableTrackedXml, /<w:tc>[\s\S]*?<w:del\b[\s\S]*?<w:delText>Pending<\/w:delText>[\s\S]*?<w:ins\b[\s\S]*?<w:t>Approved<\/w:t>[\s\S]*?<\/w:tc>/);

const acceptedTableTrackedReplacement = await DocumentFile.finalizeRevisions(tableTrackedReplacement, {
  mode: "accept",
  expectedSourceSha256: tableTrackedSha256,
});
assert.equal(acceptedTableTrackedReplacement.metadata.revisionFinalization.insertionCount, 2);
assert.equal(acceptedTableTrackedReplacement.metadata.revisionFinalization.deletionCount, 2);
assert.equal((await DocumentFile.importDocx(acceptedTableTrackedReplacement)).blocks[tableBlockIndex].getCell(1, 1).value, "Approved");

const rejectedTableTrackedReplacement = await DocumentFile.finalizeRevisions(tableTrackedReplacement, {
  mode: "reject",
  expectedSourceSha256: tableTrackedSha256,
});
assert.equal((await DocumentFile.importDocx(rejectedTableTrackedReplacement)).blocks[tableBlockIndex].getCell(1, 1).value, "Pending");
await assert.rejects(
  () => DocumentFile.addTrackedReplacement(firstDocx, {
    expectedSourceSha256: firstDocxSha256,
    target: { kind: "paragraph", blockIndex: 0 },
    targetBlockIndex: 0,
    expectedText: document.blocks[0].text,
    search: "OpenChestnut",
    replacement: "OpenChestnut native",
    author: "Reviewer",
  }),
  /either target or targetBlockIndex, not both/,
);
await assert.rejects(
  () => DocumentFile.addTrackedReplacement(firstDocx, {
    expectedSourceSha256: firstDocxSha256,
    target: { kind: "tableCell", blockIndex: tableBlockIndex, row: -1, column: 1 },
    expectedText: "Pending",
    search: "Pending",
    replacement: "Approved",
    author: "Reviewer",
  }),
  /row and column must be unsigned 32-bit physical indexes/,
);
await assert.rejects(
  () => DocumentFile.addTrackedReplacement(firstDocx, {
    expectedSourceSha256: "0".repeat(64),
    targetBlockIndex: 0,
    expectedText: document.blocks[0].text,
    search: "OpenChestnut",
    replacement: "OpenChestnut native",
    author: "Reviewer",
  }),
  (error) => error?.code === "document_source_hash_mismatch",
);
const acceptedRevisions = await DocumentFile.finalizeRevisions(firstDocx, {
  mode: "accept",
  expectedSourceSha256: firstDocxSha256,
});
assert.deepEqual(acceptedRevisions.metadata.revisionFinalization, {
  mode: "accept",
  sourceSha256: firstDocxSha256,
  outputSha256: createHash("sha256").update(Buffer.from(await acceptedRevisions.arrayBuffer())).digest("hex"),
  insertionCount: 1,
  deletionCount: 1,
  trackingBefore: true,
  trackingAfter: false,
  changedParts: ["word/document.xml", "word/settings.xml"],
});
const acceptedDocument = await DocumentFile.importDocx(acceptedRevisions);
assert.equal(acceptedDocument.settings.trackRevisions, false);
assert.equal(acceptedDocument.blocks.some((block) => block.kind === "change"), false);
assert.equal(acceptedDocument.blocks.some((block) => block.text === "Added wording"), true);
assert.equal(acceptedDocument.blocks.some((block) => block.text === "Removed wording"), false);

const rejectedRevisions = await DocumentFile.finalizeRevisions(firstDocx, {
  mode: "reject",
  keepTracking: true,
  expectedSourceSha256: firstDocxSha256,
});
assert.deepEqual(rejectedRevisions.metadata.revisionFinalization.changedParts, ["word/document.xml"]);
assert.equal(rejectedRevisions.metadata.revisionFinalization.trackingAfter, true);
const rejectedDocument = await DocumentFile.importDocx(rejectedRevisions);
assert.equal(rejectedDocument.settings.trackRevisions, true);
assert.equal(rejectedDocument.blocks.some((block) => block.kind === "change"), false);
assert.equal(rejectedDocument.blocks.some((block) => block.text === "Added wording"), false);
assert.equal(rejectedDocument.blocks.some((block) => block.text === "Removed wording"), true);
assert.deepEqual(Buffer.from(await firstDocx.arrayBuffer()), firstDocxBytes, "revision finalization must not mutate its source blob");
await assert.rejects(
  () => DocumentFile.finalizeRevisions(firstDocx, { mode: "accept", expectedSourceSha256: "0".repeat(64) }),
  (error) => error?.code === "document_source_hash_mismatch",
);
await assert.rejects(
  () => DocumentFile.finalizeRevisions(firstDocx, { mode: "all", expectedSourceSha256: firstDocxSha256 }),
  /mode must be accept or reject/i,
);
await assert.rejects(
  () => DocumentFile.finalizeRevisions(acceptedRevisions, {
    mode: "accept",
    expectedSourceSha256: acceptedRevisions.metadata.revisionFinalization.outputSha256,
  }),
  (error) => error?.code === "document_revisions_not_found",
);
const imported = await DocumentFile.importDocx(firstDocx);
assert.equal(imported.defaultRunStyle.fontFamily, "Aptos");
assert.equal(imported.defaultRunStyle.fontSize, 11);
assert.equal(imported.styles.values().some((style) => style.id === "BodyAccent" && style.basedOn === "Normal"), true);
const importedFormatted = imported.blocks.find((block) => block.text === "Bold and colored");
assert.equal(importedFormatted?.kind, "paragraph");
assert.equal(importedFormatted?.paragraphFormat.alignment, "center");
assert.equal(importedFormatted?.runs.length, 2);
assert.equal(importedFormatted?.runs[0].style.bold, true);
assert.equal(importedFormatted?.runs[0].style.fontSize, 15);
assert.equal(importedFormatted?.runs[1].style.italic, true);
assert.equal(importedFormatted?.runs[1].style.underline, true);
assert.equal(importedFormatted?.runs[1].style.color, "#cc0000");
assert.equal(imported.contentControls.length, 1);
assert.equal(imported.contentControls[0].tag, "CUSTOMER_NAME");
assert.equal(imported.contentControls[0].alias, "Customer name");
assert.equal(imported.contentControls[0].text, "Ada Lovelace");
assert.ok(Number.isInteger(imported.contentControls[0].nativeId));
assert.equal(imported.blocks.filter((block) => block.kind === "listItem").length, 2);
assert.equal(imported.blocks.find((block) => block.kind === "table")?.values[1][1], "Pending");
assert.equal(imported.blocks.find((block) => block.kind === "hyperlink")?.url, hyperlink.url);
assert.equal(imported.blocks.find((block) => block.kind === "field")?.instruction, "PAGE");
assert.deepEqual(imported.blocks.filter((block) => block.kind === "change").map((block) => [block.changeType, block.text, block.author]), [
  ["insert", "Added wording", "Reviewer"],
  ["delete", "Removed wording", "Reviewer"],
]);
assert.equal(imported.blocks.filter((block) => block.kind === "image").length, 2);
assert.equal(imported.blocks.some((block) => block.kind === "image" && block.dataUrl.startsWith("data:image/png;base64,")), true);
assert.equal(imported.blocks.some((block) => block.kind === "image" && block.dataUrl.startsWith("data:image/jpeg;base64,")), true);
assert.equal(imported.blocks.find((block) => block.kind === "section")?.orientation, "landscape");
assert.equal(imported.settings.evenAndOddHeaders, true);
assert.equal(imported.settings.trackRevisions, true);
assert.equal(imported.sectionSettings[0]?.differentFirstPage, true);
assert.deepEqual(imported.headers.map((item) => item.referenceType), ["default", "first", "even"]);
assert.deepEqual(imported.footers.map((item) => item.referenceType), ["default", "first", "even"]);
assert.equal(imported.footers[0].fieldInstruction, "PAGE");
assert.equal(imported.comments.length, 1);
assert.equal(imported.bookmarks.length, 1);
assert.equal(imported.bookmarks[0].name, "SecondSection");
assert.equal(imported.bookmarks[0].targetId, imported.blocks.find((block) => block.text === "Second-section evidence.")?.id);
assert.equal(imported.blocks.find((block) => block.kind === "hyperlink" && block.anchor === "SecondSection")?.text, "Jump to second-section evidence");
assert.deepEqual(imported.notes.map((note) => [note.kind, note.text, note.nativeId]), [
  ["footnote", "Source-free footnote", 1],
  ["endnote", "Source-free endnote", 1],
]);
assert.equal(imported.notes.every((note) => imported.resolve(note.id) === note), true);

importedFormatted.text = "Bold and edited";
importedFormatted.runs[0].text = "Bold ";
importedFormatted.runs[1].text = "and edited";
importedFormatted.runs[1].style.color = "#008844";
const importedBullet = imported.blocks.find((block) => block.kind === "listItem" && block.listType === "bullet");
importedBullet.text = "Inspect the edited semantic model.";
const importedTable = imported.blocks.find((block) => block.kind === "table");
importedTable.values[1][1] = "Pass";
const importedLink = imported.blocks.find((block) => block.kind === "hyperlink");
importedLink.url = "https://learn.microsoft.com/office/open-xml/word-processing";
importedLink.tooltip = "Edited target";
importedLink.history = false;
const importedField = imported.blocks.find((block) => block.kind === "field");
importedField.instruction = "NUMPAGES";
importedField.display = "2";
const importedInsertion = imported.blocks.find((block) => block.kind === "change" && block.changeType === "insert");
importedInsertion.text = "Edited insertion";
importedInsertion.author = "Lead reviewer";
importedInsertion.date = "2026-07-17T08:30:00Z";
const importedPng = imported.blocks.find((block) => block.kind === "image" && block.alt === "PNG approval mark");
importedPng.alt = "Edited PNG approval mark";
importedPng.widthPx = 56;
const importedSection = imported.blocks.find((block) => block.kind === "section");
importedSection.margins.left = 1200;
imported.comments[0].author = "Lead reviewer";
imported.comments[0].initials = "LR";
imported.comments[0].text = "Release evidence approved.";
imported.notes[0].text = "Edited footnote";
imported.notes[1].text = "Edited endnote";
assert.deepEqual(imported.fillContentControls({ CUSTOMER_NAME: "Grace Hopper" }), { updated: 1, matchedTags: ["CUSTOMER_NAME"], missingTags: [] });

const secondDocx = await DocumentFile.exportDocx(imported);
const roundTrip = await DocumentFile.importDocx(secondDocx);
const roundTripFormatted = roundTrip.blocks.find((block) => block.text === "Bold and edited");
assert.equal(roundTripFormatted?.runs.length, 2);
assert.equal(roundTripFormatted?.runs[0].style.bold, true);
assert.equal(roundTripFormatted?.runs[1].style.italic, true);
assert.equal(roundTripFormatted?.runs[1].style.color, "#008844");
assert.equal(roundTrip.blocks.some((block) => block.kind === "listItem" && block.text === "Inspect the edited semantic model."), true);
assert.equal(roundTrip.blocks.find((block) => block.kind === "table")?.values[1][1], "Pass");
assert.equal(roundTrip.blocks.find((block) => block.kind === "hyperlink")?.history, false);
assert.equal(roundTrip.blocks.find((block) => block.kind === "field")?.instruction, "NUMPAGES");
assert.deepEqual(roundTrip.blocks.filter((block) => block.kind === "change").map((block) => [block.changeType, block.text, block.author]), [
  ["insert", "Edited insertion", "Lead reviewer"],
  ["delete", "Removed wording", "Reviewer"],
]);
assert.equal(roundTrip.blocks.find((block) => block.kind === "image" && block.alt === "Edited PNG approval mark")?.widthPx, 56);
assert.equal(roundTrip.blocks.find((block) => block.kind === "section")?.margins.left, 1200);
assert.equal(roundTrip.comments[0].author, "Lead reviewer");
assert.equal(roundTrip.comments[0].text, "Release evidence approved.");
assert.equal(roundTrip.bookmarks[0].name, "SecondSection");
assert.equal(roundTrip.blocks.some((block) => block.kind === "hyperlink" && block.anchor === "SecondSection"), true);
assert.deepEqual(roundTrip.notes.map((note) => [note.kind, note.text]), [
  ["footnote", "Edited footnote"],
  ["endnote", "Edited endnote"],
]);
assert.equal(roundTrip.contentControls[0].text, "Grace Hopper");
assert.equal(roundTrip.resolve(roundTrip.contentControls[0].targetId)?.text, "Customer: Grace Hopper.");
assert.equal(roundTrip.verify({ visualQa: true }).ok, true);

const blockControlDocument = DocumentModel.create({ name: "Block content-control profile", blocks: [] });
const blockControlParagraph = blockControlDocument.addBlockTextContentControl("Executive summary", {
  blockId: "executive-summary-paragraph",
  id: "executive-summary-control",
  tag: "EXECUTIVE_SUMMARY",
  alias: "Executive summary",
  styleId: "Normal",
  paragraphFormat: { keepNext: true },
  runStyle: { bold: true, color: "#1D4ED8" },
});
const blockControl = blockControlDocument.contentControls[0];
assert.equal(blockControl.placement, "block");
assert.equal(blockControl.runIndex, undefined);
assert.equal(blockControl.targetId, blockControlParagraph.id);
assert.equal(blockControlDocument.resolve(blockControl.id).targetId, blockControlParagraph.id);
assert.equal(blockControlDocument.resolve(blockControl.targetId), blockControlParagraph);
assert.match(blockControlDocument.inspect({ kind: "contentControl" }).ndjson, /"placement":"block"/);
assert.deepEqual(blockControlDocument.fillContentControls({ EXECUTIVE_SUMMARY: "Updated executive summary" }), { updated: 1, matchedTags: ["EXECUTIVE_SUMMARY"], missingTags: [] });
assert.equal(blockControlParagraph.text, "Updated executive summary");
assert.equal(blockControlParagraph.runs[0].text, "Updated executive summary");
assert.equal(blockControlDocument.verify().ok, true);
const blockControlDocx = await DocumentFile.exportDocx(blockControlDocument);
const blockControlZip = await JSZip.loadAsync(await blockControlDocx.arrayBuffer());
const blockControlXml = await blockControlZip.file("word/document.xml").async("text");
assert.match(blockControlXml, /<w:body>\s*<w:sdt>[\s\S]*?<w:tag w:val="EXECUTIVE_SUMMARY"\s*\/>[\s\S]*?<w:text\s*\/>[\s\S]*?<w:sdtContent>\s*<w:p>[\s\S]*Updated executive summary[\s\S]*?<\/w:p>\s*<\/w:sdtContent>\s*<\/w:sdt>/);
const importedBlockControlDocument = await DocumentFile.importDocx(blockControlDocx);
const importedBlockControl = importedBlockControlDocument.contentControls[0];
assert.equal(importedBlockControl.placement, "block");
assert.equal(importedBlockControl.tag, "EXECUTIVE_SUMMARY");
assert.equal(importedBlockControl.text, "Updated executive summary");
assert.equal(importedBlockControlDocument.blocks[0].runs[0].style.bold, true);
assert.equal(importedBlockControlDocument.blocks[0].runs[0].style.color, "#1d4ed8");
const unchangedBlockControlDocx = await DocumentFile.exportDocx(importedBlockControlDocument);
assert.deepEqual(Buffer.from(await unchangedBlockControlDocx.arrayBuffer()), Buffer.from(await blockControlDocx.arrayBuffer()), "unchanged imported block content control must preserve source bytes");
importedBlockControlDocument.fillContentControls({ EXECUTIVE_SUMMARY: "Final executive summary" });
importedBlockControl.tag = "SUMMARY";
importedBlockControl.alias = "Summary";
const roundTripBlockControlDocument = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedBlockControlDocument));
assert.equal(roundTripBlockControlDocument.blocks[0].text, "Final executive summary");
assert.equal(roundTripBlockControlDocument.contentControls[0].tag, "SUMMARY");
assert.equal(roundTripBlockControlDocument.contentControls[0].alias, "Summary");
const blockControlTopologyTamper = await DocumentFile.importDocx(blockControlDocx);
delete blockControlTopologyTamper.blocks[0].blockContentControl;
await assert.rejects(
  () => DocumentFile.exportDocx(blockControlTopologyTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => blockControlDocument.addBlockTextContentControl("Invalid", { tag: "INVALID_BLOCK_CHECKBOX", controlType: "checkbox", checked: false }),
  /only plain text/i,
);
assert.throws(
  () => blockControlDocument.addBlockTextContentControl("Invalid", { tag: "INVALID_BLOCK_ALIAS", alias: "" }),
  /non-empty alias/i,
);
assert.throws(
  () => blockControlDocument.addParagraph("Invalid", { blockContentControl: { tag: "INVALID_BLOCK_RUNS" }, runs: [{ text: "One" }, { text: "Two" }] }),
  /exactly one ordinary paragraph run/i,
);

const tableCellControlDocument = DocumentModel.create({ name: "Table-cell content-control profile", blocks: [] });
tableCellControlDocument.applyDesignPreset("report");
const tableCellControlTable = tableCellControlDocument.addTable({
  id: "owner-table",
  values: [["Field", "Value"], ["Owner", "Ada Lovelace"]],
});
const tableCellControl = tableCellControlTable.getCell(1, 1).addTextContentControl({
  id: "table-owner-control",
  tag: "TABLE_OWNER",
  alias: "Table owner",
});
assert.equal(tableCellControl.placement, "tableCell");
assert.equal(tableCellControl.targetId, "owner-table/cell/1/1");
assert.equal(tableCellControl.row, 1);
assert.equal(tableCellControl.column, 1);
assert.equal(tableCellControl.text, "Ada Lovelace");
assert.equal(tableCellControlTable.getCell(1, 1).contentControl.id, tableCellControl.id);
assert.equal(tableCellControlDocument.resolve(tableCellControl.id).targetId, tableCellControl.targetId);
assert.equal(tableCellControlDocument.resolve(tableCellControl.targetId).contentControl.id, tableCellControl.id);
assert.match(tableCellControlDocument.inspect({ kind: "contentControl" }).ndjson, /"placement":"tableCell"/);
assert.deepEqual(tableCellControlDocument.fillContentControls({ TABLE_OWNER: "Grace Hopper" }), { updated: 1, matchedTags: ["TABLE_OWNER"], missingTags: [] });
assert.equal(tableCellControlTable.values[1][1], "Grace Hopper");
assert.equal(tableCellControlDocument.verify().ok, true);
const tableCellControlDocx = await DocumentFile.exportDocx(tableCellControlDocument);
const tableCellControlZip = await JSZip.loadAsync(await tableCellControlDocx.arrayBuffer());
const tableCellControlXml = await tableCellControlZip.file("word/document.xml").async("text");
assert.match(tableCellControlXml, /<w:tc>[\s\S]*?<w:sdt>[\s\S]*?<w:tag w:val="TABLE_OWNER"\s*\/>[\s\S]*?<w:text\s*\/>[\s\S]*?<w:sdtContent>\s*<w:p>[\s\S]*?Grace Hopper[\s\S]*?<\/w:p>\s*<\/w:sdtContent>\s*<\/w:sdt>[\s\S]*?<\/w:tc>/);
const importedTableCellControlDocument = await DocumentFile.importDocx(tableCellControlDocx);
const importedTableCellControl = importedTableCellControlDocument.contentControls.find((control) => control.tag === "TABLE_OWNER");
assert.equal(importedTableCellControl.placement, "tableCell");
assert.equal(importedTableCellControl.text, "Grace Hopper");
assert.ok(Number.isInteger(importedTableCellControl.nativeId));
assert.equal(importedTableCellControlDocument.resolve(importedTableCellControl.targetId).value, "Grace Hopper");
const unchangedTableCellControlDocx = await DocumentFile.exportDocx(importedTableCellControlDocument);
assert.deepEqual(Buffer.from(await unchangedTableCellControlDocx.arrayBuffer()), Buffer.from(await tableCellControlDocx.arrayBuffer()), "unchanged imported table-cell content control must preserve source bytes");
importedTableCellControlDocument.fillContentControls({ TABLE_OWNER: "Katherine Johnson" });
importedTableCellControl.tag = "OWNER";
importedTableCellControl.alias = "Owner";
const roundTripTableCellControlDocument = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedTableCellControlDocument));
const roundTripTableCellControl = roundTripTableCellControlDocument.contentControls.find((control) => control.tag === "OWNER");
assert.equal(roundTripTableCellControl.text, "Katherine Johnson");
assert.equal(roundTripTableCellControl.alias, "Owner");
assert.equal(roundTripTableCellControlDocument.resolve(roundTripTableCellControl.targetId).value, "Katherine Johnson");
const tableCellControlTopologyTamper = await DocumentFile.importDocx(tableCellControlDocx);
delete tableCellControlTopologyTamper.blocks.find((block) => block.kind === "table").cells.find((cell) => cell.row === 1 && cell.column === 1).contentControl;
await assert.rejects(
  () => DocumentFile.exportDocx(tableCellControlTopologyTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
const importedOrdinaryTableDocument = await DocumentFile.importDocx(firstDocx);
const importedOrdinaryTable = importedOrdinaryTableDocument.blocks.find((block) => block.kind === "table");
assert.throws(
  () => importedOrdinaryTable.getCell(0, 0).addTextContentControl({ tag: "NEW_IMPORTED_CONTROL", alias: "New imported control" }),
  /cannot add a content control to an imported table.*topology is source-bound/i,
);
assert.throws(
  () => tableCellControlTable.getCell(0, 1).addTextContentControl({ tag: "INVALID_TABLE_CHECKBOX", alias: "Invalid", controlType: "checkbox", checked: false }),
  /text content-control creation cannot use type checkbox/i,
);
assert.throws(
  () => tableCellControlDocument.addTable({ values: [["A", "B"], ["C"]] }).getCell(0, 0).addTextContentControl({ tag: "RAGGED", alias: "Ragged" }),
  /must be rectangular/i,
);
assert.throws(
  () => tableCellControlTable.getCell(9, 9).addTextContentControl({ tag: "OUT_OF_RANGE", alias: "Out of range" }),
  /existing physical cell/i,
);
assert.throws(
  () => tableCellControlTable.getCell(0.5, 0).addTextContentControl({ tag: "FRACTIONAL", alias: "Fractional" }),
  /existing physical cell/i,
);

const typedTableControlDocument = DocumentModel.create({ name: "Typed table-cell content-control profile", blocks: [] });
typedTableControlDocument.applyDesignPreset("report");
const typedTableControlTable = typedTableControlDocument.addTable({
  id: "typed-table-controls",
  values: [
    ["Field", "Value"],
    ["Approved", "pending"],
    ["Priority", "pending"],
    ["Contact", "pending"],
    ["Review date", "pending"],
  ],
});
const typedTableChoices = [
  { displayText: "Low", value: "low" },
  { displayText: "High", value: "high" },
];
const typedTableCheckbox = typedTableControlTable.getCell(1, 1).addCheckboxContentControl(false, {
  id: "table-approved-control",
  tag: "TABLE_APPROVED",
  alias: "Table approved",
});
const typedTableDropdown = typedTableControlTable.getCell(2, 1).addDropdownContentControl(typedTableChoices, {
  id: "table-priority-control",
  tag: "TABLE_PRIORITY",
  alias: "Table priority",
  selectedValue: "low",
});
const typedTableComboBox = typedTableControlTable.getCell(3, 1).addComboBoxContentControl(typedTableChoices, {
  id: "table-contact-control",
  tag: "TABLE_CONTACT",
  alias: "Table contact",
  value: "High",
});
const typedTableDate = typedTableControlTable.getCell(4, 1).addDateContentControl("2026-07-22", {
  id: "table-review-date-control",
  tag: "TABLE_REVIEW_DATE",
  alias: "Table review date",
});
assert.deepEqual(
  [typedTableCheckbox, typedTableDropdown, typedTableComboBox, typedTableDate].map((control) => [control.controlType, control.placement, control.row, control.column, control.text]),
  [
    ["checkbox", "tableCell", 1, 1, "☐"],
    ["dropdown", "tableCell", 2, 1, "Low"],
    ["comboBox", "tableCell", 3, 1, "High"],
    ["date", "tableCell", 4, 1, "2026-07-22"],
  ],
);
assert.deepEqual(typedTableControlTable.values.map((row) => row[1]), ["Value", "☐", "Low", "High", "2026-07-22"]);
assert.deepEqual(typedTableControlDocument.setCheckboxContentControls({ TABLE_APPROVED: true }), { updated: 1, matchedTags: ["TABLE_APPROVED"], missingTags: [] });
assert.deepEqual(typedTableControlDocument.setDropdownContentControls({ TABLE_PRIORITY: "high" }), { updated: 1, matchedTags: ["TABLE_PRIORITY"], missingTags: [] });
assert.deepEqual(typedTableControlDocument.setComboBoxContentControls({ TABLE_CONTACT: "Pager duty" }), { updated: 1, matchedTags: ["TABLE_CONTACT"], missingTags: [] });
assert.deepEqual(typedTableControlDocument.setDateContentControls({ TABLE_REVIEW_DATE: "2028-02-29" }), { updated: 1, matchedTags: ["TABLE_REVIEW_DATE"], missingTags: [] });
assert.deepEqual(typedTableControlTable.values.map((row) => row[1]), ["Value", "☒", "High", "Pager duty", "2028-02-29"]);
assert.equal(typedTableControlDocument.verify().ok, true);
const typedTableControlDocx = await DocumentFile.exportDocx(typedTableControlDocument);
const typedTableControlZip = await JSZip.loadAsync(await typedTableControlDocx.arrayBuffer());
const typedTableControlXml = await typedTableControlZip.file("word/document.xml").async("text");
assert.match(typedTableControlXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_APPROVED"\s*\/>[\s\S]*?<w14:checkbox>[\s\S]*?<w14:checked w14:val="1"\s*\/>[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
assert.match(typedTableControlXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_PRIORITY"\s*\/>[\s\S]*?<w:dropDownList w:lastValue="high">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
assert.match(typedTableControlXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_CONTACT"\s*\/>[\s\S]*?<w:comboBox w:lastValue="Pager duty">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
assert.match(typedTableControlXml, /<w:tc>[\s\S]*?<w:tag w:val="TABLE_REVIEW_DATE"\s*\/>[\s\S]*?<w:date w:fullDate="2028-02-29T00:00:00Z">[\s\S]*?<\/w:sdt>[\s\S]*?<\/w:tc>/);
const importedTypedTableControlDocument = await DocumentFile.importDocx(typedTableControlDocx);
assert.deepEqual(
  importedTypedTableControlDocument.contentControls.map((control) => [control.tag, control.controlType, control.placement, control.text]),
  [
    ["TABLE_APPROVED", "checkbox", "tableCell", "☒"],
    ["TABLE_PRIORITY", "dropdown", "tableCell", "High"],
    ["TABLE_CONTACT", "comboBox", "tableCell", "Pager duty"],
    ["TABLE_REVIEW_DATE", "date", "tableCell", "2028-02-29"],
  ],
);
assert.equal(importedTypedTableControlDocument.contentControls.every((control) => Number.isInteger(control.nativeId)), true);
const unchangedTypedTableControlDocx = await DocumentFile.exportDocx(importedTypedTableControlDocument);
assert.deepEqual(Buffer.from(await unchangedTypedTableControlDocx.arrayBuffer()), Buffer.from(await typedTableControlDocx.arrayBuffer()), "unchanged imported typed table-cell content controls must preserve source bytes");
const importedTypedTableDropdown = importedTypedTableControlDocument.contentControls.find((control) => control.tag === "TABLE_PRIORITY");
importedTypedTableDropdown.tag = "TABLE_PRIORITY_FINAL";
importedTypedTableDropdown.alias = "Final table priority";
assert.deepEqual(importedTypedTableControlDocument.setCheckboxContentControls({ TABLE_APPROVED: false }), { updated: 1, matchedTags: ["TABLE_APPROVED"], missingTags: [] });
assert.deepEqual(importedTypedTableControlDocument.setDropdownContentControls({ TABLE_PRIORITY_FINAL: "low" }), { updated: 1, matchedTags: ["TABLE_PRIORITY_FINAL"], missingTags: [] });
assert.deepEqual(importedTypedTableControlDocument.setComboBoxContentControls({ TABLE_CONTACT: "Email" }), { updated: 1, matchedTags: ["TABLE_CONTACT"], missingTags: [] });
assert.deepEqual(importedTypedTableControlDocument.setDateContentControls({ TABLE_REVIEW_DATE: "2030-12-31" }), { updated: 1, matchedTags: ["TABLE_REVIEW_DATE"], missingTags: [] });
const roundTripTypedTableControlDocument = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedTypedTableControlDocument));
assert.deepEqual(roundTripTypedTableControlDocument.blocks.find((block) => block.kind === "table").values.map((row) => row[1]), ["Value", "☐", "Low", "Email", "2030-12-31"]);
assert.equal(roundTripTypedTableControlDocument.contentControls.find((control) => control.tag === "TABLE_PRIORITY_FINAL").alias, "Final table priority");
const typedTableChoiceTamper = await DocumentFile.importDocx(typedTableControlDocx);
typedTableChoiceTamper.blocks.find((block) => block.kind === "table").cells.find((cell) => cell.row === 2 && cell.column === 1).contentControl.choices[0].displayText = "Minor";
await assert.rejects(
  () => DocumentFile.exportDocx(typedTableChoiceTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => typedTableControlTable.getCell(0, 0).addCheckboxContentControl("yes", { tag: "INVALID_TABLE_BOOLEAN", alias: "Invalid boolean" }),
  /checked state must be boolean/i,
);
assert.throws(
  () => typedTableControlTable.getCell(0, 0).addDropdownContentControl(typedTableChoices, { tag: "INVALID_TABLE_CHOICE", alias: "Invalid choice", selectedValue: "missing" }),
  /does not match a choice value/i,
);

const checkboxDocument = DocumentModel.create({ name: "Checkbox content-control profile", blocks: [] });
const checkboxParagraph = checkboxDocument.addParagraph("Terms: ");
checkboxParagraph.addCheckboxContentControl(false, {
  id: "terms-accepted-control",
  tag: "TERMS_ACCEPTED",
  alias: "Terms accepted",
  style: { fontFamily: "Aptos", fontSize: 12 },
});
checkboxParagraph.addRun(" I agree.");
const checkboxControl = checkboxDocument.contentControls[0];
assert.equal(checkboxControl.controlType, "checkbox");
assert.equal(checkboxControl.checked, false);
assert.equal(checkboxControl.text, "☐");
assert.equal(checkboxParagraph.text, "Terms: ☐ I agree.");
assert.throws(() => { checkboxControl.text = "x"; }, /set checked instead/i);
assert.throws(() => checkboxDocument.fillContentControls({ TERMS_ACCEPTED: "yes" }), /Unknown document content-control tag/i);
assert.throws(() => checkboxDocument.setCheckboxContentControls({ TERMS_ACCEPTED: true, MISSING: false }), /Unknown document checkbox content-control tag/i);
assert.equal(checkboxControl.checked, false, "strict checkbox updates must fail before mutation");
assert.deepEqual(checkboxDocument.setCheckboxContentControls({ TERMS_ACCEPTED: true }), { updated: 1, matchedTags: ["TERMS_ACCEPTED"], missingTags: [] });
assert.equal(checkboxControl.checked, true);
assert.equal(checkboxParagraph.text, "Terms: ☒ I agree.");
assert.equal(checkboxDocument.inspect({ kind: "contentControl" }).ndjson.includes('"controlType":"checkbox"'), true);
assert.equal(checkboxDocument.verify().ok, true);
const checkboxDocx = await DocumentFile.exportDocx(checkboxDocument);
const importedCheckboxDocument = await DocumentFile.importDocx(checkboxDocx);
const importedCheckbox = importedCheckboxDocument.contentControls[0];
assert.equal(importedCheckbox.controlType, "checkbox");
assert.equal(importedCheckbox.checked, true);
assert.equal(importedCheckbox.text, "☒");
assert.ok(Number.isInteger(importedCheckbox.nativeId));
assert.deepEqual(importedCheckboxDocument.setCheckboxContentControls({ TERMS_ACCEPTED: false }), { updated: 1, matchedTags: ["TERMS_ACCEPTED"], missingTags: [] });
const editedCheckboxDocx = await DocumentFile.exportDocx(importedCheckboxDocument);
const roundTripCheckboxDocument = await DocumentFile.importDocx(editedCheckboxDocx);
assert.equal(roundTripCheckboxDocument.contentControls[0].checked, false);
assert.equal(roundTripCheckboxDocument.contentControls[0].text, "☐");
assert.equal(roundTripCheckboxDocument.resolve(roundTripCheckboxDocument.contentControls[0].targetId).text, "Terms: ☐ I agree.");
const checkboxTypeTamper = await DocumentFile.importDocx(checkboxDocx);
checkboxTypeTamper.resolve(checkboxTypeTamper.contentControls[0].targetId).runs[checkboxTypeTamper.contentControls[0].runIndex].contentControl.controlType = "text";
await assert.rejects(
  () => DocumentFile.exportDocx(checkboxTypeTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => checkboxDocument.addParagraph("", { runs: [{ text: "X", contentControl: { tag: "INVALID_CHECKBOX", controlType: "checkbox", checked: false } }] }),
  /text is codec-owned/i,
);

const dropdownDocument = DocumentModel.create({ name: "Drop-down content-control profile", blocks: [] });
const dropdownParagraph = dropdownDocument.addParagraph("Priority: ");
dropdownParagraph.addDropdownContentControl([
  { displayText: "Low", value: "low" },
  { displayText: "Medium", value: "medium" },
  { displayText: "High", value: "high" },
], {
  id: "priority-control",
  tag: "PRIORITY",
  alias: "Priority",
  selectedValue: "medium",
  style: { fontFamily: "Aptos", fontSize: 12 },
});
dropdownParagraph.addRun(".");
const dropdownControl = dropdownDocument.contentControls[0];
assert.equal(dropdownControl.controlType, "dropdown");
assert.equal(dropdownControl.selectedValue, "medium");
assert.equal(dropdownControl.text, "Medium");
assert.equal(dropdownParagraph.text, "Priority: Medium.");
assert.deepEqual(dropdownControl.choices, [
  { displayText: "Low", value: "low" },
  { displayText: "Medium", value: "medium" },
  { displayText: "High", value: "high" },
]);
const defensiveChoices = dropdownControl.choices;
defensiveChoices[0].displayText = "Changed copy";
assert.equal(dropdownControl.choices[0].displayText, "Low");
assert.throws(() => { dropdownControl.text = "High"; }, /set selectedValue instead/i);
assert.throws(() => dropdownDocument.fillContentControls({ PRIORITY: "high" }), /Unknown document content-control tag/i);
assert.throws(() => dropdownDocument.setCheckboxContentControls({ PRIORITY: true }), /Unknown document checkbox content-control tag/i);
assert.throws(() => dropdownDocument.setDropdownContentControls({ PRIORITY: "high", MISSING: "low" }), /Unknown document drop-down content-control tag/i);
assert.equal(dropdownControl.selectedValue, "medium", "strict drop-down updates must fail before mutation");
assert.throws(() => dropdownDocument.setDropdownContentControls({ PRIORITY: "urgent" }), /does not match a choice value/i);
assert.equal(dropdownControl.selectedValue, "medium", "invalid drop-down selections must fail before mutation");
assert.throws(() => dropdownDocument.setDropdownContentControls({ PRIORITY: 1 }), /selectedValue must be a string/i);
assert.equal(dropdownControl.selectedValue, "medium", "non-string drop-down selections must fail before mutation");
assert.deepEqual(dropdownDocument.setDropdownContentControls({ PRIORITY: "high" }), { updated: 1, matchedTags: ["PRIORITY"], missingTags: [] });
assert.equal(dropdownControl.selectedValue, "high");
assert.equal(dropdownControl.text, "High");
assert.equal(dropdownParagraph.text, "Priority: High.");
assert.match(dropdownDocument.inspect({ kind: "contentControl" }).ndjson, /"controlType":"dropdown"/);
assert.match(dropdownDocument.inspect({ kind: "contentControl" }).ndjson, /"selectedValue":"high"/);
assert.equal(dropdownDocument.verify().ok, true);
const dropdownDocx = await DocumentFile.exportDocx(dropdownDocument);
const importedDropdownDocument = await DocumentFile.importDocx(dropdownDocx);
const importedDropdown = importedDropdownDocument.contentControls[0];
assert.equal(importedDropdown.controlType, "dropdown");
assert.equal(importedDropdown.selectedValue, "high");
assert.equal(importedDropdown.text, "High");
assert.deepEqual(importedDropdown.choices.map((choice) => choice.value), ["low", "medium", "high"]);
assert.ok(Number.isInteger(importedDropdown.nativeId));
assert.deepEqual(importedDropdownDocument.setDropdownContentControls({ PRIORITY: "low" }), { updated: 1, matchedTags: ["PRIORITY"], missingTags: [] });
const editedDropdownDocx = await DocumentFile.exportDocx(importedDropdownDocument);
const roundTripDropdownDocument = await DocumentFile.importDocx(editedDropdownDocx);
assert.equal(roundTripDropdownDocument.contentControls[0].selectedValue, "low");
assert.equal(roundTripDropdownDocument.contentControls[0].text, "Low");
assert.equal(roundTripDropdownDocument.resolve(roundTripDropdownDocument.contentControls[0].targetId).text, "Priority: Low.");
const dropdownChoiceTamper = await DocumentFile.importDocx(dropdownDocx);
dropdownChoiceTamper.resolve(dropdownChoiceTamper.contentControls[0].targetId).runs[dropdownChoiceTamper.contentControls[0].runIndex].contentControl.choices[0].displayText = "Routine";
await assert.rejects(
  () => DocumentFile.exportDocx(dropdownChoiceTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
const dropdownTypeTamper = await DocumentFile.importDocx(dropdownDocx);
dropdownTypeTamper.resolve(dropdownTypeTamper.contentControls[0].targetId).runs[dropdownTypeTamper.contentControls[0].runIndex].contentControl.controlType = "text";
await assert.rejects(
  () => DocumentFile.exportDocx(dropdownTypeTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => dropdownDocument.addParagraph("", { runs: [{ contentControl: { tag: "INVALID_DROPDOWN", controlType: "dropdown", choices: [{ displayText: "Same", value: "a" }, { displayText: "Same", value: "b" }] } }] }),
  /must be unique/i,
);
assert.throws(
  () => dropdownDocument.addParagraph("", { runs: [{ text: "Visible override", contentControl: { tag: "INVALID_DROPDOWN_TEXT", controlType: "dropdown", choices: ["Canonical"] } }] }),
  /text is codec-owned/i,
);
assert.throws(
  () => dropdownDocument.addParagraph("", { runs: [{ contentControl: { tag: "INVALID_DROPDOWN_VALUE", controlType: "dropdown", choices: [{ displayText: "One", value: 1 }] } }] }),
  /displayText and value must be strings/i,
);

const comboBoxDocument = DocumentModel.create({ name: "Combo-box content-control profile", blocks: [] });
const comboBoxParagraph = comboBoxDocument.addParagraph("Contact method: ");
comboBoxParagraph.addComboBoxContentControl([
  { displayText: "Email", value: "email" },
  { displayText: "Phone call", value: "phone" },
], {
  id: "contact-method-control",
  tag: "CONTACT_METHOD",
  alias: "Contact method",
  value: "email",
  style: { fontFamily: "Aptos", fontSize: 12 },
});
comboBoxParagraph.addRun(".");
const comboBoxControl = comboBoxDocument.contentControls[0];
assert.equal(comboBoxControl.controlType, "comboBox");
assert.equal(comboBoxControl.value, "email");
assert.equal(comboBoxControl.text, "Email");
assert.equal(comboBoxParagraph.text, "Contact method: Email.");
assert.deepEqual(comboBoxControl.choices, [
  { displayText: "Email", value: "email" },
  { displayText: "Phone call", value: "phone" },
]);
const defensiveComboBoxChoices = comboBoxControl.choices;
defensiveComboBoxChoices[0].displayText = "Changed copy";
assert.equal(comboBoxControl.choices[0].displayText, "Email");
assert.throws(() => { comboBoxControl.text = "Pager duty"; }, /set value instead/i);
assert.throws(() => comboBoxDocument.fillContentControls({ CONTACT_METHOD: "Pager duty" }), /Unknown document content-control tag/i);
assert.throws(() => comboBoxDocument.setDropdownContentControls({ CONTACT_METHOD: "email" }), /Unknown document drop-down content-control tag/i);
assert.throws(() => comboBoxDocument.setComboBoxContentControls({ CONTACT_METHOD: "Pager duty", MISSING: "Other" }), /Unknown document combo-box content-control tag/i);
assert.equal(comboBoxControl.value, "email", "strict combo-box updates must fail before mutation");
assert.throws(() => comboBoxDocument.setComboBoxContentControls({ CONTACT_METHOD: 1 }), /value must be a string/i);
assert.throws(() => comboBoxDocument.setComboBoxContentControls({ CONTACT_METHOD: "" }), /1 through 255 characters/i);
assert.equal(comboBoxControl.value, "email", "invalid combo-box values must fail before mutation");
assert.deepEqual(comboBoxDocument.setComboBoxContentControls({ CONTACT_METHOD: "Pager duty" }), { updated: 1, matchedTags: ["CONTACT_METHOD"], missingTags: [] });
assert.equal(comboBoxControl.value, "Pager duty");
assert.equal(comboBoxControl.text, "Pager duty");
assert.equal(comboBoxParagraph.text, "Contact method: Pager duty.");
assert.match(comboBoxDocument.inspect({ kind: "contentControl" }).ndjson, /"controlType":"comboBox"/);
assert.match(comboBoxDocument.inspect({ kind: "contentControl" }).ndjson, /"value":"Pager duty"/);
assert.equal(comboBoxDocument.verify().ok, true);
const comboBoxDocx = await DocumentFile.exportDocx(comboBoxDocument);
const importedComboBoxDocument = await DocumentFile.importDocx(comboBoxDocx);
const importedComboBox = importedComboBoxDocument.contentControls[0];
assert.equal(importedComboBox.controlType, "comboBox");
assert.equal(importedComboBox.value, "Pager duty");
assert.equal(importedComboBox.text, "Pager duty");
assert.deepEqual(importedComboBox.choices.map((choice) => choice.value), ["email", "phone"]);
assert.ok(Number.isInteger(importedComboBox.nativeId));
assert.deepEqual(importedComboBoxDocument.setComboBoxContentControls({ CONTACT_METHOD: "phone" }), { updated: 1, matchedTags: ["CONTACT_METHOD"], missingTags: [] });
assert.equal(importedComboBoxDocument.contentControls[0].text, "Phone call");
const editedComboBoxDocx = await DocumentFile.exportDocx(importedComboBoxDocument);
const roundTripComboBoxDocument = await DocumentFile.importDocx(editedComboBoxDocx);
assert.equal(roundTripComboBoxDocument.contentControls[0].value, "phone");
assert.equal(roundTripComboBoxDocument.contentControls[0].text, "Phone call");
assert.equal(roundTripComboBoxDocument.resolve(roundTripComboBoxDocument.contentControls[0].targetId).text, "Contact method: Phone call.");
const comboBoxChoiceTamper = await DocumentFile.importDocx(comboBoxDocx);
comboBoxChoiceTamper.resolve(comboBoxChoiceTamper.contentControls[0].targetId).runs[comboBoxChoiceTamper.contentControls[0].runIndex].contentControl.choices[0].displayText = "Electronic mail";
await assert.rejects(
  () => DocumentFile.exportDocx(comboBoxChoiceTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
const comboBoxTypeTamper = await DocumentFile.importDocx(comboBoxDocx);
comboBoxTypeTamper.resolve(comboBoxTypeTamper.contentControls[0].targetId).runs[comboBoxTypeTamper.contentControls[0].runIndex].contentControl.controlType = "dropdown";
await assert.rejects(
  () => DocumentFile.exportDocx(comboBoxTypeTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => comboBoxDocument.addParagraph("", { runs: [{ text: "Visible override", contentControl: { tag: "INVALID_COMBO_TEXT", controlType: "comboBox", choices: ["Canonical"], value: "Custom" } }] }),
  /text is codec-owned/i,
);
assert.throws(
  () => comboBoxDocument.addParagraph("", { runs: [{ contentControl: { tag: "INVALID_COMBO_CHOICES", controlType: "comboBox", choices: [{ displayText: "Same", value: "a" }, { displayText: "Same", value: "b" }] } }] }),
  /must be unique/i,
);

const dateDocument = DocumentModel.create({ name: "Date content-control profile", blocks: [] });
const dateParagraph = dateDocument.addParagraph("Review date: ");
dateParagraph.addDateContentControl("2026-07-21", {
  id: "review-date-control",
  tag: "REVIEW_DATE",
  alias: "Review date",
  style: { fontFamily: "Aptos", fontSize: 12 },
});
dateParagraph.addRun(".");
const dateControl = dateDocument.contentControls[0];
assert.equal(dateControl.controlType, "date");
assert.equal(dateControl.dateValue, "2026-07-21");
assert.equal(dateControl.text, "2026-07-21");
assert.equal(dateParagraph.text, "Review date: 2026-07-21.");
assert.throws(() => { dateControl.text = "July 21, 2026"; }, /set dateValue instead/i);
assert.throws(() => dateDocument.fillContentControls({ REVIEW_DATE: "2026-08-01" }), /Unknown document content-control tag/i);
assert.throws(() => dateDocument.setComboBoxContentControls({ REVIEW_DATE: "2026-08-01" }), /Unknown document combo-box content-control tag/i);
assert.throws(() => dateDocument.setDateContentControls({ REVIEW_DATE: "2026-08-01", MISSING: "2026-08-02" }), /Unknown document date content-control tag/i);
assert.equal(dateControl.dateValue, "2026-07-21", "strict date updates must fail before mutation");
for (const invalid of ["2026-7-21", "2026-02-29", "2026-04-31", "0000-01-01", new Date("2026-07-21T00:00:00Z")]) {
  assert.throws(() => dateDocument.setDateContentControls({ REVIEW_DATE: invalid }), /dateValue must/i);
  assert.equal(dateControl.dateValue, "2026-07-21", "invalid date updates must fail before mutation");
}
assert.deepEqual(dateDocument.setDateContentControls({ REVIEW_DATE: "2028-02-29" }), { updated: 1, matchedTags: ["REVIEW_DATE"], missingTags: [] });
assert.equal(dateControl.dateValue, "2028-02-29");
assert.equal(dateControl.text, "2028-02-29");
assert.equal(dateParagraph.text, "Review date: 2028-02-29.");
assert.match(dateDocument.inspect({ kind: "contentControl" }).ndjson, /"controlType":"date"/);
assert.match(dateDocument.inspect({ kind: "contentControl" }).ndjson, /"dateValue":"2028-02-29"/);
assert.equal(dateDocument.verify().ok, true);
const dateDocx = await DocumentFile.exportDocx(dateDocument);
const importedDateDocument = await DocumentFile.importDocx(dateDocx);
const importedDate = importedDateDocument.contentControls[0];
assert.equal(importedDate.controlType, "date");
assert.equal(importedDate.dateValue, "2028-02-29");
assert.equal(importedDate.text, "2028-02-29");
assert.ok(Number.isInteger(importedDate.nativeId));
assert.deepEqual(importedDateDocument.setDateContentControls({ REVIEW_DATE: "2027-12-31" }), { updated: 1, matchedTags: ["REVIEW_DATE"], missingTags: [] });
const editedDateDocx = await DocumentFile.exportDocx(importedDateDocument);
const roundTripDateDocument = await DocumentFile.importDocx(editedDateDocx);
assert.equal(roundTripDateDocument.contentControls[0].dateValue, "2027-12-31");
assert.equal(roundTripDateDocument.contentControls[0].text, "2027-12-31");
assert.equal(roundTripDateDocument.resolve(roundTripDateDocument.contentControls[0].targetId).text, "Review date: 2027-12-31.");
const dateTypeTamper = await DocumentFile.importDocx(dateDocx);
dateTypeTamper.resolve(dateTypeTamper.contentControls[0].targetId).runs[dateTypeTamper.contentControls[0].runIndex].contentControl.controlType = "text";
await assert.rejects(
  () => DocumentFile.exportDocx(dateTypeTamper),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);
assert.throws(
  () => dateDocument.addParagraph("", { runs: [{ text: "July 21, 2026", contentControl: { tag: "INVALID_DATE_TEXT", controlType: "date", dateValue: "2026-07-21" } }] }),
  /text is codec-owned/i,
);

const modernCommentDocument = DocumentModel.create({ name: "Modern comment thread", blocks: [] });
const modernCommentTarget = modernCommentDocument.addParagraph("Review the bounded modern comment thread.");
const modernRoot = modernCommentDocument.addComment(modernCommentTarget, "Please confirm the evidence.", {
  author: "Lead reviewer",
  initials: "LR",
  date: "2026-07-19T08:00:00Z",
  resolved: false,
  paraId: "11111111",
  durableId: "33333333",
  dateUtc: "2026-07-19T08:00:00Z",
  person: { providerId: "provider-a", userId: "lead@example.test" },
});
const modernReply = modernCommentDocument.replyToComment(modernRoot, "Evidence confirmed.", {
  author: "Second reviewer",
  initials: "SR",
  date: "2026-07-19T08:05:00Z",
  paraId: "22222222",
  durableId: "44444444",
  dateUtc: "2026-07-19T08:05:00Z",
  person: { providerId: "provider-b", userId: "second@example.test" },
});
assert.equal(modernReply.parentId, modernRoot.id);
const modernDocx = await DocumentFile.exportDocx(modernCommentDocument);
const importedModernComments = await DocumentFile.importDocx(modernDocx);
assert.equal(importedModernComments.comments.length, 2);
assert.equal(importedModernComments.comments[0].resolved, false);
assert.equal(importedModernComments.comments[0].paraId, "11111111");
assert.equal(importedModernComments.comments[0].durableId, "33333333");
assert.deepEqual(importedModernComments.comments[0].person, { providerId: "provider-a", userId: "lead@example.test" });
assert.equal(importedModernComments.comments[1].parentId, importedModernComments.comments[0].id);
assert.equal(importedModernComments.comments[1].targetId, importedModernComments.comments[0].targetId);
assert.match(importedModernComments.inspect({ kind: "comment" }).ndjson, /"parentId":"document\/comment\/1"/);
importedModernComments.comments[0].text = "Resolved after public-facade review.";
importedModernComments.comments[0].resolve();
importedModernComments.comments[1].text = "Reply retained with the root.";
const editedModernDocx = await DocumentFile.exportDocx(importedModernComments);
const roundTripModernComments = await DocumentFile.importDocx(editedModernDocx);
assert.equal(roundTripModernComments.comments[0].resolved, true);
assert.equal(roundTripModernComments.comments[0].text, "Resolved after public-facade review.");
assert.equal(roundTripModernComments.comments[1].text, "Reply retained with the root.");
assert.equal(roundTripModernComments.comments[1].parentId, roundTripModernComments.comments[0].id);
roundTripModernComments.comments[0].reopen();
assert.equal(roundTripModernComments.comments[0].resolved, false);

const explicitOpenCommentDocument = DocumentModel.create({ name: "Explicit open modern comment", blocks: [] });
const explicitOpenTarget = explicitOpenCommentDocument.addParagraph("Explicit resolved=false selects the modern graph.");
explicitOpenCommentDocument.addComment(explicitOpenTarget, "Open modern root", {
  author: "Reviewer",
  resolved: false,
});
const explicitOpenDocx = await DocumentFile.exportDocx(explicitOpenCommentDocument);
const importedExplicitOpen = await DocumentFile.importDocx(explicitOpenDocx);
assert.equal(importedExplicitOpen.comments[0].resolved, false);
assert.match(importedExplicitOpen.comments[0].paraId, /^[0-9A-F]{8}$/);
assert.ok(Number.parseInt(importedExplicitOpen.comments[0].paraId, 16) > 0);
assert.ok(Number.parseInt(importedExplicitOpen.comments[0].paraId, 16) < 0x80000000);
assert.equal(importedExplicitOpen.comments[0]._resolvedSpecified, true);

const invalidPackageParaIdZip = await JSZip.loadAsync(Buffer.from(await modernDocx.arrayBuffer()));
for (const partPath of ["word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml"]) {
  const part = invalidPackageParaIdZip.file(partPath);
  if (part) invalidPackageParaIdZip.file(partPath, (await part.async("text")).replaceAll("11111111", "00000000"));
}
const invalidPackageParaIdInspection = await DocumentFile.inspectDocx(await invalidPackageParaIdZip.generateAsync({ type: "uint8array" }));
assert.equal(invalidPackageParaIdInspection.ok, false);
assert.equal(invalidPackageParaIdInspection.issues.some((issue) => issue.type === "docxCommentParaIdMissing"), true);
assert.equal(invalidPackageParaIdInspection.issues.some((issue) => issue.type === "docxCommentExParaIdInvalid"), true);
assert.equal(invalidPackageParaIdInspection.issues.some((issue) => issue.type === "docxCommentIdParaIdInvalid"), true);

const nestedCommentDocument = DocumentModel.create({ name: "Nested comment rejection", blocks: [] });
const nestedTarget = nestedCommentDocument.addParagraph("Nested replies fail closed.");
const nestedRoot = nestedCommentDocument.addComment(nestedTarget, "Root", { author: "Reviewer" });
const directReply = nestedCommentDocument.replyToComment(nestedRoot, "Direct", { author: "Reviewer" });
nestedCommentDocument.replyToComment(directReply, "Nested", { author: "Reviewer" });
await assert.rejects(
  () => DocumentFile.exportDocx(nestedCommentDocument),
  (error) => error?.code === "unsupported_document_comment_thread" && /nested reply/i.test(error.message),
);

const importedModernMetadataEdit = await DocumentFile.importDocx(modernDocx);
importedModernMetadataEdit.comments[0].person.userId = "tampered@example.test";
await assert.rejects(
  () => DocumentFile.exportDocx(importedModernMetadataEdit),
  (error) => error?.code === "unsupported_document_comment_edit" && /source-bound/i.test(error.message),
);

const importedModernTopologyEdit = await DocumentFile.importDocx(modernDocx);
importedModernTopologyEdit.replyToComment(importedModernTopologyEdit.comments[0], "A new imported reply is outside the fixed topology.", { author: "Reviewer" });
await assert.rejects(
  () => DocumentFile.exportDocx(importedModernTopologyEdit),
  (error) => error?.code === "document_comment_topology_changed",
);

const invalidDurableCommentDocument = DocumentModel.create({ name: "Invalid durable comment", blocks: [] });
const invalidDurableTarget = invalidDurableCommentDocument.addParagraph("Invalid durable IDs fail closed.");
invalidDurableCommentDocument.addComment(invalidDurableTarget, "Invalid durable ID", {
  author: "Reviewer",
  resolved: false,
  durableId: "FFFFFFFF",
});
await assert.rejects(
  () => DocumentFile.exportDocx(invalidDurableCommentDocument),
  (error) => error?.code === "invalid_document_comment" && /00000001.*7FFFFFFE/i.test(error.message),
);

for (const paraId of ["00000000", "80000000"]) {
  const invalidParaCommentDocument = DocumentModel.create({ name: "Invalid paragraph comment identity", blocks: [] });
  const invalidParaTarget = invalidParaCommentDocument.addParagraph("Paragraph identities must remain inside the Open XML range.");
  invalidParaCommentDocument.addComment(invalidParaTarget, "Invalid paragraph identity", {
    author: "Reviewer",
    resolved: false,
    paraId,
  });
  assert.equal(invalidParaCommentDocument.verify().issues.some((issue) => issue.type === "invalidCommentParaId"), true);
  await assert.rejects(
    () => DocumentFile.exportDocx(invalidParaCommentDocument),
    (error) => error?.code === "invalid_document_comment" && /00000001.*7FFFFFFF/i.test(error.message),
  );
}

const inconsistentPeopleDocument = DocumentModel.create({ name: "Inconsistent people", blocks: [] });
const inconsistentPeopleTarget = inconsistentPeopleDocument.addParagraph("People metadata must remain author-consistent.");
const inconsistentPeopleRoot = inconsistentPeopleDocument.addComment(inconsistentPeopleTarget, "Root", {
  author: "Same reviewer",
  resolved: false,
  person: { providerId: "directory", userId: "reviewer@example.test" },
});
inconsistentPeopleDocument.replyToComment(inconsistentPeopleRoot, "Reply", { author: "Same reviewer" });
await assert.rejects(
  () => DocumentFile.exportDocx(inconsistentPeopleDocument),
  (error) => error?.code === "invalid_document_comment" && /inconsistent people metadata/i.test(error.message),
);

const importedWithChangedContentControlTopology = await DocumentFile.importDocx(firstDocx);
const removedControl = importedWithChangedContentControlTopology.contentControls[0];
delete importedWithChangedContentControlTopology.resolve(removedControl.targetId).runs[removedControl.runIndex].contentControl;
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithChangedContentControlTopology),
  (error) => error?.code === "document_content_control_topology_changed" && /source-bound/i.test(error.message),
);

const importedWithRenamedBookmark = await DocumentFile.importDocx(firstDocx);
importedWithRenamedBookmark.bookmarks[0].name = "RenamedSection";
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithRenamedBookmark),
  (error) => error?.code === "unsupported_document_bookmark_edit" && /source-bound/i.test(error.message),
);

const invalidBookmarkRange = DocumentModel.create({ blocks: [
  { kind: "paragraph", text: "Start" },
  { kind: "paragraph", text: "End" },
] });
invalidBookmarkRange.addBookmark(invalidBookmarkRange.blocks[0], "CrossBlock", { endTarget: invalidBookmarkRange.blocks[1] });
await assert.rejects(
  () => DocumentFile.exportDocx(invalidBookmarkRange),
  (error) => error?.code === "invalid_document_bookmark" && /exactly one block/i.test(error.message),
);

const importedWithChangedRevisionKind = await DocumentFile.importDocx(firstDocx);
importedWithChangedRevisionKind.blocks.find((block) => block.kind === "change").changeType = "delete";
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithChangedRevisionKind),
  (error) => error?.code === "unsupported_document_edit" && /insertion\/deletion kind/i.test(error.message),
);

const bibliographyDocument = DocumentModel.create({
  name: "Bounded bibliography and citation",
  bibliography: { selectedStyle: "\\APASixthEditionOfficeOnline.xsl", styleName: "APA" },
  blocks: [],
});
bibliographyDocument.addBibliographySource({
  id: "bibliography/AgentSource",
  tag: "AgentSource",
  sourceType: "Book",
  title: "Sketch of the Analytical Engine",
  year: "1843",
  publisher: "Scientific Memoirs",
  authors: [{ first: "Ada", last: "Lovelace" }],
});
bibliographyDocument.addCitation("(Lovelace, 1843)", { tag: "AgentSource" }, { id: "citation/agent-source" });
const bibliographyDocx = await DocumentFile.exportDocx(bibliographyDocument);
const importedBibliography = await DocumentFile.importDocx(bibliographyDocx);
assert.equal(importedBibliography.bibliography.styleName, "APA");
assert.equal(importedBibliography.bibliographySources.length, 1);
assert.equal(importedBibliography.bibliographySources[0].title, "Sketch of the Analytical Engine");
assert.deepEqual(importedBibliography.bibliographySources[0].authors, [{ first: "Ada", middle: "", last: "Lovelace" }]);
assert.equal(importedBibliography.blocks[0].kind, "citation");
assert.equal(importedBibliography.blocks[0].metadata.tag, "AgentSource");
assert.equal(importedBibliography.bookmarks[0].targetId, importedBibliography.blocks[0].id);
importedBibliography.bibliographySources[0].title = "Notes on the Analytical Engine";
importedBibliography.bibliographySources[0].authors[0].first = "Augusta Ada";
importedBibliography.blocks[0].text = "(Lovelace, 1843, revised)";
const roundTripBibliography = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedBibliography));
assert.equal(roundTripBibliography.bibliographySources[0].title, "Notes on the Analytical Engine");
assert.equal(roundTripBibliography.bibliographySources[0].authors[0].first, "Augusta Ada");
assert.equal(roundTripBibliography.blocks[0].text, "(Lovelace, 1843, revised)");

roundTripBibliography.bibliographySources[0].tag = "RenamedSource";
roundTripBibliography.blocks[0].metadata.tag = "RenamedSource";
await assert.rejects(
  () => DocumentFile.exportDocx(roundTripBibliography),
  (error) => new Set(["unsupported_document_bibliography_edit", "unsupported_document_edit"]).has(error?.code) && /source-bound/i.test(error.message),
);

const tocDocument = DocumentModel.create({
  name: "Bounded native TOC field",
  blocks: [
    { kind: "paragraph", styleId: "Title", text: "Native TOC workflow" },
    { kind: "paragraph", styleId: "Heading1", text: "First section" },
    { kind: "paragraph", styleId: "Heading2", text: "Nested section" },
  ],
});
const tocField = tocDocument.addTableOfContents({ levels: "1-3", display: "Refresh fields before delivery" });
assert.equal(tocField.complex, true);
assert.equal(tocField.instruction, 'TOC \\o "1-3" \\h \\z \\u');
assert.equal(tocDocument.settings.updateFields, true);
const tocDocx = await DocumentFile.exportDocx(tocDocument);
const importedToc = await DocumentFile.importDocx(tocDocx);
const importedTocField = importedToc.blocks.find((block) => block.kind === "field");
assert.equal(importedToc.settings.updateFields, true);
assert.equal(importedTocField?.complex, true);
assert.equal(importedTocField?.instruction, 'TOC \\o "1-3" \\h \\z \\u');
assert.equal(importedTocField?.display, "Refresh fields before delivery");
assert.match(importedToc.inspect({ kind: "field,settings", maxChars: 4_000 }).ndjson, /"complex":true/);
importedTocField.instruction = 'TOC \\o "1-4" \\h \\z \\u';
importedTocField.display = "Update this TOC in Word";
importedToc.setSettings({ updateFields: false });
const editedTocDocx = await DocumentFile.exportDocx(importedToc);
const roundTripToc = await DocumentFile.importDocx(editedTocDocx);
assert.equal(roundTripToc.settings.updateFields, false);
assert.equal(roundTripToc.blocks.find((block) => block.kind === "field")?.instruction, 'TOC \\o "1-4" \\h \\z \\u');
assert.equal(roundTripToc.blocks.find((block) => block.kind === "field")?.display, "Update this TOC in Word");
assert.throws(() => tocDocument.addTableOfContents({ levels: "4-2" }), /ascending range/);

const inlineFieldDocument = DocumentModel.create({ name: "Inline fields", blocks: [] });
const inlineFieldParagraph = inlineFieldDocument.addParagraph("", { name: "inline-field-caption", styleId: "Caption" });
inlineFieldParagraph.addRun("Figure ");
inlineFieldParagraph.addField("SEQ Figure \\* ARABIC", "0", { bookmarkName: "fig1", style: { bold: true } });
inlineFieldParagraph.addRun(": Revenue. See ");
inlineFieldParagraph.addField("REF fig1 \\h", "0");
inlineFieldParagraph.addRun(".");
assert.equal(inlineFieldParagraph.text, "Figure 0: Revenue. See 0.");
assert.equal(inlineFieldParagraph.runs[1].inlineField.instruction, "SEQ Figure \\* ARABIC");
assert.equal(inlineFieldParagraph.runs[1].inlineField.bookmarkName, "fig1");
const inlineFieldDocx = await DocumentFile.exportDocx(inlineFieldDocument);
const importedInlineFieldDocument = await DocumentFile.importDocx(inlineFieldDocx);
const importedInlineFieldParagraph = importedInlineFieldDocument.blocks[0];
assert.equal(importedInlineFieldParagraph.runs.length, 5);
assert.equal(importedInlineFieldParagraph.runs[1].inlineField.instruction, "SEQ Figure \\* ARABIC");
assert.equal(importedInlineFieldParagraph.runs[1].inlineField.bookmarkName, "fig1");
assert.equal(importedInlineFieldParagraph.runs[1].inlineField.bookmarkNativeId, 0);
assert.equal(importedInlineFieldParagraph.runs[3].inlineField.instruction, "REF fig1 \\h");
const inlineMaterializationDryRun = importedInlineFieldDocument.materializeFields({ dryRun: true });
assert.equal(inlineMaterializationDryRun.updated, 0);
assert.equal(inlineMaterializationDryRun.wouldUpdate, 2);
assert.equal(inlineMaterializationDryRun.seqFields, 1);
assert.equal(inlineMaterializationDryRun.refFields, 1);
assert.equal(importedInlineFieldParagraph.text, "Figure 0: Revenue. See 0.");
assert.deepEqual(importedInlineFieldDocument.materializeFields(), {
  dryRun: false,
  updated: 2,
  wouldUpdate: 2,
  seqFields: 1,
  refFields: 1,
  skippedPageReferences: 0,
  missingBookmarks: [],
  changes: inlineMaterializationDryRun.changes,
});
importedInlineFieldParagraph.runs[2].text = ": Updated revenue. See ";
importedInlineFieldParagraph.text = importedInlineFieldParagraph.runs.map((run) => run.text).join("");
const editedInlineFieldDocx = await DocumentFile.exportDocx(importedInlineFieldDocument);
const roundTripInlineFieldDocument = await DocumentFile.importDocx(editedInlineFieldDocx);
assert.equal(roundTripInlineFieldDocument.blocks[0].text, "Figure 1: Updated revenue. See 1.");
assert.equal(roundTripInlineFieldDocument.blocks[0].runs[1].inlineField.bookmarkName, "fig1");
importedInlineFieldParagraph.runs[1].inlineField.bookmarkName = "fig2";
await assert.rejects(
  () => DocumentFile.exportDocx(importedInlineFieldDocument),
  (error) => error?.code === "document_inline_field_topology_changed" && /source-bound/i.test(error.message),
);
importedInlineFieldParagraph.runs[1].inlineField.bookmarkName = "fig1";
importedInlineFieldParagraph.runs[3].inlineField.instruction = "REF fig2 \\h";
await assert.rejects(
  () => DocumentFile.exportDocx(importedInlineFieldDocument),
  (error) => error?.code === "document_inline_field_topology_changed" && /source-bound/i.test(error.message),
);
const invalidInlineFieldDocument = DocumentModel.create({ blocks: [] });
invalidInlineFieldDocument.addParagraph("", { runs: [{ text: "0", inlineField: { instruction: "SEQ Figure \\* ROMAN" } }] });
await assert.rejects(
  () => DocumentFile.exportDocx(invalidInlineFieldDocument),
  (error) => error?.code === "invalid_document_inline_field" && /canonical SEQ/i.test(error.message),
);
const invalidInlineBookmarkDocument = DocumentModel.create({ blocks: [] });
invalidInlineBookmarkDocument.addParagraph("", { runs: [{ text: "0", inlineField: { instruction: "REF fig1 \\h", bookmarkName: "fig1" } }] });
await assert.rejects(
  () => DocumentFile.exportDocx(invalidInlineBookmarkDocument),
  (error) => error?.code === "invalid_document_inline_field" && /bookmark only a canonical SEQ/i.test(error.message),
);
const duplicateInlineBookmarkDocument = DocumentModel.create({ blocks: [] });
duplicateInlineBookmarkDocument.addParagraph("", { runs: [{ text: "0", inlineField: { instruction: "SEQ Figure \\* ARABIC", bookmarkName: "fig1" } }] });
const duplicateBookmarkTarget = duplicateInlineBookmarkDocument.addParagraph("Target");
duplicateInlineBookmarkDocument.addBookmark(duplicateBookmarkTarget, "FIG1");
await assert.rejects(
  () => DocumentFile.exportDocx(duplicateInlineBookmarkDocument),
  (error) => error?.code === "invalid_document_bookmark" && /duplicated/i.test(error.message),
);
const unresolvedFieldDocument = DocumentModel.create({ blocks: [] });
const unresolvedFieldParagraph = unresolvedFieldDocument.addParagraph("");
unresolvedFieldParagraph.addField("SEQ Figure \\* ARABIC", "0", { bookmarkName: "fig1" });
unresolvedFieldParagraph.addField("REF missingTarget \\h", "0");
assert.throws(() => unresolvedFieldDocument.materializeFields(), /cannot resolve bookmark.*missingTarget/i);
assert.equal(unresolvedFieldParagraph.text, "00", "strict field materialization must be transactional");
assert.deepEqual(unresolvedFieldDocument.materializeFields({ strict: false }).missingBookmarks, ["missingTarget"]);
assert.equal(unresolvedFieldParagraph.text, "10", "non-strict materialization may update resolvable fields while reporting missing targets");
assert.throws(() => unresolvedFieldDocument.materializeFields({ types: ["PAGEREF"] }), /requires a real pagination host/i);
const multiSequenceDocument = DocumentModel.create({ blocks: [] });
const firstFigure = multiSequenceDocument.addParagraph("");
firstFigure.addField("SEQ Figure \\* ARABIC", "0", { bookmarkName: "fig1" });
const secondFigure = multiSequenceDocument.addParagraph("");
secondFigure.addField("SEQ Figure \\* ARABIC", "0", { bookmarkName: "fig2" });
const secondFigureReference = multiSequenceDocument.addParagraph("");
secondFigureReference.addField("REF FIG2 \\h", "0");
assert.equal(multiSequenceDocument.materializeFields().updated, 3);
assert.deepEqual([firstFigure.text, secondFigure.text, secondFigureReference.text], ["1", "2", "2"]);

const invalidComplexField = DocumentModel.create({ blocks: [] });
invalidComplexField.addField('TOC \\o "1-3" \\p "custom separator"', "Unsafe switches", { complex: true });
await assert.rejects(
  () => DocumentFile.exportDocx(invalidComplexField),
  (error) => error?.code === "invalid_document_field" && /canonical bounded profile/i.test(error.message),
);

const trackedSettings = DocumentModel.create({
  name: "Tracked document settings",
  settings: { trackRevisions: true },
  blocks: [{ kind: "paragraph", text: "Tracking is independent from existing revisions." }],
});
const trackedSettingsDocx = await DocumentFile.exportDocx(trackedSettings);
const importedTrackedSettings = await DocumentFile.importDocx(trackedSettingsDocx);
assert.equal(importedTrackedSettings.settings.trackRevisions, true);
importedTrackedSettings.setSettings({ trackRevisions: false });
const untrackedSettingsDocx = await DocumentFile.exportDocx(importedTrackedSettings);
assert.equal((await DocumentFile.importDocx(untrackedSettingsDocx)).settings.trackRevisions, false);

const mirroredSettings = DocumentModel.create({
  name: "Facing-page mirror margins",
  settings: { mirrorMargins: true },
  blocks: [{ kind: "paragraph", text: "Inside and outside margins alternate on facing pages." }],
});
const mirroredDocx = await DocumentFile.exportDocx(mirroredSettings);
const mirroredZip = await JSZip.loadAsync(await mirroredDocx.arrayBuffer());
const mirroredSettingsXml = await mirroredZip.file("word/settings.xml").async("text");
assert.match(mirroredSettingsXml, /<w:mirrorMargins\s*\/>/);
const importedMirroredSettings = await DocumentFile.importDocx(mirroredDocx);
assert.equal(importedMirroredSettings.settings.mirrorMargins, true);
assert.match(importedMirroredSettings.inspect({ kind: "settings", maxChars: 2_000 }).ndjson, /"mirrorMargins":true/);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedMirroredSettings)).arrayBuffer()),
  Buffer.from(await mirroredDocx.arrayBuffer()),
  "an unchanged canonical mirrorMargins setting must preserve the source package exactly",
);
importedMirroredSettings.setSettings({ mirrorMargins: false });
const unmirroredDocx = await DocumentFile.exportDocx(importedMirroredSettings);
const unmirroredZip = await JSZip.loadAsync(await unmirroredDocx.arrayBuffer());
assert.doesNotMatch(await unmirroredZip.file("word/settings.xml").async("text"), /mirrorMargins/);
assert.equal((await DocumentFile.importDocx(unmirroredDocx)).settings.mirrorMargins, false);

const disabledMirrorSource = await DocumentFile.patchDocx(mirroredDocx, [{
  path: "word/settings.xml",
  xml: mirroredSettingsXml.replace(/<w:mirrorMargins\b[^>]*\/>/, '<w:mirrorMargins w:val="0"/>'),
}]);
const importedDisabledMirror = await DocumentFile.importDocx(disabledMirrorSource);
assert.equal(importedDisabledMirror.settings.mirrorMargins, false);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedDisabledMirror)).arrayBuffer()),
  Buffer.from(await disabledMirrorSource.arrayBuffer()),
  "an unchanged canonical mirrorMargins false value must preserve the source package exactly",
);
importedDisabledMirror.setSettings({ mirrorMargins: true });
assert.equal((await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDisabledMirror))).settings.mirrorMargins, true);

const irregularMirrorSource = await DocumentFile.patchDocx(mirroredDocx, [{
  path: "word/settings.xml",
  xml: mirroredSettingsXml.replace(/<w:mirrorMargins\b[^>]*\/>/, '<w:mirrorMargins w:compatFlag="retained"/>'),
}]);
const importedIrregularMirror = await DocumentFile.importDocx(irregularMirrorSource);
assert.equal(importedIrregularMirror.settings.mirrorMargins, false);
assert.deepEqual(await changedZipParts(irregularMirrorSource, await DocumentFile.exportDocx(importedIrregularMirror)), [],
  "irregular mirrorMargins markup must remain source-owned when left untouched");
importedIrregularMirror.setSettings({ updateFields: true });
const unrelatedMirrorEdit = await DocumentFile.exportDocx(importedIrregularMirror);
assert.match(
  await (await JSZip.loadAsync(await unrelatedMirrorEdit.arrayBuffer())).file("word/settings.xml").async("text"),
  /<w:mirrorMargins w:compatFlag="retained"\s*\/>/,
);
importedIrregularMirror.setSettings({ updateFields: false, mirrorMargins: true });
await assert.rejects(
  () => DocumentFile.exportDocx(importedIrregularMirror),
  (error) => error?.code === "unsupported_document_settings_edit" && /irregular mirrorMargins markup/i.test(error.message),
);

for (const [label, replacement] of [
  ["duplicate", '<w:mirrorMargins/><w:mirrorMargins/>'],
  ["child-bearing", '<w:mirrorMargins><w:compat/></w:mirrorMargins>'],
]) {
  const structuralSource = await DocumentFile.patchDocx(mirroredDocx, [{
    path: "word/settings.xml",
    xml: mirroredSettingsXml.replace(/<w:mirrorMargins\b[^>]*\/>/, replacement),
  }]);
  const importedStructuralMirror = await DocumentFile.importDocx(structuralSource);
  assert.equal(importedStructuralMirror.settings.mirrorMargins, false, `${label} markup must not project as canonical mirror margins`);
  assert.deepEqual(
    Buffer.from(await (await DocumentFile.exportDocx(importedStructuralMirror)).arrayBuffer()),
    Buffer.from(await structuralSource.arrayBuffer()),
    `${label} markup must remain byte-exact when no semantic edit is requested`,
  );
  importedStructuralMirror.setSettings({ updateFields: true });
  await assert.rejects(
    () => DocumentFile.exportDocx(importedStructuralMirror),
    (error) => error?.code === "unsupported_document_settings_edit" && /sibling document settings/i.test(error.message),
    `${label} markup must block a same-part settings rewrite that cannot prove exact preservation`,
  );
  importedStructuralMirror.setSettings({ updateFields: false, mirrorMargins: true });
  await assert.rejects(
    () => DocumentFile.exportDocx(importedStructuralMirror),
    (error) => error?.code === "unsupported_document_settings_edit" && /irregular mirrorMargins markup/i.test(error.message),
    `${label} markup must fail closed on semantic replacement`,
  );
}

const bindingGutterDocument = DocumentModel.create({
  name: "Top-edge binding gutter",
  settings: { mirrorMargins: true, gutterAtTop: true },
  blocks: [
    { kind: "paragraph", text: "Leave space for a top-edge binding." },
    {
      kind: "section",
      breakType: "nextPage",
      pageSize: { widthTwips: 12240, heightTwips: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440, gutter: 720 },
    },
  ],
});
const bindingGutterDocx = await DocumentFile.exportDocx(bindingGutterDocument);
const bindingGutterZip = await JSZip.loadAsync(await bindingGutterDocx.arrayBuffer());
const bindingGutterSettingsXml = await bindingGutterZip.file("word/settings.xml").async("text");
const bindingGutterDocumentXml = await bindingGutterZip.file("word/document.xml").async("text");
assert.match(bindingGutterSettingsXml, /<w:gutterAtTop\s*\/>/);
assert.match(bindingGutterDocumentXml, /<w:pgMar\b(?=[^>]*w:gutter="720")[^>]*\/>/);
const importedBindingGutter = await DocumentFile.importDocx(bindingGutterDocx);
const importedBindingSection = importedBindingGutter.blocks.find((block) => block.kind === "section");
assert.equal(importedBindingGutter.settings.gutterAtTop, true);
assert.equal(importedBindingSection?.margins.gutter, 720);
assert.equal(importedBindingSection?.editable, true);
assert.match(importedBindingGutter.inspect({ kind: "settings,section", maxChars: 4_000 }).ndjson, /"gutterAtTop":true/);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedBindingGutter)).arrayBuffer()),
  Buffer.from(await bindingGutterDocx.arrayBuffer()),
  "an unchanged canonical top-edge gutter must preserve the source package exactly",
);
importedBindingGutter.setSettings({ gutterAtTop: false });
importedBindingSection.margins.gutter = 900;
const sideBindingGutterDocx = await DocumentFile.exportDocx(importedBindingGutter);
const sideBindingGutterZip = await JSZip.loadAsync(await sideBindingGutterDocx.arrayBuffer());
assert.doesNotMatch(await sideBindingGutterZip.file("word/settings.xml").async("text"), /gutterAtTop/);
assert.match(await sideBindingGutterZip.file("word/document.xml").async("text"), /<w:pgMar\b(?=[^>]*w:gutter="900")[^>]*\/>/);
const importedSideBindingGutter = await DocumentFile.importDocx(sideBindingGutterDocx);
assert.equal(importedSideBindingGutter.settings.gutterAtTop, false);
assert.equal(importedSideBindingGutter.blocks.find((block) => block.kind === "section")?.margins.gutter, 900);

const disabledTopGutterSource = await DocumentFile.patchDocx(bindingGutterDocx, [{
  path: "word/settings.xml",
  xml: bindingGutterSettingsXml.replace(/<w:gutterAtTop\b[^>]*\/>/, '<w:gutterAtTop w:val="0"/>'),
}]);
const importedDisabledTopGutter = await DocumentFile.importDocx(disabledTopGutterSource);
assert.equal(importedDisabledTopGutter.settings.gutterAtTop, false);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedDisabledTopGutter)).arrayBuffer()),
  Buffer.from(await disabledTopGutterSource.arrayBuffer()),
  "an unchanged canonical gutterAtTop false value must preserve the source package exactly",
);
importedDisabledTopGutter.setSettings({ gutterAtTop: true });
assert.equal((await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDisabledTopGutter))).settings.gutterAtTop, true);

const irregularTopGutterSource = await DocumentFile.patchDocx(bindingGutterDocx, [{
  path: "word/settings.xml",
  xml: bindingGutterSettingsXml.replace(/<w:gutterAtTop\b[^>]*\/>/, '<w:gutterAtTop w:compatFlag="retained"/>'),
}]);
const importedIrregularTopGutter = await DocumentFile.importDocx(irregularTopGutterSource);
const irregularTopGutterSection = importedIrregularTopGutter.blocks.find((block) => block.kind === "section");
assert.equal(importedIrregularTopGutter.settings.gutterAtTop, false);
assert.equal(irregularTopGutterSection?.editable, false, "irregular page-margin mode settings must make section geometry read-only");
assert.match(importedIrregularTopGutter.inspect({ kind: "section", maxChars: 2_000 }).ndjson, /"editable":false/);
importedIrregularTopGutter.blocks.find((block) => block.kind === "paragraph")?.replaceText("top-edge", "retained top-edge");
const unrelatedTopGutterEdit = await DocumentFile.exportDocx(importedIrregularTopGutter);
const unrelatedTopGutterZip = await JSZip.loadAsync(await unrelatedTopGutterEdit.arrayBuffer());
assert.equal(
  await unrelatedTopGutterZip.file("word/settings.xml").async("text"),
  await (await JSZip.loadAsync(await irregularTopGutterSource.arrayBuffer())).file("word/settings.xml").async("text"),
  "an unrelated body edit must preserve irregular gutterAtTop settings bytes",
);
irregularTopGutterSection.margins.gutter = 1080;
await assert.rejects(
  () => DocumentFile.exportDocx(importedIrregularTopGutter),
  (error) => error?.code === "unsupported_document_edit" && /source-bound and read-only/i.test(error.message),
);
irregularTopGutterSection.margins.gutter = 720;
importedIrregularTopGutter.setSettings({ gutterAtTop: true });
await assert.rejects(
  () => DocumentFile.exportDocx(importedIrregularTopGutter),
  (error) => error?.code === "unsupported_document_settings_edit" && /irregular gutterAtTop markup/i.test(error.message),
);

const impossibleBindingGutter = DocumentModel.create({
  settings: { gutterAtTop: true },
  blocks: [
    { kind: "paragraph", text: "Invalid binding geometry" },
    { kind: "section", pageSize: { widthTwips: 5000, heightTwips: 5000 }, margins: { top: 1000, right: 500, bottom: 1000, left: 500, gutter: 3000 } },
  ],
});
assert.equal(impossibleBindingGutter.verify().issues.some((issue) => issue.type === "sectionMarginsExceedPage"), true);
await assert.rejects(
  () => DocumentFile.exportDocx(impossibleBindingGutter),
  (error) => error?.code === "invalid_document_section" && /binding gutter must leave a positive page content area/i.test(error.message),
);

const equalWidthColumnsDocument = DocumentModel.create({
  name: "Equal-width section columns",
  blocks: [
    { kind: "paragraph", text: "Flow this brief through two equal columns." },
    {
      kind: "section",
      breakType: "continuous",
      pageSize: { widthTwips: 12240, heightTwips: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      columns: { count: 2, spacing: 720, separator: true },
    },
  ],
});
const equalWidthColumnsDocx = await DocumentFile.exportDocx(equalWidthColumnsDocument);
const equalWidthColumnsZip = await JSZip.loadAsync(await equalWidthColumnsDocx.arrayBuffer());
const equalWidthColumnsXml = await equalWidthColumnsZip.file("word/document.xml").async("text");
assert.match(equalWidthColumnsXml, /<w:cols\b(?=[^>]*w:equalWidth="(?:true|1)")(?=[^>]*w:num="2")(?=[^>]*w:space="720")(?=[^>]*w:sep="(?:true|1)")[^>]*\/>/);
const importedEqualWidthColumns = await DocumentFile.importDocx(equalWidthColumnsDocx);
const importedEqualWidthSection = importedEqualWidthColumns.blocks.find((block) => block.kind === "section");
assert.deepEqual(importedEqualWidthSection?.columns, { count: 2, spacing: 720, separator: true });
assert.equal(importedEqualWidthSection?.editable, true);
assert.match(importedEqualWidthColumns.inspect({ kind: "section", maxChars: 2_000 }).ndjson, /"columns":\{"count":2,"spacing":720,"separator":true\}/);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedEqualWidthColumns)).arrayBuffer()),
  Buffer.from(await equalWidthColumnsDocx.arrayBuffer()),
  "unchanged canonical equal-width columns must preserve the source package exactly",
);
importedEqualWidthSection.columns.count = 3;
importedEqualWidthSection.columns.spacing = 360;
importedEqualWidthSection.columns.separator = false;
const editedEqualWidthColumnsDocx = await DocumentFile.exportDocx(importedEqualWidthColumns);
const editedEqualWidthColumns = await DocumentFile.importDocx(editedEqualWidthColumnsDocx);
const editedEqualWidthSection = editedEqualWidthColumns.blocks.find((block) => block.kind === "section");
assert.deepEqual(editedEqualWidthSection?.columns, { count: 3, spacing: 360, separator: false });
editedEqualWidthSection.columns = undefined;
const removedEqualWidthColumnsDocx = await DocumentFile.exportDocx(editedEqualWidthColumns);
const removedEqualWidthColumnsXml = await (await JSZip.loadAsync(await removedEqualWidthColumnsDocx.arrayBuffer())).file("word/document.xml").async("text");
assert.doesNotMatch(removedEqualWidthColumnsXml, /<w:cols\b/);
assert.equal((await DocumentFile.importDocx(removedEqualWidthColumnsDocx)).blocks.find((block) => block.kind === "section")?.columns, undefined);

const implicitEqualWidthColumnsSource = await DocumentFile.patchDocx(equalWidthColumnsDocx, [{
  path: "word/document.xml",
  xml: equalWidthColumnsXml.replace(/(<w:cols\b[^>]*?)\s+w:equalWidth="(?:true|1)"/, "$1"),
}]);
const importedImplicitEqualWidthColumns = await DocumentFile.importDocx(implicitEqualWidthColumnsSource);
assert.equal(importedImplicitEqualWidthColumns.blocks.find((block) => block.kind === "section")?.editable, true);
assert.deepEqual(importedImplicitEqualWidthColumns.blocks.find((block) => block.kind === "section")?.columns, { count: 2, spacing: 720, separator: true });
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedImplicitEqualWidthColumns)).arrayBuffer()),
  Buffer.from(await implicitEqualWidthColumnsSource.arrayBuffer()),
  "an omitted equalWidth attribute must retain its source bytes and canonical true semantics",
);

const customWidthColumnsDocument = DocumentModel.create({
  name: "Custom-width section columns",
  blocks: [
    { kind: "paragraph", text: "Flow this brief through asymmetric columns." },
    {
      kind: "section",
      breakType: "continuous",
      pageSize: { widthTwips: 12240, heightTwips: 15840 },
      margins: { top: 1440, right: 1440, bottom: 1440, left: 1440 },
      columns: { definitions: [{ width: 3000, spacing: 720 }, { width: 5000, spacing: 0 }], separator: true },
    },
  ],
});
const customWidthColumnsSource = await DocumentFile.exportDocx(customWidthColumnsDocument);
const customWidthColumnsXml = await (await JSZip.loadAsync(await customWidthColumnsSource.arrayBuffer())).file("word/document.xml").async("text");
const customWidthColumnsFragment = customWidthColumnsXml.match(/<w:cols\b[\s\S]*?<\/w:cols>/)?.[0];
const normalizedCustomColumnsMarkup = (xml) => xml.replace(/<w:(cols|col)\b([^>]*?)(\/?)>/g, (_match, name, attributes, closed) => {
  const sorted = [...attributes.matchAll(/[^\s=]+="[^"]*"/g)].map((match) => match[0]).sort();
  return `<w:${name}${sorted.length ? ` ${sorted.join(" ")}` : ""}${closed ? "/" : ""}>`;
});
assert.match(customWidthColumnsFragment, /<w:cols\b(?=[^>]*w:equalWidth="(?:false|0)")(?=[^>]*w:sep="(?:true|1)")[^>]*>/);
assert.doesNotMatch(customWidthColumnsFragment.match(/<w:cols\b[^>]*>/)?.[0] || "", /\bw:(?:num|space)=/);
assert.match(customWidthColumnsFragment, /<w:col\b(?=[^>]*w:w="3000")(?=[^>]*w:space="720")[^>]*\/>[\s\S]*<w:col\b(?=[^>]*w:w="5000")[^>]*\/>/);
const importedCustomWidthColumns = await DocumentFile.importDocx(customWidthColumnsSource);
const customWidthSection = importedCustomWidthColumns.blocks.find((block) => block.kind === "section");
assert.equal(customWidthSection?.editable, true);
assert.deepEqual(customWidthSection?.columns, { definitions: [{ width: 3000, spacing: 720 }, { width: 5000, spacing: 0 }], separator: true });
assert.match(importedCustomWidthColumns.inspect({ kind: "section", maxChars: 2_000 }).ndjson, /"definitions":\[\{"width":3000,"spacing":720\},\{"width":5000,"spacing":0\}\]/);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedCustomWidthColumns)).arrayBuffer()),
  Buffer.from(await customWidthColumnsSource.arrayBuffer()),
  "unchanged canonical custom-width columns must preserve the source package exactly",
);
const customWidthParagraph = importedCustomWidthColumns.blocks.find((block) => block.kind === "paragraph");
customWidthParagraph.replaceText("brief", "retained brief");
const customWidthBodyEdit = await DocumentFile.exportDocx(importedCustomWidthColumns);
const customWidthBodyEditXml = await (await JSZip.loadAsync(await customWidthBodyEdit.arrayBuffer())).file("word/document.xml").async("text");
assert.equal(
  normalizedCustomColumnsMarkup(customWidthBodyEditXml.match(/<w:cols\b[\s\S]*?<\/w:cols>/)?.[0]),
  normalizedCustomColumnsMarkup(customWidthColumnsFragment),
  "an unrelated body edit must preserve custom-width column markup",
);
customWidthSection.columns.definitions[0] = { width: 3200, spacing: 360 };
customWidthSection.columns.definitions[1] = { width: 5680, spacing: 120 };
customWidthSection.columns.separator = false;
const editedCustomWidthColumnsDocx = await DocumentFile.exportDocx(importedCustomWidthColumns);
const editedCustomWidthColumns = await DocumentFile.importDocx(editedCustomWidthColumnsDocx);
const editedCustomWidthSection = editedCustomWidthColumns.blocks.find((block) => block.kind === "section");
assert.deepEqual(editedCustomWidthSection?.columns, { definitions: [{ width: 3200, spacing: 360 }, { width: 5680, spacing: 120 }], separator: false });
editedCustomWidthSection.columns = undefined;
const removedCustomWidthColumnsDocx = await DocumentFile.exportDocx(editedCustomWidthColumns);
const removedCustomWidthColumnsXml = await (await JSZip.loadAsync(await removedCustomWidthColumnsDocx.arrayBuffer())).file("word/document.xml").async("text");
assert.doesNotMatch(removedCustomWidthColumnsXml, /<w:cols\b/);

const canonicalColumnsMarkup = equalWidthColumnsXml.match(/<w:cols\b[^>]*\/>/)?.[0];
const normalizedColumnTags = (xml) => [...xml.matchAll(/<w:cols\b([^>]*)\/>/g)]
  .map((match) => match[1].trim().split(/\s+/).sort());
for (const [label, markup] of [
  ["duplicate", `${canonicalColumnsMarkup}${canonicalColumnsMarkup}`],
  ["extension-bearing", canonicalColumnsMarkup.replace("<w:cols", '<w:cols w:compatFlag="retained"')],
]) {
  const irregularColumnsSource = await DocumentFile.patchDocx(equalWidthColumnsDocx, [{
    path: "word/document.xml",
    xml: equalWidthColumnsXml.replace(canonicalColumnsMarkup, markup),
  }]);
  const irregularColumnsSourceXml = await (await JSZip.loadAsync(await irregularColumnsSource.arrayBuffer())).file("word/document.xml").async("text");
  const irregularColumnsSectionMarkup = irregularColumnsSourceXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0];
  const importedIrregularColumns = await DocumentFile.importDocx(irregularColumnsSource);
  const irregularColumnsSection = importedIrregularColumns.blocks.find((block) => block.kind === "section");
  assert.equal(irregularColumnsSection?.editable, false, `${label} columns must keep section geometry source-owned`);
  assert.equal(irregularColumnsSection?.columns, undefined);
  importedIrregularColumns.blocks.find((block) => block.kind === "paragraph")?.replaceText("brief", `${label} brief`);
  const unrelatedIrregularColumnsEdit = await DocumentFile.exportDocx(importedIrregularColumns);
  const unrelatedIrregularColumnsXml = await (await JSZip.loadAsync(await unrelatedIrregularColumnsEdit.arrayBuffer())).file("word/document.xml").async("text");
  const unrelatedIrregularColumnsSectionMarkup = unrelatedIrregularColumnsXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0];
  assert.equal(
    unrelatedIrregularColumnsSectionMarkup.replace(/<w:cols\b[^>]*\/>/g, ""),
    irregularColumnsSectionMarkup.replace(/<w:cols\b[^>]*\/>/g, ""),
    `an unrelated body edit must preserve the section siblings around ${label} columns`,
  );
  assert.deepEqual(
    normalizedColumnTags(unrelatedIrregularColumnsSectionMarkup),
    normalizedColumnTags(irregularColumnsSectionMarkup),
    `an unrelated body edit must preserve every ${label} columns attribute`,
  );
  irregularColumnsSection.columns = { count: 2, spacing: 720, separator: false };
  await assert.rejects(
    () => DocumentFile.exportDocx(importedIrregularColumns),
    (error) => error?.code === "unsupported_document_edit" && /source-bound and read-only/i.test(error.message),
  );
}

for (const [label, markup] of [
  ["equal-width root attributes", customWidthColumnsFragment.replace("<w:cols", '<w:cols w:num="2"')],
  ["extension-bearing definition", customWidthColumnsFragment.replace(/<w:col(?=\s)/, '<w:col w:compatFlag="retained"')],
  ["missing custom width", customWidthColumnsFragment.replace(/(<w:col\b[^>]*?)\s+w:w="5000"/, "$1")],
]) {
  const irregularCustomColumnsSource = await DocumentFile.patchDocx(customWidthColumnsSource, [{
    path: "word/document.xml",
    xml: customWidthColumnsXml.replace(customWidthColumnsFragment, markup),
  }]);
  const importedIrregularCustomColumns = await DocumentFile.importDocx(irregularCustomColumnsSource);
  const irregularCustomColumnsSection = importedIrregularCustomColumns.blocks.find((block) => block.kind === "section");
  assert.equal(irregularCustomColumnsSection?.editable, false, `${label} must keep custom column geometry source-owned`);
  assert.equal(irregularCustomColumnsSection?.columns, undefined);
  importedIrregularCustomColumns.blocks.find((block) => block.kind === "paragraph")?.replaceText("brief", `${label} brief`);
  const irregularCustomColumnsEdit = await DocumentFile.exportDocx(importedIrregularCustomColumns);
  const irregularCustomColumnsEditXml = await (await JSZip.loadAsync(await irregularCustomColumnsEdit.arrayBuffer())).file("word/document.xml").async("text");
  assert.equal(
    normalizedCustomColumnsMarkup(irregularCustomColumnsEditXml.match(/<w:cols\b[\s\S]*?<\/w:cols>/)?.[0]),
    normalizedCustomColumnsMarkup(markup),
    `an unrelated body edit must preserve ${label} markup`,
  );
  irregularCustomColumnsSection.columns = { definitions: [{ width: 3000, spacing: 720 }, { width: 5640, spacing: 0 }], separator: false };
  await assert.rejects(
    () => DocumentFile.exportDocx(importedIrregularCustomColumns),
    (error) => error?.code === "unsupported_document_edit" && /source-bound and read-only/i.test(error.message),
  );
}

const impossibleSectionColumns = DocumentModel.create({
  blocks: [
    { kind: "paragraph", text: "Invalid equal-width columns" },
    { kind: "section", pageSize: { widthTwips: 5000, heightTwips: 5000 }, margins: { top: 500, right: 500, bottom: 500, left: 500 }, columns: { count: 3, spacing: 2000, separator: false } },
  ],
});
assert.equal(impossibleSectionColumns.verify().issues.some((issue) => issue.type === "sectionColumnsExceedPage"), true);
await assert.rejects(
  () => DocumentFile.exportDocx(impossibleSectionColumns),
  (error) => error?.code === "invalid_document_section" && /column spacing must leave positive width/i.test(error.message),
);
const tooManySectionColumns = DocumentModel.create({ blocks: [
  { kind: "paragraph", text: "Too many columns" },
  { kind: "section", columns: { count: 46, spacing: 0, separator: false } },
] });
assert.equal(tooManySectionColumns.verify().issues.some((issue) => issue.type === "invalidSectionColumns"), true);
await assert.rejects(() => DocumentFile.exportDocx(tooManySectionColumns), /column count must be 1 through 45/i);
assert.throws(
  () => DocumentModel.create({ blocks: [{ kind: "section", columns: "two" }] }),
  /section columns must be an object/i,
);
assert.throws(
  () => DocumentModel.create({ blocks: [{ kind: "section", columns: { count: 2, widths: [3000, 5000] } }] }),
  /unsupported document section column properties: widths/i,
);
assert.throws(
  () => DocumentModel.create({ blocks: [{ kind: "section", columns: { count: 2, definitions: [{ width: 3000 }, { width: 5000 }] } }] }),
  /cannot combine definitions with equal-width count or spacing/i,
);
assert.throws(
  () => DocumentModel.create({ blocks: [{ kind: "section", columns: { definitions: [{ width: 3000, spaceAfter: 720 }, { width: 5000 }] } }] }),
  /unsupported document section column definition properties.*spaceAfter/i,
);
const impossibleCustomSectionColumns = DocumentModel.create({ blocks: [
  { kind: "paragraph", text: "Custom columns that exceed the page" },
  { kind: "section", pageSize: { widthTwips: 5000, heightTwips: 5000 }, margins: { top: 500, right: 500, bottom: 500, left: 500 }, columns: { definitions: [{ width: 2500, spacing: 500 }, { width: 1500, spacing: 0 }] } },
] });
assert.equal(impossibleCustomSectionColumns.verify().issues.some((issue) => issue.type === "sectionColumnsExceedPage"), true);
await assert.rejects(() => DocumentFile.exportDocx(impossibleCustomSectionColumns), /custom column widths and spacing must fit within the page content width/i);
const zeroWidthCustomSectionColumn = DocumentModel.create({ blocks: [
  { kind: "paragraph", text: "Zero-width custom column" },
  { kind: "section", columns: { definitions: [{ width: 0, spacing: 720 }, { width: 5000, spacing: 0 }] } },
] });
assert.equal(zeroWidthCustomSectionColumn.verify().issues.some((issue) => issue.type === "invalidSectionColumns"), true);
await assert.rejects(() => DocumentFile.exportDocx(zeroWidthCustomSectionColumn), /custom column widths must be 1 through 31680/i);

const sectionPageNumberingDocument = DocumentModel.create({
  name: "Section page numbering",
  blocks: [
    { kind: "paragraph", text: "Front matter uses a bounded page-number restart." },
    {
      kind: "section",
      breakType: "nextPage",
      pageNumbering: { start: 1, format: "lowerRoman" },
      columns: { count: 2, spacing: 720, separator: false },
    },
  ],
});
const sectionPageNumberingDocx = await DocumentFile.exportDocx(sectionPageNumberingDocument);
const sectionPageNumberingXml = await (await JSZip.loadAsync(await sectionPageNumberingDocx.arrayBuffer())).file("word/document.xml").async("text");
const canonicalPageNumberingMarkup = sectionPageNumberingXml.match(/<w:pgNumType\b[^>]*\/>/)?.[0];
const normalizedPageNumberingTags = (xml) => [...xml.matchAll(/<w:pgNumType\b([^>]*)\/>/g)]
  .map((match) => match[1].trim().split(/\s+/).filter(Boolean).sort());
assert.match(canonicalPageNumberingMarkup, /<w:pgNumType\b(?=[^>]*w:start="1")(?=[^>]*w:fmt="lowerRoman")[^>]*\/>/);
assert.match(sectionPageNumberingXml, /<w:pgNumType\b[^>]*\/>[\s\S]*<w:cols\b/, "pgNumType must precede columns in the native section-property sequence");
const importedSectionPageNumbering = await DocumentFile.importDocx(sectionPageNumberingDocx);
const sectionPageNumbering = importedSectionPageNumbering.blocks.find((block) => block.kind === "section");
assert.equal(sectionPageNumbering?.editable, true);
assert.deepEqual(sectionPageNumbering?.pageNumbering, { start: 1, format: "lowerRoman" });
assert.match(importedSectionPageNumbering.inspect({ kind: "section", maxChars: 2_000 }).ndjson, /"pageNumbering":\{"start":1,"format":"lowerRoman"\}/);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedSectionPageNumbering)).arrayBuffer()),
  Buffer.from(await sectionPageNumberingDocx.arrayBuffer()),
  "unchanged canonical section page numbering must preserve the source package exactly",
);
importedSectionPageNumbering.blocks.find((block) => block.kind === "paragraph")?.replaceText("Front matter", "Retained front matter");
const pageNumberingBodyEdit = await DocumentFile.exportDocx(importedSectionPageNumbering);
const pageNumberingBodyEditXml = await (await JSZip.loadAsync(await pageNumberingBodyEdit.arrayBuffer())).file("word/document.xml").async("text");
assert.equal(pageNumberingBodyEditXml.match(/<w:pgNumType\b[^>]*\/>/)?.[0], canonicalPageNumberingMarkup, "an unrelated body edit must preserve page-number markup");
sectionPageNumbering.pageNumbering = { start: 7, format: "upperLetter" };
const editedSectionPageNumberingDocx = await DocumentFile.exportDocx(importedSectionPageNumbering);
const editedSectionPageNumbering = await DocumentFile.importDocx(editedSectionPageNumberingDocx);
const editedPageNumberingSection = editedSectionPageNumbering.blocks.find((block) => block.kind === "section");
assert.deepEqual(editedPageNumberingSection?.pageNumbering, { start: 7, format: "upperLetter" });
editedPageNumberingSection.pageNumbering = { format: "decimal" };
const continuedDecimalDocx = await DocumentFile.exportDocx(editedSectionPageNumbering);
const continuedDecimalXml = await (await JSZip.loadAsync(await continuedDecimalDocx.arrayBuffer())).file("word/document.xml").async("text");
assert.match(continuedDecimalXml, /<w:pgNumType\b(?=[^>]*w:fmt="decimal")[^>]*\/>/);
assert.doesNotMatch(continuedDecimalXml.match(/<w:pgNumType\b[^>]*\/>/)?.[0] || "", /\bw:start=/);
assert.deepEqual((await DocumentFile.importDocx(continuedDecimalDocx)).blocks.find((block) => block.kind === "section")?.pageNumbering, { format: "decimal" });
editedPageNumberingSection.pageNumbering = { start: 0 };
assert.deepEqual((await DocumentFile.importDocx(await DocumentFile.exportDocx(editedSectionPageNumbering))).blocks.find((block) => block.kind === "section")?.pageNumbering, { start: 0 });
editedPageNumberingSection.pageNumbering = undefined;
const removedPageNumberingDocx = await DocumentFile.exportDocx(editedSectionPageNumbering);
assert.doesNotMatch(await (await JSZip.loadAsync(await removedPageNumberingDocx.arrayBuffer())).file("word/document.xml").async("text"), /<w:pgNumType\b/);

for (const [label, markup] of [
  ["duplicate", `${canonicalPageNumberingMarkup}${canonicalPageNumberingMarkup}`],
  ["extension-bearing", canonicalPageNumberingMarkup.replace("<w:pgNumType", '<w:pgNumType w:compatFlag="retained"')],
  ["chapter-numbered", canonicalPageNumberingMarkup.replace("/>", ' w:chapStyle="1" w:chapSep="hyphen"/>')],
  ["unsupported format", canonicalPageNumberingMarkup.replace('w:fmt="lowerRoman"', 'w:fmt="ordinal"')],
  ["empty", "<w:pgNumType/>"],
]) {
  const irregularPageNumberingSource = await DocumentFile.patchDocx(sectionPageNumberingDocx, [{
    path: "word/document.xml",
    xml: sectionPageNumberingXml.replace(canonicalPageNumberingMarkup, markup),
  }]);
  const importedIrregularPageNumbering = await DocumentFile.importDocx(irregularPageNumberingSource);
  const irregularPageNumberingSection = importedIrregularPageNumbering.blocks.find((block) => block.kind === "section");
  assert.equal(irregularPageNumberingSection?.editable, false, `${label} page numbering must keep section geometry source-owned`);
  assert.equal(irregularPageNumberingSection?.pageNumbering, undefined);
  importedIrregularPageNumbering.blocks.find((block) => block.kind === "paragraph")?.replaceText("Front matter", `${label} front matter`);
  const unrelatedIrregularPageNumberingEdit = await DocumentFile.exportDocx(importedIrregularPageNumbering);
  const unrelatedIrregularPageNumberingXml = await (await JSZip.loadAsync(await unrelatedIrregularPageNumberingEdit.arrayBuffer())).file("word/document.xml").async("text");
  const editedSectionMarkup = unrelatedIrregularPageNumberingXml.match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0];
  const sourceSectionMarkup = (await (await JSZip.loadAsync(await irregularPageNumberingSource.arrayBuffer())).file("word/document.xml").async("text")).match(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/)?.[0];
  assert.equal(
    editedSectionMarkup.replace(/<w:pgNumType\b[^>]*\/>/g, ""),
    sourceSectionMarkup.replace(/<w:pgNumType\b[^>]*\/>/g, ""),
    `an unrelated body edit must preserve section siblings around ${label} page numbering`,
  );
  assert.deepEqual(normalizedPageNumberingTags(editedSectionMarkup), normalizedPageNumberingTags(sourceSectionMarkup), `an unrelated body edit must preserve every ${label} page-number attribute`);
  irregularPageNumberingSection.pageNumbering = { start: 1, format: "decimal" };
  await assert.rejects(
    () => DocumentFile.exportDocx(importedIrregularPageNumbering),
    (error) => error?.code === "unsupported_document_edit" && /source-bound and read-only/i.test(error.message),
  );
}

assert.throws(() => DocumentModel.create({ blocks: [{ kind: "section", pageNumbering: "roman" }] }), /pageNumbering must be an object/i);
assert.throws(() => DocumentModel.create({ blocks: [{ kind: "section", pageNumbering: { start: 1, chapterStyle: 1 } }] }), /unsupported document section pageNumbering properties: chapterStyle/i);
for (const [pageNumbering, message] of [
  [{}, /requires a start value or supported format/i],
  [{ start: -1 }, /unsigned 32-bit integer|start must be an integer from 0 through 2147483647/i],
  [{ start: 2_147_483_648 }, /start must not exceed 2147483647/i],
  [{ format: "ordinal" }, /unsupported document section page-number format/i],
]) {
  const invalid = DocumentModel.create({ blocks: [{ kind: "paragraph", text: "Invalid page numbering" }, { kind: "section", pageNumbering }] });
  assert.equal(invalid.verify().issues.some((issue) => issue.type === "invalidSectionPageNumbering"), true);
  await assert.rejects(() => DocumentFile.exportDocx(invalid), message);
}

const protectedSettings = DocumentModel.create({
  name: "Passwordless document protection",
  settings: { documentProtection: "readOnly" },
  blocks: [{ kind: "paragraph", text: "Bounded editing restriction." }],
});
assert.deepEqual(protectedSettings.settings.documentProtection, {
  edit: "readOnly",
  enforcement: true,
  formatting: false,
});
assert.throws(
  () => protectedSettings.setSettings({ documentProtection: { edit: "readOnly", password: "secret" } }),
  /Password hashing is intentionally unsupported/,
);
const protectedDocx = await DocumentFile.exportDocx(protectedSettings);
const protectedZip = await JSZip.loadAsync(await protectedDocx.arrayBuffer());
const protectedXml = await protectedZip.file("word/settings.xml").async("text");
assert.match(protectedXml, /<w:documentProtection(?=[^>]*w:edit="readOnly")(?=[^>]*w:enforcement="true")(?=[^>]*w:formatting="false")[^>]*\/>/);
const importedProtectedSettings = await DocumentFile.importDocx(protectedDocx);
assert.deepEqual(importedProtectedSettings.settings.documentProtection, protectedSettings.settings.documentProtection);
assert.deepEqual(
  Buffer.from(await (await DocumentFile.exportDocx(importedProtectedSettings)).arrayBuffer()),
  Buffer.from(await protectedDocx.arrayBuffer()),
  "an unchanged source-bound protection setting must preserve the source package exactly",
);
importedProtectedSettings.setSettings({ documentProtection: { edit: "comments", enforcement: false, formatting: true } });
const commentsProtectedDocx = await DocumentFile.exportDocx(importedProtectedSettings);
const commentsProtectedRoundTrip = await DocumentFile.importDocx(commentsProtectedDocx);
assert.deepEqual(commentsProtectedRoundTrip.settings.documentProtection, {
  edit: "comments",
  enforcement: false,
  formatting: true,
});
commentsProtectedRoundTrip.setSettings({ documentProtection: "none" });
const explicitNoneDocx = await DocumentFile.exportDocx(commentsProtectedRoundTrip);
assert.equal((await DocumentFile.importDocx(explicitNoneDocx)).settings.documentProtection.edit, "none");
commentsProtectedRoundTrip.setSettings({ documentProtection: false });
const unprotectedDocx = await DocumentFile.exportDocx(commentsProtectedRoundTrip);
const unprotectedZip = await JSZip.loadAsync(await unprotectedDocx.arrayBuffer());
assert.doesNotMatch(await unprotectedZip.file("word/settings.xml").async("text"), /documentProtection/);
assert.equal((await DocumentFile.importDocx(unprotectedDocx)).settings.documentProtection, null);

const verifierSource = await DocumentFile.patchDocx(protectedDocx, [{
  path: "word/settings.xml",
  xml: protectedXml.replace(/<w:documentProtection\b([^>]*)\/>/, '<w:documentProtection$1 w:hash="AA=="/>'),
}]);
const importedVerifierSource = await DocumentFile.importDocx(verifierSource);
assert.equal(importedVerifierSource.settings.documentProtection, null);
const preservedVerifierSource = await DocumentFile.exportDocx(importedVerifierSource);
assert.deepEqual(await changedZipParts(verifierSource, preservedVerifierSource), [],
  "unsupported password-verifier markup must remain source-owned without changing any package part");
assert.match(
  await (await JSZip.loadAsync(await preservedVerifierSource.arrayBuffer())).file("word/settings.xml").async("text"),
  /w:hash="AA=="/,
);
importedVerifierSource.setSettings({ documentProtection: "forms" });
await assert.rejects(
  () => DocumentFile.exportDocx(importedVerifierSource),
  (error) => error?.code === "unsupported_document_protection_edit" && /password verifiers/i.test(error.message),
);

const unsupportedTableStyle = DocumentModel.create({ name: "Unsupported table style", blocks: [] });
unsupportedTableStyle.styles.add("ComparisonTable", { name: "Comparison Table", type: "table" });
unsupportedTableStyle.addTable({
  styleId: "ComparisonTable",
  values: [["Option", "Status"], ["Pilot", "Ready"]],
});
await assert.rejects(
  () => DocumentFile.exportDocx(unsupportedTableStyle),
  (error) => error?.code === "unsupported_document_features" && /cannot materialize custom table style ComparisonTable/i.test(error.message),
);

const importedWithAddedBookmark = await DocumentFile.importDocx(firstDocx);
importedWithAddedBookmark.addBookmark(importedWithAddedBookmark.blocks[0], "AddedBookmark");
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithAddedBookmark),
  (error) => error?.code === "document_bookmark_topology_changed" && /bookmark topology/i.test(error.message),
);

const importedWithMovedNote = await DocumentFile.importDocx(firstDocx);
importedWithMovedNote.notes[0].targetId = importedWithMovedNote.notes[1].targetId;
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithMovedNote),
  (error) => error?.code === "unsupported_document_note_edit" && /source-bound/i.test(error.message),
);

const importedWithoutEndnote = await DocumentFile.importDocx(firstDocx);
importedWithoutEndnote.notes.pop();
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithoutEndnote),
  (error) => error?.code === "document_note_topology_changed" && /note topology/i.test(error.message),
);

const invalidNoteTarget = DocumentModel.create({ blocks: [{ kind: "table", values: [["Not a paragraph"]] }] });
invalidNoteTarget.addFootnote(invalidNoteTarget.blocks[0], "Invalid target");
await assert.rejects(
  () => DocumentFile.exportDocx(invalidNoteTarget),
  (error) => error?.code === "invalid_document_note" && /paragraph or list item/i.test(error.message),
);

const importedWithoutSourceSnapshot = await DocumentFile.importDocx(firstDocx);
const documentState = importedWithoutSourceSnapshot[Symbol.for("open-office-artifact-tool.open-chestnut-document-state")];
documentState.opaqueOpc.sourcePackage = undefined;
await assert.rejects(
  () => DocumentFile.exportDocx(importedWithoutSourceSnapshot),
  (error) => error?.code === "missing_source_package",
);

assert.throws(
  () => DocumentModel.create({ blocks: [{ kind: "unknown-office-block", text: "Do not coerce me." }] }),
  /Unsupported document block kind unknown-office-block/,
);

assert.equal(bullet.kind, "listItem");
assert.equal(numbered.kind, "listItem");
assert.equal(field.kind, "field");
assert.equal(jpegImage.kind, "image");

console.log("document tests ok");
