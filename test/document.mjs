import assert from "node:assert/strict";

import { DocumentFile, DocumentModel } from "../src/index.mjs";
import { DocumentFile as DocumentFileModule, DocumentModel as DocumentModelModule } from "../src/document/index.mjs";

assert.strictEqual(DocumentModel, DocumentModelModule);
assert.strictEqual(DocumentFile, DocumentFileModule);

const png = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
const jpeg = "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQAAAQABAAD/4gHYSUNDX1BST0ZJTEUAAQEAAAHIAAAAAAQwAABtbnRyUkdCIFhZWiAH4AABAAEAAAAAAABhY3NwAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAQAA9tYAAQAAAADTLQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAlkZXNjAAAA8AAAACRyWFlaAAABFAAAABRnWFlaAAABKAAAABRiWFlaAAABPAAAABR3dHB0AAABUAAAABRyVFJDAAABZAAAAChnVFJDAAABZAAAAChiVFJDAAABZAAAAChjcHJ0AAABjAAAADxtbHVjAAAAAAAAAAEAAAAMZW5VUwAAAAgAAAAcAHMAUgBHAEJYWVogAAAAAAAAb6IAADj1AAADkFhZWiAAAAAAAABimQAAt4UAABjaWFlaIAAAAAAAACSgAAAPhAAAts9YWVogAAAAAAAA9tYAAQAAAADTLXBhcmEAAAAAAAQAAAACZmYAAPKnAAANWQAAE9AAAApbAAAAAAAAAABtbHVjAAAAAAAAAAEAAAAMZW5VUwAAACAAAAAcAEcAbwBvAGcAbABlACAASQBuAGMALgAgADIAMAAxADb/2wBDAAYEBQYFBAYGBQYHBwYIChAKCgkJChQODwwQFxQYGBcUFhYaHSUfGhsjHBYWICwgIyYnKSopGR8tMC0oMCUoKSj/2wBDAQcHBwoIChMKChMoGhYaKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCgoKCj/wAARCAACAAIDASIAAhEBAxEB/8QAFQABAQAAAAAAAAAAAAAAAAAAAAX/xAAUEAEAAAAAAAAAAAAAAAAAAAAA/8QAFQEBAQAAAAAAAAAAAAAAAAAABQf/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oADAMBAAIRAxEAPwCQAHTd/9k=";

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
document.addParagraph("Second-section evidence.", { name: "second-section", styleId: "Normal" });

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
  kind: "document,paragraph,listItem,table,comment,header,footer,hyperlink,field,image,section,style,layout",
  maxChars: 24_000,
}).ndjson;
for (const expected of [
  "OpenChestnut document profile",
  "BodyAccent",
  "Bold and colored",
  "readiness-table",
  "Default header",
  "First-page header",
  "Even-page header",
  "PNG approval mark",
  "JPEG approval mark",
  "landscape-section",
  "Confirm the release evidence",
]) assert.match(inspect, new RegExp(expected));
assert.equal(document.resolve(formatted.id), formatted);
assert.equal(document.resolve(table.id).getCell(1, 1).value, "Pending");
assert.equal(document.resolve(pngImage.id).alt, "PNG approval mark");
assert.equal(document.resolve(section.id).orientation, "landscape");
assert.equal(document.resolve(comment.id).targetId, commentTarget.id);
assert.equal(document.resolve(defaultHeader.id).referenceType, "default");
assert.equal(document.resolve(firstHeader.id).referenceType, "first");
assert.equal(document.resolve(evenHeader.id).referenceType, "even");
assert.equal(document.resolve(defaultFooter.id).fieldInstruction, "PAGE");
assert.equal(document.resolve(firstFooter.id).referenceType, "first");
assert.equal(document.resolve(evenFooter.id).referenceType, "even");
const modelVerification = document.verify({ visualQa: true });
assert.equal(modelVerification.ok, true, JSON.stringify(modelVerification.issues));

const firstDocx = await DocumentFile.exportDocx(document);
assert.equal(firstDocx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
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
assert.equal(imported.blocks.filter((block) => block.kind === "listItem").length, 2);
assert.equal(imported.blocks.find((block) => block.kind === "table")?.values[1][1], "Pending");
assert.equal(imported.blocks.find((block) => block.kind === "hyperlink")?.url, hyperlink.url);
assert.equal(imported.blocks.find((block) => block.kind === "field")?.instruction, "PAGE");
assert.equal(imported.blocks.filter((block) => block.kind === "image").length, 2);
assert.equal(imported.blocks.some((block) => block.kind === "image" && block.dataUrl.startsWith("data:image/png;base64,")), true);
assert.equal(imported.blocks.some((block) => block.kind === "image" && block.dataUrl.startsWith("data:image/jpeg;base64,")), true);
assert.equal(imported.blocks.find((block) => block.kind === "section")?.orientation, "landscape");
assert.equal(imported.settings.evenAndOddHeaders, true);
assert.equal(imported.sectionSettings[0]?.differentFirstPage, true);
assert.deepEqual(imported.headers.map((item) => item.referenceType), ["default", "first", "even"]);
assert.deepEqual(imported.footers.map((item) => item.referenceType), ["default", "first", "even"]);
assert.equal(imported.footers[0].fieldInstruction, "PAGE");
assert.equal(imported.comments.length, 1);

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
const importedPng = imported.blocks.find((block) => block.kind === "image" && block.alt === "PNG approval mark");
importedPng.alt = "Edited PNG approval mark";
importedPng.widthPx = 56;
const importedSection = imported.blocks.find((block) => block.kind === "section");
importedSection.margins.left = 1200;
imported.comments[0].author = "Lead reviewer";
imported.comments[0].initials = "LR";
imported.comments[0].text = "Release evidence approved.";

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
assert.equal(roundTrip.blocks.find((block) => block.kind === "image" && block.alt === "Edited PNG approval mark")?.widthPx, 56);
assert.equal(roundTrip.blocks.find((block) => block.kind === "section")?.margins.left, 1200);
assert.equal(roundTrip.comments[0].author, "Lead reviewer");
assert.equal(roundTrip.comments[0].text, "Release evidence approved.");
assert.equal(roundTrip.verify({ visualQa: true }).ok, true);

const unsupported = DocumentModel.create({
  name: "Unsupported advanced authoring",
  blocks: [{ kind: "paragraph", styleId: "Normal", text: "Advanced source" }],
});
unsupported.addCitation("Unsupported citation", { tag: "AdvancedSource", title: "Advanced source" });
await assert.rejects(
  () => DocumentFile.exportDocx(unsupported),
  /cannot author or edit these DOCX features:/i,
);

const unsupportedSettings = DocumentModel.create({
  name: "Unsupported document settings",
  settings: { trackRevisions: true },
  blocks: [{ kind: "paragraph", text: "Tracked authoring must fail." }],
});
await assert.rejects(
  () => DocumentFile.exportDocx(unsupportedSettings),
  (error) => error?.code === "unsupported_document_features" && /revision tracking/i.test(error.message),
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
  (error) => error?.code === "unsupported_document_features" && /bookmarks/i.test(error.message),
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
