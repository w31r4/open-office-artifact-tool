import assert from "node:assert/strict";
import JSZip from "jszip";
import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
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
const replacementPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const gif = "data:image/gif;base64,R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==";

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
const canonicalThread = workbook.comments.addThread({ cell: sheet.getRange("D2") }, "Canonical threaded comment");
canonicalThread.addReply("Canonical direct reply", { author: "Reviewer", date: "2026-07-17T09:30:00.000Z" });
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
const xlsxZip = await JSZip.loadAsync(xlsx.bytes);
const threadedPart = Object.keys(xlsxZip.files).find((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
assert.ok(threadedPart);
const threadedXml = await xlsxZip.file(threadedPart).async("text");
assert.match(threadedXml, /parentId="\{[0-9A-F-]+\}"/);
assert.match(threadedXml, /Canonical direct reply/);
const importedWorkbook = await importXlsxWithOpenChestnut(xlsx);
const importedSheet = importedWorkbook.worksheets.getItem("Core");
assert.equal(importedSheet.getRange("B3").values[0][0], "B");
assert.ok(importedSheet.getRange("A2").values[0][0] > 40_000, "Date must cross the wire as an Excel serial");
assert.equal(importedSheet.getRange("D2:D3").dataValidation.type, "list");
assert.equal(importedSheet.conditionalFormattings.items.length, 1);
assert.equal(importedWorkbook.comments.threads.length, 1);
assert.equal(importedWorkbook.comments.threads[0].comments.length, 2);
assert.equal(importedWorkbook.comments.threads[0].comments[1].text, "Canonical direct reply");
const importedReply = importedWorkbook.comments.threads[0].comments[1];
const rootParentId = importedReply.parentId;
importedReply.parentId = importedReply.id;
await assert.rejects(exportXlsxWithOpenChestnut(importedWorkbook), /nested or branched reply graph/i);
importedReply.parentId = rootParentId;
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
  settings: { documentProtection: "comments" },
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
document.addParagraph("", { runs: [
  { text: "Owner: " },
  { text: "Ada", contentControl: { id: "owner-control", tag: "OWNER", alias: "Owner" } },
] });
const approvalParagraph = document.addParagraph("Approved: ");
approvalParagraph.addCheckboxContentControl(false, { id: "approval-control", tag: "APPROVED", alias: "Approved" });
const priorityParagraph = document.addParagraph("Priority: ");
priorityParagraph.addDropdownContentControl([
  { displayText: "Low", value: "low" },
  { displayText: "Medium", value: "medium" },
  { displayText: "High", value: "high" },
], { id: "priority-control", tag: "PRIORITY", alias: "Priority", selectedValue: "medium" });
const contactParagraph = document.addParagraph("Contact method: ");
contactParagraph.addComboBoxContentControl([
  { displayText: "Email", value: "email" },
  { displayText: "Phone call", value: "phone" },
], { id: "contact-method-control", tag: "CONTACT_METHOD", alias: "Contact method", value: "email" });
const reviewDateParagraph = document.addParagraph("Review date: ");
reviewDateParagraph.addDateContentControl("2026-07-21", { id: "review-date-control", tag: "REVIEW_DATE", alias: "Review date" });
document.addBlockTextContentControl("Executive summary", {
  blockId: "executive-summary-paragraph",
  id: "executive-summary-control",
  tag: "EXECUTIVE_SUMMARY",
  alias: "Executive summary",
  runStyle: { bold: true, color: "#1D4ED8" },
});
document.addHeader("Confidential", { referenceType: "default", sectionIndex: 0 });
document.addFooter("Page ", { referenceType: "default", sectionIndex: 0, fieldInstruction: "PAGE" });
document.addField("PAGE", "1");
document.addHyperlink("Evidence", "https://example.com/evidence");
document.addListItem("First action", { listType: "number", numberingId: 7 });
document.addListItem("Picture action one", {
  listType: "bullet",
  numberingId: 8,
  abstractNumberingId: 8,
  pictureBullet: { dataUrl: png, sizePt: 12, alt: "Action marker" },
});
document.addListItem("Picture action two", {
  listType: "bullet",
  numberingId: 8,
  abstractNumberingId: 8,
  pictureBullet: { dataUrl: png, sizePt: 12, alt: "Action marker" },
});
document.addListItem("External marker action", {
  listType: "bullet",
  level: 1,
  numberingId: 8,
  abstractNumberingId: 8,
  pictureBullet: { uri: "https://example.test/list-marker.png", sizePt: 10, alt: "External action marker" },
});
document.addTable({ values: [["Metric", "Value"], ["Revenue", "42"]], widthDxa: 9000, columnWidthsDxa: [3600, 5400], styleId: "TableGrid" });
document.addImage({
  name: "Logo",
  dataUrl: png,
  alt: "One pixel logo",
  widthPx: 32,
  heightPx: 32,
  placement: {
    type: "floating",
    horizontal: { relativeTo: "margin", offsetPx: 36 },
    vertical: { relativeTo: "paragraph", offsetPx: 10 },
    wrap: "square",
    wrapSide: "bothSides",
    distanceFromTextPx: { top: 0, right: 12, bottom: 4, left: 12 },
  },
});
document.addSection({ breakType: "nextPage", orientation: "landscape", pageSize: { widthTwips: 15840, heightTwips: 12240 }, margins: { top: 720, right: 720, bottom: 720, left: 720 } });
document.addComment(bodyParagraph, "Review body paragraph", { author: "Reviewer", initials: "RV", date: "2026-07-16T08:00:00Z" });

const docx = await exportDocxWithOpenChestnut(document);
assert.equal(docx.metadata.codec, "open-chestnut");
assert.deepEqual([...docx.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
const docxZip = await JSZip.loadAsync(docx.bytes);
assert.ok(docxZip.file("word/document.xml"));
assert.ok(docxZip.file("word/styles.xml"));
assert.match(await docxZip.file("word/settings.xml").async("text"), /<w:documentProtection(?=[^>]*w:edit="comments")(?=[^>]*w:enforcement="true")[^>]*\/>/);
assert.match(await docxZip.file("word/document.xml").async("text"), /<w:sdt>[\s\S]*<w:tag w:val="OWNER"\s*\/>[\s\S]*<w:text\s*\/>[\s\S]*Ada[\s\S]*<\/w:sdt>/);
const docxXml = await docxZip.file("word/document.xml").async("text");
assert.match(docxXml, /<wp:anchor(?=[^>]*behindDoc="0")(?=[^>]*allowOverlap="0")[^>]*>/);
assert.match(docxXml, /<wp:positionH relativeFrom="margin"><wp:posOffset>342900<\/wp:posOffset><\/wp:positionH>/);
assert.match(docxXml, /<wp:positionV relativeFrom="paragraph"><wp:posOffset>95250<\/wp:posOffset><\/wp:positionV>/);
assert.match(docxXml, /<wp:wrapSquare wrapText="bothSides"\s*\/>/);
assert.match(docxXml, /<w:tag w:val="APPROVED"\s*\/>[\s\S]*<w14:checkbox>[\s\S]*<w14:checked w14:val="0"\s*\/>[\s\S]*☐/);
assert.match(docxXml, /<w14:checkedState(?=[^>]*w14:val="2612")(?=[^>]*w14:font="MS Gothic")[^>]*\/>/);
assert.match(docxXml, /<w14:uncheckedState(?=[^>]*w14:val="2610")(?=[^>]*w14:font="MS Gothic")[^>]*\/>/);
assert.match(docxXml, /<w:tag w:val="PRIORITY"\s*\/>[\s\S]*<w:dropDownList w:lastValue="medium">[\s\S]*<w:listItem(?=[^>]*w:displayText="Low")(?=[^>]*w:value="low")[^>]*\/>[\s\S]*Medium/);
assert.match(docxXml, /<w:tag w:val="CONTACT_METHOD"\s*\/>[\s\S]*<w:comboBox w:lastValue="email">[\s\S]*<w:listItem(?=[^>]*w:displayText="Phone call")(?=[^>]*w:value="phone")[^>]*\/>[\s\S]*Email/);
assert.match(docxXml, /<w:tag w:val="REVIEW_DATE"\s*\/>[\s\S]*<w:date w:fullDate="2026-07-21T00:00:00Z">[\s\S]*<w:dateFormat w:val="yyyy-MM-dd"\s*\/>[\s\S]*<w:lid w:val="en-US"\s*\/>[\s\S]*<w:storeMappedDataAs w:val="date"\s*\/>[\s\S]*<w:calendar w:val="gregorian"\s*\/>[\s\S]*2026-07-21/);
assert.match(docxXml, /<w:sdt>[\s\S]*?<w:tag w:val="EXECUTIVE_SUMMARY"\s*\/>[\s\S]*?<w:text\s*\/>[\s\S]*?<w:sdtContent>\s*<w:p>[\s\S]*Executive summary[\s\S]*?<\/w:p>\s*<\/w:sdtContent>\s*<\/w:sdt>/);
const numberingXml = await docxZip.file("word/numbering.xml").async("text");
assert.equal((numberingXml.match(/<w:numPicBullet\b/g) || []).length, 2);
assert.equal((numberingXml.match(/<w:lvlPicBulletId\b/g) || []).length, 2);
assert.match(numberingXml, /<v:shape(?=[^>]*style="width:12pt;height:12pt")(?=[^>]*alt="Action marker")(?=[^>]*o:bullet="(?:t|true)")[^>]*>/);
assert.ok(Object.keys(docxZip.files).some((part) => /(?:^|\/)media\/[^/]+\.png$/.test(part)));
const importedDocument = await importDocxWithOpenChestnut(docx);
assert.deepEqual(importedDocument.settings.documentProtection, { edit: "comments", enforcement: true, formatting: false });
assert.equal(importedDocument.defaultRunStyle.fontFamily, "Aptos");
assert.ok(importedDocument.styles.get("CoreHeading"));
assert.equal(importedDocument.blocks[0].text, "Quarterly brief");
assert.equal(importedDocument.blocks[0].runs[1].style.italic, true);
assert.equal(importedDocument.headers[0].text, "Confidential");
assert.equal(importedDocument.footers[0].fieldInstruction, "PAGE");
assert.equal(importedDocument.comments.length, 1);
assert.equal(importedDocument.contentControls[0].tag, "OWNER");
assert.equal(importedDocument.contentControls[1].controlType, "checkbox");
assert.equal(importedDocument.contentControls[1].checked, false);
assert.equal(importedDocument.contentControls[2].controlType, "dropdown");
assert.equal(importedDocument.contentControls[2].selectedValue, "medium");
assert.deepEqual(importedDocument.contentControls[2].choices.map((choice) => choice.value), ["low", "medium", "high"]);
assert.equal(importedDocument.contentControls[3].controlType, "comboBox");
assert.equal(importedDocument.contentControls[3].value, "email");
assert.deepEqual(importedDocument.contentControls[3].choices.map((choice) => choice.value), ["email", "phone"]);
assert.equal(importedDocument.contentControls[4].controlType, "date");
assert.equal(importedDocument.contentControls[4].dateValue, "2026-07-21");
assert.equal(importedDocument.contentControls[5].placement, "block");
assert.equal(importedDocument.contentControls[5].tag, "EXECUTIVE_SUMMARY");
const importedPictureItems = importedDocument.blocks.filter((block) => block.kind === "listItem" && block.pictureBullet);
assert.equal(importedPictureItems.length, 3);
assert.deepEqual(importedPictureItems[0].pictureBullet, {
  dataUrl: png,
  uri: undefined,
  widthPt: 12,
  heightPt: 12,
  alt: "Action marker",
});
assert.deepEqual(importedPictureItems[0].pictureBullet, importedPictureItems[1].pictureBullet);
assert.equal(importedPictureItems[2].pictureBullet.uri, "https://example.test/list-marker.png");
const importedDocumentImage = importedDocument.blocks.find((block) => block.kind === "image");
const importedDocumentImageIndex = importedDocument.blocks.indexOf(importedDocumentImage);
assert.equal(importedDocumentImage.placement.type, "floating");
assert.deepEqual(importedDocumentImage.placement.horizontal, { relativeTo: "margin", offsetPx: 36 });
assert.deepEqual(importedDocumentImage.placement.vertical, { relativeTo: "paragraph", offsetPx: 10 });
assert.equal(importedDocumentImage.placement.wrap, "square");
assert.equal(importedDocumentImage.placement.wrapSide, "bothSides");
assert.deepEqual(importedDocumentImage.placement.distanceFromTextPx, { top: 0, right: 12, bottom: 4, left: 12 });
const unchangedDocx = await exportDocxWithOpenChestnut(importedDocument);
assert.deepEqual(Buffer.from(unchangedDocx.bytes), Buffer.from(docx.bytes));

const gifPictureDocument = DocumentModel.create({ blocks: [] });
gifPictureDocument.addListItem("GIF marker", {
  listType: "bullet",
  numberFormat: "bullet",
  levelText: "•",
  numberingId: 71,
  abstractNumberingId: 71,
  pictureBullet: { dataUrl: gif, widthPt: 11, heightPt: 9, alt: "GIF marker" },
});
const gifPictureRoundtrip = await importDocxWithOpenChestnut(await exportDocxWithOpenChestnut(gifPictureDocument));
assert.deepEqual(gifPictureRoundtrip.blocks[0].pictureBullet, {
  dataUrl: gif,
  uri: undefined,
  widthPt: 11,
  heightPt: 9,
  alt: "GIF marker",
});

const conflictingPictureDocument = DocumentModel.create({ blocks: [] });
for (const [text, dataUrl] of [["Blue marker", png], ["Green marker", replacementPng]]) {
  conflictingPictureDocument.addListItem(text, {
    listType: "bullet",
    numberFormat: "bullet",
    levelText: "•",
    numberingId: 72,
    abstractNumberingId: 72,
    pictureBullet: { dataUrl, sizePt: 12, alt: `${text} image` },
  });
}
await assert.rejects(
  exportDocxWithOpenChestnut(conflictingPictureDocument),
  (error) => error?.code === "invalid_document_numbering" && /conflicting definitions/i.test(error.message),
);

const invalidPictureDocument = DocumentModel.create({ blocks: [] });
invalidPictureDocument.addListItem("Invalid marker", {
  listType: "bullet",
  numberFormat: "bullet",
  levelText: "•",
  pictureBullet: "data:image/png;base64,AAAA",
});
await assert.rejects(
  exportDocxWithOpenChestnut(invalidPictureDocument),
  (error) => error?.code === "invalid_document_picture_bullet" && /do not match image\/png|valid PNG/i.test(error.message),
);

const ambiguousPictureDocument = DocumentModel.create({ blocks: [] });
const ambiguousPictureItem = ambiguousPictureDocument.addListItem("Ambiguous marker", {
  listType: "bullet",
  numberFormat: "bullet",
  levelText: "•",
  pictureBullet: { dataUrl: gif, sizePt: 12, alt: "Ambiguous marker" },
});
ambiguousPictureItem.pictureBullet.uri = "https://example.test/also-a-marker.gif";
await assert.rejects(
  exportDocxWithOpenChestnut(ambiguousPictureDocument),
  (error) => error?.code === "invalid_document_picture_bullet" && /exactly one embedded dataUrl or external uri/i.test(error.message),
);

const noncanonicalPictureDocument = DocumentModel.create({ blocks: [] });
noncanonicalPictureDocument.addListItem("Noncanonical marker", {
  listType: "bullet",
  numberFormat: "bullet",
  levelText: "•",
  pictureBullet: { dataUrl: gif.replace(/=+$/, ""), sizePt: 12, alt: "Noncanonical marker" },
});
await assert.rejects(
  exportDocxWithOpenChestnut(noncanonicalPictureDocument),
  (error) => error?.code === "invalid_document_picture_bullet" && /valid decoded bytes/i.test(error.message),
);

const partialPictureBulletEdit = await importDocxWithOpenChestnut(docx);
partialPictureBulletEdit.blocks.find((block) => block.text === "Picture action one").pictureBullet.alt = "Partial marker edit";
await assert.rejects(
  exportDocxWithOpenChestnut(partialPictureBulletEdit),
  (error) => error?.code === "unsupported_document_edit" && /coherently/i.test(error.message),
);

const pictureSourceKindTransition = await importDocxWithOpenChestnut(docx);
for (const block of pictureSourceKindTransition.blocks.filter((item) => item.pictureBullet?.dataUrl)) {
  block.pictureBullet = {
    uri: "https://example.test/replacement-marker.png",
    widthPt: block.pictureBullet.widthPt,
    heightPt: block.pictureBullet.heightPt,
    alt: block.pictureBullet.alt,
  };
}
await assert.rejects(
  exportDocxWithOpenChestnut(pictureSourceKindTransition),
  (error) => error?.code === "unsupported_document_edit" && /source topology is source-bound/i.test(error.message),
);

const inlineTransitionDocument = await importDocxWithOpenChestnut(docx);
delete inlineTransitionDocument.blocks.find((block) => block.kind === "image").placement;
await assert.rejects(
  exportDocxWithOpenChestnut(inlineTransitionDocument),
  (error) => error?.code === "unsupported_document_image_edit" && /inline and floating/i.test(error.message),
);

const unsupportedAnchorXml = docxXml.replace('behindDoc="0"', 'behindDoc="1"');
assert.notEqual(unsupportedAnchorXml, docxXml);
const unsupportedAnchorDocx = await DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: unsupportedAnchorXml }]);
const unsupportedAnchorDocument = await importDocxWithOpenChestnut(unsupportedAnchorDocx);
const unsupportedAnchorBlock = unsupportedAnchorDocument.blocks[importedDocumentImageIndex];
assert.equal(unsupportedAnchorBlock.kind, "paragraph");
assert.equal(unsupportedAnchorBlock.textEditable, false);
assert.match(unsupportedAnchorBlock.name, /Preserved p/i);
const preservedUnsupportedAnchor = await exportDocxWithOpenChestnut(unsupportedAnchorDocument);
assert.deepEqual(Buffer.from(preservedUnsupportedAnchor.bytes), Buffer.from(unsupportedAnchorDocx.bytes));
unsupportedAnchorBlock.text = "Unsafe semantic replacement";
await assert.rejects(
  exportDocxWithOpenChestnut(unsupportedAnchorDocument),
  (error) => error?.code === "unsupported_document_edit" && /read-only/i.test(error.message),
);
assert.deepEqual(importedDocument.fillContentControls({ OWNER: "Grace" }), { updated: 1, matchedTags: ["OWNER"], missingTags: [] });
assert.deepEqual(importedDocument.setCheckboxContentControls({ APPROVED: true }), { updated: 1, matchedTags: ["APPROVED"], missingTags: [] });
assert.deepEqual(importedDocument.setDropdownContentControls({ PRIORITY: "high" }), { updated: 1, matchedTags: ["PRIORITY"], missingTags: [] });
assert.deepEqual(importedDocument.setComboBoxContentControls({ CONTACT_METHOD: "Pager duty" }), { updated: 1, matchedTags: ["CONTACT_METHOD"], missingTags: [] });
assert.deepEqual(importedDocument.setDateContentControls({ REVIEW_DATE: "2028-02-29" }), { updated: 1, matchedTags: ["REVIEW_DATE"], missingTags: [] });
assert.deepEqual(importedDocument.fillContentControls({ EXECUTIVE_SUMMARY: "Updated executive summary" }), { updated: 1, matchedTags: ["EXECUTIVE_SUMMARY"], missingTags: [] });
importedDocumentImage.placement = {
  type: "floating",
  horizontal: { relativeTo: "page", offsetPx: 48 },
  vertical: { relativeTo: "margin", offsetPx: 16 },
  wrap: "topAndBottom",
  distanceFromTextPx: { top: 6, right: 0, bottom: 8, left: 0 },
};
importedDocument.blocks[2].text = "Edited through OpenChestnut.";
importedDocument.blocks[2].runs = [{ text: importedDocument.blocks[2].text, style: {} }];
for (const block of importedPictureItems.slice(0, 2)) {
  block.pictureBullet.dataUrl = replacementPng;
  block.pictureBullet.widthPt = 15;
  block.pictureBullet.heightPt = 14;
  block.pictureBullet.alt = "Updated action marker";
}
const docx2 = await exportDocxWithOpenChestnut(importedDocument);
const importedDocument2 = await importDocxWithOpenChestnut(docx2);
assert.equal(importedDocument2.blocks[2].text, "Edited through OpenChestnut.");
assert.equal(importedDocument2.contentControls[0].text, "Grace");
assert.equal(importedDocument2.contentControls[1].checked, true);
assert.equal(importedDocument2.contentControls[2].selectedValue, "high");
assert.equal(importedDocument2.contentControls[2].text, "High");
assert.equal(importedDocument2.contentControls[3].value, "Pager duty");
assert.equal(importedDocument2.contentControls[3].text, "Pager duty");
assert.equal(importedDocument2.contentControls[4].dateValue, "2028-02-29");
assert.equal(importedDocument2.contentControls[4].text, "2028-02-29");
assert.equal(importedDocument2.contentControls[5].placement, "block");
assert.equal(importedDocument2.contentControls[5].text, "Updated executive summary");
const editedPictureItems = importedDocument2.blocks.filter((block) => block.kind === "listItem" && block.pictureBullet);
assert.equal(editedPictureItems[0].pictureBullet.dataUrl, replacementPng);
assert.equal(editedPictureItems[0].pictureBullet.widthPt, 15);
assert.equal(editedPictureItems[0].pictureBullet.heightPt, 14);
assert.equal(editedPictureItems[0].pictureBullet.alt, "Updated action marker");
assert.deepEqual(editedPictureItems[0].pictureBullet, editedPictureItems[1].pictureBullet);
assert.equal(editedPictureItems[2].pictureBullet.uri, "https://example.test/list-marker.png");
const editedDocumentImage = importedDocument2.blocks.find((block) => block.kind === "image");
assert.deepEqual(editedDocumentImage.placement.horizontal, { relativeTo: "page", offsetPx: 48 });
assert.deepEqual(editedDocumentImage.placement.vertical, { relativeTo: "margin", offsetPx: 16 });
assert.equal(editedDocumentImage.placement.wrap, "topAndBottom");
assert.equal(editedDocumentImage.placement.wrapSide, undefined);
assert.deepEqual(editedDocumentImage.placement.distanceFromTextPx, { top: 6, right: 0, bottom: 8, left: 0 });
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
const authoredPptxZip = await JSZip.loadAsync(pptx.bytes);
const authoredSlideXml = await authoredPptxZip.file("ppt/slides/slide1.xml").async("text");
const backgroundFillSlideXml = authoredSlideXml.replace("<p:sp>", '<p:sp useBgFill="1">');
assert.notEqual(backgroundFillSlideXml, authoredSlideXml);
const backgroundFillPptx = await PresentationFile.patchPptx(pptx, [{ path: "ppt/slides/slide1.xml", xml: backgroundFillSlideXml }]);
const backgroundFillPresentation = await importPptxWithOpenChestnut(backgroundFillPptx);
const backgroundFillShape = backgroundFillPresentation.slides.getItem(0).shapes.items[0];
assert.equal(backgroundFillShape.useBackgroundFill, true);
assert.throws(() => { backgroundFillShape.useBackgroundFill = false; }, TypeError);
assert.match(backgroundFillShape.toSvg(), /fill="#ffffff"/i);
const backgroundFillRoundTrip = await exportPptxWithOpenChestnut(backgroundFillPresentation);
const backgroundFillRoundTripZip = await JSZip.loadAsync(backgroundFillRoundTrip.bytes);
assert.equal(await backgroundFillRoundTripZip.file("ppt/slides/slide1.xml").async("text"), backgroundFillSlideXml);
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

for (const options of [
  { codec: "open-chestnut" },
  { codec: "javascript" },
  { allowLossy: true },
  { preferNative: true },
  { relativeDateAsOf: "2026-07-16" },
]) {
  await assert.rejects(
    SpreadsheetFile.exportXlsx(workbook, options),
    /does not accept option|only Office codec|lossy fallback/i,
  );
}
await assert.rejects(PresentationFile.importPptx(new Uint8Array(), "open-chestnut"), /options must be an object/i);
await assert.rejects(DocumentFile.exportDocx(document, { allowLossy: true }), /does not accept option|lossy fallback/i);
await assert.rejects(DocumentFile.importDocx(docx, { preferNative: true }), /does not accept option|only Office codec/i);
await assert.rejects(PresentationFile.exportPptx(presentation, { codec: "javascript" }), /does not accept option|only Office codec/i);
await assert.rejects(PresentationFile.importPptx(pptx, { allowLossy: true }), /does not accept option|lossy fallback/i);

console.log("OpenChestnut protocol 2 canonical core smoke ok");
