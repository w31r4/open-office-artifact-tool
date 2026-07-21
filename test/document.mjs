import assert from "node:assert/strict";
import { createHash } from "node:crypto";

import { DocumentFile, DocumentModel } from "../src/index.mjs";
import { DocumentFile as DocumentFileModule, DocumentModel as DocumentModelModule } from "../src/document/index.mjs";

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
assert.equal(importedExplicitOpen.comments[0]._resolvedSpecified, true);

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
