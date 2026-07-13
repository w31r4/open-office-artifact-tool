import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, FileBlob, renderArtifact } from "../src/index.mjs";

const pictureBulletPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAABAAAAAQCAYAAAAf8/9hAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAAHUlEQVR4nGNQOhr3nxLMMGrA/9EwiBsNg6PDIgwAUQdEH39xn2wAAAAASUVORK5CYII=";

const document = DocumentModel.create({
  name: "Research memo",
  paragraphs: ["Research memo", "This document exercises the clean-room DOCX facade."],
  theme: {
    name: "Research Theme",
    colors: { accent1: "#336699", accent2: "#cc3300" },
    fonts: { major: "Source Serif 4", minor: "Aptos", majorEastAsia: "Noto Serif CJK SC", majorComplexScript: "Noto Naskh Arabic" },
  },
  defaultRunStyle: { fontFamily: "Default Serif", fontSize: 19, italic: true, color: "#111111" },
  settings: { updateFields: true, mirrorMargins: true },
});
document.setSettings({ trackRevisions: true, documentProtection: { edit: "trackedChanges" } });
document.applyDesignPreset("report");
document.styles.add("Callout", { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos", color: "#0f172a" });
document.styles.add("RiskCallout", { name: "Risk Callout", basedOn: "Callout", italic: true, color: "#b91c1c" });
document.styles.add("CascadeParagraphBase", { name: "Cascade Paragraph Base", basedOn: "Normal", fontSize: 26, bold: true });
document.styles.add("CascadeParagraph", { name: "Cascade Paragraph", basedOn: "CascadeParagraphBase", themeColor: "accent1" });
document.styles.add("CascadeCharacterBase", { name: "Cascade Character Base", type: "character", italic: true, themeColor: "accent2" });
document.styles.add("CascadeCharacter", { name: "Cascade Character", type: "character", basedOn: "CascadeCharacterBase", bold: false, fontTheme: "majorHAnsi" });
const header = document.addHeader("Confidential research memo", { name: "default-header" });
const footer = document.addFooter("Page footer", { name: "default-footer" });
const firstHeader = document.addHeader("First-page research memo", { name: "first-header", referenceType: "first" });
const evenFooter = document.addFooter("Even-page footer", { name: "even-footer", referenceType: "even" });
const heading = document.addParagraph("Findings", { styleId: "Heading1", name: "findings-heading" });
const riskCallout = document.addParagraph("Risk callout inherits bold styling.", { styleId: "RiskCallout", name: "risk-callout" });
const runParagraph = document.addParagraph("", {
  name: "run-styled-paragraph",
  styleId: "Normal",
  runs: [
    { text: "Run styled ", style: { bold: true, color: "#0ea5e9" } },
    { text: "paragraph", style: { italic: true, color: "#f97316" } },
  ],
});
const themedRunParagraph = document.addParagraph("", {
  name: "theme-run-paragraph",
  styleId: "Normal",
  runs: [
    { text: "Theme Latin ", style: { fontTheme: "majorHAnsi", themeColor: "accent1", themeTint: "80", bold: true, boldComplexScript: false, fontSize: 28, fontSizeComplexScript: 34 } },
    { text: "中文 ", style: { fontTheme: "minorHAnsi", fontThemeEastAsia: "majorEastAsia", themeColor: "accent2", themeShade: "BF", italic: true } },
    { text: "العربية", style: { fontTheme: "minorHAnsi", fontThemeComplexScript: "majorBidi", bold: false, boldComplexScript: true, italic: false, italicComplexScript: true, fontSize: 22, fontSizeComplexScript: 32 } },
  ],
});
const singleThemeRunParagraph = document.addParagraph("", { name: "single-theme-run", runs: [{ text: "Single theme run", style: { fontTheme: "majorHAnsi", themeColor: "accent1" } }] });
const cascadeParagraph = document.addParagraph("", {
  name: "run-style-cascade",
  styleId: "CascadeParagraph",
  runs: [
    { text: "Direct override", style: { runStyleId: "CascadeCharacter", italic: false, color: "#123456" } },
    { text: " Character theme", style: { runStyleId: "CascadeCharacter" } },
  ],
});
const findingsBookmark = document.addBookmark(heading, "FindingsSection", { endTarget: riskCallout, nativeId: 42 });
const internalHyperlink = document.addHyperlink("Jump to findings", findingsBookmark, { name: "findings-link", tooltip: "Open the findings section", history: false });
const hyperlink = document.addHyperlink("w31r4 research note", "https://example.com/research", { name: "research-link" });
const field = document.addField("PAGE", "1", { name: "page-field" });
const uncitedBibliographySource = document.addBibliographySource({ tag: "OpenXmlSdk", sourceType: "InternetSite", title: "Open XML SDK documentation", corporateAuthor: "Microsoft", year: "2026", url: "https://learn.microsoft.com/office/open-xml/open-xml-sdk" });
const citation = document.addCitation("Source: Market brief", { tag: "MarketBrief26", source: "Market brief", sourceType: "Report", title: "Market brief", authors: [{ first: "Ada", last: "Analyst" }], year: "2026", publisher: "Example Research", url: "https://example.com/brief", page: 2 }, { name: "market-citation" });
const logo = document.addImage({
  name: "memo-logo",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  alt: "Memo logo",
  widthPx: 96,
  heightPx: 64,
});
const landscapeSection = document.addSection({
  name: "landscape-appendix",
  breakType: "nextPage",
  orientation: "landscape",
  pageSize: { widthTwips: 15840, heightTwips: 12240 },
  margins: { top: 720, right: 900, bottom: 720, left: 900 },
});
const openingHeader = document.addHeader("Opening-section header", { name: "opening-header", sectionIndex: 0 });
const openingFooter = document.addFooter("Opening-section footer", { name: "opening-footer", sectionIndex: 0 });
const insertion = document.addInsertion("Inserted reviewer clarification.", { author: "Reviewer", date: "2026-07-11T00:00:00.000Z", name: "tracked-insert" });
const deletion = document.addDeletion("Remove stale claim.", { author: "Reviewer", date: "2026-07-11T00:05:00.000Z", name: "tracked-delete" });
const bullet = document.addListItem("Use real numbering definitions", { listType: "bullet", name: "numbering-rule" });
const numbered = document.addListItem("Render and verify", { listType: "number", name: "render-step" });
const table = document.addTable({
  name: "evidence-table",
  styleId: "TableGrid",
  widthDxa: 9360,
  indentDxa: 120,
  columnWidthsDxa: [3600, 5760],
  headerFill: "F2F4F7",
  values: [["Area", "Status"], ["DOCX styles", "partial"], ["Comments", "roundtrip"]],
});
table.getCell(2, 1).value = "anchored";
const tableCellBookmark = document.addBookmark(table.getCell(1, 0), "EvidenceCells", { endTarget: table.getCell(2, 1), nativeId: 43 });
const customListParent = document.addListItem("Lettered evidence group", { listType: "number", level: 0, numberFormat: "upperLetter", start: 2, levelText: "%1)", numberingId: 42, name: "lettered-evidence" });
const customListChild = document.addListItem("Nested roman evidence", { listType: "number", level: 1, numberFormat: "lowerRoman", start: 3, levelText: "%1.%2)", numberingId: 42, name: "roman-evidence" });
const pictureBullet = document.addListItem("Picture bullet evidence", { name: "picture-bullet-evidence", pictureBullet: { dataUrl: pictureBulletPng, widthPt: 10, heightPt: 10, alt: "Green status marker" } });
const comment = document.addComment(heading, "Check this heading before final export.", { author: "Reviewer", initials: "RV", date: "2026-07-11T00:10:00.000Z", dateUtc: "2026-07-11T08:10:00+08:00", durableId: "0000A001", person: { providerId: "None", userId: "reviewer@example.test" }, resolved: true });
const tableComment = document.addComment(table, "Review the evidence table.", { author: "R&D Analyst", initials: "RA", date: "2026-07-11T00:15:00.000Z" });
const linkComment = document.addComment(hyperlink, "Verify the native hyperlink target.", { author: "Link Reviewer", initials: "LR", date: "2026-07-11T00:18:00.000Z" });
const commentReply = document.replyToComment(comment, "Heading language is now approved.", { author: "Editor", initials: "ED", date: "2026-07-11T00:20:00.000Z", durableId: "0000A004", person: { providerId: "None", userId: "editor@example.test" } });
assert.equal(commentReply.parentId, comment.id);
assert.equal(commentReply.targetId, comment.targetId);

const inspect = document.inspect({ kind: "document,theme,settings,paragraph,table,tableCell,bookmark,bibliographySource,comment,style,listItem,header,footer,hyperlink,field,citation,image,section,change,layout", maxChars: 32000 }).ndjson;
assert.match(inspect, /Research memo/);
assert.match(inspect, /"kind":"document"/);
assert.match(inspect, /"designPreset":"report"/);
assert.match(inspect, /Research Theme/);
assert.match(inspect, /Source Serif 4/);
assert.match(inspect, /theme-run-paragraph/);
assert.match(inspect, /"resolvedColor":"#99b3cc"/);
assert.equal(document.resolve(`${document.id}/theme`), document.theme);
assert.match(inspect, /"kind":"settings"/);
assert.match(inspect, /"trackRevisions":true/);
assert.equal(document.resolve(`${document.id}/settings`), document.settings);
assert.match(inspect, /"kind":"layout"/);
assert.match(inspect, /findings-heading/);
assert.match(inspect, /research-link/);
assert.match(inspect, /https:\/\/example.com\/research/);
assert.match(inspect, /findings-link/);
assert.match(inspect, /FindingsSection/);
assert.match(inspect, /page-field/);
assert.match(inspect, /market-citation/);
assert.match(inspect, /memo-logo/);
assert.match(inspect, /Memo logo/);
assert.match(inspect, /picture-bullet-evidence/);
assert.match(inspect, /Green status marker/);
assert.match(await (await document.render()).text(), /data-picture-bullet="embedded"/);
assert.throws(() => document.addListItem("Invalid picture number", { listType: "number", pictureBullet: logo.dataUrl }), /requires listType and numberFormat to be bullet/);
assert.match(inspect, /landscape-appendix/);
assert.match(inspect, /landscape/);
assert.match(inspect, /tracked-insert/);
assert.match(inspect, /Inserted reviewer clarification/);
assert.match(inspect, /tracked-delete/);
assert.match(inspect, /Remove stale claim/);
assert.match(inspect, /numbering-rule/);
assert.match(inspect, /render-step/);
assert.match(inspect, /lettered-evidence/);
assert.match(inspect, /lowerRoman/);
assert.match(inspect, /Confidential research memo/);
assert.match(inspect, /Page footer/);
assert.match(inspect, /First-page research memo/);
assert.match(inspect, /Even-page footer/);
assert.match(inspect, /Opening-section header/);
assert.match(inspect, /Opening-section footer/);
assert.equal(firstHeader.referenceType, "first");
assert.equal(evenFooter.referenceType, "even");
assert.equal(openingHeader.sectionIndex, 0);
assert.equal(openingFooter.sectionIndex, 0);
assert.match(inspect, /evidence-table/);
assert.match(inspect, /Check this heading/);
assert.match(inspect, /Review the evidence table/);
assert.match(inspect, /R&D Analyst/);
assert.match(inspect, /Verify the native hyperlink target/);
assert.match(inspect, /Callout/);
assert.match(inspect, /Risk Callout/);
assert.match(inspect, /Risk callout inherits bold styling/);
assert.match(inspect, /run-styled-paragraph/);
assert.match(inspect, /Run styled paragraph/);
assert.match(inspect, /"runs"/);
assert.match(inspect, /0ea5e9/);
assert.match(inspect, /"effectiveStyle"/);
assert.equal(document.resolve(heading.id).styleId, "Heading1");
assert.equal(document.resolve(runParagraph.id).text, "Run styled paragraph");
assert.equal(document.resolve(runParagraph.id).runs[1].style.italic, true);
assert.equal(document.resolve(themedRunParagraph.id).runs[0].style.resolvedFontFamily, "Source Serif 4");
assert.equal(document.resolve(themedRunParagraph.id).runs[0].style.resolvedColor, "#99b3cc");
assert.equal(document.resolve(themedRunParagraph.id).runs[1].style.resolvedFontFamilyEastAsia, "Noto Serif CJK SC");
assert.equal(document.resolve(themedRunParagraph.id).runs[1].style.resolvedColor, "#992600");
assert.equal(document.resolve(themedRunParagraph.id).runs[2].style.resolvedFontFamilyComplexScript, "Noto Naskh Arabic");
assert.equal(document.toProto().blocks.find((item) => item.name === "single-theme-run")?.runs[0].style.themeColor, "accent1");
assert.equal(document.defaultRunStyle.fontFamily, "Default Serif");
assert.equal(document.toProto().defaultRunStyle.fontSize, 19);
assert.equal(document.resolve(riskCallout.id).styleId, "RiskCallout");
assert.equal(document.styles.effective("RiskCallout").bold, true);
assert.equal(document.styles.effective("RiskCallout").italic, true);
assert.equal(document.styles.effective("RiskCallout").color, "#b91c1c");
assert.match(document.help("document.styles.effective").ndjson, /basedOn inheritance/);
const headingTextRange = document.resolve(`${heading.id}/text`);
assert.equal(headingTextRange.text, "Findings");
const documentTextRangeInspect = document.inspect({ kind: "textRange", target: `${heading.id}/text`, include: "text,parentId", maxChars: 4000 }).ndjson;
assert.match(documentTextRangeInspect, /Findings/);
assert.match(documentTextRangeInspect, new RegExp(heading.id.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
assert.equal(document.resolve(hyperlink.id).url, "https://example.com/research");
assert.equal(document.resolve(internalHyperlink.id).anchor, "FindingsSection");
assert.equal(document.resolve(findingsBookmark.id), findingsBookmark);
assert.equal(document.resolve("FindingsSection"), findingsBookmark);
assert.equal(findingsBookmark.targetId, heading.id);
assert.equal(findingsBookmark.endTargetId, riskCallout.id);
assert.equal(findingsBookmark.nativeId, 42);
assert.equal(tableCellBookmark.targetId, `${table.id}/cell/1/0`);
assert.equal(tableCellBookmark.endTargetId, `${table.id}/cell/2/1`);
assert.equal(document.resolve(tableCellBookmark.targetId)?.value, "DOCX styles");
assert.equal(document.resolve(tableCellBookmark.endTargetId)?.value, "anchored");
assert.match(inspect, /"kind":"tableCell"/);
assert.equal(document.resolve(field.id).instruction, "PAGE");
assert.equal(document.resolve(citation.id).metadata.page, 2);
assert.equal(document.resolve("MarketBrief26")?.title, "Market brief");
assert.equal(document.resolve(uncitedBibliographySource.id), uncitedBibliographySource);
assert.match(inspect, /"kind":"bibliographySource"/);
assert.match(inspect, /Open XML SDK documentation/);
assert.equal(document.resolve(logo.id).alt, "Memo logo");
assert.equal(document.resolve(logo.id).widthPx, 96);
assert.equal(document.resolve(landscapeSection.id).orientation, "landscape");
assert.equal(document.resolve(landscapeSection.id).margins.left, 900);
assert.equal(document.resolve(insertion.id).changeType, "insert");
assert.equal(document.resolve(insertion.id).author, "Reviewer");
assert.equal(document.resolve(deletion.id).changeType, "delete");
assert.equal(document.resolve(deletion.id).date, "2026-07-11T00:05:00.000Z");
assert.equal(document.resolve(bullet.id).listType, "bullet");
assert.equal(document.resolve(numbered.id).listType, "number");
assert.equal(document.resolve(customListParent.id).numberFormat, "upperLetter");
assert.equal(document.resolve(customListChild.id).level, 1);
assert.equal(document.resolve(customListChild.id).start, 3);
assert.equal(document.resolve(header.id).text, "Confidential research memo");
assert.equal(document.resolve(footer.id).text, "Page footer");
assert.equal(document.resolve(table.id).getCell(2, 1).value, "anchored");
assert.equal(document.resolve(comment.id).author, "Reviewer");
assert.equal(document.resolve(comment.id).initials, "RV");
assert.equal(document.resolve(comment.id).date, "2026-07-11T00:10:00.000Z");
assert.equal(document.resolve(tableComment.id).targetId, table.id);
assert.equal(document.resolve(linkComment.id).targetId, hyperlink.id);
const targetedDocumentInspect = document.inspect({ kind: "paragraph,table,image,comment", target: logo.id, maxChars: 4000 }).ndjson;
assert.match(targetedDocumentInspect, /Memo logo/);
assert.doesNotMatch(targetedDocumentInspect, /evidence-table/);
const shapedDocumentInspect = document.inspect({ kind: "paragraph", target: heading.id, include: "text,style", exclude: "comments", maxChars: 4000 }).ndjson;
assert.match(shapedDocumentInspect, /Findings/);
assert.match(shapedDocumentInspect, /Heading1/);
assert.doesNotMatch(shapedDocumentInspect, /Check this heading/);
const targetedCommentInspect = document.inspect({ kind: "comment,paragraph", target: comment.id, maxChars: 4000 }).ndjson;
assert.match(targetedCommentInspect, /Check this heading/);
assert.doesNotMatch(targetedCommentInspect, /findings-heading/);
assert.match(document.help("document.addParagraph").ndjson, /runStyleId/);
assert.match(document.help("document.addTable").ndjson, /Word-style table/);
assert.match(document.help("document.addListItem").ndjson, /numbering definitions/);
assert.match(document.help("document.addListItem").ndjson, /numberFormat/);
assert.match(document.help("document.addHeader").ndjson, /DOCX header/);
assert.match(document.help("document.addHeader").ndjson, /sectionIndex/);
assert.match(document.help("document.addComment").ndjson, /initials/);
assert.match(document.help("document.addComment").ndjson, /w:date/);
assert.match(document.help("document.addComment").ndjson, /commentsExtended/);
assert.match(document.help("document.replyToComment").ndjson, /paraIdParent/);
assert.match(document.help("document.addHyperlink").ndjson, /w:hyperlink/);
assert.match(document.help("document.addBookmark").ndjson, /bookmark range/);
assert.match(document.help("document.addField").ndjson, /w:fldSimple/);
assert.match(document.help("document.addCitation").ndjson, /structured metadata/);
assert.match(document.help("document.addImage").ndjson, /native DOCX media/);
assert.match(document.help("document.addSection").ndjson, /w:sectPr/);
assert.match(document.help("document.addInsertion").ndjson, /w:ins/);
assert.match(document.help("document.addDeletion").ndjson, /w:del/);
assert.match(document.help("document.setSectionSettings").ndjson, /different-first-page/);
assert.match(document.help("document.applyDesignPreset").ndjson, /design preset/);
assert.match(document.help("document.layoutJson").ndjson, /layout JSON/);
assert.equal(document.verify({ visualQa: true }).ok, true);
const inheritedHeaderDocument = DocumentModel.create({ blocks: [] });
const inheritedDefaultHeader = inheritedHeaderDocument.addHeader("Section zero default header", { sectionIndex: 0 });
const inheritedFirstHeader = inheritedHeaderDocument.addHeader("Section zero first header", { sectionIndex: 0, referenceType: "first" });
const inheritedEvenHeader = inheritedHeaderDocument.addHeader("Section zero even header", { sectionIndex: 0, referenceType: "even" });
const inheritedDefaultFooter = inheritedHeaderDocument.addFooter("Section zero default footer", { sectionIndex: 0 });
const inheritedEvenFooter = inheritedHeaderDocument.addFooter("Section zero even footer", { sectionIndex: 0, referenceType: "even" });
inheritedHeaderDocument.addParagraph("Section zero body");
inheritedHeaderDocument.addSection({ breakType: "nextPage" });
for (let index = 0; index < 4; index += 1) inheritedHeaderDocument.addParagraph(`Section one body ${index + 1}`);
const inheritedHeaderLayout = inheritedHeaderDocument.layoutJson({ pageHeight: 90, margin: 10 });
assert.equal(inheritedHeaderLayout.pages.length, 3);
assert.deepEqual(inheritedHeaderLayout.pages[0].headers, [inheritedFirstHeader.id]);
assert.deepEqual(inheritedHeaderLayout.pages[0].footers, []);
assert.equal(inheritedHeaderLayout.pages[0].header.referenceType, "first");
assert.equal(inheritedHeaderLayout.pages[0].header.inherited, false);
assert.deepEqual(inheritedHeaderLayout.pages[1].headers, [inheritedDefaultHeader.id]);
assert.deepEqual(inheritedHeaderLayout.pages[1].footers, [inheritedDefaultFooter.id]);
assert.equal(inheritedHeaderLayout.pages[1].sectionIndex, 1);
assert.equal(inheritedHeaderLayout.pages[1].pageInSection, 1);
assert.equal(inheritedHeaderLayout.pages[1].header.inherited, true);
assert.equal(inheritedHeaderLayout.pages[1].header.sourceSectionIndex, 0);
assert.deepEqual(inheritedHeaderLayout.pages[2].headers, [inheritedEvenHeader.id]);
assert.deepEqual(inheritedHeaderLayout.pages[2].footers, [inheritedEvenFooter.id]);
assert.equal(inheritedHeaderLayout.pages[2].pageInSection, 2);
assert.equal(inheritedHeaderLayout.pages[2].header.referenceType, "even");
assert.equal(inheritedHeaderLayout.pages[2].header.inherited, true);
inheritedHeaderDocument.setSectionSettings(0, { differentFirstPage: false });
assert.deepEqual(inheritedHeaderDocument.layoutJson({ pageHeight: 90, margin: 10 }).pages[0].headers, [inheritedDefaultHeader.id]);
inheritedHeaderDocument.setSectionSettings(0, { differentFirstPage: true });
assert.deepEqual(inheritedHeaderDocument.layoutJson({ pageHeight: 90, margin: 10 }).pages[0].headers, [inheritedFirstHeader.id]);
assert.throws(() => inheritedHeaderDocument.setSectionSettings(2, { differentFirstPage: true }), /must be an integer from 0 through 1/);

const dormantVariantDocument = DocumentModel.create({
  blocks: [{ kind: "paragraph", text: "Dormant header variants" }],
  settings: { evenAndOddHeaders: false },
  sectionSettings: [{ sectionIndex: 0, differentFirstPage: false }],
  headers: [
    { text: "Active default header", referenceType: "default" },
    { text: "Dormant first header", referenceType: "first", variantActive: false },
    { text: "Dormant even header", referenceType: "even", variantActive: false },
  ],
});
assert.deepEqual(dormantVariantDocument.layoutJson().pages[0].headers, [dormantVariantDocument.headers[0].id]);
const dormantVariantDocx = await DocumentFile.exportDocx(dormantVariantDocument);
const dormantVariantZip = await JSZip.loadAsync(new Uint8Array(await dormantVariantDocx.arrayBuffer()));
const dormantVariantXml = await dormantVariantZip.file("word/document.xml").async("text");
assert.match(dormantVariantXml, /<w:headerReference w:type="first"/);
assert.match(dormantVariantXml, /<w:headerReference w:type="even"/);
assert.doesNotMatch(dormantVariantXml, /<w:titlePg\/>/);
assert.ok(!dormantVariantZip.file("word/settings.xml") || !/<w:evenAndOddHeaders/.test(await dormantVariantZip.file("word/settings.xml").async("text")));
dormantVariantZip.remove("word/open-office-artifact.json");
const dormantNativeDocx = new FileBlob(await dormantVariantZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: dormantVariantDocx.type });
const dormantNativeDocument = await DocumentFile.importDocx(dormantNativeDocx, { preferNative: true });
assert.equal(dormantNativeDocument.settings.evenAndOddHeaders, false);
assert.equal(dormantNativeDocument.sectionSettings[0].differentFirstPage, false);
assert.equal(dormantNativeDocument.headers.find((item) => item.referenceType === "first")?.variantActive, false);
assert.equal(dormantNativeDocument.headers.find((item) => item.referenceType === "even")?.variantActive, false);
assert.deepEqual(dormantNativeDocument.layoutJson().pages[0].headers, [dormantNativeDocument.headers.find((item) => item.referenceType === "default").id]);
const dormantSecondExportZip = await JSZip.loadAsync(new Uint8Array(await (await DocumentFile.exportDocx(dormantNativeDocument)).arrayBuffer()));
assert.doesNotMatch(await dormantSecondExportZip.file("word/document.xml").async("text"), /<w:titlePg\/>/);
assert.ok(!dormantSecondExportZip.file("word/settings.xml") || !/<w:evenAndOddHeaders/.test(await dormantSecondExportZip.file("word/settings.xml").async("text")));
const alternateSectionPrefixZip = await JSZip.loadAsync(new Uint8Array(await dormantNativeDocx.arrayBuffer()));
const alternateSectionPrefixXml = dormantVariantXml
  .replace("<w:document ", '<w:document xmlns:s="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ')
  .replace(/<w:sectPr\b[\s\S]*?<\/w:sectPr>/g, (sectionXml) => sectionXml.replaceAll("<w:", "<s:").replaceAll("</w:", "</s:").replaceAll(" w:", " s:"));
alternateSectionPrefixZip.file("word/document.xml", alternateSectionPrefixXml);
const alternateSectionPrefixDocx = new FileBlob(await alternateSectionPrefixZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: dormantVariantDocx.type });
const alternateSectionPrefixNative = await DocumentFile.importDocx(alternateSectionPrefixDocx, { preferNative: true });
assert.equal(alternateSectionPrefixNative.sectionSettings[0].differentFirstPage, false);
assert.deepEqual(alternateSectionPrefixNative.headers.map((item) => item.referenceType).sort(), ["default", "even", "first"]);
const invalidBookmarkDocument = DocumentModel.create({ paragraphs: ["First", "Second"] });
invalidBookmarkDocument.addBookmark(invalidBookmarkDocument.blocks[1], "ReversedRange", { endTarget: invalidBookmarkDocument.blocks[0] });
invalidBookmarkDocument.addBookmark(invalidBookmarkDocument.blocks[0], "DuplicateName");
invalidBookmarkDocument.addBookmark(invalidBookmarkDocument.blocks[1], "DuplicateName");
invalidBookmarkDocument.addHyperlink("Missing target", undefined, { anchor: "UnknownBookmark" });
const invalidBookmarkIssues = invalidBookmarkDocument.verify().issues;
assert.ok(invalidBookmarkIssues.some((issue) => issue.type === "reversedBookmarkRange"));
assert.ok(invalidBookmarkIssues.some((issue) => issue.type === "duplicateBookmarkName"));
assert.ok(invalidBookmarkIssues.some((issue) => issue.type === "missingHyperlinkAnchor"));
await assert.rejects(() => DocumentFile.exportDocx(invalidBookmarkDocument), /Duplicate DOCX bookmark name/);
const reversedBookmarkDocument = DocumentModel.create({ paragraphs: ["First", "Second"] });
reversedBookmarkDocument.addBookmark(reversedBookmarkDocument.blocks[1], "ReversedOnly", { endTarget: reversedBookmarkDocument.blocks[0] });
await assert.rejects(() => DocumentFile.exportDocx(reversedBookmarkDocument), /end target precedes its start target/);
const tableBookmarkDocument = DocumentModel.create({ blocks: [] });
const bookmarkTable = tableBookmarkDocument.addTable({ values: [["Unsupported"]] });
tableBookmarkDocument.addBookmark(bookmarkTable, "TableBookmark");
assert.ok(tableBookmarkDocument.verify().issues.some((issue) => issue.type === "unsupportedBookmarkTarget"));
await assert.rejects(() => DocumentFile.exportDocx(tableBookmarkDocument), /must identify a table cell/);
const validTableBookmarkDocument = DocumentModel.create({ blocks: [] });
const validBookmarkTable = validTableBookmarkDocument.addTable({ values: [["A", "B"], ["C", "D"]] });
validTableBookmarkDocument.addBookmark(validBookmarkTable.getCell(0, 1), "ValidTableCells", { endTarget: validBookmarkTable.getCell(1, 1) });
assert.ok(!validTableBookmarkDocument.verify().issues.some((issue) => /Bookmark/.test(issue.message)));
validTableBookmarkDocument.addBookmark(validBookmarkTable.getCell(1, 1), "ReversedTableCells", { endTarget: validBookmarkTable.getCell(0, 1) });
assert.ok(validTableBookmarkDocument.verify().issues.some((issue) => issue.type === "reversedBookmarkRange"));
const invalidCellBookmarkDocument = DocumentModel.create({ blocks: [] });
const invalidCellTable = invalidCellBookmarkDocument.addTable({ values: [["Only"]] });
invalidCellBookmarkDocument.addBookmark(invalidCellTable.getCell(4, 0), "InvalidTableCell");
assert.ok(invalidCellBookmarkDocument.verify().issues.some((issue) => issue.type === "invalidBookmarkTableCell"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCellBookmarkDocument), /row 4 is out of range/);
const duplicateBibliographyDocument = DocumentModel.create({ paragraphs: ["Bibliography fixture"] });
assert.throws(() => duplicateBibliographyDocument.addBibliographySource({ tag: "BadType", sourceType: "Unknown", title: "Invalid" }), /unsupported SourceType Unknown/);
assert.throws(() => duplicateBibliographyDocument.addBibliographySource({ tag: "X".repeat(256), sourceType: "Book", title: "Invalid" }), /at most 255 characters/);
assert.throws(() => duplicateBibliographyDocument.addBibliographySource({ tag: "AmbiguousAuthor", sourceType: "Book", authors: ["Ada Analyst"], corporateAuthor: "Example Corp" }), /cannot combine personal authors/);
duplicateBibliographyDocument.addBibliographySource({ tag: "DuplicateSource", sourceType: "Book", title: "First" });
duplicateBibliographyDocument.addBibliographySource({ tag: "DuplicateSource", sourceType: "Report", title: "Second" });
assert.ok(duplicateBibliographyDocument.verify().issues.some((issue) => issue.type === "duplicateBibliographyTag"));
await assert.rejects(() => DocumentFile.exportDocx(duplicateBibliographyDocument), /Duplicate DOCX bibliography source tag DuplicateSource/);
const spacedCitationDocument = DocumentModel.create({ blocks: [] });
spacedCitationDocument.addBibliographySource({ tag: "Space Tag", sourceType: "Book", title: "Quoted field tag" });
spacedCitationDocument.addCitation("Quoted citation", { tag: "Space Tag" });
const spacedCitationDocx = await DocumentFile.exportDocx(spacedCitationDocument);
const spacedCitationZip = await JSZip.loadAsync(new Uint8Array(await spacedCitationDocx.arrayBuffer()));
assert.match(await spacedCitationZip.file("word/document.xml").async("text"), /w:instr="CITATION &quot;Space Tag&quot;"/);
spacedCitationZip.remove("word/open-office-artifact.json");
const spacedCitationImported = await DocumentFile.importDocx(new FileBlob(await spacedCitationZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: spacedCitationDocx.type }), { preferNative: true });
assert.equal(spacedCitationImported.blocks[0].metadata.tag, "Space Tag");
assert.equal(spacedCitationImported.resolve("Space Tag")?.title, "Quoted field tag");
const invalidSectionDocument = DocumentModel.create({ paragraphs: ["Invalid section fixture"] });
invalidSectionDocument.addHeader("Out of range", { sectionIndex: 1 });
assert.ok(invalidSectionDocument.verify().issues.some((issue) => issue.type === "invalidHeaderFooterSection"));
await assert.rejects(() => DocumentFile.exportDocx(invalidSectionDocument), /sectionIndex must be an integer from 0 through 0/);
const invalidCommentDocument = DocumentModel.create({ paragraphs: ["Invalid comment fixture"] });
invalidCommentDocument.addComment(invalidCommentDocument.blocks[0], "Bad timestamp", { date: "not-a-date" });
assert.ok(invalidCommentDocument.verify().issues.some((issue) => issue.type === "invalidCommentDate"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCommentDocument), /date must be a valid date string/);
const invalidCommentParentDocument = DocumentModel.create({ paragraphs: ["Invalid comment parent fixture"] });
invalidCommentParentDocument.addComment(invalidCommentParentDocument.blocks[0], "Missing parent", { id: "comment/child", parentId: "comment/missing" });
assert.ok(invalidCommentParentDocument.verify().issues.some((issue) => issue.type === "missingCommentParent"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCommentParentDocument), /missing parent comment/);
const cyclicCommentDocument = DocumentModel.create({ paragraphs: ["Cyclic comment fixture"] });
cyclicCommentDocument.addComment(cyclicCommentDocument.blocks[0], "First", { id: "comment/first", parentId: "comment/second" });
cyclicCommentDocument.addComment(cyclicCommentDocument.blocks[0], "Second", { id: "comment/second", parentId: "comment/first" });
await assert.rejects(() => DocumentFile.exportDocx(cyclicCommentDocument), /cyclic parent chain/);
const invalidCommentParaIdDocument = DocumentModel.create({ paragraphs: ["Invalid paraId fixture"] });
invalidCommentParaIdDocument.addComment(invalidCommentParaIdDocument.blocks[0], "Bad paraId", { paraId: "123" });
assert.ok(invalidCommentParaIdDocument.verify().issues.some((issue) => issue.type === "invalidCommentParaId"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCommentParaIdDocument), /exactly eight hexadecimal digits/);
const invalidCommentDurableIdDocument = DocumentModel.create({ paragraphs: ["Invalid durableId fixture"] });
invalidCommentDurableIdDocument.addComment(invalidCommentDurableIdDocument.blocks[0], "Bad durableId", { durableId: "7FFFFFFF" });
assert.ok(invalidCommentDurableIdDocument.verify().issues.some((issue) => issue.type === "invalidCommentDurableId"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCommentDurableIdDocument), /greater than 0 and less than 7FFFFFFF/);
const invalidCommentDateUtcDocument = DocumentModel.create({ paragraphs: ["Invalid dateUtc fixture"] });
invalidCommentDateUtcDocument.addComment(invalidCommentDateUtcDocument.blocks[0], "Bad dateUtc", { dateUtc: "not-a-date" });
assert.ok(invalidCommentDateUtcDocument.verify().issues.some((issue) => issue.type === "invalidCommentDateUtc"));
await assert.rejects(() => DocumentFile.exportDocx(invalidCommentDateUtcDocument), /dateUtc must be a valid date string/);
const invalidPlaceholderReplyDocument = DocumentModel.create({ paragraphs: ["Invalid placeholder reply fixture"] });
const placeholderRoot = invalidPlaceholderReplyDocument.addComment(invalidPlaceholderReplyDocument.blocks[0], "Root");
invalidPlaceholderReplyDocument.replyToComment(placeholderRoot, "Reply", { intelligentPlaceholder: true });
assert.ok(invalidPlaceholderReplyDocument.verify().issues.some((issue) => issue.type === "invalidCommentReplyPlaceholder"));
await assert.rejects(() => DocumentFile.exportDocx(invalidPlaceholderReplyDocument), /must not be an intelligent placeholder/);
const invalidListDocument = DocumentModel.create({ paragraphs: ["Invalid list fixture"] });
invalidListDocument.addListItem("Too deep", { level: 9, start: 0 });
assert.ok(invalidListDocument.verify().issues.some((issue) => issue.type === "invalidListLevel"));
const missingRunStyleDocument = DocumentModel.create({ blocks: [{ kind: "paragraph", text: "", runs: [{ text: "Missing character style", style: { runStyleId: "MissingCharacterStyle" } }] }] });
assert.ok(missingRunStyleDocument.verify().issues.some((issue) => issue.type === "unknownRunStyle" && issue.runStyleId === "MissingCharacterStyle"));
assert.throws(() => DocumentModel.create({ blocks: [{ kind: "paragraph", text: "", runs: [{ text: "Invalid character style", style: { runStyleId: "" } }] }] }), /runStyleId must be a non-empty string/);
assert.ok(invalidListDocument.verify().issues.some((issue) => issue.type === "invalidListStart"));
await assert.rejects(() => DocumentFile.exportDocx(invalidListDocument), /level must be an integer from 0 through 8/);

const preview = await document.render();
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /Research memo/);
assert.match(svg, /Run styled /);
assert.match(svg, /#0ea5e9/);
assert.match(svg, /#f97316/);
assert.match(svg, /#99b3cc/);
assert.match(svg, /#992600/);
assert.match(svg, /Source Serif 4/);
assert.match(svg, /Noto Serif CJK SC/);
assert.match(svg, /Noto Naskh Arabic/);
assert.match(svg, /Risk callout inherits bold styling/);
assert.match(svg, /#b91c1c/);
assert.match(svg, /DOCX styles/);
assert.match(svg, /Opening-section header/);
assert.doesNotMatch(svg, /Confidential research memo/);
assert.match(svg, /w31r4 research note/);
assert.match(svg, /PAGE/);
assert.match(svg, /Source: Market brief/);
assert.match(svg, /Memo logo/);
assert.match(svg, /section break: nextPage landscape/);
assert.match(svg, /Inserted reviewer clarification/);
assert.match(svg, /Remove stale claim/);
assert.match(svg, /tracked insert by Reviewer/);
assert.match(svg, /Render and verify/);
const layoutBlob = await document.render({ format: "layout" });
assert.equal(layoutBlob.type, "application/vnd.open-office-artifact.layout+json");
const layout = JSON.parse(await layoutBlob.text());
assert.equal(layout.document.designPreset, "report");
assert.ok(layout.pages.length >= 1);
assert.equal(layout.elements.find((element) => element.id === riskCallout.id).effectiveStyle.bold, true);
assert.equal(layout.elements.find((element) => element.id === riskCallout.id).effectiveStyle.italic, true);
const themedRunLayout = layout.elements.find((element) => element.id === themedRunParagraph.id);
assert.equal(themedRunLayout.runs[0].style.effectiveColor, "#99b3cc");
assert.equal(themedRunLayout.runs[1].style.effectiveFontFamily, "Noto Serif CJK SC");
assert.equal(themedRunLayout.runs[2].style.effectiveBold, true);
assert.equal(themedRunLayout.runs[2].style.effectiveItalic, true);
assert.equal(themedRunLayout.runs[2].style.effectiveFontSize, 32);
const cascadeRunLayout = layout.elements.find((element) => element.id === cascadeParagraph.id).runs;
assert.equal(cascadeRunLayout[0].style.runStyleId, "CascadeCharacter");
assert.equal(cascadeRunLayout[0].style.effectiveColor, "#123456");
assert.equal(cascadeRunLayout[0].style.effectiveItalic, false);
assert.equal(cascadeRunLayout[0].style.effectiveBold, false);
assert.equal(cascadeRunLayout[0].style.effectiveFontSize, 26);
assert.equal(cascadeRunLayout[0].style.effectiveFontFamily, "Source Serif 4");
assert.equal(cascadeRunLayout[1].style.effectiveColor, "#cc3300");
assert.equal(cascadeRunLayout[1].style.effectiveItalic, true);
assert.ok(layout.elements.some((element) => element.id === table.id));
const targetedLayoutBlob = await document.render({ format: "layout", target: table.id });
assert.equal(targetedLayoutBlob.metadata.artifactKind, "document");
assert.equal(targetedLayoutBlob.metadata.target, table.id);
const targetedLayout = JSON.parse(await targetedLayoutBlob.text());
assert.deepEqual(targetedLayout.elements.map((element) => element.id), [table.id]);
assert.deepEqual(targetedLayout.pages.map((page) => page.page), [targetedLayout.elements[0].page]);
assert.equal(targetedLayout.slice.matchedElements, 1);
const contextualLayout = document.layoutJson({ target: table.id, before: 1 });
assert.deepEqual(contextualLayout.elements.map((element) => element.id), [numbered.id, table.id]);
assert.equal(contextualLayout.slice.returnedElements, 2);
const textSlicedLayout = document.layoutJson({ search: "anchored" });
assert.deepEqual(textSlicedLayout.elements.map((element) => element.id), [table.id]);
const commentSlicedLayout = document.layoutJson({ target: comment.id });
assert.deepEqual(commentSlicedLayout.elements.map((element) => element.id), [heading.id]);
const pageSlicedLayout = document.layoutJson({ target: `${document.id}/page/${targetedLayout.elements[0].page}` });
assert.ok(pageSlicedLayout.elements.length > 0);
assert.ok(pageSlicedLayout.elements.every((element) => element.page === targetedLayout.elements[0].page));
const textRangeSlicedLayout = document.layoutJson({ target: `${heading.id}/text` });
assert.deepEqual(textRangeSlicedLayout.elements.map((element) => element.id), [heading.id]);
const docxSourceBlob = await document.render({ format: "docx", source: "docx" });
assert.equal(docxSourceBlob.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const docxSourceZip = await JSZip.loadAsync(new Uint8Array(await docxSourceBlob.arrayBuffer()));
assert.ok(docxSourceZip.file("word/document.xml"));
let docxRendererSawInput = false;
const renderedDocxPdf = await renderArtifact(document, {
  format: "pdf",
  source: "docx",
  renderer: async ({ input, inputType, outputType, artifactKind }) => {
    docxRendererSawInput = true;
    assert.equal(artifactKind, "document");
    assert.equal(inputType, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
    assert.equal(outputType, "application/pdf");
    const zip = await JSZip.loadAsync(new Uint8Array(await input.arrayBuffer()));
    assert.ok(zip.file("word/document.xml"));
    return new FileBlob("%PDF-docx-render", { type: outputType, metadata: { renderer: "mock-docx" } });
  },
});
assert.equal(docxRendererSawInput, true);
assert.equal(renderedDocxPdf.type, "application/pdf");
assert.equal(renderedDocxPdf.metadata.renderSource, "docx");
assert.match(await renderedDocxPdf.text(), /%PDF-docx-render/);
assert.match(document.help("document.render").ndjson, /source: 'docx'/);
assert.match(document.help("DocumentFile.inspectDocx").ndjson, /DOCX package/);
assert.match(document.help("DocumentFile.patchDocx").ndjson, /path traversal/);
assert.match(document.help("DocumentFile.patchDocx").ndjson, /comment anchors/);
assert.match(document.help("DocumentFile.patchDocx").ndjson, /numbering assignments/);
assert.match(document.help("DocumentFile.patchDocx").ndjson, /settings mutations/);
assert.match(document.help("document.setSettings").ndjson, /passwordless/);

const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const docxBytes = new Uint8Array(await docx.arrayBuffer());
const zip = await JSZip.loadAsync(docxBytes);
const documentXml = await zip.file("word/document.xml").async("text");
const commentsXml = await zip.file("word/comments.xml").async("text");
const commentsExtendedXml = await zip.file("word/commentsExtended.xml").async("text");
const commentsIdsXml = await zip.file("word/commentsIds.xml").async("text");
const commentsExtensibleXml = await zip.file("word/commentsExtensible.xml").async("text");
const peopleXml = await zip.file("word/people.xml").async("text");
const numberingXml = await zip.file("word/numbering.xml").async("text");
const numberingRelsXml = await zip.file("word/_rels/numbering.xml.rels").async("text");
const stylesXml = await zip.file("word/styles.xml").async("text");
const themeXml = await zip.file("word/theme/theme1.xml").async("text");
const bibliographyXml = await zip.file("customXml/item1.xml").async("text");
assert.match(bibliographyXml, /<b:Sources xmlns:b="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/bibliography">/);
assert.match(bibliographyXml, /<b:Tag>MarketBrief26<\/b:Tag><b:SourceType>Report<\/b:SourceType>/);
assert.match(bibliographyXml, /<b:Author><b:Author><b:NameList><b:Person><b:Last>Analyst<\/b:Last><b:First>Ada<\/b:First><\/b:Person><\/b:NameList><\/b:Author><\/b:Author>/);
assert.match(bibliographyXml, /<b:Title>Market brief<\/b:Title>/);
assert.match(bibliographyXml, /<b:Publisher>Example Research<\/b:Publisher>/);
assert.match(bibliographyXml, /<b:Tag>OpenXmlSdk<\/b:Tag><b:SourceType>InternetSite<\/b:SourceType>/);
assert.match(bibliographyXml, /<b:Corporate>Microsoft<\/b:Corporate>/);
assert.match(themeXml, /name="Research Theme"/);
assert.match(themeXml, /<a:accent1><a:srgbClr val="336699"\/><\/a:accent1>/);
assert.match(themeXml, /<a:majorFont><a:latin typeface="Source Serif 4"\/><a:ea typeface="Noto Serif CJK SC"\/><a:cs typeface="Noto Naskh Arabic"\/>/);
assert.match(stylesXml, /w:styleId="RiskCallout"/);
assert.match(stylesXml, /<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Default Serif"/);
assert.match(stylesXml, /w:styleId="CascadeCharacter"/);
assert.match(stylesXml, /<w:basedOn w:val="CascadeCharacterBase"\/>/);
assert.match(stylesXml, /<w:basedOn w:val="Callout"\/>/);
assert.match(stylesXml, /<w:i\/>/);
assert.match(stylesXml, /<w:color w:val="b91c1c"\/>/);
assert.match(documentXml, /<w:t>Run styled <\/w:t>/);
assert.match(documentXml, /<w:color w:val="0ea5e9"\/>/);
assert.match(documentXml, /<w:t>paragraph<\/w:t>/);
assert.match(documentXml, /<w:rStyle w:val="CascadeCharacter"\/>/);
assert.match(documentXml, /<w:color w:val="f97316"\/>/);
assert.match(documentXml, /<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi"\/>/);
assert.match(documentXml, /<w:b\/><w:bCs w:val="0"\/>/);
assert.match(documentXml, /<w:color w:val="99b3cc" w:themeColor="accent1" w:themeTint="80"\/>/);
assert.match(documentXml, /<w:sz w:val="28"\/><w:szCs w:val="34"\/>/);
assert.match(documentXml, /w:eastAsiaTheme="majorEastAsia"/);
assert.match(documentXml, /<w:color w:val="992600" w:themeColor="accent2" w:themeShade="BF"\/>/);
assert.match(documentXml, /w:cstheme="majorBidi"/);
assert.match(documentXml, /<w:b w:val="0"\/><w:bCs\/><w:i w:val="0"\/><w:iCs\/>/);
assert.match(documentXml, /<w:bookmarkStart w:id="42" w:name="FindingsSection"\/>/);
assert.match(documentXml, /<w:bookmarkEnd w:id="42"\/>/);
assert.match(documentXml, /<w:bookmarkStart w:id="43" w:name="EvidenceCells"\/><w:r><w:t>DOCX styles<\/w:t>/);
assert.match(documentXml, /<w:t>anchored<\/w:t><\/w:r><w:bookmarkEnd w:id="43"\/>/);
assert.match(documentXml, /<w:hyperlink w:anchor="FindingsSection" w:history="0" w:tooltip="Open the findings section">/);
assert.match(documentXml, /<w:fldSimple w:instr="CITATION MarketBrief26"><w:r><w:t>Source: Market brief \(Market brief\)<\/w:t><\/w:r><\/w:fldSimple>/);
assert.match(documentXml, /<a:blip r:embed="rIdImage1"\/>/);
assert.match(documentXml, /<wp:docPr[^>]*name="memo-logo"[^>]*descr="Memo logo"/);
assert.match(documentXml, /<w:type w:val="nextPage"\/>/);
assert.match(documentXml, /<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/);
assert.match(documentXml, /<w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900"/);
assert.match(documentXml, /<w:ins\b[^>]*w:author="Reviewer"/);
assert.match(documentXml, /<w:t>Inserted reviewer clarification\.<\/w:t>/);
assert.match(documentXml, /<w:del\b[^>]*w:author="Reviewer"/);
assert.match(documentXml, /<w:delText>Remove stale claim\.<\/w:delText>/);
assert.match(documentXml, /<w:tblInd w:w="120" w:type="dxa"\/>/);
assert.match(documentXml, /<w:tblLayout w:type="fixed"\/>/);
assert.match(documentXml, /<w:tblBorders>/);
assert.match(documentXml, /<w:gridCol w:w="3600"\/><w:gridCol w:w="5760"\/>/);
assert.match(documentXml, /<w:tcW w:w="3600" w:type="dxa"\/>/);
assert.match(documentXml, /<w:tcW w:w="5760" w:type="dxa"\/>/);
assert.match(documentXml, /<w:shd w:val="clear" w:color="auto" w:fill="F2F4F7"\/>/);
assert.match(documentXml, /<w:rPr><w:b\/><\/w:rPr><w:t>Area<\/w:t>/);
assert.match(commentsXml, /w:id="0" w:author="Reviewer" w:initials="RV" w:date="2026-07-11T00:10:00\.000Z"/);
assert.match(commentsXml, /w:id="1" w:author="R&amp;D Analyst" w:initials="RA" w:date="2026-07-11T00:15:00\.000Z"/);
assert.match(commentsXml, /w:id="2" w:author="Link Reviewer" w:initials="LR" w:date="2026-07-11T00:18:00\.000Z"/);
assert.match(commentsXml, /w:id="0"[\s\S]*?<w:p w14:paraId="00000001">/);
assert.match(commentsXml, /w:id="3"[\s\S]*?<w:p w14:paraId="00000004">[\s\S]*?Heading language is now approved/);
assert.match(commentsExtendedXml, /<w15:commentEx w15:paraId="00000001" w15:done="1"\/>/);
assert.match(commentsExtendedXml, /<w15:commentEx w15:paraId="00000004" w15:paraIdParent="00000001" w15:done="0"\/>/);
assert.match(commentsIdsXml, /<w16cid:commentId w16cid:paraId="00000001" w16cid:durableId="0000A001"\/>/);
assert.match(commentsIdsXml, /<w16cid:commentId w16cid:paraId="00000004" w16cid:durableId="0000A004"\/>/);
assert.match(commentsExtensibleXml, /<w16cex:commentExtensible w16cex:durableId="0000A001" w16cex:dateUtc="2026-07-11T00:10:00\.000Z"\/>/);
assert.match(peopleXml, /<w15:person w15:author="Reviewer"><w15:presenceInfo w15:providerId="None" w15:userId="reviewer@example\.test"\/><\/w15:person>/);
assert.match(peopleXml, /<w15:person w15:author="Editor"><w15:presenceInfo w15:providerId="None" w15:userId="editor@example\.test"\/><\/w15:person>/);
assert.doesNotMatch(documentXml, /<w:commentRangeStart w:id="3"\/>/);
const exportedTableXml = /<w:tbl[\s\S]*?<\/w:tbl>/.exec(documentXml)?.[0] || "";
assert.match(exportedTableXml, /<w:commentRangeStart w:id="1"\/>/);
assert.match(exportedTableXml, /<w:commentRangeEnd w:id="1"\/>/);
assert.match(exportedTableXml, /<w:commentReference w:id="1"\/>/);
const exportedHyperlinkParagraph = [...documentXml.matchAll(/<w:p>[\s\S]*?<\/w:p>/g)].map((match) => match[0]).find((paragraph) => paragraph.includes("w31r4 research note")) || "";
assert.match(exportedHyperlinkParagraph, /<w:commentRangeStart w:id="2"\/>/);
assert.match(exportedHyperlinkParagraph, /<w:commentRangeEnd w:id="2"\/>/);
assert.match(exportedHyperlinkParagraph, /<w:commentReference w:id="2"\/>/);
assert.match(numberingXml, /<w:abstractNum w:abstractNumId="3">/);
assert.match(numberingXml, /<w:lvl w:ilvl="0"><w:start w:val="2"\/><w:numFmt w:val="upperLetter"\/><w:lvlText w:val="%1\)"/);
assert.match(numberingXml, /<w:lvl w:ilvl="1"><w:start w:val="3"\/><w:numFmt w:val="lowerRoman"\/><w:lvlText w:val="%1\.%2\)"/);
assert.match(documentXml, new RegExp(`<w:numPr><w:ilvl w:val="1"/><w:numId w:val="3"/></w:numPr>[\\s\\S]*?Nested roman evidence`));
assert.match(numberingXml, /<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shapetype/);
assert.match(numberingXml, /<v:shape id="_x0000_i1025"[^>]*style="width:10pt;height:10pt" o:bullet="t">/);
assert.match(numberingXml, /<v:imagedata r:id="rIdPictureBullet1" o:title="Green status marker"\/>/);
assert.match(numberingXml, /<w:abstractNum w:abstractNumId="4">[\s\S]*?<w:lvlPicBulletId w:val="0"\/>/);
assert.match(documentXml, new RegExp(`<w:numPr><w:ilvl w:val="0"/><w:numId w:val="4"/></w:numPr>[\\s\\S]*?Picture bullet evidence`));
assert.match(numberingRelsXml, /Id="rIdPictureBullet1"[^>]*relationships\/image[^>]*Target="media\/image2\.png"/);
assert.ok(zip.file("word/media/image2.png"));

const missingPictureBulletRelationshipZip = await JSZip.loadAsync(docxBytes);
missingPictureBulletRelationshipZip.file("word/_rels/numbering.xml.rels", numberingRelsXml.replace(/<Relationship\b[^>]*Id="rIdPictureBullet1"[^>]*\/>/, ""));
await assert.rejects(
  async () => DocumentFile.importDocx(new FileBlob(await missingPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type }), { preferNative: true }),
  /picture bullet 0 references missing image relationship rIdPictureBullet1/,
);
const wrongPictureBulletRelationshipZip = await JSZip.loadAsync(docxBytes);
wrongPictureBulletRelationshipZip.file("word/_rels/numbering.xml.rels", numberingRelsXml.replace(/relationships\/image/, "relationships/hyperlink"));
await assert.rejects(
  async () => DocumentFile.importDocx(new FileBlob(await wrongPictureBulletRelationshipZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type }), { preferNative: true }),
  /picture bullet 0 references missing image relationship rIdPictureBullet1/,
);
const missingPictureBulletPartZip = await JSZip.loadAsync(docxBytes);
missingPictureBulletPartZip.remove("word/media/image2.png");
await assert.rejects(
  async () => DocumentFile.importDocx(new FileBlob(await missingPictureBulletPartZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type }), { preferNative: true }),
  /picture bullet 0 relationship rIdPictureBullet1 targets missing part word\/media\/image2\.png/,
);
const alternatePictureBulletPrefixZip = await JSZip.loadAsync(docxBytes);
const alternatePictureBulletPrefixXml = numberingXml
  .replace(/\bwp:/g, "wordDrawing:")
  .replace(/\bpic:/g, "picture:")
  .replace(/\bw:/g, "word:")
  .replace(/\br:/g, "relationships:")
  .replace(/\ba:/g, "drawing:")
  .replace(/\bv:/g, "vector:")
  .replace(/\bo:/g, "office:");
alternatePictureBulletPrefixZip.file("word/numbering.xml", alternatePictureBulletPrefixXml);
const alternatePictureBulletPrefixDocument = await DocumentFile.importDocx(new FileBlob(await alternatePictureBulletPrefixZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type }), { preferNative: true });
assert.equal(alternatePictureBulletPrefixDocument.blocks.find((item) => item.text === "Picture bullet evidence")?.pictureBullet?.alt, "Green status marker");

const drawingPictureBulletZip = await JSZip.loadAsync(docxBytes);
const drawingPictureBulletXml = numberingXml
  .replace('xmlns:o="urn:schemas-microsoft-com:office:office"', 'xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')
  .replace(/<w:numPicBullet\b[^>]*>[\s\S]*?<\/w:numPicBullet>/, '<w:numPicBullet w:numPicBulletId="0"><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="127000" cy="127000"/><wp:docPr id="1" name="Drawing picture bullet" descr="DrawingML status marker"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="1" name="Drawing picture bullet"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="rIdPictureBullet1"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="127000" cy="127000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:numPicBullet>');
drawingPictureBulletZip.file("word/numbering.xml", drawingPictureBulletXml);
const drawingPictureBulletDocument = await DocumentFile.importDocx(new FileBlob(await drawingPictureBulletZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type }), { preferNative: true });
assert.equal(drawingPictureBulletDocument.blocks.find((item) => item.text === "Picture bullet evidence")?.pictureBullet?.alt, "DrawingML status marker");

const externalPictureBulletDocument = DocumentModel.create({ paragraphs: ["External picture bullet"] });
externalPictureBulletDocument.addListItem("Non-fetched marker", { pictureBullet: { uri: "https://example.com/status.png", widthPt: 9, alt: "External status" } });
const externalPictureBulletDocx = await DocumentFile.exportDocx(externalPictureBulletDocument);
const externalPictureBulletZip = await JSZip.loadAsync(new Uint8Array(await externalPictureBulletDocx.arrayBuffer()));
assert.match(await externalPictureBulletZip.file("word/numbering.xml").async("text"), /<v:imagedata r:id="rIdPictureBullet1" o:title="External status"\/>/);
assert.match(await externalPictureBulletZip.file("word\/_rels\/numbering.xml.rels").async("text"), /Target="https:\/\/example\.com\/status\.png" TargetMode="External"/);
assert.equal((await DocumentFile.importDocx(externalPictureBulletDocx, { preferNative: true })).blocks.find((item) => item.text === "Non-fetched marker")?.pictureBullet?.uri, "https://example.com/status.png");
const multiplePictureBulletDocument = DocumentModel.create({ paragraphs: ["Multiple picture bullets"] });
multiplePictureBulletDocument.addListItem("Green marker", { pictureBullet: pictureBulletPng });
multiplePictureBulletDocument.addListItem("Dot marker", { pictureBullet: logo.dataUrl });
const multiplePictureBulletZip = await JSZip.loadAsync(new Uint8Array(await (await DocumentFile.exportDocx(multiplePictureBulletDocument)).arrayBuffer()));
const multiplePictureBulletXml = await multiplePictureBulletZip.file("word/numbering.xml").async("text");
assert.match(multiplePictureBulletXml, /<v:shapetype id="_x0000_t75"/);
assert.match(multiplePictureBulletXml, /<v:shapetype id="_x0000_t76"/);
assert.match(multiplePictureBulletXml, /<v:shape id="_x0000_i1026" type="#_x0000_t76"/);
const resizedPictureBulletDocument = DocumentModel.create({ paragraphs: ["Resized picture bullets"] });
resizedPictureBulletDocument.addListItem("Small marker", { pictureBullet: { dataUrl: pictureBulletPng, widthPt: 8, heightPt: 9, alt: "Small status" } });
resizedPictureBulletDocument.addListItem("Large marker", { pictureBullet: { dataUrl: pictureBulletPng, widthPt: 16, heightPt: 17, alt: "Large status" } });
const resizedPictureBulletDocx = await DocumentFile.exportDocx(resizedPictureBulletDocument);
const resizedPictureBulletZip = await JSZip.loadAsync(new Uint8Array(await resizedPictureBulletDocx.arrayBuffer()));
const resizedPictureBulletXml = await resizedPictureBulletZip.file("word/numbering.xml").async("text");
const resizedPictureBulletRels = await resizedPictureBulletZip.file("word/_rels/numbering.xml.rels").async("text");
assert.equal((resizedPictureBulletXml.match(/<w:numPicBullet\b/g) || []).length, 2);
assert.match(resizedPictureBulletXml, /style="width:8pt;height:9pt"[^>]*>[\s\S]*?o:title="Small status"/);
assert.match(resizedPictureBulletXml, /style="width:16pt;height:17pt"[^>]*>[\s\S]*?o:title="Large status"/);
assert.equal((resizedPictureBulletRels.match(/relationships\/image/g) || []).length, 1);
const resizedPictureBulletRoundTrip = await DocumentFile.importDocx(resizedPictureBulletDocx, { preferNative: true });
assert.equal(resizedPictureBulletRoundTrip.blocks.find((item) => item.text === "Small marker")?.pictureBullet?.widthPt, 8);
assert.equal(resizedPictureBulletRoundTrip.blocks.find((item) => item.text === "Large marker")?.pictureBullet?.alt, "Large status");
const conflictingPictureBulletDocument = DocumentModel.create({ paragraphs: ["Conflicting picture bullets"] });
conflictingPictureBulletDocument.addListItem("Embedded", { numberingId: 99, pictureBullet: pictureBulletPng });
conflictingPictureBulletDocument.addListItem("External", { numberingId: 99, pictureBullet: "https://example.com/status.png" });
await assert.rejects(() => DocumentFile.exportDocx(conflictingPictureBulletDocument), /numbering 99 level 0 has conflicting definitions/);
const conflictingPictureBulletPresentationDocument = DocumentModel.create({ paragraphs: ["Conflicting picture bullet presentation"] });
conflictingPictureBulletPresentationDocument.addListItem("Small", { numberingId: 100, pictureBullet: { dataUrl: pictureBulletPng, widthPt: 8 } });
conflictingPictureBulletPresentationDocument.addListItem("Large", { numberingId: 100, pictureBullet: { dataUrl: pictureBulletPng, widthPt: 16 } });
await assert.rejects(() => DocumentFile.exportDocx(conflictingPictureBulletPresentationDocument), /numbering 100 level 0 has conflicting definitions/);

const documentRelsXml = await zip.file("word/_rels/document.xml.rels").async("text");
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/theme" Target="theme\/theme1\.xml"/);
assert.match(documentRelsXml, /Id="rIdImage1"/);
assert.match(documentRelsXml, /Target="media\/image1\.png"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/header" Target="header2\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/footer" Target="footer2\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/header" Target="header3\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/footer" Target="footer3\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/settings" Target="settings\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/commentsExtended" Target="commentsExtended\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.microsoft\.com\/office\/2016\/09\/relationships\/commentsIds" Target="commentsIds\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.microsoft\.com\/office\/2018\/08\/relationships\/commentsExtensible" Target="commentsExtensible\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/people" Target="people\.xml"/);
assert.match(documentRelsXml, /Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/customXml" Target="\.\.\/customXml\/item1\.xml"/);
assert.equal((documentRelsXml.match(/relationships\/hyperlink/g) || []).length, 1);
assert.doesNotMatch(documentRelsXml, /FindingsSection/);
assert.match(documentXml, /<w:headerReference w:type="default" r:id="[^"]+"\/>/);
assert.match(documentXml, /<w:headerReference w:type="first" r:id="[^"]+"\/>/);
assert.match(documentXml, /<w:footerReference w:type="even" r:id="[^"]+"\/>/);
assert.match(documentXml, /<w:titlePg\/>/);
assert.match(await zip.file("word/header2.xml").async("text"), /First-page research memo/);
assert.match(await zip.file("word/footer2.xml").async("text"), /Even-page footer/);
assert.match(await zip.file("word/header3.xml").async("text"), /Opening-section header/);
assert.match(await zip.file("word/footer3.xml").async("text"), /Opening-section footer/);
const relationshipIdForTarget = (rels, target) => [...rels.matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map((match) => match[0]).find((tag) => new RegExp(`\\bTarget=["']${target.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}["']`).test(tag))?.match(/\bId=["']([^"']+)["']/)?.[1];
const finalHeaderRelId = relationshipIdForTarget(documentRelsXml, "header1.xml");
const finalFirstHeaderRelId = relationshipIdForTarget(documentRelsXml, "header2.xml");
const openingHeaderRelId = relationshipIdForTarget(documentRelsXml, "header3.xml");
const openingFooterRelId = relationshipIdForTarget(documentRelsXml, "footer3.xml");
assert.ok(finalHeaderRelId && finalFirstHeaderRelId && openingHeaderRelId && openingFooterRelId);
const exportedSections = [...documentXml.matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)].map((match) => match[0]);
assert.equal(exportedSections.length, 2);
assert.match(exportedSections[0], new RegExp(`r:id="${openingHeaderRelId}"`));
assert.match(exportedSections[0], new RegExp(`r:id="${openingFooterRelId}"`));
assert.doesNotMatch(exportedSections[0], new RegExp(`r:id="${finalHeaderRelId}"`));
assert.match(exportedSections[1], new RegExp(`r:id="${finalHeaderRelId}"`));
assert.match(exportedSections[1], new RegExp(`r:id="${finalFirstHeaderRelId}"`));
assert.doesNotMatch(exportedSections[1], new RegExp(`r:id="${openingHeaderRelId}"`));
assert.doesNotMatch(exportedSections[0], /<w:titlePg\/>/);
assert.match(exportedSections[1], /<w:titlePg\/>/);
assert.match(await zip.file("word/settings.xml").async("text"), /<w:evenAndOddHeaders\/>/);
assert.match(await zip.file("word/settings.xml").async("text"), /<w:mirrorMargins\/>[\s\S]*?<w:trackRevisions\/>[\s\S]*?<w:documentProtection w:edit="trackedChanges" w:enforcement="1" w:formatting="0"\/>[\s\S]*?<w:evenAndOddHeaders\/>[\s\S]*?<w:updateFields\/>/);
const docxMediaBytes = await zip.file("word/media/image1.png").async("uint8array");
assert.ok(docxMediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
assert.match(contentTypesXml, /PartName="\/word\/settings\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.settings\+xml"/);
assert.match(contentTypesXml, /PartName="\/word\/commentsExtended\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.commentsExtended\+xml"/);
assert.match(contentTypesXml, /PartName="\/word\/commentsIds\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.commentsIds\+xml"/);
assert.match(contentTypesXml, /PartName="\/word\/commentsExtensible\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.commentsExtensible\+xml"/);
assert.match(contentTypesXml, /PartName="\/word\/people\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.people\+xml"/);
assert.match(contentTypesXml, /PartName="\/word\/theme\/theme1\.xml" ContentType="application\/vnd\.openxmlformats-officedocument\.theme\+xml"/);
const packageInspect = await DocumentFile.inspectDocx(docx, { includeText: true, maxChars: 12000 });
assert.equal(packageInspect.ok, true);
assert.ok(packageInspect.parts.some((part) => part.path === "word/document.xml"));
assert.match(packageInspect.ndjson, /docxPart/);
assert.match(packageInspect.ndjson, /word\/styles\.xml/);
assert.ok(packageInspect.records[0].uncompressedBytes > 0);
assert.ok(packageInspect.records[0].relationshipReferences > 0);
assert.equal(packageInspect.records[0].relationshipReferenceIssues, 0);
assert.ok(packageInspect.parts.some((part) => part.path === "word/document.xml" && part.contentType.includes("wordprocessingml.document.main+xml")));
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: documentXml.replace('w:anchor="FindingsSection"', 'w:anchor="MissingSection"') }]),
  /docxHyperlinkAnchorNotFound/,
);
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: documentXml.replace('<w:bookmarkEnd w:id="42"/>', "") }]),
  /docxBookmarkEndMissing/,
);
const reversedBookmarkXml = documentXml
  .replace('<w:bookmarkStart w:id="42" w:name="FindingsSection"/>', "__BOOKMARK_START__")
  .replace('<w:bookmarkEnd w:id="42"/>', '<w:bookmarkStart w:id="42" w:name="FindingsSection"/>')
  .replace("__BOOKMARK_START__", '<w:bookmarkEnd w:id="42"/>');
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: reversedBookmarkXml }]),
  /docxBookmarkRangeReversed/,
);
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "customXml/item1.xml", remove: true }]),
  /docxCitationSourceMissing/,
);
const orphanBibliographyXml = '<?xml version="1.0"?><b:Sources xmlns:b="http://schemas.openxmlformats.org/officeDocument/2006/bibliography"><b:Source><b:Tag>Orphan</b:Tag><b:SourceType>Book</b:SourceType><b:Title>Orphan source</b:Title></b:Source></b:Sources>';
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "customXml/orphan-sources.xml", xml: orphanBibliographyXml }]),
  /docxBibliographyMultipleParts|docxBibliographyRelationshipMissing/,
);
const duplicateBookmarkNameXml = documentXml.replace("</w:body>", '<w:p><w:bookmarkStart w:id="99" w:name="FindingsSection"/><w:r><w:t>Duplicate</w:t></w:r><w:bookmarkEnd w:id="99"/></w:p></w:body>');
await assert.rejects(
  () => DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: duplicateBookmarkNameXml }]),
  /docxBookmarkNameDuplicate/,
);
const complexFieldDocumentXml = documentXml.replace(/<w:fldSimple\b[^>]*>[\s\S]*?<\/w:fldSimple>/, '<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> NUMPAGES \\* MERGEFORMAT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>3</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>');
const entityHyperlinkRelsXml = documentRelsXml.replace('Target="https://example.com/research"', 'Target="https://example.com/research?a=1&amp;b=2"');
const complexNativeDocx = await DocumentFile.patchDocx(docx, [
  { path: "word/document.xml", xml: complexFieldDocumentXml },
  { path: "word/_rels/document.xml.rels", xml: entityHyperlinkRelsXml },
]);
const complexNativeDocument = await DocumentFile.importDocx(complexNativeDocx, { preferNative: true });
assert.equal(complexNativeDocument.blocks.find((item) => item.kind === "field")?.instruction, "NUMPAGES \\* MERGEFORMAT");
assert.equal(complexNativeDocument.blocks.find((item) => item.kind === "field")?.display, "3");
assert.equal(complexNativeDocument.blocks.find((item) => item.kind === "hyperlink" && item.url)?.url, "https://example.com/research?a=1&b=2");
const overriddenNumberingXml = numberingXml.replace(/(<w:num w:numId="3"><w:abstractNumId w:val="3"\/>)(<\/w:num>)/, '$1<w:lvlOverride w:ilvl="1"><w:startOverride w:val="7"/></w:lvlOverride>$2');
const relocatedConfigDocx = await DocumentFile.patchDocx(docx, [
  { path: "word/styles.xml", remove: true },
  { path: "word/numbering.xml", remove: true },
  { path: "word/config/styles-custom.xml", xml: stylesXml, recipe: { kind: "styles", source: "word/document.xml", id: "rIdCustomStyles" } },
  { path: "word/config/numbering-custom.xml", xml: overriddenNumberingXml, recipe: { kind: "numbering", source: "word/document.xml", id: "rIdCustomNumbering" } },
  { path: "word/config/_rels/numbering-custom.xml.rels", xml: numberingRelsXml.replace('Target="media/image2.png"', 'Target="../media/image2.png"') },
]);
const relocatedConfigInspect = await DocumentFile.inspectDocx(relocatedConfigDocx);
assert.equal(relocatedConfigInspect.ok, true);
assert.ok(relocatedConfigInspect.parts.some((part) => part.path === "word/config/styles-custom.xml" && part.contentType.endsWith("wordprocessingml.styles+xml")));
assert.ok(relocatedConfigInspect.parts.some((part) => part.path === "word/config/numbering-custom.xml" && part.contentType.endsWith("wordprocessingml.numbering+xml")));
const relocatedConfigNative = await DocumentFile.importDocx(relocatedConfigDocx, { preferNative: true });
assert.equal(relocatedConfigNative.styles.effective("RiskCallout").bold, true);
assert.equal(relocatedConfigNative.blocks.find((item) => item.text === "Nested roman evidence")?.numberFormat, "lowerRoman");
assert.equal(relocatedConfigNative.blocks.find((item) => item.text === "Nested roman evidence")?.start, 7);
assert.equal(relocatedConfigNative.blocks.find((item) => item.text === "Nested roman evidence")?.levelText, "%1.%2)");
const brokenDocxReferenceXml = documentXml.replace(/<\/w:body>\s*<\/w:document>\s*$/, '<w:p><w:hyperlink xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="rIdMissingSourceReference"><w:r><w:t>Broken link</w:t></w:r></w:hyperlink></w:p></w:body></w:document>');
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/document.xml", xml: brokenDocxReferenceXml }]), /invalid OOXML package.*relationshipReferenceIdNotFound/);
const missingReferencePartXml = '<source xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rIdMissingRelationshipPart"/>';
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "customXml/source-reference.xml", xml: missingReferencePartXml }]), /invalid OOXML package.*relationshipReferencePartNotFound/);
const invalidReferencePartDocx = await DocumentFile.patchDocx(docx, [{ path: "customXml/source-reference.xml", xml: missingReferencePartXml }], { validateResult: false });
const invalidReferencePartInspect = await DocumentFile.inspectDocx(invalidReferencePartDocx);
assert.ok(invalidReferencePartInspect.issues.some((issue) => issue.type === "relationshipReferencePartNotFound" && issue.path === "customXml/source-reference.xml"));
const patchedDocx = await DocumentFile.patchDocx(docx, [{ path: "customXml/review-note.xml", text: "<review>ok</review>" }]);
assert.equal(patchedDocx.type, docx.type);
assert.equal(patchedDocx.metadata.patchedParts, 1);
assert.equal(patchedDocx.metadata.validated, true);
assert.equal(patchedDocx.metadata.validationIssues, 0);
const patchedInspect = await DocumentFile.inspectDocx(patchedDocx, { includeText: true, maxChars: 12000 });
assert.match(patchedInspect.ndjson, /customXml\/review-note\.xml/);
assert.match(patchedInspect.ndjson, /&lt;review&gt;ok&lt;\/review&gt;|<review>ok<\/review>/);
const relocatedCommentsDocx = await DocumentFile.patchDocx(docx, [
  { path: "word/comments.xml", remove: true },
  { path: "word/review/comments-relocated.xml", xml: commentsXml, recipe: { kind: "comments", source: "word/document.xml", id: "rIdRelocatedComments" } },
]);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsExtended.xml", xml: commentsExtendedXml.replace('w15:paraIdParent="00000001"', 'w15:paraIdParent="DEADBEEF"') }]), /docxCommentExParentNotFound/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsExtended.xml", xml: commentsExtendedXml.replace('w15:paraId="00000004"', 'w15:paraId="00000001"') }]), /docxCommentExParaIdDuplicate/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsIds.xml", xml: commentsIdsXml.replace('w16cid:durableId="0000A004"', 'w16cid:durableId="0000A001"') }]), /docxCommentDurableIdDuplicate/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsIds.xml", xml: commentsIdsXml.replace('w16cid:durableId="0000A001"', 'w16cid:durableId="7FFFFFFF"') }]), /docxCommentDurableIdInvalid/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsIds.xml", xml: commentsIdsXml.replace(/<w16cid:commentId w16cid:paraId="00000001"[^>]*\/>/, "") }]), /docxCommentIdMappingMissing/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsExtensible.xml", xml: commentsExtensibleXml.replace('w16cex:durableId="0000A001"', 'w16cex:durableId="0000BEEF"') }]), /docxCommentExtensibleReferenceNotFound/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsExtensible.xml", xml: commentsExtensibleXml.replace('w16cex:dateUtc="2026-07-11T00:10:00.000Z"', 'w16cex:dateUtc="2026-07-11T08:10:00+08:00"') }]), /docxCommentExtensibleDateUtcInvalid/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/commentsExtensible.xml", xml: commentsExtensibleXml.replace('w16cex:durableId="0000A004"', 'w16cex:durableId="0000A004" w16cex:intelligentPlaceholder="1"') }]), /docxCommentExtensibleReplyPlaceholderInvalid/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/people.xml", xml: peopleXml.replace('w15:author="Editor"', 'w15:author="Reviewer"') }]), /docxPersonAuthorDuplicate/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/people.xml", xml: peopleXml.replace('w15:author="Editor"', 'w15:author="Unknown Reviewer"') }]), /docxPersonAuthorNotFound/);
const relocatedCommentsExtendedDocx = await DocumentFile.patchDocx(relocatedCommentsDocx, [
  { path: "word/commentsExtended.xml", remove: true },
  { path: "word/review/comments-extended-relocated.xml", xml: commentsExtendedXml, recipe: { kind: "commentsExtended", source: "word/document.xml", id: "rIdRelocatedCommentsExtended" } },
]);
const relocatedCommentsExtendedInspect = await DocumentFile.inspectDocx(relocatedCommentsExtendedDocx);
assert.equal(relocatedCommentsExtendedInspect.ok, true);
assert.ok(relocatedCommentsExtendedInspect.parts.some((part) => part.path === "word/review/comments-extended-relocated.xml" && part.contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"));
const relocatedCommentsExtendedZip = await JSZip.loadAsync(new Uint8Array(await relocatedCommentsExtendedDocx.arrayBuffer()));
assert.match(await relocatedCommentsExtendedZip.file("word/_rels/document.xml.rels").async("text"), /Id="rIdRelocatedCommentsExtended"[^>]*Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/commentsExtended"[^>]*Target="review\/comments-extended-relocated\.xml"/);
const relocatedCommentsExtendedNative = await DocumentFile.importDocx(relocatedCommentsExtendedDocx, { preferNative: true });
const relocatedExtendedRoot = relocatedCommentsExtendedNative.comments.find((item) => item.text === "Check this heading before final export.");
const relocatedExtendedReply = relocatedCommentsExtendedNative.comments.find((item) => item.text === "Heading language is now approved.");
assert.equal(relocatedExtendedRoot?.resolved, true);
assert.equal(relocatedExtendedReply?.parentId, relocatedExtendedRoot?.id);
const relocatedModernCommentPartsDocx = await DocumentFile.patchDocx(relocatedCommentsExtendedDocx, [
  { path: "word/commentsIds.xml", remove: true },
  { path: "word/commentsExtensible.xml", remove: true },
  { path: "word/people.xml", remove: true },
  { path: "word/review/comments-ids-relocated.xml", xml: commentsIdsXml, recipe: { kind: "commentsIds", source: "word/document.xml", id: "rIdRelocatedCommentsIds" } },
  { path: "word/review/comments-extensible-relocated.xml", xml: commentsExtensibleXml, recipe: { kind: "commentsExtensible", source: "word/document.xml", id: "rIdRelocatedCommentsExtensible" } },
  { path: "word/review/people-relocated.xml", xml: peopleXml, recipe: { kind: "people", source: "word/document.xml", id: "rIdRelocatedPeople" } },
]);
const relocatedModernInspect = await DocumentFile.inspectDocx(relocatedModernCommentPartsDocx);
assert.equal(relocatedModernInspect.ok, true, JSON.stringify(relocatedModernInspect.issues));
assert.ok(relocatedModernInspect.parts.some((part) => part.path === "word/review/comments-ids-relocated.xml" && part.contentType.endsWith("commentsIds+xml")));
assert.ok(relocatedModernInspect.parts.some((part) => part.path === "word/review/comments-extensible-relocated.xml" && part.contentType.endsWith("commentsExtensible+xml")));
assert.ok(relocatedModernInspect.parts.some((part) => part.path === "word/review/people-relocated.xml" && part.contentType.endsWith("people+xml")));
const relocatedModernNative = await DocumentFile.importDocx(relocatedModernCommentPartsDocx, { preferNative: true });
const relocatedModernRoot = relocatedModernNative.comments.find((item) => item.text === "Check this heading before final export.");
assert.equal(relocatedModernRoot?.durableId, "0000A001");
assert.equal(relocatedModernRoot?.dateUtc, "2026-07-11T00:10:00.000Z");
assert.deepEqual(relocatedModernRoot?.person, { providerId: "None", userId: "reviewer@example.test" });
const relocatedCommentsInspect = await DocumentFile.inspectDocx(relocatedCommentsDocx);
assert.equal(relocatedCommentsInspect.ok, true);
assert.ok(relocatedCommentsInspect.parts.some((part) => part.path === "word/review/comments-relocated.xml" && part.contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"));
const relocatedCommentsZip = await JSZip.loadAsync(new Uint8Array(await relocatedCommentsDocx.arrayBuffer()));
assert.match(await relocatedCommentsZip.file("word/_rels/document.xml.rels").async("text"), /Id="rIdRelocatedComments"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/comments"[^>]*Target="review\/comments-relocated\.xml"/);
const relocatedCommentsNative = await DocumentFile.importDocx(relocatedCommentsDocx, { preferNative: true });
assert.equal(relocatedCommentsNative.comments.find((item) => item.text === "Check this heading before final export.")?.author, "Reviewer");
assert.equal(relocatedCommentsNative.comments.find((item) => item.text === "Check this heading before final export.")?.initials, "RV");
assert.equal(relocatedCommentsNative.comments.find((item) => item.text === "Review the evidence table.")?.author, "R&D Analyst");
assert.equal(relocatedCommentsNative.comments.find((item) => item.text === "Review the evidence table.")?.date, "2026-07-11T00:15:00.000Z");
assert.equal(relocatedCommentsNative.resolve(relocatedCommentsNative.comments.find((item) => item.text === "Review the evidence table.")?.targetId)?.kind, "table");
assert.equal(relocatedCommentsNative.resolve(relocatedCommentsNative.comments.find((item) => item.text === "Verify the native hyperlink target.")?.targetId)?.kind, "hyperlink");
const relocatedCommentRoot = relocatedCommentsNative.comments.find((item) => item.text === "Check this heading before final export.");
const relocatedCommentReply = relocatedCommentsNative.comments.find((item) => item.text === "Heading language is now approved.");
assert.equal(relocatedCommentRoot?.resolved, true);
assert.equal(relocatedCommentRoot?.paraId, "00000001");
assert.equal(relocatedCommentReply?.parentId, relocatedCommentRoot?.id);
assert.equal(relocatedCommentReply?.targetId, relocatedCommentRoot?.targetId);
const recipeHeaderDocx = await DocumentFile.patchDocx(docx, [{
  path: "word/headerReview.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Review header</w:t></w:r></w:p></w:hdr>',
  recipe: { kind: "header", source: "word/document.xml", id: "rIdReviewHeader", sourceReference: { type: "first" } },
}]);
assert.equal(recipeHeaderDocx.metadata.recipesApplied, 1);
assert.equal(recipeHeaderDocx.metadata.sourceReferencesUpdated, 1);
const recipeHeaderInspect = await DocumentFile.inspectDocx(recipeHeaderDocx);
assert.equal(recipeHeaderInspect.ok, true);
assert.ok(recipeHeaderInspect.parts.some((part) => part.path === "word/headerReview.xml" && part.contentType === "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"));
const recipeHeaderZip = await JSZip.loadAsync(new Uint8Array(await recipeHeaderDocx.arrayBuffer()));
assert.match(await recipeHeaderZip.file("word/_rels/document.xml.rels").async("text"), /Id="rIdReviewHeader"[^>]*Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/header"[^>]*Target="headerReview\.xml"/);
assert.match(await recipeHeaderZip.file("word/document.xml").async("text"), /<w:headerReference\b[^>]*w:type="first"[^>]*r:id="rIdReviewHeader"/);
const recipeNativeDocument = await DocumentFile.importDocx(recipeHeaderDocx, { preferNative: true });
const importedReviewHeader = recipeNativeDocument.headers.find((item) => item.text === "Review header");
assert.ok(importedReviewHeader);
assert.equal(importedReviewHeader.referenceType, "first");
assert.equal(importedReviewHeader.relationshipId, "rIdReviewHeader");
assert.equal(importedReviewHeader.partPath, "word/headerReview.xml");
assert.equal(importedReviewHeader.sectionIndex, 1);
const recipeOpeningHeaderDocx = await DocumentFile.patchDocx(recipeHeaderDocx, [{
  path: "word/headerOpening.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Patched opening header</w:t></w:r></w:p></w:hdr>',
  recipe: { kind: "header", source: "word/document.xml", id: "rIdOpeningHeader", sourceReference: { type: "default", sectionIndex: 0 } },
}]);
const recipeOpeningZip = await JSZip.loadAsync(new Uint8Array(await recipeOpeningHeaderDocx.arrayBuffer()));
const recipeOpeningXml = await recipeOpeningZip.file("word/document.xml").async("text");
const recipeOpeningSections = [...recipeOpeningXml.matchAll(/<w:sectPr\b[^>]*>[\s\S]*?<\/w:sectPr>/g)].map((match) => match[0]);
assert.match(recipeOpeningSections[0], /r:id="rIdOpeningHeader"/);
assert.doesNotMatch(recipeOpeningSections[1], /r:id="rIdOpeningHeader"/);
assert.match(recipeOpeningSections[1], new RegExp(`r:id="${finalHeaderRelId}"`));
const recipeOpeningNative = await DocumentFile.importDocx(recipeOpeningHeaderDocx, { preferNative: true });
assert.equal(recipeOpeningNative.headers.find((item) => item.text === "Patched opening header")?.sectionIndex, 0);
await assert.rejects(() => DocumentFile.patchDocx(recipeHeaderDocx, [{
  path: "word/headerInvalidSection.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p/></w:hdr>',
  recipe: { kind: "header", source: "word/document.xml", id: "rIdInvalidSection", sourceReference: { sectionIndex: 2 } },
}]), /sectionIndex must be an integer from 0 through 1/);
const removedRecipeHeaderDocx = await DocumentFile.patchDocx(recipeHeaderDocx, [{ path: "word/headerReview.xml", remove: true, recipe: { kind: "header", source: "word/document.xml", id: "rIdReviewHeader", sourceReference: { type: "first" } } }]);
assert.equal(removedRecipeHeaderDocx.metadata.sourceReferencesUpdated, 1);
assert.equal((await DocumentFile.inspectDocx(removedRecipeHeaderDocx)).ok, true);
const removedRecipeHeaderZip = await JSZip.loadAsync(new Uint8Array(await removedRecipeHeaderDocx.arrayBuffer()));
assert.doesNotMatch(await removedRecipeHeaderZip.file("word/document.xml").async("text"), /r:id="rIdReviewHeader"/);
const relatedCustomDocx = await DocumentFile.patchDocx(docx, [
  { path: "customXml/source.xml", xml: "<source/>" },
  { path: "customXml/target.xml", xml: "<target/>", relationship: { source: "customXml/source.xml", id: "rIdTarget", type: "urn:open-office:test-target" } },
]);
const relatedCustomZip = await JSZip.loadAsync(new Uint8Array(await relatedCustomDocx.arrayBuffer()));
assert.ok(relatedCustomZip.file("customXml/_rels/source.xml.rels"));
const removedCustomSource = await DocumentFile.patchDocx(relatedCustomDocx, [{ path: "customXml/source.xml", remove: true }]);
const removedCustomSourceZip = await JSZip.loadAsync(new Uint8Array(await removedCustomSource.arrayBuffer()));
assert.equal(removedCustomSourceZip.file("customXml/_rels/source.xml.rels"), null);
assert.equal((await DocumentFile.inspectDocx(removedCustomSource)).ok, true);
const rootRelationshipsXml = await zip.file("_rels/.rels").async("text");
const firstRootRelationship = rootRelationshipsXml.match(/<Relationship\b[^>]*\/?\s*>/)?.[0];
assert.ok(firstRootRelationship);
const duplicateRootRelationshipXml = rootRelationshipsXml.replace(/<\/Relationships>/, `${firstRootRelationship}</Relationships>`);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "_rels/.rels", text: duplicateRootRelationshipXml }]), /invalid OOXML package.*duplicateRelationshipId/);
const invalidDuplicateDocx = await DocumentFile.patchDocx(docx, [{ path: "_rels/.rels", text: duplicateRootRelationshipXml }], { validateResult: false });
const duplicateInspect = await DocumentFile.inspectDocx(invalidDuplicateDocx);
assert.equal(duplicateInspect.ok, false);
assert.ok(duplicateInspect.issues.some((issue) => issue.type === "duplicateRelationshipId"));
const danglingContentTypeXml = contentTypesXml.replace(/<\/Types>/, '<Override PartName="/customXml/missing.xml" ContentType="application/xml"/></Types>');
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "[Content_Types].xml", text: danglingContentTypeXml }]), /invalid OOXML package.*contentTypeTargetNotFound/);
const orphanRelationshipsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="urn:test" Target="../../word/document.xml"/></Relationships>';
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "customXml/_rels/missing.xml.rels", text: orphanRelationshipsXml }]), /invalid OOXML package.*relationshipSourceNotFound/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "../evil.xml", text: "bad" }]), /Unsafe DOCX part path/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "customXml/large.txt", text: "12345" }], { maxPatchBytes: 4 }), /exceeds maxPatchBytes/);
await assert.rejects(() => DocumentFile.inspectDocx(docx, { maxTotalBytes: 1 }), /maxTotalBytes/);
const nativeOnlyZip = await JSZip.loadAsync(docxBytes);
nativeOnlyZip.remove("word/open-office-artifact.json");
const nativeOnlyDocx = new FileBlob(await nativeOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const nativeOnlyLoaded = await DocumentFile.importDocx(nativeOnlyDocx);
assert.equal(nativeOnlyLoaded.theme.name, "Research Theme");
assert.equal(nativeOnlyLoaded.theme.colors.accent1, "#336699");
assert.equal(nativeOnlyLoaded.theme.fonts.major, "Source Serif 4");
const nativeThemeRuns = nativeOnlyLoaded.blocks.find((item) => item.text === "Theme Latin 中文 العربية")?.runs;
assert.equal(nativeThemeRuns?.[0].style.fontTheme, "majorHAnsi");
assert.equal(nativeThemeRuns?.[0].style.themeTint, "80");
assert.equal(nativeThemeRuns?.[0].style.resolvedColor, "#99b3cc");
assert.equal(nativeThemeRuns?.[0].style.boldComplexScript, false);
assert.equal(nativeThemeRuns?.[0].style.fontSizeComplexScript, 34);
assert.equal(nativeThemeRuns?.[1].style.fontThemeEastAsia, "majorEastAsia");
assert.equal(nativeThemeRuns?.[1].style.fontHint, "eastAsia");
assert.equal(nativeThemeRuns?.[1].style.resolvedFontFamilyEastAsia, "Noto Serif CJK SC");
assert.equal(nativeThemeRuns?.[1].style.themeShade, "BF");
assert.equal(nativeThemeRuns?.[2].style.fontThemeComplexScript, "majorBidi");
assert.equal(nativeThemeRuns?.[2].style.fontHint, "cs");
assert.equal(nativeThemeRuns?.[2].style.boldComplexScript, true);
assert.equal(nativeThemeRuns?.[2].style.italicComplexScript, true);
assert.equal(nativeOnlyLoaded.defaultRunStyle.fontFamily, "Default Serif");
assert.equal(nativeOnlyLoaded.defaultRunStyle.fontSize, 19);
const nativeCascadeRuns = nativeOnlyLoaded.blocks.find((item) => item.text === "Direct override Character theme")?.runs;
assert.equal(nativeCascadeRuns?.[0].style.runStyleId, "CascadeCharacter");
assert.equal(nativeCascadeRuns?.[0].style.resolvedColor, "#123456");
assert.equal(nativeCascadeRuns?.[0].style.italic, false);
assert.equal(nativeCascadeRuns?.[0].style.bold, false);
assert.equal(nativeCascadeRuns?.[0].style.fontSize, 26);
assert.equal(nativeCascadeRuns?.[0].style.resolvedFontFamily, "Source Serif 4");
assert.equal(nativeCascadeRuns?.[1].style.resolvedColor, "#cc3300");
assert.equal(nativeCascadeRuns?.[1].style.italic, true);
assert.equal(nativeOnlyLoaded.headers.find((item) => item.text === "Confidential research memo")?.referenceType, "default");
assert.equal(nativeOnlyLoaded.headers.find((item) => item.text === "First-page research memo")?.referenceType, "first");
assert.equal(nativeOnlyLoaded.headers.find((item) => item.text === "First-page research memo")?.partPath, "word/header2.xml");
assert.equal(nativeOnlyLoaded.footers.find((item) => item.text === "Page footer")?.referenceType, "default");
assert.equal(nativeOnlyLoaded.footers.find((item) => item.text === "Even-page footer")?.referenceType, "even");
assert.equal(nativeOnlyLoaded.footers.find((item) => item.text === "Even-page footer")?.partPath, "word/footer2.xml");
assert.equal(nativeOnlyLoaded.headers.find((item) => item.text === "Opening-section header")?.sectionIndex, 0);
assert.equal(nativeOnlyLoaded.headers.find((item) => item.text === "Confidential research memo")?.sectionIndex, 1);
assert.equal(nativeOnlyLoaded.footers.find((item) => item.text === "Opening-section footer")?.sectionIndex, 0);
assert.equal(nativeOnlyLoaded.footers.find((item) => item.text === "Page footer")?.sectionIndex, 1);
assert.equal(nativeOnlyLoaded.comments.find((item) => item.text === "Check this heading before final export.")?.author, "Reviewer");
assert.equal(nativeOnlyLoaded.comments.find((item) => item.text === "Check this heading before final export.")?.initials, "RV");
assert.equal(nativeOnlyLoaded.comments.find((item) => item.text === "Review the evidence table.")?.author, "R&D Analyst");
assert.equal(nativeOnlyLoaded.comments.find((item) => item.text === "Review the evidence table.")?.date, "2026-07-11T00:15:00.000Z");
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyLoaded.comments.find((item) => item.text === "Review the evidence table.")?.targetId)?.kind, "table");
const nativeCommentRoot = nativeOnlyLoaded.comments.find((item) => item.text === "Check this heading before final export.");
const nativeCommentReply = nativeOnlyLoaded.comments.find((item) => item.text === "Heading language is now approved.");
assert.equal(nativeCommentRoot?.resolved, true);
assert.equal(nativeCommentRoot?.paraId, "00000001");
assert.equal(nativeCommentRoot?.durableId, "0000A001");
assert.equal(nativeCommentRoot?.dateUtc, "2026-07-11T00:10:00.000Z");
assert.deepEqual(nativeCommentRoot?.person, { providerId: "None", userId: "reviewer@example.test" });
assert.equal(nativeCommentReply?.paraId, "00000004");
assert.equal(nativeCommentReply?.durableId, "0000A004");
assert.deepEqual(nativeCommentReply?.person, { providerId: "None", userId: "editor@example.test" });
assert.equal(nativeCommentReply?.parentId, nativeCommentRoot?.id);
assert.equal(nativeCommentReply?.targetId, nativeCommentRoot?.targetId);
const nativeCommentRoundtripZip = await JSZip.loadAsync(new Uint8Array(await (await DocumentFile.exportDocx(nativeOnlyLoaded)).arrayBuffer()));
assert.match(await nativeCommentRoundtripZip.file("word/commentsExtended.xml").async("text"), /w15:paraId="00000004" w15:paraIdParent="00000001"/);
assert.match(await nativeCommentRoundtripZip.file("word/commentsIds.xml").async("text"), /w16cid:paraId="00000001" w16cid:durableId="0000A001"/);
assert.match(await nativeCommentRoundtripZip.file("word/commentsExtensible.xml").async("text"), /w16cex:durableId="0000A001" w16cex:dateUtc="2026-07-11T00:10:00\.000Z"/);
assert.match(await nativeCommentRoundtripZip.file("word/people.xml").async("text"), /w15:author="Reviewer"[\s\S]*?w15:userId="reviewer@example\.test"/);
const nativeOnlyHyperlink = nativeOnlyLoaded.blocks.find((item) => item.kind === "hyperlink" && item.url);
assert.equal(nativeOnlyHyperlink?.text, "w31r4 research note");
assert.equal(nativeOnlyHyperlink?.url, "https://example.com/research");
assert.ok(nativeOnlyHyperlink?.relationshipId);
const nativeOnlyInternalHyperlink = nativeOnlyLoaded.blocks.find((item) => item.kind === "hyperlink" && item.anchor === "FindingsSection");
const nativeOnlyBookmark = nativeOnlyLoaded.bookmarks.find((item) => item.name === "FindingsSection");
const nativeOnlyTableCellBookmark = nativeOnlyLoaded.bookmarks.find((item) => item.name === "EvidenceCells");
assert.equal(nativeOnlyInternalHyperlink?.text, "Jump to findings");
assert.equal(nativeOnlyInternalHyperlink?.history, false);
assert.equal(nativeOnlyInternalHyperlink?.tooltip, "Open the findings section");
assert.equal(nativeOnlyBookmark?.nativeId, 42);
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyBookmark?.targetId)?.text, "Findings");
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyBookmark?.endTargetId)?.text, "Risk callout inherits bold styling.");
assert.equal(nativeOnlyLoaded.resolve("FindingsSection"), nativeOnlyBookmark);
assert.equal(nativeOnlyTableCellBookmark?.nativeId, 43);
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyTableCellBookmark?.targetId)?.value, "DOCX styles");
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyTableCellBookmark?.endTargetId)?.value, "anchored");
assert.equal(nativeOnlyTableCellBookmark?.target?.type, "tableCell");
assert.equal(nativeOnlyLoaded.blocks.find((item) => item.kind === "field")?.instruction, "PAGE");
assert.equal(nativeOnlyLoaded.blocks.find((item) => item.kind === "field")?.display, "1");
assert.equal(nativeOnlyLoaded.blocks.find((item) => item.text === "Lettered evidence group")?.numberFormat, "upperLetter");
assert.equal(nativeOnlyLoaded.blocks.find((item) => item.text === "Nested roman evidence")?.numberFormat, "lowerRoman");
assert.equal(nativeOnlyLoaded.blocks.find((item) => item.text === "Nested roman evidence")?.level, 1);
const nativePictureBullet = nativeOnlyLoaded.blocks.find((item) => item.text === "Picture bullet evidence")?.pictureBullet;
assert.match(nativePictureBullet?.dataUrl || "", /^data:image\/png;base64,/);
assert.equal(nativePictureBullet?.widthPt, 10);
assert.equal(nativePictureBullet?.heightPt, 10);
assert.equal(nativePictureBullet?.alt, "Green status marker");
const nativeOnlyCitation = nativeOnlyLoaded.blocks.find((item) => item.kind === "citation");
assert.match(nativeOnlyCitation?.text || "", /Source: Market brief/);
assert.match(nativeOnlyCitation?.metadata?.bookmark || "", /^OpenOfficeCitation_/);
assert.equal(nativeOnlyCitation?.metadata?.tag, "MarketBrief26");
assert.equal(nativeOnlyCitation?.metadata?.sourceType, "Report");
assert.equal(nativeOnlyCitation?.metadata?.title, "Market brief");
assert.deepEqual(nativeOnlyCitation?.metadata?.authors, [{ first: "Ada", last: "Analyst", middle: "" }]);
assert.equal(nativeOnlyLoaded.bibliographySources.length, 2);
assert.equal(nativeOnlyLoaded.resolve("OpenXmlSdk")?.corporateAuthor, "Microsoft");
const relocatedBibliographyZip = await JSZip.loadAsync(docxBytes);
relocatedBibliographyZip.remove("word/open-office-artifact.json");
relocatedBibliographyZip.remove("customXml/item1.xml");
relocatedBibliographyZip.file("customXml/references/sources.xml", bibliographyXml.replace("xmlns:b=", "xmlns:ref=").replaceAll("<b:", "<ref:").replaceAll("</b:", "</ref:"));
relocatedBibliographyZip.file("word/_rels/document.xml.rels", documentRelsXml.replace('Target="../customXml/item1.xml"', 'Target="../customXml/references/sources.xml"'));
const relocatedBibliographyDocx = new FileBlob(await relocatedBibliographyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
assert.equal((await DocumentFile.inspectDocx(relocatedBibliographyDocx)).ok, true);
const relocatedBibliographyDocument = await DocumentFile.importDocx(relocatedBibliographyDocx, { preferNative: true });
assert.equal(relocatedBibliographyDocument.resolve("MarketBrief26")?.authors[0].last, "Analyst");
assert.equal(relocatedBibliographyDocument.blocks.find((item) => item.kind === "citation")?.metadata?.publisher, "Example Research");
assert.equal(nativeOnlyLoaded.resolve(nativeOnlyLoaded.comments.find((item) => item.text === "Verify the native hyperlink target.")?.targetId)?.kind, "hyperlink");
const alternateCommentPrefixZip = await JSZip.loadAsync(docxBytes);
alternateCommentPrefixZip.remove("word/open-office-artifact.json");
alternateCommentPrefixZip.file("word/comments.xml", commentsXml.replace("xmlns:w14=", "xmlns:c14=").replaceAll("w14:", "c14:"));
alternateCommentPrefixZip.file("word/commentsExtended.xml", commentsExtendedXml.replace("xmlns:w15=", "xmlns:ce=").replaceAll("w15:", "ce:"));
alternateCommentPrefixZip.file("word/commentsIds.xml", commentsIdsXml.replace("xmlns:w16cid=", "xmlns:ci=").replaceAll("w16cid:", "ci:"));
alternateCommentPrefixZip.file("word/commentsExtensible.xml", commentsExtensibleXml.replace("xmlns:w16cex=", "xmlns:cx=").replaceAll("w16cex:", "cx:"));
alternateCommentPrefixZip.file("word/people.xml", peopleXml.replace("xmlns:w15=", "xmlns:pe=").replaceAll("w15:", "pe:"));
const alternateCommentPrefixDocx = new FileBlob(await alternateCommentPrefixZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const alternateCommentPrefixInspect = await DocumentFile.inspectDocx(alternateCommentPrefixDocx);
assert.equal(alternateCommentPrefixInspect.ok, true, JSON.stringify(alternateCommentPrefixInspect.issues));
const alternateCommentPrefixLoaded = await DocumentFile.importDocx(alternateCommentPrefixDocx, { preferNative: true });
const alternateCommentRoot = alternateCommentPrefixLoaded.comments.find((item) => item.text === "Check this heading before final export.");
const alternateCommentReply = alternateCommentPrefixLoaded.comments.find((item) => item.text === "Heading language is now approved.");
assert.equal(alternateCommentRoot?.resolved, true);
assert.equal(alternateCommentRoot?.durableId, "0000A001");
assert.deepEqual(alternateCommentRoot?.person, { providerId: "None", userId: "reviewer@example.test" });
assert.equal(alternateCommentReply?.parentId, alternateCommentRoot?.id);
const relocatedThemeZip = await JSZip.loadAsync(docxBytes);
relocatedThemeZip.remove("word/open-office-artifact.json");
relocatedThemeZip.remove("word/theme/theme1.xml");
const alternatePrefixThemeXml = themeXml.replaceAll("xmlns:a=", "xmlns:d=").replaceAll("<a:", "<d:").replaceAll("</a:", "</d:");
relocatedThemeZip.file("word/config/theme-custom.xml", alternatePrefixThemeXml);
relocatedThemeZip.file("word/_rels/document.xml.rels", documentRelsXml.replace('Target="theme/theme1.xml"', 'Target="config/theme-custom.xml"'));
relocatedThemeZip.file("[Content_Types].xml", contentTypesXml.replace('PartName="/word/theme/theme1.xml"', 'PartName="/word/config/theme-custom.xml"'));
const relocatedThemeDocx = new FileBlob(await relocatedThemeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const relocatedThemeNative = await DocumentFile.importDocx(relocatedThemeDocx, { preferNative: true });
assert.equal(relocatedThemeNative.theme.name, "Research Theme");
assert.equal(relocatedThemeNative.theme.fonts.majorEastAsia, "Noto Serif CJK SC");
assert.equal(relocatedThemeNative.blocks.find((item) => item.text === "Theme Latin 中文 العربية")?.runs[0].style.resolvedColor, "#99b3cc");
const alternateStylesPrefixZip = await JSZip.loadAsync(docxBytes);
alternateStylesPrefixZip.remove("word/open-office-artifact.json");
alternateStylesPrefixZip.file("word/styles.xml", stylesXml.replaceAll("xmlns:w=", "xmlns:s=").replaceAll("<w:", "<s:").replaceAll("</w:", "</s:").replaceAll(" w:", " s:"));
const alternateStylesPrefixDocx = new FileBlob(await alternateStylesPrefixZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const alternateStylesPrefixLoaded = await DocumentFile.importDocx(alternateStylesPrefixDocx, { preferNative: true });
assert.equal(alternateStylesPrefixLoaded.defaultRunStyle.fontFamily, "Default Serif");
assert.equal(alternateStylesPrefixLoaded.styles.get("CascadeCharacter")?.basedOn, "CascadeCharacterBase");
assert.equal(alternateStylesPrefixLoaded.blocks.find((item) => item.text === "Direct override Character theme")?.runs[1].style.resolvedColor, "#cc3300");
const sharedHeaderDocumentXml = documentXml.replace(new RegExp(`(<w:sectPr\\b[^>]*>[\\s\\S]*?<w:headerReference\\b[^>]*r:id=")${openingHeaderRelId}("[^>]*\\/>[\\s\\S]*?<\\/w:sectPr>)`), `$1${finalHeaderRelId}$2`);
const sharedHeaderDocx = await DocumentFile.patchDocx(docx, [
  { path: "word/document.xml", xml: sharedHeaderDocumentXml },
  { path: "word/header3.xml", remove: true },
]);
const sharedHeaderNative = await DocumentFile.importDocx(sharedHeaderDocx, { preferNative: true });
assert.deepEqual(sharedHeaderNative.headers.filter((item) => item.text === "Confidential research memo").map((item) => item.sectionIndex).sort(), [0, 1]);
const nativeOnlyInspect = nativeOnlyLoaded.inspect({ kind: "theme,image,section,change,style,paragraph,table", maxChars: 24000 }).ndjson;
assert.match(nativeOnlyInspect, /Memo logo/);
assert.match(nativeOnlyInspect, /memo-logo/);
assert.match(nativeOnlyInspect, /landscape/);
assert.match(nativeOnlyInspect, /15840/);
assert.match(nativeOnlyInspect, /Inserted reviewer clarification/);
assert.match(nativeOnlyInspect, /Remove stale claim/);
assert.match(nativeOnlyInspect, /Risk Callout/);
assert.match(nativeOnlyInspect, /Run styled paragraph/);
assert.match(nativeOnlyInspect, /0ea5e9/);
assert.match(nativeOnlyInspect, /"basedOn":"Callout"/);
assert.match(nativeOnlyInspect, /Research Theme/);
assert.match(nativeOnlyInspect, /Noto Naskh Arabic/);
assert.equal(nativeOnlyLoaded.styles.effective("RiskCallout").bold, true);
assert.equal(nativeOnlyLoaded.styles.effective("RiskCallout").italic, true);
const reexportedNativeTheme = await DocumentFile.exportDocx(nativeOnlyLoaded);
const reexportedNativeThemeZip = await JSZip.loadAsync(new Uint8Array(await reexportedNativeTheme.arrayBuffer()));
assert.match(await reexportedNativeThemeZip.file("word/numbering.xml").async("text"), /<w:lvlPicBulletId w:val="0"\/>/);
assert.match(await reexportedNativeThemeZip.file("word\/_rels\/numbering.xml.rels").async("text"), /relationships\/image/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /w:cstheme="majorBidi"/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /<w:hyperlink w:anchor="FindingsSection" w:history="0" w:tooltip="Open the findings section">/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /<w:bookmarkStart w:id="42" w:name="FindingsSection"\/>/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /<w:bookmarkStart w:id="43" w:name="EvidenceCells"\/>/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /w:instr="CITATION MarketBrief26"/);
assert.match(await reexportedNativeThemeZip.file("customXml/item1.xml").async("text"), /<b:Tag>OpenXmlSdk<\/b:Tag>/);
assert.match(await reexportedNativeThemeZip.file("word/document.xml").async("text"), /<w:rStyle w:val="CascadeCharacter"\/>/);
assert.match(await reexportedNativeThemeZip.file("word/styles.xml").async("text"), /<w:docDefaults>/);
assert.match(await reexportedNativeThemeZip.file("word/theme/theme1.xml").async("text"), /typeface="Noto Naskh Arabic"/);
const nativeOnlyTable = nativeOnlyLoaded.blocks.find((block) => block.kind === "table");
assert.deepEqual(nativeOnlyTable.columnWidthsDxa, [3600, 5760]);
assert.equal(nativeOnlyTable.widthDxa, 9360);
assert.equal(nativeOnlyTable.indentDxa, 120);
assert.equal(nativeOnlyTable.headerFill, "F2F4F7");
assert.equal(nativeOnlyTable.borderColor, "D9D9D9");
assert.equal(nativeOnlyTable.borderSize, 4);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.docx`);
await docx.save(out);
const loaded = await DocumentFile.importDocx(await FileBlob.load(out));
assert.equal(loaded.blocks.find((item) => item.name === "single-theme-run")?.runs[0].style.themeColor, "accent1");
const loadedInspect = loaded.inspect({ kind: "paragraph,table,bookmark,comment,listItem,header,footer,hyperlink,field,citation,image,section,change", maxChars: 12000 }).ndjson;
assert.match(loadedInspect, /clean-room DOCX facade/);
assert.match(loadedInspect, /Risk callout inherits bold styling/);
assert.match(loadedInspect, /Run styled paragraph/);
assert.match(loadedInspect, /DOCX styles/);
assert.match(loadedInspect, /anchored/);
assert.match(loadedInspect, /Check this heading/);
assert.match(loadedInspect, /w31r4 research note/);
assert.match(loadedInspect, /https:\/\/example.com\/research/);
assert.match(loadedInspect, /FindingsSection/);
assert.match(loadedInspect, /page-field/);
assert.match(loadedInspect, /Market brief/);
assert.match(loadedInspect, /memo-logo/);
assert.match(loadedInspect, /Memo logo/);
assert.match(loadedInspect, /landscape-appendix/);
assert.match(loadedInspect, /landscape/);
assert.match(loadedInspect, /Inserted reviewer clarification/);
assert.match(loadedInspect, /Remove stale claim/);
assert.match(loadedInspect, /Use real numbering definitions/);
assert.match(loadedInspect, /Render and verify/);
assert.match(loadedInspect, /Confidential research memo/);
assert.match(loadedInspect, /Page footer/);
console.log("document smoke ok");
