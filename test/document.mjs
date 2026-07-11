import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, FileBlob } from "../src/index.mjs";

const document = DocumentModel.create({
  name: "Research memo",
  paragraphs: ["Research memo", "This document exercises the clean-room DOCX facade."],
});
document.applyDesignPreset("report");
document.styles.add("Callout", { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos" });
const header = document.addHeader("Confidential research memo", { name: "default-header" });
const footer = document.addFooter("Page footer", { name: "default-footer" });
const heading = document.addParagraph("Findings", { styleId: "Heading1", name: "findings-heading" });
const hyperlink = document.addHyperlink("w31r4 research note", "https://example.com/research", { name: "research-link" });
const field = document.addField("PAGE", "1", { name: "page-field" });
const citation = document.addCitation("Source: Market brief", { source: "Market brief", url: "https://example.com/brief", page: 2 }, { name: "market-citation" });
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
const insertion = document.addInsertion("Inserted reviewer clarification.", { author: "Reviewer", date: "2026-07-11T00:00:00.000Z", name: "tracked-insert" });
const deletion = document.addDeletion("Remove stale claim.", { author: "Reviewer", date: "2026-07-11T00:05:00.000Z", name: "tracked-delete" });
const bullet = document.addListItem("Use real numbering definitions", { listType: "bullet", name: "numbering-rule" });
const numbered = document.addListItem("Render and verify", { listType: "number", name: "render-step" });
const table = document.addTable({
  name: "evidence-table",
  styleId: "TableGrid",
  values: [["Area", "Status"], ["DOCX styles", "partial"], ["Comments", "roundtrip"]],
});
table.getCell(2, 1).value = "anchored";
const comment = document.addComment(heading, "Check this heading before final export.", { author: "Reviewer" });

const inspect = document.inspect({ kind: "document,paragraph,table,comment,style,listItem,header,footer,hyperlink,field,citation,image,section,change,layout", maxChars: 16000 }).ndjson;
assert.match(inspect, /Research memo/);
assert.match(inspect, /"kind":"document"/);
assert.match(inspect, /"designPreset":"report"/);
assert.match(inspect, /"kind":"layout"/);
assert.match(inspect, /findings-heading/);
assert.match(inspect, /research-link/);
assert.match(inspect, /https:\/\/example.com\/research/);
assert.match(inspect, /page-field/);
assert.match(inspect, /market-citation/);
assert.match(inspect, /memo-logo/);
assert.match(inspect, /Memo logo/);
assert.match(inspect, /landscape-appendix/);
assert.match(inspect, /landscape/);
assert.match(inspect, /tracked-insert/);
assert.match(inspect, /Inserted reviewer clarification/);
assert.match(inspect, /tracked-delete/);
assert.match(inspect, /Remove stale claim/);
assert.match(inspect, /numbering-rule/);
assert.match(inspect, /render-step/);
assert.match(inspect, /Confidential research memo/);
assert.match(inspect, /Page footer/);
assert.match(inspect, /evidence-table/);
assert.match(inspect, /Check this heading/);
assert.match(inspect, /Callout/);
assert.equal(document.resolve(heading.id).styleId, "Heading1");
assert.equal(document.resolve(hyperlink.id).url, "https://example.com/research");
assert.equal(document.resolve(field.id).instruction, "PAGE");
assert.equal(document.resolve(citation.id).metadata.page, 2);
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
assert.equal(document.resolve(header.id).text, "Confidential research memo");
assert.equal(document.resolve(footer.id).text, "Page footer");
assert.equal(document.resolve(table.id).getCell(2, 1).value, "anchored");
assert.equal(document.resolve(comment.id).author, "Reviewer");
const targetedDocumentInspect = document.inspect({ kind: "paragraph,table,image,comment", target: logo.id, maxChars: 4000 }).ndjson;
assert.match(targetedDocumentInspect, /Memo logo/);
assert.doesNotMatch(targetedDocumentInspect, /evidence-table/);
const targetedCommentInspect = document.inspect({ kind: "comment,paragraph", target: comment.id, maxChars: 4000 }).ndjson;
assert.match(targetedCommentInspect, /Check this heading/);
assert.doesNotMatch(targetedCommentInspect, /findings-heading/);
assert.match(document.help("document.addTable").ndjson, /Word-style table/);
assert.match(document.help("document.addListItem").ndjson, /numbering definitions/);
assert.match(document.help("document.addHeader").ndjson, /DOCX header/);
assert.match(document.help("document.addHyperlink").ndjson, /w:hyperlink/);
assert.match(document.help("document.addField").ndjson, /w:fldSimple/);
assert.match(document.help("document.addCitation").ndjson, /structured metadata/);
assert.match(document.help("document.addImage").ndjson, /native DOCX media/);
assert.match(document.help("document.addSection").ndjson, /w:sectPr/);
assert.match(document.help("document.addInsertion").ndjson, /w:ins/);
assert.match(document.help("document.addDeletion").ndjson, /w:del/);
assert.match(document.help("document.applyDesignPreset").ndjson, /design preset/);
assert.match(document.help("document.layoutJson").ndjson, /layout JSON/);
assert.equal(document.verify({ visualQa: true }).ok, true);

const preview = await document.render();
assert.equal(preview.type, "image/svg+xml");
const svg = await preview.text();
assert.match(svg, /Research memo/);
assert.match(svg, /DOCX styles/);
assert.match(svg, /Confidential research memo/);
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
assert.ok(layout.elements.some((element) => element.id === table.id));

const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const docxBytes = new Uint8Array(await docx.arrayBuffer());
const zip = await JSZip.loadAsync(docxBytes);
const documentXml = await zip.file("word/document.xml").async("text");
assert.match(documentXml, /<a:blip r:embed="rIdImage1"\/>/);
assert.match(documentXml, /<wp:docPr[^>]*name="memo-logo"[^>]*descr="Memo logo"/);
assert.match(documentXml, /<w:type w:val="nextPage"\/>/);
assert.match(documentXml, /<w:pgSz w:w="15840" w:h="12240" w:orient="landscape"\/>/);
assert.match(documentXml, /<w:pgMar w:top="720" w:right="900" w:bottom="720" w:left="900"/);
assert.match(documentXml, /<w:ins\b[^>]*w:author="Reviewer"/);
assert.match(documentXml, /<w:t>Inserted reviewer clarification\.<\/w:t>/);
assert.match(documentXml, /<w:del\b[^>]*w:author="Reviewer"/);
assert.match(documentXml, /<w:delText>Remove stale claim\.<\/w:delText>/);
const documentRelsXml = await zip.file("word/_rels/document.xml.rels").async("text");
assert.match(documentRelsXml, /Id="rIdImage1"/);
assert.match(documentRelsXml, /Target="media\/image1\.png"/);
const docxMediaBytes = await zip.file("word/media/image1.png").async("uint8array");
assert.ok(docxMediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
const nativeOnlyZip = await JSZip.loadAsync(docxBytes);
nativeOnlyZip.remove("word/open-office-artifact.json");
const nativeOnlyDocx = new FileBlob(await nativeOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const nativeOnlyLoaded = await DocumentFile.importDocx(nativeOnlyDocx);
const nativeOnlyInspect = nativeOnlyLoaded.inspect({ kind: "image,section,change", maxChars: 12000 }).ndjson;
assert.match(nativeOnlyInspect, /Memo logo/);
assert.match(nativeOnlyInspect, /memo-logo/);
assert.match(nativeOnlyInspect, /landscape/);
assert.match(nativeOnlyInspect, /15840/);
assert.match(nativeOnlyInspect, /Inserted reviewer clarification/);
assert.match(nativeOnlyInspect, /Remove stale claim/);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.docx`);
await docx.save(out);
const loaded = await DocumentFile.importDocx(await FileBlob.load(out));
const loadedInspect = loaded.inspect({ kind: "paragraph,table,comment,listItem,header,footer,hyperlink,field,citation,image,section,change", maxChars: 12000 }).ndjson;
assert.match(loadedInspect, /clean-room DOCX facade/);
assert.match(loadedInspect, /DOCX styles/);
assert.match(loadedInspect, /anchored/);
assert.match(loadedInspect, /Check this heading/);
assert.match(loadedInspect, /w31r4 research note/);
assert.match(loadedInspect, /https:\/\/example.com\/research/);
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
