import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, FileBlob, renderArtifact } from "../src/index.mjs";

const document = DocumentModel.create({
  name: "Research memo",
  paragraphs: ["Research memo", "This document exercises the clean-room DOCX facade."],
});
document.applyDesignPreset("report");
document.styles.add("Callout", { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos", color: "#0f172a" });
document.styles.add("RiskCallout", { name: "Risk Callout", basedOn: "Callout", italic: true, color: "#b91c1c" });
const header = document.addHeader("Confidential research memo", { name: "default-header" });
const footer = document.addFooter("Page footer", { name: "default-footer" });
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
  widthDxa: 9360,
  indentDxa: 120,
  columnWidthsDxa: [3600, 5760],
  headerFill: "F2F4F7",
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
const shapedDocumentInspect = document.inspect({ kind: "paragraph", target: heading.id, include: "text,style", exclude: "comments", maxChars: 4000 }).ndjson;
assert.match(shapedDocumentInspect, /Findings/);
assert.match(shapedDocumentInspect, /Heading1/);
assert.doesNotMatch(shapedDocumentInspect, /Check this heading/);
const targetedCommentInspect = document.inspect({ kind: "comment,paragraph", target: comment.id, maxChars: 4000 }).ndjson;
assert.match(targetedCommentInspect, /Check this heading/);
assert.doesNotMatch(targetedCommentInspect, /findings-heading/);
assert.match(document.help("document.addParagraph").ndjson, /run-level styles/);
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
assert.match(svg, /Run styled /);
assert.match(svg, /#0ea5e9/);
assert.match(svg, /#f97316/);
assert.match(svg, /Risk callout inherits bold styling/);
assert.match(svg, /#b91c1c/);
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
assert.equal(layout.elements.find((element) => element.id === riskCallout.id).effectiveStyle.bold, true);
assert.equal(layout.elements.find((element) => element.id === riskCallout.id).effectiveStyle.italic, true);
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
assert.match(document.help("DocumentFile.inspectDocx").ndjson, /DOCX zip package/);
assert.match(document.help("DocumentFile.patchDocx").ndjson, /path traversal/);

const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const docxBytes = new Uint8Array(await docx.arrayBuffer());
const zip = await JSZip.loadAsync(docxBytes);
const documentXml = await zip.file("word/document.xml").async("text");
const stylesXml = await zip.file("word/styles.xml").async("text");
assert.match(stylesXml, /w:styleId="RiskCallout"/);
assert.match(stylesXml, /<w:basedOn w:val="Callout"\/>/);
assert.match(stylesXml, /<w:i\/>/);
assert.match(stylesXml, /<w:color w:val="b91c1c"\/>/);
assert.match(documentXml, /<w:t>Run styled <\/w:t>/);
assert.match(documentXml, /<w:color w:val="0ea5e9"\/>/);
assert.match(documentXml, /<w:t>paragraph<\/w:t>/);
assert.match(documentXml, /<w:color w:val="f97316"\/>/);
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
const documentRelsXml = await zip.file("word/_rels/document.xml.rels").async("text");
assert.match(documentRelsXml, /Id="rIdImage1"/);
assert.match(documentRelsXml, /Target="media\/image1\.png"/);
const docxMediaBytes = await zip.file("word/media/image1.png").async("uint8array");
assert.ok(docxMediaBytes.byteLength > 10);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
const packageInspect = await DocumentFile.inspectDocx(docx, { includeText: true, maxChars: 12000 });
assert.ok(packageInspect.parts.some((part) => part.path === "word/document.xml"));
assert.match(packageInspect.ndjson, /docxPart/);
assert.match(packageInspect.ndjson, /word\/styles\.xml/);
assert.ok(packageInspect.records[0].uncompressedBytes > 0);
assert.ok(packageInspect.parts.some((part) => part.path === "word/document.xml" && part.contentType.includes("wordprocessingml.document.main+xml")));
const patchedDocx = await DocumentFile.patchDocx(docx, [{ path: "customXml/review-note.xml", text: "<review>ok</review>" }]);
assert.equal(patchedDocx.type, docx.type);
assert.equal(patchedDocx.metadata.patchedParts, 1);
const patchedInspect = await DocumentFile.inspectDocx(patchedDocx, { includeText: true, maxChars: 12000 });
assert.match(patchedInspect.ndjson, /customXml\/review-note\.xml/);
assert.match(patchedInspect.ndjson, /&lt;review&gt;ok&lt;\/review&gt;|<review>ok<\/review>/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "../evil.xml", text: "bad" }]), /Unsafe DOCX part path/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "customXml/large.txt", text: "12345" }], { maxPatchBytes: 4 }), /exceeds maxPatchBytes/);
await assert.rejects(() => DocumentFile.inspectDocx(docx, { maxTotalBytes: 1 }), /maxTotalBytes/);
const nativeOnlyZip = await JSZip.loadAsync(docxBytes);
nativeOnlyZip.remove("word/open-office-artifact.json");
const nativeOnlyDocx = new FileBlob(await nativeOnlyZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), { type: docx.type });
const nativeOnlyLoaded = await DocumentFile.importDocx(nativeOnlyDocx);
const nativeOnlyInspect = nativeOnlyLoaded.inspect({ kind: "image,section,change,style,paragraph,table", maxChars: 16000 }).ndjson;
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
assert.equal(nativeOnlyLoaded.styles.effective("RiskCallout").bold, true);
assert.equal(nativeOnlyLoaded.styles.effective("RiskCallout").italic, true);
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
const loadedInspect = loaded.inspect({ kind: "paragraph,table,comment,listItem,header,footer,hyperlink,field,citation,image,section,change", maxChars: 12000 }).ndjson;
assert.match(loadedInspect, /clean-room DOCX facade/);
assert.match(loadedInspect, /Risk callout inherits bold styling/);
assert.match(loadedInspect, /Run styled paragraph/);
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
