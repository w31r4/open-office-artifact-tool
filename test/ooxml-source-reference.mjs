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

async function zipOf(file) {
  return JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
}

// Explicit package patching remains a low-level operation. It is never selected
// by an Office facade as a codec or fallback.
const document = DocumentModel.create({ blocks: [] });
document.addParagraph("Patch this DOCX");
const baseDocx = await DocumentFile.exportDocx(document);
assert.equal((await DocumentFile.inspectDocx(baseDocx)).ok, true);

const docxZip = await zipOf(baseDocx);
const documentXml = await docxZip.file("word/document.xml").async("text");
assert.match(documentXml, /Patch this DOCX/);
const patchedDocx = await DocumentFile.patchDocx(baseDocx, [{
  path: "word/document.xml",
  xml: documentXml.replace("Patch this DOCX", "Patched DOCX"),
}]);
assert.equal((await DocumentFile.importDocx(patchedDocx)).blocks[0].text, "Patched DOCX");

const headerDocx = await DocumentFile.patchDocx(patchedDocx, [{
  path: "word/headerPatch.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Patched header</w:t></w:r></w:p></w:hdr>',
  recipe: {
    kind: "header",
    source: "word/document.xml",
    id: "rIdPatchedHeader",
    sourceReference: { type: "default" },
  },
}]);
const importedHeaderDocument = await DocumentFile.importDocx(headerDocx);
assert.equal(importedHeaderDocument.headers.find((header) => header.referenceType === "default")?.text, "Patched header");
assert.equal((await DocumentFile.inspectDocx(headerDocx)).ok, true);

await assert.rejects(
  () => DocumentFile.patchDocx(baseDocx, [{ path: "../escape.xml", xml: "<escape/>" }]),
  /unsafe|traversal|outside|invalid/i,
);
await assert.rejects(
  () => DocumentFile.patchDocx(baseDocx, [{ path: "_rels/.rels", remove: true }]),
  /invalid OOXML package|missingRootRelationships|missing/i,
);

const workbook = Workbook.create();
workbook.worksheets.add("Main").getRange("A1:B2").values = [["Metric", "Value"], ["Revenue", 120]];
const baseXlsx = await SpreadsheetFile.exportXlsx(workbook);
assert.equal((await SpreadsheetFile.inspectXlsx(baseXlsx)).ok, true);

const xlsxZip = await zipOf(baseXlsx);
const sheetXml = await xlsxZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(sheetXml, />120</);
const patchedXlsx = await SpreadsheetFile.patchXlsx(baseXlsx, [{
  path: "xl/worksheets/sheet1.xml",
  xml: sheetXml.replace(">120<", ">125<"),
}]);
assert.equal((await SpreadsheetFile.importXlsx(patchedXlsx)).worksheets.getItem("Main").getRange("B2").values[0][0], 125);

const addedSheetXlsx = await SpreadsheetFile.patchXlsx(patchedXlsx, [{
  path: "xl/worksheets/patched-sheet.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Added</t></is></c></row></sheetData></worksheet>',
  recipe: {
    kind: "worksheet",
    source: "xl/workbook.xml",
    id: "rIdPatchedSheet",
    sourceReference: { name: "Patched Added" },
  },
}]);
const importedAddedSheet = (await SpreadsheetFile.importXlsx(addedSheetXlsx)).worksheets.getItem("Patched Added");
assert.ok(importedAddedSheet);
assert.equal(importedAddedSheet.getRange("A1").values[0][0], "Added");
assert.equal((await SpreadsheetFile.inspectXlsx(addedSheetXlsx)).ok, true);

const presentation = Presentation.create();
presentation.slides.add({ name: "Patch slide" }).shapes.add({
  name: "patch-text",
  geometry: "textbox",
  text: "Patch this PPTX",
  position: { left: 40, top: 40, width: 400, height: 80 },
});
const basePptx = await PresentationFile.exportPptx(presentation);
assert.equal((await PresentationFile.inspectPptx(basePptx)).ok, true);

const pptxZip = await zipOf(basePptx);
const slideXml = await pptxZip.file("ppt/slides/slide1.xml").async("text");
assert.match(slideXml, /Patch this PPTX/);
const patchedPptx = await PresentationFile.patchPptx(basePptx, [{
  path: "ppt/slides/slide1.xml",
  xml: slideXml.replace("Patch this PPTX", "Patched PPTX"),
}]);
const importedPresentation = await PresentationFile.importPptx(patchedPptx);
assert.equal(importedPresentation.slides.getItem(0).shapes.items[0].text.value, "Patched PPTX");
assert.equal((await PresentationFile.inspectPptx(patchedPptx)).ok, true);

console.log("OOXML inspect/patch tests passed");
