import assert from "node:assert/strict";
import fs from "node:fs/promises";
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
  inspectOoxmlPackage,
  ooxmlResolveRelationshipTarget,
  ooxmlSafePartPath,
  patchOoxmlPackage,
} from "../src/ooxml/package.mjs";

const packageModule = await import("../src/ooxml/package.mjs");
assert.deepEqual(Object.keys(packageModule).sort(), [
  "inspectOoxmlPackage",
  "ooxmlResolveRelationshipTarget",
  "ooxmlSafePartPath",
  "patchOoxmlPackage",
]);
assert.equal(typeof inspectOoxmlPackage, "function");
assert.equal(typeof patchOoxmlPackage, "function");
assert.equal(ooxmlSafePartPath("word\\media/image1.png", "DOCX"), "word/media/image1.png");
assert.equal(ooxmlResolveRelationshipTarget("ppt/slides/slide1.xml", "../media/image1.png"), "ppt/media/image1.png");
assert.throws(() => ooxmlSafePartPath("../escape.xml", "DOCX"), /unsafe/i);
assert.throws(() => ooxmlSafePartPath("/absolute.xml", "DOCX"), /unsafe/i);

const rootSource = await fs.readFile(new URL("../src/index.mjs", import.meta.url), "utf8");
const packageSource = await fs.readFile(new URL("../src/ooxml/package.mjs", import.meta.url), "utf8");
assert.doesNotMatch(rootSource, /from ["']jszip["']/, "root compatibility barrel must not own JSZip package mechanics");
assert.doesNotMatch(rootSource, /function ooxmlPackageRecords\b/, "root compatibility barrel must not own the OOXML package engine");
assert.doesNotMatch(packageSource, /(?:^|\/)index\.mjs["']/, "OOXML package engine must remain a dependency leaf");

async function zipOf(file) {
  return JSZip.loadAsync(new Uint8Array(await file.arrayBuffer()));
}

async function bytesOf(zip) {
  return zip.generateAsync({ type: "uint8array" });
}

function corruptCentralDirectoryCrc(bytes, entryName) {
  const result = new Uint8Array(bytes);
  const view = new DataView(result.buffer, result.byteOffset, result.byteLength);
  const name = Buffer.from(entryName, "utf8");
  for (let offset = 0; offset + 46 <= result.byteLength; offset += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) continue;
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const end = offset + 46 + nameLength + extraLength + commentLength;
    if (end > result.byteLength) break;
    const candidate = result.subarray(offset + 46, offset + 46 + nameLength);
    if (!Buffer.from(candidate).equals(name)) {
      offset = end - 1;
      continue;
    }
    view.setUint32(offset + 16, view.getUint32(offset + 16, true) ^ 1, true);
    return result;
  }
  throw new Error(`Could not locate ZIP central-directory entry ${entryName}.`);
}

// Explicit package patching remains a low-level operation. It is never selected
// by an Office facade as a codec or fallback.
const document = DocumentModel.create({ blocks: [] });
document.addParagraph("Patch this DOCX");
const baseDocx = await DocumentFile.exportDocx(document);
assert.equal((await DocumentFile.inspectDocx(baseDocx, { verifyCrc32: true })).ok, true);

const missingMainPartZip = await zipOf(baseDocx);
missingMainPartZip.remove("word/document.xml");
const missingMainPartInspection = await DocumentFile.inspectDocx(
  await bytesOf(missingMainPartZip),
  { verifyCrc32: true },
);
assert.equal(missingMainPartInspection.ok, false);
assert.ok(
  missingMainPartInspection.issues.some(
    (issue) => issue.type === "missingOfficeDocumentPart",
  ),
);

const invalidContentTypeZip = await zipOf(baseDocx);
const contentTypesXml = await invalidContentTypeZip.file("[Content_Types].xml").async("text");
invalidContentTypeZip.file(
  "[Content_Types].xml",
  contentTypesXml.replace(
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
    "application/octet-stream",
  ),
);
const invalidContentTypeInspection = await DocumentFile.inspectDocx(
  await bytesOf(invalidContentTypeZip),
  { verifyCrc32: true },
);
assert.equal(invalidContentTypeInspection.ok, false);
assert.ok(
  invalidContentTypeInspection.issues.some(
    (issue) => issue.type === "invalidOfficeDocumentContentType",
  ),
);

const invalidRootRelationshipZip = await zipOf(baseDocx);
const rootRelationshipsXml = await invalidRootRelationshipZip.file("_rels/.rels").async("text");
invalidRootRelationshipZip.file(
  "_rels/.rels",
  rootRelationshipsXml.replace(
    /Target="\/?word\/document\.xml"/u,
    'Target="/word/not-the-main-document.xml"',
  ),
);
const invalidRootRelationshipInspection = await DocumentFile.inspectDocx(
  await bytesOf(invalidRootRelationshipZip),
  { verifyCrc32: true },
);
assert.equal(invalidRootRelationshipInspection.ok, false);
assert.ok(
  invalidRootRelationshipInspection.issues.some(
    (issue) => issue.type === "invalidOfficeDocumentRelationship",
  ),
);

const externalRootRelationshipZip = await zipOf(baseDocx);
const externalRootRelationshipsXml = await externalRootRelationshipZip.file("_rels/.rels").async("text");
externalRootRelationshipZip.file(
  "_rels/.rels",
  externalRootRelationshipsXml.replace(
    /Target="\/?word\/document\.xml"/u,
    'Target="https://example.invalid/document.xml" TargetMode="External"',
  ),
);
const externalRootRelationshipInspection = await DocumentFile.inspectDocx(
  await bytesOf(externalRootRelationshipZip),
  { verifyCrc32: true },
);
assert.equal(externalRootRelationshipInspection.ok, false);
assert.ok(
  externalRootRelationshipInspection.issues.some(
    (issue) => issue.type === "invalidOfficeDocumentRelationship",
  ),
);

const crcCorruptedDocx = corruptCentralDirectoryCrc(
  new Uint8Array(await baseDocx.arrayBuffer()),
  "word/document.xml",
);
await assert.rejects(
  () => DocumentFile.inspectDocx(crcCorruptedDocx, { verifyCrc32: true }),
  /crc32|corrupted zip|corrupt/i,
);

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
