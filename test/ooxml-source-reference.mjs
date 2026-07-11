import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";

import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "open-office-artifact-tool";
import { createLibreOfficeRenderer } from "open-office-artifact-tool/renderers/libreoffice";
import { createPopplerRenderer } from "open-office-artifact-tool/renderers/poppler";

function commandExists(command) {
  return spawnSync(process.platform === "win32" ? "where" : "which", [command], { encoding: "utf8", shell: false }).status === 0;
}

const document = DocumentModel.create({ paragraphs: ["Source reference native check"] });
const docx = await DocumentFile.patchDocx(await DocumentFile.exportDocx(document), [{
  path: "word/headerNative.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>',
  recipe: { kind: "header", source: "word/document.xml", sourceReference: true },
}]);
assert.equal(docx.metadata.sourceReferencesUpdated, 1);
assert.equal((await DocumentFile.inspectDocx(docx)).ok, true);

const workbook = Workbook.create();
workbook.worksheets.add("Main").getRange("A1:B2").values = [["Metric", "Value"], ["Revenue", 120]];
const xlsx = await SpreadsheetFile.patchXlsx(await SpreadsheetFile.exportXlsx(workbook), [{
  path: "xl/worksheets/sheetNative.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Native sheet</t></is></c></row></sheetData></worksheet>',
  recipe: { kind: "worksheet", source: "xl/workbook.xml", sourceReference: { name: "Native Added" } },
}]);
assert.equal(xlsx.metadata.sourceReferencesUpdated, 1);
assert.equal((await SpreadsheetFile.inspectXlsx(xlsx)).ok, true);
assert.ok((await SpreadsheetFile.importXlsx(xlsx)).worksheets.getItem("Native Added"));

const presentation = Presentation.create();
presentation.slides.add().shapes.add({ text: "Source reference native check", position: { left: 40, top: 40, width: 400, height: 80 } });
const pptx = await PresentationFile.patchPptx(await PresentationFile.exportPptx(presentation), [{
  path: "ppt/slides/slideNative.xml",
  xml: '<?xml version="1.0" encoding="UTF-8"?><p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld></p:sld>',
  recipe: { kind: "slide", source: "ppt/presentation.xml", sourceReference: true },
}]);
assert.equal(pptx.metadata.sourceReferencesUpdated, 1);
assert.equal((await PresentationFile.inspectPptx(pptx)).ok, true);
assert.equal((await PresentationFile.importPptx(pptx)).slides.items.length, 2);

const nativeAvailable = commandExists("soffice") && commandExists("pdftoppm");
if (nativeAvailable) {
  const libreOffice = createLibreOfficeRenderer({ timeoutMs: 60_000 });
  const poppler = createPopplerRenderer({ dpi: 96, timeoutMs: 60_000 });
  for (const [artifactKind, blob] of [["document", docx], ["workbook", xlsx], ["presentation", pptx]]) {
    const pdf = await libreOffice({ input: blob, inputType: blob.type, outputType: "application/pdf", format: "pdf", artifactKind });
    assert.equal(pdf.type, "application/pdf");
    assert.ok(pdf.bytes.length > 100);
    const png = await poppler({ input: pdf, inputType: pdf.type, outputType: "image/png", format: "png", artifactKind, pageIndex: 0 });
    assert.equal(png.type, "image/png");
    assert.deepEqual([...png.bytes.slice(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  }
}

console.log(`OOXML source reference smoke ok${nativeAvailable ? " (native)" : " (native skipped)"}`);
