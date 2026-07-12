import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import JSZip from "jszip";

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
  recipe: { kind: "header", source: "word/document.xml", sourceReference: { type: "first" } },
}]);
assert.equal(docx.metadata.sourceReferencesUpdated, 1);
assert.equal((await DocumentFile.inspectDocx(docx)).ok, true);
const importedPatchedDocument = await DocumentFile.importDocx(docx, { preferNative: true });
const importedNativeHeader = importedPatchedDocument.headers.find((item) => item.text === "Native header");
assert.ok(importedNativeHeader);
assert.equal(importedNativeHeader.referenceType, "first");
assert.equal(importedNativeHeader.partPath, "word/headerNative.xml");
const docxZip = await JSZip.loadAsync(new Uint8Array(await docx.arrayBuffer()));
assert.match(await docxZip.file("word/document.xml").async("text"), /<w:headerReference\b[^>]*w:type="first"[^>]*\/>[\s\S]*?<w:titlePg\/>/);

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
const drawingPartPath = "xl/custom/drawings/agent-drawing.xml";
const drawingXml = '<?xml version="1.0" encoding="UTF-8"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"/>';
const chartXml = (title) => `<?xml version="1.0" encoding="UTF-8"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:p><a:r><a:t>${title}</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:barChart><c:ser><c:tx><c:v>Value</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>Revenue</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>120</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>`;
const drawingXlsx = await SpreadsheetFile.patchXlsx(xlsx, [
  { path: "customXml/open-office-artifact.json", remove: true },
  { path: drawingPartPath, xml: drawingXml, recipe: { kind: "drawing", source: "xl/worksheets/sheet1.xml", id: "rIdAgentDrawing", sourceReference: true } },
  {
    path: "xl/custom/media/agent-status.png",
    bytes: Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64"),
    recipe: {
      kind: "image",
      source: drawingPartPath,
      id: "rIdAgentImage",
      sourceReference: { name: "Agent status", alt: "Green agent status", objectId: 2, anchor: { type: "oneCell", from: { row: 4, col: 0 }, extent: { widthPx: 48, heightPx: 48 } } },
    },
  },
  {
    path: "xl/custom/charts/agent-absolute.xml",
    xml: chartXml("Absolute review"),
    recipe: {
      kind: "chart",
      source: drawingPartPath,
      id: "rIdAgentAbsoluteChart",
      sourceReference: { name: "Absolute review chart", objectId: 3, anchor: { type: "absolute", position: { leftPx: 180, topPx: 40 }, extent: { widthPx: 260, heightPx: 160 } } },
    },
  },
  {
    path: "xl/custom/charts/agent-two-cell.xml",
    xml: chartXml("Two-cell review"),
    recipe: {
      kind: "chart",
      source: drawingPartPath,
      id: "rIdAgentTwoCellChart",
      sourceReference: { name: "Two-cell review chart", objectId: 4, anchor: { type: "twoCell", from: { row: 10, col: 0 }, to: { row: 18, col: 5 }, extent: { widthPx: 320, heightPx: 180 } } },
    },
  },
]);
assert.equal(drawingXlsx.metadata.sourceReferencesUpdated, 4);
assert.equal((await SpreadsheetFile.inspectXlsx(drawingXlsx)).ok, true);
const drawingXlsxZip = await JSZip.loadAsync(new Uint8Array(await drawingXlsx.arrayBuffer()));
assert.match(await drawingXlsxZip.file("xl/worksheets/sheet1.xml").async("text"), /<drawing r:id="rIdAgentDrawing"\/>/);
const patchedDrawingXml = await drawingXlsxZip.file(drawingPartPath).async("text");
assert.match(patchedDrawingXml, /<xdr:oneCellAnchor>/);
assert.match(patchedDrawingXml, /<xdr:absoluteAnchor>/);
assert.match(patchedDrawingXml, /<xdr:twoCellAnchor>/);
assert.match(patchedDrawingXml, /r:embed="rIdAgentImage"/);
assert.match(patchedDrawingXml, /r:id="rIdAgentAbsoluteChart"/);
const importedDrawingWorkbook = await SpreadsheetFile.importXlsx(drawingXlsx);
const importedDrawingSheet = importedDrawingWorkbook.worksheets.getItem("Main");
assert.equal(importedDrawingSheet.images.items.length, 1);
assert.equal(importedDrawingSheet.images.items[0].alt, "Green agent status");
assert.equal(importedDrawingSheet.charts.items.length, 2);
assert.deepEqual(importedDrawingSheet.charts.items.find((chart) => chart.title === "Absolute review").position, { left: 220, top: 80, width: 260, height: 160 });
assert.ok(importedDrawingSheet.charts.items.find((chart) => chart.title === "Two-cell review").position.width > 200);
await assert.rejects(() => SpreadsheetFile.patchXlsx(drawingXlsx, [{ path: "xl/custom/charts/missing-anchor.xml", xml: chartXml("Missing anchor"), recipe: { kind: "chart", source: drawingPartPath, sourceReference: true } }]), /requires an explicit anchor/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(drawingXlsx, [{ path: "xl/custom/charts/duplicate-object.xml", xml: chartXml("Duplicate object"), recipe: { kind: "chart", source: drawingPartPath, sourceReference: { objectId: 2, anchor: { type: "oneCell", from: { row: 20, col: 0 }, extent: { widthPx: 240, heightPx: 140 } } } } }]), /objectId 2 already exists/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(drawingXlsx, [{ path: "xl/custom/drawings/second.xml", xml: drawingXml, recipe: { kind: "drawing", source: "xl/worksheets/sheet1.xml", sourceReference: true } }]), /without another drawing reference/);
const drawingObjectsRemoved = await SpreadsheetFile.patchXlsx(drawingXlsx, [
  { path: "xl/custom/media/agent-status.png", remove: true, recipe: { kind: "image", source: drawingPartPath, id: "rIdAgentImage", sourceReference: true } },
  { path: "xl/custom/charts/agent-absolute.xml", remove: true, recipe: { kind: "chart", source: drawingPartPath, id: "rIdAgentAbsoluteChart", sourceReference: true } },
  { path: "xl/custom/charts/agent-two-cell.xml", remove: true, recipe: { kind: "chart", source: drawingPartPath, id: "rIdAgentTwoCellChart", sourceReference: true } },
]);
assert.equal((await SpreadsheetFile.inspectXlsx(drawingObjectsRemoved)).ok, true);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await drawingObjectsRemoved.arrayBuffer()))).file(drawingPartPath).async("text"), /(?:oneCell|twoCell|absolute)Anchor/);
const drawingRemoved = await SpreadsheetFile.patchXlsx(drawingObjectsRemoved, [{ path: drawingPartPath, remove: true, recipe: { kind: "drawing", source: "xl/worksheets/sheet1.xml", id: "rIdAgentDrawing", sourceReference: true } }]);
assert.equal((await SpreadsheetFile.inspectXlsx(drawingRemoved)).ok, true);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await drawingRemoved.arrayBuffer()))).file("xl/worksheets/sheet1.xml").async("text"), /<drawing\b/);

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
  for (const [artifactKind, blob] of [["document", docx], ["workbook", drawingXlsx], ["presentation", pptx]]) {
    const pdf = await libreOffice({ input: blob, inputType: blob.type, outputType: "application/pdf", format: "pdf", artifactKind });
    assert.equal(pdf.type, "application/pdf");
    assert.ok(pdf.bytes.length > 100);
    const png = await poppler({ input: pdf, inputType: pdf.type, outputType: "image/png", format: "png", artifactKind, pageIndex: 0 });
    assert.equal(png.type, "image/png");
    assert.deepEqual([...png.bytes.slice(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  }
}

console.log(`OOXML source reference smoke ok${nativeAvailable ? " (native)" : " (native skipped)"}`);
