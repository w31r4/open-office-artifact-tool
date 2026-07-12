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

const document = DocumentModel.create({ paragraphs: ["Source reference native check", "Second review paragraph"] });
document.addTable({ values: [["Gate", "Status"], ["Native", "Pass"]], name: "source-reference-table" });
const commentsPartPath = "word/review/comments-agent.xml";
const commentsPartXml = '<?xml version="1.0" encoding="UTF-8"?><w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:comment w:id="7" w:author="QA Agent" w:initials="QA"><w:p><w:r><w:t>Review the opening paragraph.</w:t></w:r></w:p></w:comment><w:comment w:id="8" w:author="Maintainer" w:initials="MT"><w:p><w:r><w:t>Confirm the native table cell.</w:t></w:r></w:p></w:comment></w:comments>';
const baseDocx = await DocumentFile.exportDocx(document);
const docx = await DocumentFile.patchDocx(baseDocx, [
  {
    path: "word/headerNative.xml",
    xml: '<?xml version="1.0" encoding="UTF-8"?><w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:p><w:r><w:t>Native header</w:t></w:r></w:p></w:hdr>',
    recipe: { kind: "header", source: "word/document.xml", sourceReference: { type: "first" } },
  },
  {
    path: commentsPartPath,
    xml: commentsPartXml,
    recipe: {
      kind: "comments",
      source: "word/document.xml",
      id: "rIdAgentComments",
      sourceReference: {
        anchors: [
          { commentId: 7, target: { type: "block", index: 0 } },
          { commentId: 8, target: { type: "tableCell", tableIndex: 0, rowIndex: 1, columnIndex: 1 } },
        ],
      },
    },
  },
]);
assert.equal(docx.metadata.sourceReferencesUpdated, 2);
assert.equal((await DocumentFile.inspectDocx(docx)).ok, true);
const importedPatchedDocument = await DocumentFile.importDocx(docx, { preferNative: true });
const importedNativeHeader = importedPatchedDocument.headers.find((item) => item.text === "Native header");
assert.ok(importedNativeHeader);
assert.equal(importedNativeHeader.referenceType, "first");
assert.equal(importedNativeHeader.partPath, "word/headerNative.xml");
const importedParagraphComment = importedPatchedDocument.comments.find((comment) => comment.text === "Review the opening paragraph.");
const importedTableComment = importedPatchedDocument.comments.find((comment) => comment.text === "Confirm the native table cell.");
assert.equal(importedParagraphComment?.author, "QA Agent");
assert.equal(importedPatchedDocument.resolve(importedParagraphComment?.targetId)?.text, "Source reference native check");
assert.equal(importedTableComment?.author, "Maintainer");
assert.equal(importedPatchedDocument.resolve(importedTableComment?.targetId)?.kind, "table");
const docxZip = await JSZip.loadAsync(new Uint8Array(await docx.arrayBuffer()));
assert.match(await docxZip.file("word/document.xml").async("text"), /<w:headerReference\b[^>]*w:type="first"[^>]*\/>[\s\S]*?<w:titlePg\/>/);
const anchoredDocumentXml = await docxZip.file("word/document.xml").async("text");
assert.match(anchoredDocumentXml, /<w:commentRangeStart w:id="7"\/>[\s\S]*?<w:commentRangeEnd w:id="7"\/>[\s\S]*?<w:commentReference w:id="7"\/>/);
assert.match(anchoredDocumentXml, /<w:tc>[\s\S]*?<w:t>Pass<\/w:t>[\s\S]*?<w:commentRangeEnd w:id="8"\/>[\s\S]*?<w:commentReference w:id="8"\/>[\s\S]*?<\/w:tc>/);
assert.match(await docxZip.file("word/_rels/document.xml.rels").async("text"), /Id="rIdAgentComments"[^>]*relationships\/comments[^>]*Target="review\/comments-agent\.xml"/);
await assert.rejects(() => DocumentFile.patchDocx(docx, [{ path: "word/review/duplicate-comments.xml", xml: commentsPartXml, recipe: { kind: "comments", source: "word/document.xml", sourceReference: { commentId: 7, target: { type: "block", index: 1 } } } }]), /commentId 7 already exists/);
await assert.rejects(() => DocumentFile.patchDocx(baseDocx, [{ path: "word/review/missing-id-comments.xml", xml: commentsPartXml, recipe: { kind: "comments", source: "word/document.xml", sourceReference: { target: { type: "block", index: 0 } } } }]), /commentId is required/);
await assert.rejects(() => DocumentFile.patchDocx(baseDocx, [{ path: "word/review/undeclared-id-comments.xml", xml: commentsPartXml, recipe: { kind: "comments", source: "word/document.xml", sourceReference: { commentId: 99, target: { type: "block", index: 0 } } } }]), /commentId 99 is not declared/);
await assert.rejects(() => DocumentFile.patchDocx(baseDocx, [{ path: "word/review/duplicate-part-comments.xml", xml: commentsPartXml.replace('w:id="8"', 'w:id="7"'), recipe: { kind: "comments", source: "word/document.xml", sourceReference: { commentId: 7, target: { type: "block", index: 0 } } } }]), /duplicate commentId 7/);
await assert.rejects(() => DocumentFile.patchDocx(baseDocx, [{ path: "word/review/out-of-range-comments.xml", xml: commentsPartXml, recipe: { kind: "comments", source: "word/document.xml", sourceReference: { commentId: 7, target: { type: "tableCell", tableIndex: 0, rowIndex: 9, columnIndex: 0 } } } }]), /rowIndex 9 is out of range/);
const commentsRemovedDocx = await DocumentFile.patchDocx(docx, [{ path: commentsPartPath, remove: true, recipe: { kind: "comments", source: "word/document.xml", id: "rIdAgentComments", sourceReference: true } }]);
assert.equal((await DocumentFile.inspectDocx(commentsRemovedDocx)).ok, true);
const commentsRemovedZip = await JSZip.loadAsync(new Uint8Array(await commentsRemovedDocx.arrayBuffer()));
assert.doesNotMatch(await commentsRemovedZip.file("word/document.xml").async("text"), /commentRange|commentReference/);
assert.equal((await DocumentFile.importDocx(commentsRemovedDocx, { preferNative: true })).comments.length, 0);

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
const pivotCachePath = "xl/custom/pivots/cache-definition.xml";
const pivotRecordsPath = "xl/custom/pivots/cache-records.xml";
const pivotTablePath = "xl/custom/pivots/review-pivot.xml";
const pivotXlsx = await SpreadsheetFile.patchXlsx(drawingXlsx, [
  {
    path: pivotRecordsPath,
    xml: '<?xml version="1.0" encoding="UTF-8"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="1"><r><s v="Revenue"/><n v="120"/></r></pivotCacheRecords>',
    recipe: { kind: "pivotCacheRecords", source: pivotCachePath, id: "rIdAgentPivotRecords", sourceReference: true },
  },
  {
    path: pivotCachePath,
    xml: '<?xml version="1.0" encoding="UTF-8"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" refreshOnLoad="1" recordCount="1"><cacheSource type="worksheet"><worksheetSource ref="A1:B2" sheet="Main"/></cacheSource><cacheFields count="2"><cacheField name="Metric" numFmtId="0"><sharedItems containsString="1" count="1"><s v="Revenue"/></sharedItems></cacheField><cacheField name="Value" numFmtId="0"><sharedItems containsNumber="1" count="1"><n v="120"/></sharedItems></cacheField></cacheFields></pivotCacheDefinition>',
    relationships: [
      { source: "xl/workbook.xml", id: "rIdAgentPivotCache" },
      { source: pivotTablePath, id: "rIdAgentPivotTableCache" },
    ],
    recipe: { kind: "pivotCacheDefinition", source: "xl/workbook.xml", id: "rIdAgentPivotCache", sourceReference: { cacheId: 17 } },
  },
  {
    path: pivotTablePath,
    xml: '<?xml version="1.0" encoding="UTF-8"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="AgentReviewPivot" cacheId="17" dataCaption="Values" updatedVersion="7" minRefreshableVersion="3"><location ref="D1:E2" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="2"><pivotField axis="axisRow" showAll="0"><items count="1"><item t="default"/></items></pivotField><pivotField dataField="1" showAll="0"><items count="1"><item t="default"/></items></pivotField></pivotFields><rowFields count="1"><field x="0"/></rowFields><dataFields count="1"><dataField name="Total value" fld="1" subtotal="sum"/></dataFields></pivotTableDefinition>',
    recipe: { kind: "pivotTable", source: "xl/worksheets/sheet1.xml", id: "rIdAgentPivotTable" },
  },
]);
assert.equal(pivotXlsx.metadata.sourceReferencesUpdated, 2);
assert.equal((await SpreadsheetFile.inspectXlsx(pivotXlsx)).ok, true);
const pivotXlsxZip = await JSZip.loadAsync(new Uint8Array(await pivotXlsx.arrayBuffer()));
assert.match(await pivotXlsxZip.file("xl/workbook.xml").async("text"), /<pivotCaches><pivotCache cacheId="17" r:id="rIdAgentPivotCache"\/><\/pivotCaches>/);
assert.match(await pivotXlsxZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text"), /Id="rIdAgentPivotTable"[^>]*relationships\/pivotTable[^>]*Target="\.\.\/custom\/pivots\/review-pivot\.xml"/);
assert.match(await pivotXlsxZip.file(pivotCachePath).async("text"), /r:id="rIdAgentPivotRecords"/);
assert.match(await pivotXlsxZip.file("xl/custom/pivots/_rels/review-pivot.xml.rels").async("text"), /Id="rIdAgentPivotTableCache"[^>]*relationships\/pivotCacheDefinition[^>]*Target="cache-definition\.xml"/);
const importedPivotWorkbook = await SpreadsheetFile.importXlsx(pivotXlsx);
const importedPivot = importedPivotWorkbook.resolve("AgentReviewPivot");
assert.ok(importedPivot);
assert.deepEqual(importedPivot.computedValues(), [["Metric", "Total value"], ["Revenue", 120]]);
assert.match(importedPivotWorkbook.inspect({ kind: "pivotTable", target: "AgentReviewPivot", maxChars: 4000 }).ndjson, /Total value/);
assert.ok(importedPivotWorkbook.worksheets.getItem("Main").layoutJson().pivots.some((pivot) => pivot.name === "AgentReviewPivot"));
assert.match(importedPivotWorkbook.worksheets.getItem("Main").toSvg(), /AgentReviewPivot/);
assert.ok(importedPivotWorkbook.verify().issues.every((issue) => !["pivotSourceInvalid", "pivotFieldMissing"].includes(issue.type)));
await assert.rejects(() => SpreadsheetFile.patchXlsx(pivotXlsx, [{ path: "xl/custom/pivots/duplicate-cache.xml", xml: '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "pivotCacheDefinition", source: "xl/workbook.xml", sourceReference: { cacheId: 17 } } }]), /cacheId 17 already exists/);
await assert.rejects(() => SpreadsheetFile.patchXlsx(pivotXlsx, [{ path: "xl/custom/pivots/missing-cache-id.xml", xml: '<pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"/>', recipe: { kind: "pivotCacheDefinition", source: "xl/workbook.xml", sourceReference: true } }]), /cacheId must be a non-negative integer/);
const pivotRecordsRemoved = await SpreadsheetFile.patchXlsx(pivotXlsx, [{ path: pivotRecordsPath, remove: true, recipe: { kind: "pivotCacheRecords", source: pivotCachePath, id: "rIdAgentPivotRecords", sourceReference: true } }]);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await pivotRecordsRemoved.arrayBuffer()))).file(pivotCachePath).async("text"), /r:id=/);
const pivotRemoved = await SpreadsheetFile.patchXlsx(pivotRecordsRemoved, [
  { path: pivotTablePath, remove: true, recipe: { kind: "pivotTable", source: "xl/worksheets/sheet1.xml", id: "rIdAgentPivotTable" } },
  { path: pivotCachePath, remove: true, recipe: { kind: "pivotCacheDefinition", source: "xl/workbook.xml", id: "rIdAgentPivotCache", sourceReference: true } },
]);
assert.equal((await SpreadsheetFile.inspectXlsx(pivotRemoved)).ok, true);
const pivotRemovedZip = await JSZip.loadAsync(new Uint8Array(await pivotRemoved.arrayBuffer()));
assert.doesNotMatch(await pivotRemovedZip.file("xl/workbook.xml").async("text"), /pivotCaches/);
assert.doesNotMatch(await pivotRemovedZip.file("xl/worksheets/_rels/sheet1.xml.rels").async("text"), /relationships\/pivotTable/);
assert.equal((await SpreadsheetFile.importXlsx(pivotRemoved)).worksheets.getItem("Main").pivotTables.items.length, 0);

const presentation = Presentation.create();
presentation.slides.add().shapes.add({ text: "Source reference native check", position: { left: 40, top: 40, width: 400, height: 80 } });
const basePptx = await PresentationFile.exportPptx(presentation);
const masterPartPath = "ppt/custom/masters/agent-master.xml";
const layoutPartPath = "ppt/custom/layouts/agent-layout.xml";
const unusedLayoutPartPath = "ppt/custom/layouts/unused-layout.xml";
const masterPartXml = '<?xml version="1.0" encoding="UTF-8"?><p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/></p:spTree></p:cSld><p:txStyles/></p:sldMaster>';
const layoutPartXml = '<?xml version="1.0" encoding="UTF-8"?><p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" type="title" preserve="1"><p:cSld name="Agent Review Layout"><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/><p:sp><p:nvSpPr><p:cNvPr id="2" name="Agent review title"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="381000" y="381000"/><a:ext cx="6096000" cy="762000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Agent review</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>';
const masterLayoutPptx = await PresentationFile.patchPptx(basePptx, [
  {
    path: masterPartPath,
    xml: masterPartXml,
    recipe: { kind: "slideMaster", source: "ppt/presentation.xml", id: "rIdAgentMaster", sourceReference: { masterId: 2_147_483_648 } },
  },
  {
    path: layoutPartPath,
    xml: layoutPartXml,
    relationships: [
      { source: masterPartPath, id: "rIdAgentLayout" },
      { source: layoutPartPath, id: "rIdAgentLayoutMaster", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "../masters/agent-master.xml" },
      { source: "ppt/slides/slide1.xml", id: "rIdAgentSlideLayout", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout", target: "../custom/layouts/agent-layout.xml" },
    ],
    recipe: { kind: "slideLayout", source: masterPartPath, id: "rIdAgentLayout", sourceReference: { layoutId: 2_147_483_649 } },
  },
  {
    path: unusedLayoutPartPath,
    xml: layoutPartXml.replace("Agent Review Layout", "Unused Review Layout"),
    relationships: [
      { source: masterPartPath, id: "rIdUnusedLayout" },
      { source: unusedLayoutPartPath, id: "rIdUnusedLayoutMaster", type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster", target: "../masters/agent-master.xml" },
    ],
    recipe: { kind: "slideLayout", source: masterPartPath, id: "rIdUnusedLayout", sourceReference: { layoutId: 2_147_483_650 } },
  },
]);
assert.equal(masterLayoutPptx.metadata.sourceReferencesUpdated, 3);
assert.equal((await PresentationFile.inspectPptx(masterLayoutPptx)).ok, true);
const masterLayoutZip = await JSZip.loadAsync(new Uint8Array(await masterLayoutPptx.arrayBuffer()));
assert.match(await masterLayoutZip.file("ppt/presentation.xml").async("text"), /<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rIdAgentMaster"\/><\/p:sldMasterIdLst>/);
assert.match(await masterLayoutZip.file(masterPartPath).async("text"), /<p:sldLayoutIdLst><p:sldLayoutId id="2147483649" r:id="rIdAgentLayout"\/><p:sldLayoutId id="2147483650" r:id="rIdUnusedLayout"\/><\/p:sldLayoutIdLst>/);
assert.match(await masterLayoutZip.file("ppt/custom/layouts/_rels/agent-layout.xml.rels").async("text"), /relationships\/slideMaster" Target="\.\.\/masters\/agent-master\.xml"/);
const importedMasterLayout = await PresentationFile.importPptx(masterLayoutPptx);
assert.equal(importedMasterLayout.layouts.items.length, 2);
assert.equal(importedMasterLayout.layouts.items[0].name, "Agent Review Layout");
assert.equal(importedMasterLayout.layouts.items[1].name, "Unused Review Layout");
assert.equal(importedMasterLayout.layouts.items[0].id, "pptx-layout-2147483649");
assert.equal(importedMasterLayout.layouts.items[0].masterId, "pptx-master-2147483648");
assert.equal(importedMasterLayout.slides.items[0].layoutId, importedMasterLayout.layouts.items[0].id);
assert.equal(importedMasterLayout.layouts.items[0].placeholders[0].name, "Agent review title");
await assert.rejects(() => PresentationFile.patchPptx(masterLayoutPptx, [{ path: "ppt/custom/masters/duplicate.xml", xml: masterPartXml, recipe: { kind: "slideMaster", source: "ppt/presentation.xml", sourceReference: { masterId: 2_147_483_648 } } }]), /masterId 2147483648 already exists/);
await assert.rejects(() => PresentationFile.patchPptx(masterLayoutPptx, [{ path: "ppt/custom/layouts/duplicate.xml", xml: layoutPartXml, recipe: { kind: "slideLayout", source: masterPartPath, sourceReference: { layoutId: 2_147_483_649 } } }]), /layoutId 2147483649 already exists/);
const layoutRemovedPptx = await PresentationFile.patchPptx(masterLayoutPptx, [{ path: layoutPartPath, remove: true, recipe: { kind: "slideLayout", source: masterPartPath, id: "rIdAgentLayout", sourceReference: true } }]);
const layoutRemovedMasterXml = await (await JSZip.loadAsync(new Uint8Array(await layoutRemovedPptx.arrayBuffer()))).file(masterPartPath).async("text");
assert.doesNotMatch(layoutRemovedMasterXml, /rIdAgentLayout"/);
assert.match(layoutRemovedMasterXml, /rIdUnusedLayout/);
const allLayoutsRemovedPptx = await PresentationFile.patchPptx(layoutRemovedPptx, [{ path: unusedLayoutPartPath, remove: true, recipe: { kind: "slideLayout", source: masterPartPath, id: "rIdUnusedLayout", sourceReference: true } }]);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await allLayoutsRemovedPptx.arrayBuffer()))).file(masterPartPath).async("text"), /sldLayoutId/);
const masterRemovedPptx = await PresentationFile.patchPptx(allLayoutsRemovedPptx, [{ path: masterPartPath, remove: true, recipe: { kind: "slideMaster", source: "ppt/presentation.xml", id: "rIdAgentMaster", sourceReference: true } }]);
assert.equal((await PresentationFile.inspectPptx(masterRemovedPptx)).ok, true);
assert.doesNotMatch(await (await JSZip.loadAsync(new Uint8Array(await masterRemovedPptx.arrayBuffer()))).file("ppt/presentation.xml").async("text"), /sldMasterId/);

const pptx = await PresentationFile.patchPptx(masterLayoutPptx, [{
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
  for (const [artifactKind, blob] of [["document", docx], ["workbook", pivotXlsx], ["presentation", pptx]]) {
    const pdf = await libreOffice({ input: blob, inputType: blob.type, outputType: "application/pdf", format: "pdf", artifactKind });
    assert.equal(pdf.type, "application/pdf");
    assert.ok(pdf.bytes.length > 100);
    const png = await poppler({ input: pdf, inputType: pdf.type, outputType: "image/png", format: "png", artifactKind, pageIndex: 0 });
    assert.equal(png.type, "image/png");
    assert.deepEqual([...png.bytes.slice(0, 8)], [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
  }
}

console.log(`OOXML source reference smoke ok${nativeAvailable ? " (native)" : " (native skipped)"}`);
