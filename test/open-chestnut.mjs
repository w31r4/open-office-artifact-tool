import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, Presentation, PresentationFile, Workbook, SpreadsheetFile } from "../src/index.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";
import { DocumentBlockSchema, DocumentFieldSchema, DocumentHyperlinkSchema, DocumentNumberingSchema, DocumentParagraphSchema, DocumentSourceBindingSchema, DocumentTableCellSchema, DocumentTableSchema, PresentationArtifactSchema, PresentationBackgroundSchema, PresentationLayoutSchema, PresentationLayoutSourceBindingSchema, PresentationMasterSchema, PresentationMasterSourceBindingSchema, PresentationMasterTextStylesSchema, PresentationPlaceholderSchema, PresentationSlideSchema, PresentationTextBodyPropertiesSchema, PresentationTextBodySchema, PresentationTextParagraphSchema, PresentationTextRunSchema } from "../src/generated/open_office/artifact/v1/office_artifact_pb.js";
import {
  OpenChestnutCodecError,
  exportDocxWithOpenChestnut,
  exportPptxWithOpenChestnut,
  exportXlsxWithOpenChestnut,
  importDocxWithOpenChestnut,
  importPptxWithOpenChestnut,
  importXlsxWithOpenChestnut,
  openChestnutStatus,
} from "../src/codecs/open-chestnut.mjs";

const legacyTabWire = toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, {
  tabStops: [{ positionEmu: 120n, alignment: "left" }],
}));
assert.equal(toBinary(DocumentBlockSchema, create(DocumentBlockSchema, { content: { case: "hyperlink", value: { text: "x", target: { case: "externalUri", value: "https://example.test" } } } }))[0], 0x6a, "Document hyperlinks must use additive block field 13.");
assert.equal(toBinary(DocumentBlockSchema, create(DocumentBlockSchema, { content: { case: "field", value: { instruction: "PAGE", display: "1" } } }))[0], 0x72, "Document fields must use additive block field 14.");
assert.equal(toBinary(DocumentSourceBindingSchema, create(DocumentSourceBindingSchema, { residualSha256: "x" }))[0], 0x2a, "Document residual hashes must use additive field 5.");
assert.equal(toBinary(DocumentHyperlinkSchema, create(DocumentHyperlinkSchema, { target: { case: "externalUri", value: "https://example.test" } }))[0], 0x12, "Document external hyperlink targets must use field 2.");
assert.equal(toBinary(DocumentFieldSchema, create(DocumentFieldSchema, { instruction: "PAGE" }))[0], 0x0a, "Document field instructions must use field 1.");
assert.equal(toBinary(DocumentParagraphSchema, create(DocumentParagraphSchema, { numbering: { numberingId: 7 } }))[0], 0x1a, "Document paragraph numbering must use additive field 3.");
assert.equal(toBinary(DocumentNumberingSchema, create(DocumentNumberingSchema, { numberingId: 7 }))[0], 0x08, "Document numbering IDs must use field 1.");
assert.deepEqual([...toBinary(DocumentTableSchema, create(DocumentTableSchema, { gridColumns: 3 }))], [0x10, 0x03], "Document table grid width must use additive field 2.");
assert.deepEqual([...toBinary(DocumentTableCellSchema, create(DocumentTableCellSchema, { columnSpan: 2 }))], [0x10, 0x02], "Document table cell column spans must use field 2.");
assert.equal(legacyTabWire[0], 0x72, "Presentation tab stops must retain repeated-message field 14.");
assert.equal(fromBinary(PresentationTextParagraphSchema, legacyTabWire).tabStops[0].positionEmu, 120n);
assert.deepEqual([...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { noTabStops: true }))], [0x78, 0x01], "Explicit tab deletion must remain additive field 15.");
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { bullet: { case: "pictureBullet", value: { source: { case: "assetId", value: "x" } } } })).slice(0, 2)],
  [0x82, 0x01],
  "Presentation picture bullets must use additive field 16.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { bulletColor: { case: "bulletColorScheme", value: "accent1" } })).slice(0, 2)],
  [0x8a, 0x01],
  "Presentation scheme marker colors must use additive field 17.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { leftMargin: { case: "marginLeftEmu", value: 1n } })).slice(0, 2)],
  [0x90, 0x01],
  "Presentation paragraph left margins must use additive field 18.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { leftMargin: { case: "noMarginLeft", value: true } })).slice(0, 2)],
  [0xa0, 0x01],
  "Explicit paragraph left-margin deletion must use additive field 20.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { lineSpacing: { case: "lineSpacingPoints", value: 18 } })).slice(0, 2)],
  [0xb1, 0x01],
  "Presentation paragraph line spacing must use additive field 22.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { lineSpacing: { case: "noLineSpacing", value: true } })).slice(0, 2)],
  [0xc0, 0x01],
  "Explicit paragraph line-spacing deletion must use additive field 24.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { defaultRunStyle: { case: "defaultRunProperties", value: { bold: true } } })).slice(0, 2)],
  [0xfa, 0x01],
  "Presentation paragraph default-run properties must use additive field 31.",
);
assert.deepEqual(
  [...toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, { defaultRunStyle: { case: "noDefaultRunProperties", value: true } })).slice(0, 3)],
  [0x80, 0x02, 0x01],
  "Explicit paragraph default-run deletion must use additive field 32.",
);
assert.equal(
  toBinary(PresentationTextBodySchema, create(PresentationTextBodySchema, { listStyles: [{ level: 1 }] }))[0],
  0x12,
  "Presentation text-body list styles must use additive field 2.",
);
assert.deepEqual(
  [...toBinary(PresentationTextBodySchema, create(PresentationTextBodySchema, { noListStyles: true }))],
  [0x18, 0x01],
  "Explicit text-body list-style deletion must use additive field 3.",
);
assert.equal(
  toBinary(PresentationTextBodySchema, create(PresentationTextBodySchema, { bodyProperties: { anchor: { case: "verticalAnchor", value: "center" } } }))[0],
  0x22,
  "Presentation text-body properties must use additive field 4.",
);
assert.deepEqual(
  [...toBinary(PresentationTextBodyPropertiesSchema, create(PresentationTextBodyPropertiesSchema, { autoFit: { case: "noAutoFitMode", value: true } }))],
  [0x70, 0x01],
  "Explicit text-body AutoFit deletion must use additive field 14.",
);
assert.equal(
  toBinary(PresentationTextBodyPropertiesSchema, create(PresentationTextBodyPropertiesSchema, { rotation: { case: "rotationAngle60000", value: 900000 } }))[0],
  0x78,
  "Presentation text-body rotation must use additive field 15.",
);
assert.deepEqual(
  [...toBinary(PresentationTextBodyPropertiesSchema, create(PresentationTextBodyPropertiesSchema, { rotation: { case: "noRotation", value: true } }))],
  [0x80, 0x01, 0x01],
  "Explicit text-body rotation deletion must use additive field 16.",
);
assert.deepEqual(
  [...toBinary(PresentationTextBodyPropertiesSchema, create(PresentationTextBodyPropertiesSchema, { uprightText: { case: "noUpright", value: true } }))],
  [0xf0, 0x01, 0x01],
  "Explicit text-body upright deletion must use additive field 30.",
);
assert.equal(toBinary(PresentationTextRunSchema, create(PresentationTextRunSchema, { content: { case: "text", value: "x" } }))[0], 0x0a, "Presentation text must retain field 1.");
assert.equal(
  toBinary(PresentationArtifactSchema, create(PresentationArtifactSchema, { masters: [{ id: "master/1" }] }))[0],
  0x32,
  "Presentation masters must use additive field 6.",
);
assert.equal(
  toBinary(PresentationArtifactSchema, create(PresentationArtifactSchema, { layouts: [{ id: "layout/1" }] }))[0],
  0x3a,
  "Presentation layouts must use additive field 7.",
);
assert.equal(
  toBinary(PresentationSlideSchema, create(PresentationSlideSchema, { layoutId: "layout/1" }))[0],
  0x2a,
  "Presentation slide layout locators must use additive field 5.",
);
assert.equal(
  toBinary(PresentationMasterTextStylesSchema, create(PresentationMasterTextStylesSchema, { deletedOtherLevels: [8] }))[0],
  0x32,
  "Presentation master other-level deletion must use field 6.",
);
assert.equal(toBinary(PresentationMasterSchema, create(PresentationMasterSchema, { background: { color: { case: "colorScheme", value: "accent1" }, kind: { case: "solid", value: true } } }))[0], 0x2a, "Presentation master backgrounds must use additive field 5.");
assert.equal(toBinary(PresentationLayoutSchema, create(PresentationLayoutSchema, { background: { color: { case: "colorRgb", value: "FFFFFF" }, kind: { case: "solid", value: true } } }))[0], 0x32, "Presentation layout backgrounds must use additive field 6.");
assert.equal(toBinary(PresentationMasterSchema, create(PresentationMasterSchema, { placeholders: [{ id: "master/1/placeholder/1" }] }))[0], 0x32, "Presentation master placeholders must use additive field 6.");
assert.equal(toBinary(PresentationLayoutSchema, create(PresentationLayoutSchema, { placeholders: [{ id: "layout/1/placeholder/1" }] }))[0], 0x3a, "Presentation layout placeholders must use additive field 7.");
assert.equal(toBinary(PresentationPlaceholderSchema, create(PresentationPlaceholderSchema, { textBody: { paragraphs: [] } }))[0], 0x2a, "Presentation placeholder text bodies must use field 5.");
assert.equal(toBinary(PresentationMasterSourceBindingSchema, create(PresentationMasterSourceBindingSchema, { backgroundSemanticSha256: "x" }))[0], 0x3a, "Presentation master background hashes must use additive field 7.");
assert.deepEqual([...toBinary(PresentationMasterSourceBindingSchema, create(PresentationMasterSourceBindingSchema, { backgroundEditable: true }))], [0x40, 0x01], "Presentation master background editability must use additive field 8.");
assert.equal(toBinary(PresentationLayoutSourceBindingSchema, create(PresentationLayoutSourceBindingSchema, { backgroundSemanticSha256: "x" }))[0], 0x2a, "Presentation layout background hashes must use additive field 5.");
assert.deepEqual([...toBinary(PresentationLayoutSourceBindingSchema, create(PresentationLayoutSourceBindingSchema, { backgroundEditable: true }))], [0x30, 0x01], "Presentation layout background editability must use additive field 6.");
assert.equal(toBinary(PresentationBackgroundSchema, create(PresentationBackgroundSchema, { color: { case: "colorScheme", value: "accent1" }, kind: { case: "styleReferenceIndex", value: 1001 } }))[0], 0x12, "Presentation background theme colors must retain field 2.");

const workbook = Workbook.create({ dateSystem: "1904" });
const summary = workbook.worksheets.add("Summary");
summary.getRange("A1:B2").values = [["Quarter", 42.5], [true, null]];
summary.getRange("B2").formulas = [["=B1*2"]];
summary.freezePanes.freezeRows(1).freezeColumns(1);
summary.showGridLines = false;
summary.columnDimensions.set(0, { width: 18, bestFit: true });
summary.rowDimensions.set(0, { height: 24 });
summary.mergeCells("A3:B3");
const details = workbook.worksheets.add("Details");
details.getRange("A1:B1").values = [["Status", "ready"]];

const concurrentWorkbook = Workbook.create();
concurrentWorkbook.worksheets.add("Concurrent").getRange("A1").values = [["cached runtime"]];
const [status, exported, concurrentExport] = await Promise.all([
  openChestnutStatus(),
  exportXlsxWithOpenChestnut(workbook),
  exportXlsxWithOpenChestnut(concurrentWorkbook),
]);
assert.deepEqual([...concurrentExport.bytes.slice(0, 2)], [0x50, 0x4b]);
assert.deepEqual([...exported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(exported.type, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
assert.equal(exported.metadata.codec, "open-chestnut");
assert.equal((await SpreadsheetFile.inspectXlsx(exported)).ok, true);

const imported = await importXlsxWithOpenChestnut(exported);
assert.equal(imported.dateSystem, "1904");
assert.equal(imported.worksheets.items.length, 2);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(imported.worksheets.getItem("Summary").freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
assert.equal(imported.worksheets.getItem("Summary").showGridLines, false);
assert.equal(imported.worksheets.getItem("Summary").columnDimensions.get(0).width, 18);
assert.deepEqual(imported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);
assert.match(imported.inspect({ kind: "workbook,sheet,formula" }).ndjson, /"dateSystem":"1904"/);
assert.equal(imported.verify().ok, true);
assert.equal(imported.resolve(imported.worksheets.getItem("Summary").id).name, "Summary");

// Open XML SDK serializes SpreadsheetML with a legal namespace prefix. Keep the
// JavaScript migration oracle able to read the same package while both codecs
// coexist, so cross-codec fixtures compare semantics instead of XML spelling.
const javascriptImported = await SpreadsheetFile.importXlsx(exported);
assert.equal(javascriptImported.dateSystem, "1904");
assert.equal(javascriptImported.worksheets.items.length, 2);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);

const secondExport = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
assert.deepEqual([...secondExport.bytes.slice(0, 2)], [0x50, 0x4b]);

const externalZip = await JSZip.loadAsync(exported.bytes);
const relationshipPath = "xl/_rels/workbook.xml.rels";
const relationships = await externalZip.file(relationshipPath).async("text");
externalZip.file(relationshipPath, relationships.replace("</Relationships>", '<Relationship Id="rIdExternal" Type="urn:open-office-artifact-tool:test" Target="https://example.invalid/data" TargetMode="External"/></Relationships>'));
const externalBytes = await externalZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const opaqueImported = await importXlsxWithOpenChestnut(externalBytes);
opaqueImported.worksheets.getItem("Summary").getRange("B1").values = [[99]];
const preserved = await exportXlsxWithOpenChestnut(opaqueImported, { recalculate: false });
assert.equal(preserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const preservedZip = await JSZip.loadAsync(preserved.bytes);
assert.match(await preservedZip.file(relationshipPath).async("text"), /Id="rIdExternal"[^>]*Target="https:\/\/example\.invalid\/data"/);
const preservedImported = await importXlsxWithOpenChestnut(preserved);
assert.equal(preservedImported.worksheets.getItem("Summary").getRange("B1").values[0][0], 99);

await assert.rejects(
  importXlsxWithOpenChestnut(exported, { limits: { maxInputBytes: 16 } }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "input_budget_exceeded",
);

const styled = Workbook.create();
const styledSheet = styled.worksheets.add("Sheet1");
styledSheet.getRange("A1:B2").values = [["Label", "Value"], ["styled", 1]];
styledSheet.getRange("A1:B1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
styledSheet.tables.add("A1:B2", true, "StyledTable").style = "TableStyleMedium4";
await assert.rejects(
  exportXlsxWithOpenChestnut(styled),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_workbook_features",
);
const styledSource = await SpreadsheetFile.exportXlsx(styled);
const styledImported = await importXlsxWithOpenChestnut(styledSource);
styledImported.worksheets.getItem("Sheet1").getRange("B2").values = [[2]];
const styledPreserved = await exportXlsxWithOpenChestnut(styledImported, { recalculate: false });
const styledRoundTrip = await SpreadsheetFile.importXlsx(styledPreserved);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("B2").values[0][0], 2);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("A1").format.font.bold, true);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").tables.items[0].name, "StyledTable");

const minimalDocument = DocumentModel.create({
  name: "OpenChestnut brief",
  blocks: [
    { kind: "paragraph", text: "Quarterly brief", styleId: "Normal", runs: [{ text: "Quarterly ", style: { bold: true } }, { text: "brief", style: { italic: true } }] },
    { kind: "table", styleId: "TableGrid", values: [["Revenue", "42"], ["Status", "Ready"]] },
  ],
});
const docxExported = await exportDocxWithOpenChestnut(minimalDocument);
assert.deepEqual([...docxExported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(docxExported.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
assert.equal(docxExported.metadata.codec, "open-chestnut");
assert.equal((await DocumentFile.inspectDocx(docxExported)).ok, true);
const docxImported = await importDocxWithOpenChestnut(docxExported);
assert.equal(docxImported.name, "Imported document");
assert.equal(docxImported.blocks[0].text, "Quarterly brief");
assert.equal(docxImported.blocks[0].runs[0].style.bold, true);
assert.deepEqual(docxImported.blocks[1].values, [["Revenue", "42"], ["Status", "Ready"]]);
assert.equal(docxImported.verify().ok, true);
docxImported.blocks[1].values[0][1] = "84";
const tableEditedDocx = await exportDocxWithOpenChestnut(docxImported);
const tableEditedRoundTrip = await DocumentFile.importDocx(tableEditedDocx, { preferNative: true });
assert.equal(tableEditedRoundTrip.blocks[1].values[0][1], "84");
const tableTopologyImported = await importDocxWithOpenChestnut(docxExported);
tableTopologyImported.blocks[1].values[0].pop();
await assert.rejects(
  exportDocxWithOpenChestnut(tableTopologyImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);
const authoredMergedDocument = DocumentModel.create({
  name: "Direct merged table",
  blocks: [{
    kind: "table",
    styleId: "TableGrid",
    gridColumns: 3,
    values: [["Merged owner", "Status"], ["", "Ready"], ["Scope", "Complete"]],
    cells: [
      { row: 0, column: 0, gridColumn: 0, columnSpan: 2, rowSpan: 2, verticalMerge: "restart" },
      { row: 0, column: 1, gridColumn: 2, columnSpan: 1, rowSpan: 1 },
      { row: 1, column: 0, gridColumn: 0, columnSpan: 2, rowSpan: 0, verticalMerge: "continue" },
      { row: 1, column: 1, gridColumn: 2, columnSpan: 1, rowSpan: 1 },
      { row: 2, column: 0, gridColumn: 0, columnSpan: 1, rowSpan: 1 },
      { row: 2, column: 1, gridColumn: 1, columnSpan: 2, rowSpan: 1 },
    ],
  }],
});
assert.equal(authoredMergedDocument.blocks[0].getCell(1, 0).editable, false);
const authoredMergedDocx = await exportDocxWithOpenChestnut(authoredMergedDocument);
const authoredMergedXml = await (await JSZip.loadAsync(authoredMergedDocx.bytes)).file("word/document.xml").async("text");
assert.match(authoredMergedXml, /<w:tblGrid><w:gridCol\s*\/><w:gridCol\s*\/><w:gridCol\s*\/><\/w:tblGrid>/);
assert.match(authoredMergedXml, /<w:gridSpan w:val="2"\s*\/>/);
assert.match(authoredMergedXml, /<w:vMerge w:val="restart"\s*\/>/);
assert.match(authoredMergedXml, /<w:vMerge w:val="continue"\s*\/>/);
const authoredMergedRoundTrip = await importDocxWithOpenChestnut(authoredMergedDocx);
assert.equal(authoredMergedRoundTrip.blocks[0].gridColumns, 3);
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(0, 0).rowSpan, 2);
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(1, 0).verticalMerge, "continue");
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(2, 1).columnSpan, 2);

const invalidMergedAuthoring = DocumentModel.create({
  blocks: [{
    kind: "table",
    gridColumns: 2,
    values: [["Owner"], [""]],
    cells: [
      { row: 0, column: 0, gridColumn: 0, columnSpan: 2, rowSpan: 3, verticalMerge: "restart" },
      { row: 1, column: 0, gridColumn: 0, columnSpan: 2, rowSpan: 0, verticalMerge: "continue" },
    ],
  }],
});
await assert.rejects(
  exportDocxWithOpenChestnut(invalidMergedAuthoring),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_document_table" && /declares rowSpan 3 but spans 2 rows/.test(error.message),
);

const mergedZip = await JSZip.loadAsync(docxExported.bytes);
const mergedDocumentXml = await mergedZip.file("word/document.xml").async("text");
const mergedTableXml = '<w:tbl><w:tblPr><w:tblStyle w:val="TableGrid"/></w:tblPr><w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/></w:tblGrid>' +
  '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Horizontal origin</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Vertical origin</w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr>' +
  '<w:tr><w:tc><w:tcPr><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>Footer</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:vMerge w:val="continue"/></w:tcPr><w:p><w:r><w:t></w:t></w:r></w:p></w:tc></w:tr></w:tbl>';
mergedZip.file("word/document.xml", mergedDocumentXml.replace(/<w:tbl>[\s\S]*?<\/w:tbl>/, mergedTableXml));
const mergedDocx = await mergedZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const mergedImported = await importDocxWithOpenChestnut(mergedDocx);
const mergedTable = mergedImported.blocks[1];
assert.equal(mergedTable.gridColumns, 3);
assert.equal(mergedTable.cells.length, 7);
assert.deepEqual(mergedTable.cells[0], { row: 0, column: 0, gridColumn: 0, columnSpan: 2, rowSpan: 1, verticalMerge: "none", editable: true });
assert.deepEqual(mergedTable.cells[1], { row: 0, column: 1, gridColumn: 2, columnSpan: 1, rowSpan: 3, verticalMerge: "restart", editable: true });
assert.deepEqual(mergedTable.cells[4], { row: 1, column: 2, gridColumn: 2, columnSpan: 1, rowSpan: 0, verticalMerge: "continue", editable: false });
assert.equal(mergedTable.getCell(0, 1).rowSpan, 3);
assert.equal(mergedTable.getCell(1, 2).editable, false);
assert.equal(mergedImported.inspect({ kind: "tableCell", target: mergedTable.getCell(0, 1).id }).ndjson.includes('"rowSpan":3'), true);
mergedTable.values[0][0] = "Edited horizontal origin";
mergedTable.values[0][1] = "Edited vertical origin";
const mergedEditedDocx = await exportDocxWithOpenChestnut(mergedImported);
const mergedEditedXml = await (await JSZip.loadAsync(mergedEditedDocx.bytes)).file("word/document.xml").async("text");
assert.match(mergedEditedXml, /<w:gridSpan w:val="2"\s*\/>/);
assert.match(mergedEditedXml, /<w:vMerge w:val="restart"\s*\/>/);
assert.match(mergedEditedXml, /Edited horizontal origin/);
assert.match(mergedEditedXml, /Edited vertical origin/);
const mergedRoundTrip = await importDocxWithOpenChestnut(mergedEditedDocx);
assert.equal(mergedRoundTrip.blocks[1].getCell(0, 1).rowSpan, 3);
mergedRoundTrip.blocks[1].values[1][2] = "Unsafe continuation";
await assert.rejects(
  exportDocxWithOpenChestnut(mergedRoundTrip),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);
mergedRoundTrip.blocks[1].values[1][2] = "";
mergedRoundTrip.blocks[1].cells[0].columnSpan = 1;
await assert.rejects(
  exportDocxWithOpenChestnut(mergedRoundTrip),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);

const numberedSourceDocument = DocumentModel.create({ blocks: [] });
numberedSourceDocument.addListItem("Preserve numbered source text", {
  listType: "number",
  numberFormat: "upperLetter",
  start: 3,
  levelText: "%1)",
  numberingId: 77,
  styleId: "Normal",
});
const numberedSourceDocx = await DocumentFile.exportDocx(numberedSourceDocument);
const numberedSourceImported = await importDocxWithOpenChestnut(numberedSourceDocx);
const numberedSourceBlock = numberedSourceImported.blocks[0];
assert.equal(numberedSourceBlock.kind, "listItem");
assert.equal(numberedSourceBlock.numberFormat, "upperLetter");
assert.equal(numberedSourceBlock.start, 3);
assert.equal(numberedSourceBlock.levelText, "%1)");
assert.ok(numberedSourceBlock.numberingId > 0);
assert.ok(numberedSourceBlock.abstractNumberingId >= 0);
numberedSourceBlock.text = "Edited numbered source text";
const numberedSourceEdited = await exportDocxWithOpenChestnut(numberedSourceImported);
const numberedSourceRoundTrip = await DocumentFile.importDocx(numberedSourceEdited, { preferNative: true });
assert.equal(numberedSourceRoundTrip.blocks[0].kind, "listItem");
assert.equal(numberedSourceRoundTrip.blocks[0].text, "Edited numbered source text");
assert.equal(numberedSourceRoundTrip.blocks[0].numberFormat, "upperLetter");

const groupedNumberingDocument = DocumentModel.create({ blocks: [] });
groupedNumberingDocument.addListItem("First grouped item", {
  listType: "number", numberFormat: "upperLetter", start: 3, levelText: "%1)", numberingId: 88,
});
groupedNumberingDocument.addListItem("Second grouped item", {
  listType: "number", numberFormat: "upperLetter", start: 3, levelText: "%1)", numberingId: 88,
});
const groupedNumberingSource = await DocumentFile.exportDocx(groupedNumberingDocument);
const groupedNumberingImported = await importDocxWithOpenChestnut(groupedNumberingSource);
for (const block of groupedNumberingImported.blocks) {
  block.numberFormat = "lowerRoman";
  block.start = 5;
  block.levelText = "%1.";
}
groupedNumberingImported.blocks[0].text = "Edited first grouped item";
const groupedNumberingEdited = await exportDocxWithOpenChestnut(groupedNumberingImported);
const groupedNumberingZip = await JSZip.loadAsync(groupedNumberingEdited.bytes);
const groupedNumberingXml = await groupedNumberingZip.file("word/numbering.xml").async("text");
assert.match(groupedNumberingXml, /<w:lvlOverride w:ilvl="0"><w:lvl w:ilvl="0">[\s\S]*?<w:start w:val="5"\s*\/>[\s\S]*?<w:numFmt w:val="lowerRoman"\s*\/>[\s\S]*?<w:lvlText w:val="%1\."\s*\/>/);
const groupedNumberingRoundTrip = await importDocxWithOpenChestnut(groupedNumberingEdited);
assert.equal(groupedNumberingRoundTrip.blocks[0].text, "Edited first grouped item");
assert.equal(groupedNumberingRoundTrip.blocks[1].text, "Second grouped item");
assert.equal(groupedNumberingRoundTrip.blocks.every((block) => block.numberFormat === "lowerRoman" && block.start === 5 && block.levelText === "%1."), true);

const partialNumberingImported = await importDocxWithOpenChestnut(groupedNumberingSource);
partialNumberingImported.blocks[0].start = 9;
await assert.rejects(
  exportDocxWithOpenChestnut(partialNumberingImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit" && /coherently/.test(error.message),
);

const inheritedNumberingZip = await JSZip.loadAsync(numberedSourceDocx.bytes);
const inheritedDocumentXml = await inheritedNumberingZip.file("word/document.xml").async("text");
inheritedNumberingZip.file("word/document.xml", inheritedDocumentXml.replace(
  /<w:pStyle w:val="Normal"\/><w:numPr>[\s\S]*?<\/w:numPr>/,
  '<w:pStyle w:val="DerivedList"/>',
));
const inheritedStylesXml = (await inheritedNumberingZip.file("word/styles.xml").async("text")).replace(
  "</w:styles>",
  '<w:style w:type="paragraph" w:styleId="BaseList"><w:name w:val="Base list"/><w:pPr><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="DerivedList"><w:name w:val="Derived list"/><w:basedOn w:val="BaseList"/><w:pPr><w:numPr><w:ilvl w:val="8"/></w:numPr></w:pPr></w:style></w:styles>',
);
inheritedNumberingZip.file("word/styles.xml", inheritedStylesXml);
const inheritedNumberingXml = (await inheritedNumberingZip.file("word/numbering.xml").async("text")).replace(
  /(<w:lvl w:ilvl="0">[\s\S]*?<w:numFmt\b[^>]*\/>)/,
  '$1<w:pStyle w:val="DerivedList"/>',
);
inheritedNumberingZip.file("word/numbering.xml", inheritedNumberingXml);
const inheritedNumberingSource = await inheritedNumberingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const inheritedNumberingImported = await importDocxWithOpenChestnut(inheritedNumberingSource);
assert.equal(inheritedNumberingImported.blocks[0].kind, "listItem");
assert.equal(inheritedNumberingImported.blocks[0].styleId, "DerivedList");
assert.equal(inheritedNumberingImported.blocks[0].numberFormat, "upperLetter");
inheritedNumberingImported.blocks[0].text = "Edited inherited numbering text";
const inheritedNumberingEdited = await exportDocxWithOpenChestnut(inheritedNumberingImported);
const inheritedNumberingEditedZip = await JSZip.loadAsync(inheritedNumberingEdited.bytes);
assert.equal(await inheritedNumberingEditedZip.file("word/styles.xml").async("text"), inheritedStylesXml);
assert.equal(await inheritedNumberingEditedZip.file("word/numbering.xml").async("text"), inheritedNumberingXml);
const inheritedNumberingRoundTrip = await DocumentFile.importDocx(inheritedNumberingEdited, { preferNative: true });
assert.equal(inheritedNumberingRoundTrip.blocks[0].kind, "listItem");
assert.equal(inheritedNumberingRoundTrip.blocks[0].text, "Edited inherited numbering text");
assert.equal(inheritedNumberingRoundTrip.blocks[0].styleId, "DerivedList");
inheritedNumberingImported.blocks[0].start = 12;
await assert.rejects(
  exportDocxWithOpenChestnut(inheritedNumberingImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);

const styleLinkedNumberingZip = await JSZip.loadAsync(numberedSourceDocx.bytes);
const styleLinkedDocumentXml = (await styleLinkedNumberingZip.file("word/document.xml").async("text")).replace(
  /<w:pStyle w:val="Normal"\/><w:numPr>[\s\S]*?<\/w:numPr>/,
  '<w:pStyle w:val="AgentListDerived"/>',
);
const styleLinkedStylesXml = (await styleLinkedNumberingZip.file("word/styles.xml").async("text")).replace(
  "</w:styles>",
  '<w:style w:type="paragraph" w:styleId="AgentListBase"><w:name w:val="Agent list base"/><w:pPr><w:numPr><w:numId w:val="6"/></w:numPr></w:pPr></w:style><w:style w:type="paragraph" w:styleId="AgentListDerived"><w:name w:val="Agent list derived"/><w:basedOn w:val="AgentListBase"/><w:pPr><w:numPr><w:ilvl w:val="8"/></w:numPr></w:pPr></w:style><w:style w:type="numbering" w:styleId="AgentNumbering"><w:name w:val="Agent numbering"/><w:pPr><w:numPr><w:numId w:val="4"/></w:numPr></w:pPr></w:style></w:styles>',
);
const styleLinkedNumberingXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="0"><w:multiLevelType w:val="multilevel"/><w:numStyleLink w:val="AgentNumbering"/></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:multiLevelType w:val="multilevel"/><w:styleLink w:val="AgentNumbering"/><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:pStyle w:val="AgentListBase"/><w:lvlText w:val="%1."/></w:lvl><w:lvl w:ilvl="2"><w:start w:val="3"/><w:numFmt w:val="upperRoman"/><w:pStyle w:val="AgentListDerived"/><w:lvlText w:val="%1.%2.%3"/></w:lvl></w:abstractNum><w:num w:numId="4"><w:abstractNumId w:val="2"/></w:num><w:num w:numId="6"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="2"><w:startOverride w:val="9"/></w:lvlOverride></w:num></w:numbering>';
styleLinkedNumberingZip.file("word/document.xml", styleLinkedDocumentXml);
styleLinkedNumberingZip.file("word/styles.xml", styleLinkedStylesXml);
styleLinkedNumberingZip.file("word/numbering.xml", styleLinkedNumberingXml);
const styleLinkedNumberingSource = await styleLinkedNumberingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const styleLinkedNumberingImported = await importDocxWithOpenChestnut(styleLinkedNumberingSource);
const styleLinkedBlock = styleLinkedNumberingImported.blocks[0];
assert.equal(styleLinkedBlock.kind, "listItem");
assert.equal(styleLinkedBlock.styleId, "AgentListDerived");
assert.equal(styleLinkedBlock.level, 2);
assert.equal(styleLinkedBlock.numberFormat, "upperRoman");
assert.equal(styleLinkedBlock.start, 9);
assert.equal(styleLinkedBlock.levelText, "%1.%2.%3");
assert.equal(styleLinkedBlock.numberingId, 6);
assert.equal(styleLinkedBlock.abstractNumberingId, 0);
assert.equal(styleLinkedBlock.numberingStyleId, "AgentNumbering");
styleLinkedBlock.text = "Edited numbering-style link text";
const styleLinkedNumberingEdited = await exportDocxWithOpenChestnut(styleLinkedNumberingImported);
const styleLinkedNumberingEditedZip = await JSZip.loadAsync(styleLinkedNumberingEdited.bytes);
assert.equal(await styleLinkedNumberingEditedZip.file("word/styles.xml").async("text"), styleLinkedStylesXml);
assert.equal(await styleLinkedNumberingEditedZip.file("word/numbering.xml").async("text"), styleLinkedNumberingXml);
const styleLinkedNumberingRoundTrip = await DocumentFile.importDocx(styleLinkedNumberingEdited, { preferNative: true });
assert.equal(styleLinkedNumberingRoundTrip.blocks[0].text, "Edited numbering-style link text");
assert.equal(styleLinkedNumberingRoundTrip.blocks[0].numberingStyleId, "AgentNumbering");
styleLinkedBlock.numberingStyleId = "UnsafeReplacement";
await assert.rejects(
  exportDocxWithOpenChestnut(styleLinkedNumberingImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);

numberedSourceBlock.level = 1;
await assert.rejects(
  exportDocxWithOpenChestnut(numberedSourceImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit",
);

const richDocument = DocumentModel.create({ name: "Source preservation", blocks: [] });
richDocument.addParagraph("Editable lead", { styleId: "Normal" });
richDocument.addHyperlink("Preserved link", "https://example.invalid/source");
const richBookmarkTarget = richDocument.addParagraph("Bookmark target", { styleId: "Normal" });
richDocument.addBookmark(richBookmarkTarget, "LeadTarget");
richDocument.addHeader("Preserved header", { sectionIndex: 0, referenceType: "default" });
const richSource = await DocumentFile.exportDocx(richDocument);
const richImported = await importDocxWithOpenChestnut(richSource);
assert.equal(richImported.blocks[0].text, "Editable lead");
assert.equal(richImported.blocks[1].text, "Preserved link");
assert.equal(richImported.blocks[1].kind, "hyperlink");
assert.equal(richImported.blocks[1].url, "https://example.invalid/source");
richImported.blocks[0].text = "Edited lead";
richImported.blocks[0].runs[0].text = "Edited lead";
richImported.blocks[1].text = "Edited link";
richImported.blocks[1].url = "https://example.invalid/updated";
richImported.blocks[1].tooltip = "Updated by the OpenChestnut codec";
richImported.blocks[1].history = false;
const richPreserved = await exportDocxWithOpenChestnut(richImported);
assert.equal(richPreserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const richRoundTrip = await DocumentFile.importDocx(richPreserved, { preferNative: true });
assert.equal(richRoundTrip.blocks[0].text, "Edited lead");
assert.equal(richRoundTrip.blocks.some((block) => block.kind === "hyperlink" && block.text === "Edited link" && block.url === "https://example.invalid/updated" && block.tooltip === "Updated by the OpenChestnut codec" && block.history === false), true);
assert.equal(richRoundTrip.headers.some((header) => header.text === "Preserved header"), true);
const internalRichImported = await importDocxWithOpenChestnut(richSource);
internalRichImported.blocks[1].text = "Jump to lead";
internalRichImported.blocks[1].anchor = "LeadTarget";
internalRichImported.blocks[1].url = "";
const internalRichPreserved = await exportDocxWithOpenChestnut(internalRichImported);
const internalRichRoundTrip = await DocumentFile.importDocx(internalRichPreserved, { preferNative: true });
assert.equal(internalRichRoundTrip.blocks.some((block) => block.kind === "hyperlink" && block.text === "Jump to lead" && block.anchor === "LeadTarget"), true);

const directHyperlinkDocument = DocumentModel.create({ blocks: [] });
directHyperlinkDocument.addHyperlink("Direct hyperlink", "https://example.invalid/direct", { tooltip: "Direct C# authoring" });
const directHyperlinkDocx = await exportDocxWithOpenChestnut(directHyperlinkDocument);
const directHyperlinkRoundTrip = await importDocxWithOpenChestnut(directHyperlinkDocx);
assert.equal(directHyperlinkRoundTrip.blocks[0].kind, "hyperlink");
assert.equal(directHyperlinkRoundTrip.blocks[0].url, "https://example.invalid/direct");
directHyperlinkDocument.blocks[0].url = "javascript:alert(1)";
await assert.rejects(
  exportDocxWithOpenChestnut(directHyperlinkDocument),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_document_hyperlink",
);

const directFieldDocument = DocumentModel.create({ blocks: [] });
directFieldDocument.addField("PAGE", "1", { styleId: "Normal" });
const directFieldDocx = await exportDocxWithOpenChestnut(directFieldDocument);
const directFieldRoundTrip = await importDocxWithOpenChestnut(directFieldDocx);
assert.equal(directFieldRoundTrip.blocks[0].kind, "field");
assert.equal(directFieldRoundTrip.blocks[0].instruction, "PAGE");
assert.equal(directFieldRoundTrip.blocks[0].display, "1");

const sourceFieldDocument = DocumentModel.create({ blocks: [] });
sourceFieldDocument.addField("PAGE", "1", { styleId: "Normal" });
const sourceFieldDocx = await DocumentFile.exportDocx(sourceFieldDocument);
const sourceFieldImported = await importDocxWithOpenChestnut(sourceFieldDocx);
assert.equal(sourceFieldImported.blocks[0].kind, "field");
sourceFieldImported.blocks[0].instruction = "NUMPAGES \\* MERGEFORMAT";
sourceFieldImported.blocks[0].display = "2";
const sourceFieldEdited = await exportDocxWithOpenChestnut(sourceFieldImported);
const sourceFieldRoundTrip = await DocumentFile.importDocx(sourceFieldEdited, { preferNative: true });
assert.equal(sourceFieldRoundTrip.blocks[0].kind, "field");
assert.equal(sourceFieldRoundTrip.blocks[0].instruction, "NUMPAGES \\* MERGEFORMAT");
assert.equal(sourceFieldRoundTrip.blocks[0].display, "2");

directFieldDocument.blocks[0].instruction = "DDEAUTO command";
await assert.rejects(
  exportDocxWithOpenChestnut(directFieldDocument),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_document_field",
);

const minimalPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
minimalPresentation.slides.add({ name: "Overview" }).shapes.add({
  name: "Title",
  geometry: "rect",
  position: { left: 60, top: 40, width: 860, height: 70 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: "OpenChestnut presentation",
});
const pptxExported = await exportPptxWithOpenChestnut(minimalPresentation);
assert.deepEqual([...pptxExported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(pptxExported.type, "application/vnd.openxmlformats-officedocument.presentationml.presentation");
assert.equal(pptxExported.metadata.codec, "open-chestnut");
assert.equal((await PresentationFile.inspectPptx(pptxExported)).ok, true);
const pptxImported = await importPptxWithOpenChestnut(pptxExported);
assert.equal(pptxImported.slides.count, 1);
assert.equal(pptxImported.slides.getItem(0).shapes.items[0].text.value, "OpenChestnut presentation");
assert.equal(pptxImported.verify().ok, true);

const masterStylePresentation = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  master: {
    id: "master/authored",
    name: "Authored Master",
    background: { fill: "accent1", mode: "reference", index: 1001 },
    textParagraphStyles: {
      title: { 0: { alignment: "center", style: { bold: true, fontSize: 40, fontFamily: "Aptos Display", color: "accent1" } } },
      body: { 1: { marginLeft: 72, indent: -24, bulletCharacter: "•", style: { fontSize: 24 } } },
    },
  },
});
masterStylePresentation.slides.add({ name: "Master styles" }).shapes.add({
  name: "Body",
  position: { left: 60, top: 80, width: 860, height: 120 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: "Master style evidence",
});
const masterStyleAuthored = await exportPptxWithOpenChestnut(masterStylePresentation);
const masterStyleSourceZip = await JSZip.loadAsync(masterStyleAuthored.bytes);
const masterPartPath = "ppt/slideMasters/slideMaster1.xml";
const masterStyleAuthoredXml = await masterStyleSourceZip.file(masterPartPath).async("text");
assert.match(masterStyleAuthoredXml, /<p:bg><p:bgRef idx="1001"><a:schemeClr\b[^>]*val="accent1"[^>]*\/><\/p:bgRef><\/p:bg>/);
assert.match(masterStyleAuthoredXml, /<p:titleStyle>[\s\S]*?<a:lvl1pPr[^>]*algn="ctr"[^>]*>[\s\S]*?<a:defRPr[^>]*sz="3000"[^>]*b="1">[\s\S]*?<a:schemeClr val="accent1"\s*\/>[\s\S]*?<a:latin typeface="Aptos Display"\s*\/>/);
assert.match(masterStyleAuthoredXml, /<p:bodyStyle>[\s\S]*?<a:lvl2pPr[^>]*marL="685800"[^>]*indent="-228600"[^>]*>[\s\S]*?<a:buChar char="•"\s*\/>/);
const masterPlaceholderXml = '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="2" name="Master Prompt"/><p:cNvSpPr/><p:nvPr><p:ph type="title" idx="0" hasCustomPrompt="1"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm rot="60000"><a:off x="762000" y="571500"/><a:ext cx="6858000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="ctr"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Master prompt</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
masterStyleSourceZip.file(masterPartPath, masterStyleAuthoredXml
  .replace(/<a:lvl1pPr\b/, '<a:lvl1pPr marR="123456"')
  .replace("</p:spTree>", `${masterPlaceholderXml}</p:spTree>`));
const layoutPartPath = "ppt/slideLayouts/slideLayout1.xml";
const masterStyleLayoutXml = await masterStyleSourceZip.file(layoutPartPath).async("text");
const layoutPlaceholderXml = '<p:sp xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><p:nvSpPr><p:cNvPr id="2" name="Layout Prompt"/><p:cNvSpPr/><p:nvPr><p:ph type="body" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="762000" y="1905000"/><a:ext cx="6858000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr anchor="t"/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>Layout prompt</a:t></a:r><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>';
masterStyleSourceZip.file(layoutPartPath, masterStyleLayoutXml
  .replace(/(<p:cSld\b[^>]*>)/, '$1<p:bg><p:bgPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="FFF7ED"/></a:solidFill><a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:bgPr></p:bg>')
  .replace("</p:spTree>", `${layoutPlaceholderXml}</p:spTree>`));
const masterStyleSource = await masterStyleSourceZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const masterStyleImported = await importPptxWithOpenChestnut(masterStyleSource);
assert.equal(masterStyleImported.masters.count, 1);
assert.equal(masterStyleImported.layouts.items.length, 1);
assert.equal(masterStyleImported.master.id, "presentation/master/1");
assert.equal(masterStyleImported.master.name, "Authored Master");
assert.deepEqual(masterStyleImported.master.background, { fill: "accent1", mode: "reference", index: 1001 });
assert.equal(masterStyleImported.master.textParagraphStyles.title[0].alignment, "center");
assert.equal(masterStyleImported.master.textParagraphStyles.title[0].style.color, "accent1");
assert.equal(masterStyleImported.master.textParagraphStyles.body[1].bulletCharacter, "•");
assert.equal(masterStyleImported.layouts.items[0].masterId, masterStyleImported.master.id);
assert.equal(masterStyleImported.layouts.items[0].type, "blank");
assert.deepEqual(masterStyleImported.layouts.items[0].background, { fill: "#fff7ed", mode: "solid" });
assert.equal(masterStyleImported.slides.getItem(0).layoutId, masterStyleImported.layouts.items[0].id);
const importedMasterPlaceholder = masterStyleImported.master.placeholders[0];
const importedLayoutPlaceholder = masterStyleImported.layouts.items[0].placeholders[0];
assert.equal(importedMasterPlaceholder.type, "title");
assert.equal(importedMasterPlaceholder.idx, 0);
assert.equal(importedMasterPlaceholder.text[0].runs[0].text, "Master prompt");
assert.equal(importedMasterPlaceholder.textBodyProperties.anchor, "center");
assert.equal(importedLayoutPlaceholder.type, "body");
assert.equal(importedLayoutPlaceholder.idx, 2);
assert.equal(importedLayoutPlaceholder.text[0].runs[0].text, "Layout prompt");
importedMasterPlaceholder.text[0].runs[0].text = "Edited master prompt";
importedLayoutPlaceholder.text[0].runs[0] = { text: "Edited layout prompt", style: {}, link: { uri: "https://example.com/layout-help", tooltip: "Layout help" } };
importedLayoutPlaceholder.textBodyProperties.anchor = "bottom";
masterStyleImported.master.textParagraphStyles.title[0].alignment = "right";
delete masterStyleImported.master.textParagraphStyles.body[1];
masterStyleImported.master.textParagraphStyles.other[2] = {
  level: 2,
  bulletImage: { uri: "https://example.com/master-marker.png", relationshipMode: "link" },
  bulletColor: "accent3",
  style: { italic: true, fontSize: 20 },
};
masterStyleImported.master.setBackground("#112233");
masterStyleImported.layouts.items[0].background = { fill: "accent2", mode: "reference", index: 1002 };
const masterStyleEdited = await exportPptxWithOpenChestnut(masterStyleImported);
const masterStyleEditedZip = await JSZip.loadAsync(masterStyleEdited.bytes);
const masterStyleEditedXml = await masterStyleEditedZip.file(masterPartPath).async("text");
assert.match(masterStyleEditedXml, /<a:lvl1pPr[^>]*marR="123456"[^>]*algn="r"[^>]*>/);
assert.match(masterStyleEditedXml, /<p:bg><p:bgPr><a:solidFill><a:srgbClr\b[^>]*val="112233"[^>]*\/><\/a:solidFill><a:effectLst\s*\/><\/p:bgPr><\/p:bg>/);
assert.doesNotMatch(masterStyleEditedXml, /<p:bodyStyle>[\s\S]*?<a:lvl2pPr/);
assert.match(masterStyleEditedXml, /<p:otherStyle>[\s\S]*?<a:lvl3pPr>[\s\S]*?<a:buClr><a:schemeClr val="accent3"\s*\/><\/a:buClr>[\s\S]*?<a:buBlip><a:blip r:link="[^"]+"\s*\/><\/a:buBlip>/);
const masterRelationships = await masterStyleEditedZip.file("ppt/slideMasters/_rels/slideMaster1.xml.rels").async("text");
assert.match(masterRelationships, /Type="[^"]+\/image" Target="https:\/\/example\.com\/master-marker\.png" TargetMode="External"/);
const masterPlaceholderEditedXml = await masterStyleEditedZip.file(masterPartPath).async("text");
assert.match(masterPlaceholderEditedXml, /<p:ph\b[^>]*type="title"[^>]*idx="0"[^>]*hasCustomPrompt="1"/);
assert.match(masterPlaceholderEditedXml, /<a:xfrm\b[^>]*rot="60000"/);
assert.match(masterPlaceholderEditedXml, /<a:t>Edited master prompt<\/a:t>/);
const layoutEditedXml = await masterStyleEditedZip.file(layoutPartPath).async("text");
assert.match(layoutEditedXml, /<p:bg><p:bgRef idx="1002"><a:schemeClr\b[^>]*val="accent2"[^>]*\/><\/p:bgRef><\/p:bg>/);
assert.match(layoutEditedXml, /<a:bodyPr\b[^>]*anchor="b"/);
assert.match(layoutEditedXml, /<a:t>Edited layout prompt<\/a:t>/);
assert.match(await masterStyleEditedZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /Type="[^"]+\/hyperlink" Target="https:\/\/example\.com\/layout-help" TargetMode="External"/);
const masterStyleRoundTrip = await importPptxWithOpenChestnut(masterStyleEdited);
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.title[0].alignment, "right");
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.body[1], undefined);
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.other[2].bulletImage.uri, "https://example.com/master-marker.png");
assert.deepEqual(masterStyleRoundTrip.master.background, { fill: "#112233", mode: "solid" });
assert.deepEqual(masterStyleRoundTrip.layouts.items[0].background, { fill: "accent2", mode: "reference", index: 1002 });
assert.equal(masterStyleRoundTrip.master.placeholders[0].text[0].runs[0].text, "Edited master prompt");
assert.equal(masterStyleRoundTrip.layouts.items[0].placeholders[0].text[0].runs[0].link.uri, "https://example.com/layout-help");
assert.equal(masterStyleRoundTrip.layouts.items[0].placeholders[0].textBodyProperties.anchor, "bottom");
const retainedPlaceholderName = importedLayoutPlaceholder.name;
importedLayoutPlaceholder.name = "Unsafe placeholder rename";
await assert.rejects(
  exportPptxWithOpenChestnut(masterStyleImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit" && /placeholder/.test(error.message),
);
importedLayoutPlaceholder.name = retainedPlaceholderName;
masterStyleImported.layouts.items[0].placeholders.pop();
await assert.rejects(
  exportPptxWithOpenChestnut(masterStyleImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "presentation_placeholder_topology_changed",
);
masterStyleImported.layouts.items[0].placeholders.push(importedLayoutPlaceholder);
const retainedMasterBackground = masterStyleImported.master.background;
masterStyleImported.master.background = undefined;
await assert.rejects(
  exportPptxWithOpenChestnut(masterStyleImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit" && /remove master/.test(error.message),
);
masterStyleImported.master.background = retainedMasterBackground;
masterStyleImported.slides.getItem(0).layoutId = "presentation/master/1/layout/missing";
await assert.rejects(
  exportPptxWithOpenChestnut(masterStyleImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "presentation_slide_layout_binding_changed",
);
masterStyleImported.slides.getItem(0).layoutId = masterStyleImported.layouts.items[0].id;
masterStyleImported.master.name = "Unsafe rename";
await assert.rejects(
  exportPptxWithOpenChestnut(masterStyleImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);

const unsupportedBackgroundZip = await JSZip.loadAsync(masterStyleAuthored.bytes);
const unsupportedBackgroundLayoutXml = await unsupportedBackgroundZip.file(layoutPartPath).async("text");
unsupportedBackgroundZip.file(layoutPartPath, unsupportedBackgroundLayoutXml.replace(/(<p:cSld\b[^>]*>)/, '$1<p:bg><p:bgPr><a:noFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/><a:effectLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"/></p:bgPr></p:bg>'));
const unsupportedBackgroundImported = await importPptxWithOpenChestnut(await unsupportedBackgroundZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
assert.equal(unsupportedBackgroundImported.layouts.items[0].background, undefined);
unsupportedBackgroundImported.layouts.items[0].background = { fill: "#ffffff", mode: "solid" };
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedBackgroundImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit" && /background is preserved/.test(error.message),
);

const unsupportedPlaceholderZip = await JSZip.loadAsync(masterStyleSource);
const unsupportedPlaceholderLayoutXml = await unsupportedPlaceholderZip.file(layoutPartPath).async("text");
unsupportedPlaceholderZip.file(layoutPartPath, unsupportedPlaceholderLayoutXml.replace("<p:txBody><a:bodyPr", "<p:txBody><a:bodyPr/><a:bodyPr"));
const unsupportedPlaceholderImported = await importPptxWithOpenChestnut(await unsupportedPlaceholderZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
unsupportedPlaceholderImported.layouts.items[0].placeholders[0].text[0].runs[0].text = "Unsafe placeholder text";
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedPlaceholderImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit" && /placeholder/.test(error.message),
);

const multiMasterPresentation = Presentation.create({
  masters: [
    { id: "master/primary", name: "Primary", textParagraphStyles: { title: { 0: { style: { fontSize: 36, color: "accent1" } } } } },
    { id: "master/secondary", name: "Secondary", textParagraphStyles: { title: { 0: { style: { fontSize: 32, color: "accent2" } } } } },
  ],
  layouts: [
    { id: "layout/primary", name: "Primary Blank", type: "blank", masterId: "master/primary" },
    { id: "layout/secondary", name: "Secondary Blank", type: "blank", masterId: "master/secondary" },
  ],
});
for (const [name, layoutId] of [["Primary slide", "layout/primary"], ["Secondary slide", "layout/secondary"]]) {
  multiMasterPresentation.slides.add({ name, layoutId }).shapes.add({
    name,
    position: { left: 60, top: 60, width: 760, height: 80 },
    fill: "#FFFFFF",
    line: { fill: "#334155", width: 1 },
    text: name,
  });
}
const multiMasterSource = await PresentationFile.exportPptx(multiMasterPresentation);
const multiMasterImported = await importPptxWithOpenChestnut(multiMasterSource);
assert.equal(multiMasterImported.masters.count, 2);
assert.equal(multiMasterImported.layouts.items.length, 2);
assert.notEqual(multiMasterImported.slides.items[0].layoutId, multiMasterImported.slides.items[1].layoutId);
assert.equal(multiMasterImported.layouts.getItem(multiMasterImported.slides.items[1].layoutId).masterId, multiMasterImported.masters.items[1].id);
multiMasterImported.masters.items[1].textParagraphStyles.title[0].style.color = "accent4";
const multiMasterEdited = await exportPptxWithOpenChestnut(multiMasterImported);
const multiMasterRoundTrip = await importPptxWithOpenChestnut(multiMasterEdited);
assert.equal(multiMasterRoundTrip.masters.items[1].textParagraphStyles.title[0].style.color, "accent4");
assert.equal(multiMasterRoundTrip.slides.items[1].layoutId, multiMasterRoundTrip.layouts.items[1].id);

const richPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const richShape = richPresentation.slides.add({ name: "Rich text" }).shapes.add({
  name: "Rich text",
  geometry: "rect",
  position: { left: 60, top: 40, width: 920, height: 180 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  textBodyProperties: { insets: { left: 8, top: 4, right: 12, bottom: 6 }, anchor: "center", wrap: "none", autoFit: "shrinkText", rotation: 15, verticalText: "vertical", verticalOverflow: "ellipsis", horizontalOverflow: "clip", columns: { count: 2, spacing: 18, rightToLeft: true }, upright: true },
  text: [
    {
      alignment: "center",
      bulletCharacter: "•",
      bulletFont: "Georgia",
      bulletColor: "accent1",
      bulletSizePercent: 1.5,
      marginLeft: 32,
      indent: -16,
      lineSpacing: 1.25,
      spaceBefore: 12,
      spaceAfterPercent: 0.5,
      style: { bold: false, italic: true, fontSize: 24, fontFamily: "Aptos", color: "accent2" },
      runs: [
        { text: "Quarterly ", style: { bold: true, fontSize: 36, fontFamily: "Aptos Display", color: "#0F172A" } },
        { text: "brief", style: { italic: true, fontSize: 36 } },
      ],
    },
    { level: 1, autoNumber: { type: "romanLcPeriod", startAt: 3 }, bulletFontFollowText: true, bulletColorFollowText: true, bulletSizeFollowText: true, runs: [{ text: "Source-bound detail", style: { fontSize: 20, color: "#475569" } }] },
    { bulletNone: true, bulletSize: 24, runs: [{ text: "Explicitly unbulleted", style: { fontSize: 20 } }] },
  ],
});
richShape.text.inheritedParagraphStyles = {
  0: { level: 0, bulletCharacter: "•", bulletColor: "accent1", marginLeft: 48, indent: -16, spaceAfter: 6, tabStops: [{ position: 180, alignment: "decimal" }], style: { fontSize: 18, color: "tx1" } },
  2: { level: 2, autoNumber: { type: "romanLcPeriod", startAt: 3 }, bulletFontFollowText: true, lineSpacing: 1.25, style: {} },
};
const richPptx = await exportPptxWithOpenChestnut(richPresentation);
const richPptxZip = await JSZip.loadAsync(richPptx.bytes);
const richPptxXml = await richPptxZip.file("ppt/slides/slide1.xml").async("text");
assert.match(richPptxXml, /<a:buClr><a:schemeClr val="accent1"\s*\/><\/a:buClr>/);
assert.match(richPptxXml, /<a:lnSpc><a:spcPct val="125000"\s*\/><\/a:lnSpc>/);
assert.match(richPptxXml, /<a:spcBef><a:spcPts val="900"\s*\/><\/a:spcBef>/);
assert.match(richPptxXml, /<a:spcAft><a:spcPct val="50000"\s*\/><\/a:spcAft>/);
assert.match(richPptxXml, /<a:defRPr[^>]*sz="1800"[^>]*b="0"[^>]*i="1">[\s\S]*?<a:schemeClr val="accent2"\s*\/>[\s\S]*?<a:latin typeface="Aptos"\s*\/>[\s\S]*?<\/a:defRPr>/);
assert.match(richPptxXml, /<a:lstStyle[^>]*>[\s\S]*?<a:lvl1pPr[^>]*marL="457200"[^>]*indent="-152400">[\s\S]*?<a:schemeClr val="accent1"\s*\/>[\s\S]*?<a:tab pos="1714500" algn="dec"\s*\/>[\s\S]*?<a:defRPr[^>]*sz="1350">[\s\S]*?<a:schemeClr val="tx1"\s*\/>[\s\S]*?<\/a:defRPr>[\s\S]*?<\/a:lvl1pPr>/);
assert.match(richPptxXml, /<a:lvl3pPr>[\s\S]*?<a:lnSpc><a:spcPct val="125000"\s*\/><\/a:lnSpc>[\s\S]*?<a:buAutoNum type="romanLcPeriod" startAt="3"\s*\/>[\s\S]*?<\/a:lvl3pPr>/);
assert.match(richPptxXml, /<a:bodyPr[^>]*rot="900000"[^>]*vertOverflow="ellipsis"[^>]*horzOverflow="clip"[^>]*vert="vert"[^>]*wrap="none"[^>]*lIns="76200"[^>]*tIns="38100"[^>]*rIns="114300"[^>]*bIns="57150"[^>]*numCol="2"[^>]*spcCol="171450"[^>]*rtlCol="1"[^>]*anchor="ctr"[^>]*upright="1"[^>]*>\s*<a:normAutofit\s*\/>\s*<\/a:bodyPr>/);
const richPptxImported = await importPptxWithOpenChestnut(richPptx);
const richImportedShape = richPptxImported.slides.getItem(0).shapes.items[0];
assert.equal(richImportedShape.text.value, "Quarterly brief\nSource-bound detail\nExplicitly unbulleted");
assert.equal(richImportedShape.text.paragraphs.length, 3);
assert.equal(richImportedShape.text.paragraphs[0].alignment, "center");
assert.equal(richImportedShape.text.paragraphs[0].bulletCharacter, "•");
assert.equal(richImportedShape.text.paragraphs[0].bulletFont, "Georgia");
assert.equal(richImportedShape.text.paragraphs[0].bulletColor, "accent1");
assert.equal(richImportedShape.text.paragraphs[0].bulletSizePercent, 1.5);
assert.equal(richImportedShape.text.paragraphs[0].marginLeft, 32);
assert.equal(richImportedShape.text.paragraphs[0].indent, -16);
assert.equal(richImportedShape.text.paragraphs[0].lineSpacing, 1.25);
assert.equal(richImportedShape.text.paragraphs[0].spaceBefore, 12);
assert.equal(richImportedShape.text.paragraphs[0].spaceAfterPercent, 0.5);
assert.deepEqual(richImportedShape.text.paragraphs[0].style, { bold: false, italic: true, fontSize: 24, fontFamily: "Aptos", color: "accent2" });
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.bold, true);
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.fontSize, 36);
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.fontFamily, "Aptos Display");
assert.equal(richImportedShape.text.paragraphs[0].runs[0].style.color, "#0F172A");
assert.equal(richImportedShape.text.paragraphs[0].runs[1].style.italic, true);
assert.deepEqual(richImportedShape.text.paragraphs[1].autoNumber, { type: "romanLcPeriod", startAt: 3 });
assert.equal(richImportedShape.text.paragraphs[1].bulletFontFollowText, true);
assert.equal(richImportedShape.text.paragraphs[1].bulletColorFollowText, true);
assert.equal(richImportedShape.text.paragraphs[1].bulletSizeFollowText, true);
assert.equal(richImportedShape.text.paragraphs[2].bulletNone, true);
assert.equal(richImportedShape.text.paragraphs[2].bulletSize, 24);
assert.deepEqual(Object.keys(richImportedShape.text.inheritedParagraphStyles), ["0", "2"]);
assert.equal(richImportedShape.text.inheritedParagraphStyles[0].bulletCharacter, "•");
assert.equal(richImportedShape.text.inheritedParagraphStyles[0].bulletColor, "accent1");
assert.equal(richImportedShape.text.inheritedParagraphStyles[0].marginLeft, 48);
assert.equal(richImportedShape.text.inheritedParagraphStyles[0].indent, -16);
assert.equal(richImportedShape.text.inheritedParagraphStyles[0].spaceAfter, 6);
assert.deepEqual(richImportedShape.text.inheritedParagraphStyles[0].tabStops, [{ position: 180, alignment: "decimal" }]);
assert.deepEqual(richImportedShape.text.inheritedParagraphStyles[0].style, { fontSize: 18, color: "tx1" });
assert.deepEqual(richImportedShape.text.inheritedParagraphStyles[2].autoNumber, { type: "romanLcPeriod", startAt: 3 });
assert.deepEqual(richImportedShape.text.bodyProperties, { insets: { left: 8, top: 4, right: 12, bottom: 6 }, anchor: "center", wrap: "none", autoFit: "shrinkText", rotation: 15, verticalText: "vertical", verticalOverflow: "ellipsis", horizontalOverflow: "clip", columns: { count: 2, spacing: 18, rightToLeft: true }, upright: true });
richImportedShape.text.inheritedParagraphStyles = {
  0: { ...richImportedShape.text.inheritedParagraphStyles[0], bulletCharacter: "◆", bulletColor: "#16A34A", marginLeft: 56 },
  8: { level: 8, autoNumber: { type: "arabicPeriod", startAt: 9 }, bulletSizeFollowText: true, lineSpacing: 18, runs: [], style: {} },
};
richImportedShape.text.bodyProperties = { insets: { left: 16, top: 10, bottom: 7 }, anchor: "bottom", wrap: "square", autoFit: "resizeShape", rotation: -30, verticalText: "vertical270", verticalOverflow: "clip", horizontalOverflow: "overflow", columns: { count: 3, spacing: 24, rightToLeft: false }, upright: false };
richImportedShape.text.paragraphs = richImportedShape.text.paragraphs.map((paragraph, paragraphIndex) => ({
  ...paragraph,
  ...(paragraphIndex === 0 ? { bulletCharacter: "◆", bulletFont: undefined, bulletFontFollowText: true, bulletColor: "accent2", bulletSizePercent: undefined, bulletSize: 24, marginLeft: 40, indent: -20, lineSpacing: 18, spaceBefore: undefined, spaceBeforePercent: 0.25, spaceAfterPercent: undefined, spaceAfter: 6, style: { bold: true, italic: false, fontSize: 26, fontFamily: "Georgia", color: "#0EA5E9" } } : {}),
  ...(paragraphIndex === 1 ? { autoNumber: { type: "arabicPeriod", startAt: 5 }, bulletFontFollowText: undefined, bulletFont: "Aptos", bulletColorFollowText: undefined, bulletColor: "#16A34A", bulletSizeFollowText: undefined, bulletSizePercent: 1.25 } : {}),
  ...(paragraphIndex === 2 ? { bulletNone: undefined, bulletCharacter: "–", bulletSize: undefined, bulletSizeFollowText: true } : {}),
  runs: paragraph.runs.map((run, runIndex) => paragraphIndex === 0 && runIndex === 0
    ? { ...run, text: "Updated ", style: { ...run.style, bold: false, color: "#2563EB" } }
    : run),
}));
const richPptxEdited = await exportPptxWithOpenChestnut(richPptxImported);
const richPptxEditedZip = await JSZip.loadAsync(richPptxEdited.bytes);
const richPptxEditedXml = await richPptxEditedZip.file("ppt/slides/slide1.xml").async("text");
assert.match(richPptxEditedXml, /<a:buClr><a:schemeClr val="accent2"\s*\/><\/a:buClr>/);
assert.match(richPptxEditedXml, /<a:lnSpc><a:spcPts val="1350"\s*\/><\/a:lnSpc>/);
assert.match(richPptxEditedXml, /<a:spcBef><a:spcPct val="25000"\s*\/><\/a:spcBef>/);
assert.match(richPptxEditedXml, /<a:spcAft><a:spcPts val="450"\s*\/><\/a:spcAft>/);
assert.match(richPptxEditedXml, /<a:defRPr[^>]*sz="1950"[^>]*b="1"[^>]*i="0">[\s\S]*?<a:srgbClr val="0EA5E9"\s*\/>[\s\S]*?<a:latin typeface="Georgia"\s*\/>[\s\S]*?<\/a:defRPr>/);
assert.match(richPptxEditedXml, /<a:lvl1pPr[^>]*marL="533400"[^>]*indent="-152400">[\s\S]*?<a:srgbClr val="16A34A"\s*\/>[\s\S]*?<a:buChar char="◆"\s*\/>[\s\S]*?<\/a:lvl1pPr>/);
assert.doesNotMatch(richPptxEditedXml, /<a:lvl3pPr/);
assert.match(richPptxEditedXml, /<a:lvl9pPr>[\s\S]*?<a:lnSpc><a:spcPts val="1350"\s*\/><\/a:lnSpc>[\s\S]*?<a:buSzTx\s*\/>[\s\S]*?<a:buAutoNum type="arabicPeriod" startAt="9"\s*\/>[\s\S]*?<\/a:lvl9pPr>/);
assert.match(richPptxEditedXml, /<a:bodyPr[^>]*rot="-1800000"[^>]*vertOverflow="clip"[^>]*horzOverflow="overflow"[^>]*vert="vert270"[^>]*wrap="square"[^>]*lIns="152400"[^>]*tIns="95250"[^>]*bIns="66675"[^>]*numCol="3"[^>]*spcCol="228600"[^>]*rtlCol="0"[^>]*anchor="b"[^>]*upright="0"[^>]*>\s*<a:spAutoFit\s*\/>\s*<\/a:bodyPr>/);
assert.doesNotMatch(richPptxEditedXml, /<a:bodyPr[^>]*\brIns=/);
const richPptxRoundTrip = await importPptxWithOpenChestnut(richPptxEdited);
const richRoundTripShape = richPptxRoundTrip.slides.getItem(0).shapes.items[0];
assert.equal(richRoundTripShape.text.value, "Updated brief\nSource-bound detail\nExplicitly unbulleted");
assert.equal(richRoundTripShape.text.paragraphs[0].runs[0].style.bold, false);
assert.equal(richRoundTripShape.text.paragraphs[0].runs[0].style.color, "#2563EB");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletCharacter, "◆");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletFontFollowText, true);
assert.equal(richRoundTripShape.text.paragraphs[0].bulletColor, "accent2");
assert.equal(richRoundTripShape.text.paragraphs[0].bulletSize, 24);
assert.equal(richRoundTripShape.text.paragraphs[0].marginLeft, 40);
assert.equal(richRoundTripShape.text.paragraphs[0].indent, -20);
assert.equal(richRoundTripShape.text.paragraphs[0].lineSpacing, 18);
assert.equal(richRoundTripShape.text.paragraphs[0].spaceBeforePercent, 0.25);
assert.equal(richRoundTripShape.text.paragraphs[0].spaceAfter, 6);
assert.deepEqual(richRoundTripShape.text.paragraphs[0].style, { bold: true, italic: false, fontSize: 26, fontFamily: "Georgia", color: "#0EA5E9" });
assert.deepEqual(richRoundTripShape.text.paragraphs[1].autoNumber, { type: "arabicPeriod", startAt: 5 });
assert.equal(richRoundTripShape.text.paragraphs[1].bulletFont, "Aptos");
assert.equal(richRoundTripShape.text.paragraphs[1].bulletColor, "#16A34A");
assert.equal(richRoundTripShape.text.paragraphs[1].bulletSizePercent, 1.25);
assert.equal(richRoundTripShape.text.paragraphs[2].bulletCharacter, "–");
assert.equal(richRoundTripShape.text.paragraphs[2].bulletSizeFollowText, true);
assert.deepEqual(Object.keys(richRoundTripShape.text.inheritedParagraphStyles), ["0", "8"]);
assert.equal(richRoundTripShape.text.inheritedParagraphStyles[0].bulletCharacter, "◆");
assert.equal(richRoundTripShape.text.inheritedParagraphStyles[0].bulletColor, "#16A34A");
assert.equal(richRoundTripShape.text.inheritedParagraphStyles[8].lineSpacing, 18);
assert.deepEqual(richRoundTripShape.text.bodyProperties, { insets: { left: 16, top: 10, bottom: 7 }, anchor: "bottom", wrap: "square", autoFit: "resizeShape", rotation: -30, verticalText: "vertical270", verticalOverflow: "clip", horizontalOverflow: "overflow", columns: { count: 3, spacing: 24, rightToLeft: false }, upright: false });

richRoundTripShape.text.inheritedParagraphStyles = {};
richRoundTripShape.text.bodyProperties = {};
richRoundTripShape.text.paragraphs = richRoundTripShape.text.paragraphs.map((paragraph, index) => index === 0
  ? { ...paragraph, marginLeft: undefined, indent: undefined, lineSpacing: undefined, spaceBefore: undefined, spaceBeforePercent: undefined, spaceAfter: undefined, spaceAfterPercent: undefined, style: {} }
  : paragraph);
const richLayoutDeleted = await exportPptxWithOpenChestnut(richPptxRoundTrip);
const richLayoutDeletedZip = await JSZip.loadAsync(richLayoutDeleted.bytes);
const richLayoutDeletedXml = await richLayoutDeletedZip.file("ppt/slides/slide1.xml").async("text");
assert.doesNotMatch(richLayoutDeletedXml, /\bmarL=/);
assert.doesNotMatch(richLayoutDeletedXml, /\bindent=/);
assert.doesNotMatch(richLayoutDeletedXml, /<a:lnSpc>/);
assert.doesNotMatch(richLayoutDeletedXml, /<a:spcBef>/);
assert.doesNotMatch(richLayoutDeletedXml, /<a:spcAft>/);
assert.doesNotMatch(richLayoutDeletedXml, /<a:defRPr/);
assert.doesNotMatch(richLayoutDeletedXml, /<a:lvl[1-9]pPr/);
assert.match(richLayoutDeletedXml, /<a:bodyPr\b[^>]*\/>/);
const richLayoutDeletedRoundTrip = await importPptxWithOpenChestnut(richLayoutDeleted);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].marginLeft, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].indent, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].lineSpacing, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].spaceBefore, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].spaceBeforePercent, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].spaceAfter, undefined);
assert.equal(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].spaceAfterPercent, undefined);
assert.deepEqual(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].style, {});
assert.deepEqual(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.inheritedParagraphStyles, {});
assert.deepEqual(richLayoutDeletedRoundTrip.slides.getItem(0).shapes.items[0].text.bodyProperties, {});

const transformedColorZip = await JSZip.loadAsync(richPptx.bytes);
const transformedColorSlidePath = "ppt/slides/slide1.xml";
const transformedColorSlideXml = await transformedColorZip.file(transformedColorSlidePath).async("text");
const transformedColorNeedle = '<a:schemeClr val="accent1" />';
const transformedColorIndex = transformedColorSlideXml.lastIndexOf(transformedColorNeedle);
assert.notEqual(transformedColorIndex, -1);
transformedColorZip.file(
  transformedColorSlidePath,
  `${transformedColorSlideXml.slice(0, transformedColorIndex)}<a:schemeClr val="accent1"><a:tint val="50000" /></a:schemeClr>${transformedColorSlideXml.slice(transformedColorIndex + transformedColorNeedle.length)}`,
);
const transformedColorBytes = await transformedColorZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const transformedColorImported = await importPptxWithOpenChestnut(transformedColorBytes);
const transformedColorShape = transformedColorImported.slides.getItem(0).shapes.items[0];
assert.equal(transformedColorShape.text.paragraphs[0].bulletColor, undefined);
const transformedColorPreserved = await exportPptxWithOpenChestnut(transformedColorImported);
const transformedColorPreservedZip = await JSZip.loadAsync(transformedColorPreserved.bytes);
assert.match(await transformedColorPreservedZip.file(transformedColorSlidePath).async("text"), /<a:schemeClr val="accent1"><a:tint val="50000"\s*\/><\/a:schemeClr>/);
transformedColorShape.text.paragraphs = transformedColorShape.text.paragraphs.map((paragraph, index) => index === 0 ? { ...paragraph, bulletColor: "accent2" } : paragraph);
await assert.rejects(
  exportPptxWithOpenChestnut(transformedColorImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);

const pictureBulletPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";
const replacementPictureBulletPng = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAusB9Wl2nGQAAAAASUVORK5CYII=";
const pictureBulletPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
pictureBulletPresentation.slides.add({ name: "Picture markers" }).shapes.add({
  name: "Embedded and external markers",
  position: { left: 60, top: 40, width: 920, height: 180 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [
    { bulletImage: { dataUrl: pictureBulletPng }, runs: ["Embedded marker"] },
    { bulletImage: { uri: "https://example.com/marker.png" }, runs: ["External marker"] },
  ],
});
pictureBulletPresentation.slides.add({ name: "Shared marker" }).shapes.add({
  name: "Shared embedded marker",
  position: { left: 60, top: 40, width: 920, height: 100 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [{ bulletImage: pictureBulletPng, runs: ["Same bytes, different owner"] }],
});
const pictureBulletPptx = await exportPptxWithOpenChestnut(pictureBulletPresentation);
const pictureBulletZip = await JSZip.loadAsync(pictureBulletPptx.bytes);
assert.equal(Object.keys(pictureBulletZip.files).filter((name) => /^ppt\/media\//.test(name)).length, 1);
assert.match(await pictureBulletZip.file("ppt/slides/slide1.xml").async("text"), /<a:buBlip><a:blip r:embed="[^"]+"[^>]*\/><\/a:buBlip>/);
assert.match(await pictureBulletZip.file("ppt/slides/slide1.xml").async("text"), /<a:buBlip><a:blip r:link="[^"]+"[^>]*\/><\/a:buBlip>/);
assert.match(await pictureBulletZip.file("ppt/slides/_rels/slide1.xml.rels").async("text"), /relationships\/image[^>]*Target="https:\/\/example\.com\/marker\.png"[^>]*TargetMode="External"/);
const pictureBulletImported = await importPptxWithOpenChestnut(pictureBulletPptx);
const importedPictureParagraphs = pictureBulletImported.slides.getItem(0).shapes.items[0].text.paragraphs;
assert.equal(importedPictureParagraphs[0].bulletImage.dataUrl, pictureBulletPng);
assert.deepEqual(importedPictureParagraphs[1].bulletImage, { uri: "https://example.com/marker.png", relationshipMode: "link" });
assert.equal(pictureBulletImported.slides.getItem(1).shapes.items[0].text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);
pictureBulletImported.slides.getItem(0).shapes.items[0].text.paragraphs = importedPictureParagraphs.map((paragraph, index) => index === 0
  ? { ...paragraph, bulletImage: { dataUrl: replacementPictureBulletPng } }
  : { ...paragraph, bulletImage: undefined, bulletNone: true });
const pictureBulletEdited = await exportPptxWithOpenChestnut(pictureBulletImported);
const pictureBulletEditedZip = await JSZip.loadAsync(pictureBulletEdited.bytes);
assert.equal(Object.keys(pictureBulletEditedZip.files).filter((name) => /^ppt\/media\//.test(name)).length, 2);
const pictureBulletRoundTrip = await importPptxWithOpenChestnut(pictureBulletEdited);
assert.equal(pictureBulletRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].bulletImage.dataUrl, replacementPictureBulletPng);
assert.equal(pictureBulletRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[1].bulletNone, true);
assert.equal(pictureBulletRoundTrip.slides.getItem(1).shapes.items[0].text.paragraphs[0].bulletImage.dataUrl, pictureBulletPng);

for (const bulletImage of [
  { dataUrl: "data:image/png;base64,YWJj" },
  { dataUrl: `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><script>alert(1)</script></svg>').toString("base64")}` },
  { dataUrl: `data:image/svg+xml;base64,${Buffer.from('<svg xmlns="http://www.w3.org/2000/svg"><style>@import url(https://example.com/style.css)</style></svg>').toString("base64")}` },
  { uri: "file:///tmp/marker.png" },
]) {
  const invalidPictureBullet = Presentation.create();
  invalidPictureBullet.slides.add().shapes.add({ text: [{ bulletImage, runs: ["invalid"] }] });
  await assert.rejects(
    exportPptxWithOpenChestnut(invalidPictureBullet),
    (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_asset",
  );
}

const inlinePresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const inlineShape = inlinePresentation.slides.add({ name: "Inline text" }).shapes.add({
  name: "Fields, breaks, and tabs",
  geometry: "rect",
  position: { left: 60, top: 40, width: 920, height: 180 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [{
    tabStops: [{ position: 120, alignment: "left" }, { position: 260, alignment: "decimal" }],
    runs: [
      "Slide\t",
      { field: { id: "{11111111-2222-4333-8444-555555555555}", type: "slidenum", text: "1" }, style: { bold: true, color: "#2563EB" } },
      { break: true, style: { fontSize: 18 } },
      "Revenue\t42.5",
    ],
  }],
});
const inlinePptx = await exportPptxWithOpenChestnut(inlinePresentation);
const inlineZip = await JSZip.loadAsync(inlinePptx.bytes);
const inlineXml = await inlineZip.file("ppt/slides/slide1.xml").async("text");
assert.match(inlineXml, /<a:tabLst><a:tab pos="1143000" algn="l"\s*\/><a:tab pos="2476500" algn="dec"\s*\/><\/a:tabLst>/);
assert.match(inlineXml, /<a:fld id="\{11111111-2222-4333-8444-555555555555\}" type="slidenum">/);
assert.match(inlineXml, /<a:br><a:rPr[^>]*sz="1350"/);
const inlineImported = await importPptxWithOpenChestnut(inlinePptx);
const inlineImportedShape = inlineImported.slides.getItem(0).shapes.items[0];
assert.equal(inlineImportedShape.text.value, "Slide\t1\nRevenue\t42.5");
assert.deepEqual(inlineImportedShape.text.paragraphs[0].tabStops, [{ position: 120, alignment: "left" }, { position: 260, alignment: "decimal" }]);
assert.deepEqual(inlineImportedShape.text.paragraphs[0].runs[1].field, { id: "{11111111-2222-4333-8444-555555555555}", type: "slidenum", text: "1" });
assert.equal(inlineImportedShape.text.paragraphs[0].runs[2].break, true);
inlineImportedShape.text.paragraphs = inlineImportedShape.text.paragraphs.map((paragraph) => ({
  ...paragraph,
  tabStops: paragraph.tabStops.map((tab, index) => index === 1 ? { ...tab, position: 280 } : tab),
  runs: paragraph.runs.map((run) => run.field ? { ...run, field: { ...run.field, text: "2" } } : run.break ? { ...run, style: { ...run.style, italic: true } } : run),
}));
const inlineEdited = await exportPptxWithOpenChestnut(inlineImported);
const inlineRoundTrip = await importPptxWithOpenChestnut(inlineEdited);
const inlineRoundTripParagraph = inlineRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0];
assert.equal(inlineRoundTrip.slides.getItem(0).shapes.items[0].text.value, "Slide\t2\nRevenue\t42.5");
assert.equal(inlineRoundTripParagraph.tabStops[1].position, 280);
assert.equal(inlineRoundTripParagraph.runs[2].style.italic, true);
inlineRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs = [{ ...inlineRoundTripParagraph, tabStops: [] }];
const inlineTabsDeleted = await exportPptxWithOpenChestnut(inlineRoundTrip);
const inlineTabsDeletedZip = await JSZip.loadAsync(inlineTabsDeleted.bytes);
assert.doesNotMatch(await inlineTabsDeletedZip.file("ppt/slides/slide1.xml").async("text"), /<a:tabLst>/);
const inlineKindChanged = await importPptxWithOpenChestnut(inlinePptx);
const inlineKindChangedShape = inlineKindChanged.slides.getItem(0).shapes.items[0];
inlineKindChangedShape.text.paragraphs = inlineKindChangedShape.text.paragraphs.map((paragraph) => ({ ...paragraph, runs: paragraph.runs.map((run) => run.field ? { text: "not a field", style: run.style } : run) }));
await assert.rejects(
  exportPptxWithOpenChestnut(inlineKindChanged),
  (error) => error instanceof OpenChestnutCodecError && error.code === "presentation_text_topology_changed",
);

const hyperlinkPresentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const hyperlinkSlide = hyperlinkPresentation.slides.add({ name: "Links" });
const hyperlinkTarget = hyperlinkPresentation.slides.add({ name: "Details" });
const hyperlinkAppendix = hyperlinkPresentation.slides.add({ name: "Appendix" });
hyperlinkSlide.shapes.add({
  name: "Run links",
  position: { left: 60, top: 40, width: 920, height: 120 },
  fill: "#FFFFFF",
  line: { fill: "#334155", width: 1 },
  text: [{ runs: [
    { text: "Guide ", link: { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank", history: false, highlightClick: true } },
    { text: "Details ", link: { slideId: hyperlinkTarget.id } },
    { text: "Next ", link: { action: "nextSlide" } },
    { text: "End", link: { action: "endShow" } },
  ] }],
});
const hyperlinkPptx = await exportPptxWithOpenChestnut(hyperlinkPresentation);
const hyperlinkZip = await JSZip.loadAsync(hyperlinkPptx.bytes);
const hyperlinkSlideXml = await hyperlinkZip.file("ppt/slides/slide1.xml").async("text");
const hyperlinkRelsXml = await hyperlinkZip.file("ppt/slides/_rels/slide1.xml.rels").async("text");
assert.match(hyperlinkSlideXml, /<a:hlinkClick[^>]*tgtFrame="_blank"[^>]*tooltip="Read the guide"[^>]*history="0"[^>]*highlightClick="1"/);
assert.match(hyperlinkSlideXml, /action="ppaction:\/\/hlinksldjump"/);
assert.match(hyperlinkSlideXml, /action="ppaction:\/\/hlinkshowjump\?jump=nextslide"/);
assert.match(hyperlinkRelsXml, /relationships\/hyperlink[^>]*Target="https:\/\/example\.com\/guide\?x=1&amp;y=2"[^>]*TargetMode="External"/);
assert.match(hyperlinkRelsXml, /relationships\/slide[^>]*Target="(?:\/ppt\/slides\/)?slide2\.xml"/);
const hyperlinkImported = await importPptxWithOpenChestnut(hyperlinkPptx);
const hyperlinkImportedShape = hyperlinkImported.slides.getItem(0).shapes.items[0];
const hyperlinkRuns = hyperlinkImportedShape.text.paragraphs[0].runs;
assert.deepEqual(hyperlinkRuns[0].link, { uri: "https://example.com/guide?x=1&y=2", tooltip: "Read the guide", targetFrame: "_blank", history: false, highlightClick: true });
assert.deepEqual(hyperlinkRuns[1].link, { slideId: hyperlinkImported.slides.getItem(1).id });
assert.deepEqual(hyperlinkRuns[2].link, { action: "nextSlide" });
hyperlinkImportedShape.text.paragraphs = hyperlinkImportedShape.text.paragraphs.map((paragraph) => ({
  ...paragraph,
  runs: paragraph.runs.map((run, index) => ({
    ...run,
    link: index === 0
      ? { uri: "https://example.com/updated", targetFrame: "_self" }
      : index === 1
        ? { slideId: hyperlinkImported.slides.getItem(2).id, tooltip: "Appendix" }
        : index === 2
          ? { action: "lastSlide", highlightClick: false }
          : undefined,
  })),
}));
const hyperlinkEdited = await exportPptxWithOpenChestnut(hyperlinkImported);
const hyperlinkRoundTrip = await importPptxWithOpenChestnut(hyperlinkEdited);
const hyperlinkRoundTripRuns = hyperlinkRoundTrip.slides.getItem(0).shapes.items[0].text.paragraphs[0].runs;
assert.deepEqual(hyperlinkRoundTripRuns[0].link, { uri: "https://example.com/updated", targetFrame: "_self" });
assert.deepEqual(hyperlinkRoundTripRuns[1].link, { slideId: hyperlinkRoundTrip.slides.getItem(2).id, tooltip: "Appendix" });
assert.deepEqual(hyperlinkRoundTripRuns[2].link, { action: "lastSlide", highlightClick: false });
assert.equal(hyperlinkRoundTripRuns[3].link, undefined);

const unsupportedHyperlinkPresentation = Presentation.create();
unsupportedHyperlinkPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "Tour", link: { customShow: "Evidence" } }] }] });
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedHyperlinkPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_features" && /custom-show hyperlink/.test(error.message),
);

richImportedShape.text.paragraphs = [
  ...richImportedShape.text.paragraphs.slice(0, 1),
  { ...richImportedShape.text.paragraphs[1], runs: [...richImportedShape.text.paragraphs[1].runs, { text: "unsafe", style: {} }] },
];
await assert.rejects(
  exportPptxWithOpenChestnut(richPptxImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "presentation_text_topology_changed",
);

const unsupportedRichPresentation = Presentation.create();
unsupportedRichPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "underline", style: { underline: "single" } }] }] });
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedRichPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_features",
);

const invalidRichPresentation = Presentation.create();
invalidRichPresentation.slides.add().shapes.add({ text: [{ runs: [{ text: "transparent", style: { color: "transparent" } }] }] });
await assert.rejects(
  exportPptxWithOpenChestnut(invalidRichPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_features",
);

const invalidBulletPresentation = Presentation.create();
const invalidBulletShape = invalidBulletPresentation.slides.add().shapes.add({ text: [{ autoNumber: "arabicPeriod", runs: ["invalid"] }] });
invalidBulletShape.text._paragraphs[0].autoNumber = { type: "not-a-scheme" };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBulletPresentation),
  (error) => error instanceof RangeError && /auto-number type/.test(error.message),
);

const invalidSchemeBulletColorPresentation = Presentation.create();
const invalidSchemeBulletColorShape = invalidSchemeBulletColorPresentation.slides.add().shapes.add({ text: [{ bulletCharacter: "•", bulletColor: "accent1", runs: ["styled"] }] });
invalidSchemeBulletColorShape.text._paragraphs[0].bulletColor = "accent7";
await assert.rejects(
  exportPptxWithOpenChestnut(invalidSchemeBulletColorPresentation),
  (error) => error instanceof TypeError && /scheme color/.test(error.message),
);

const invalidParagraphLayoutPresentation = Presentation.create();
invalidParagraphLayoutPresentation.slides.add().shapes.add({ text: [{ bulletCharacter: "•", marginLeft: 6000, indent: -20, runs: ["too wide"] }] });
await assert.rejects(
  exportPptxWithOpenChestnut(invalidParagraphLayoutPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text",
);

const invalidParagraphSpacingPresentation = Presentation.create();
invalidParagraphSpacingPresentation.slides.add().shapes.add({ text: [{ spaceBefore: 2112.01, runs: ["too much space"] }] });
await assert.rejects(
  exportPptxWithOpenChestnut(invalidParagraphSpacingPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text",
);

const invalidDefaultRunStylePresentation = Presentation.create();
const invalidDefaultRunStyleShape = invalidDefaultRunStylePresentation.slides.add().shapes.add({ text: [{ style: { color: "accent1" }, runs: ["invalid theme"] }] });
invalidDefaultRunStyleShape.text._paragraphs[0].style.color = "accent7";
await assert.rejects(
  exportPptxWithOpenChestnut(invalidDefaultRunStylePresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_features" && /paragraphStyle.color/.test(error.message),
);

const emptyListStylePresentation = Presentation.create();
const emptyListStyleShape = emptyListStylePresentation.slides.add().shapes.add({ text: "empty list style" });
emptyListStyleShape.text.inheritedParagraphStyles = { 0: { level: 0, runs: [], style: {} } };
await assert.rejects(
  exportPptxWithOpenChestnut(emptyListStylePresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /at least one modeled property/.test(error.message),
);

const invalidListLevelPresentation = Presentation.create();
const invalidListLevelShape = invalidListLevelPresentation.slides.add().shapes.add({ text: "invalid list level" });
invalidListLevelShape.text.inheritedParagraphStyles = { 9: { level: 9, bulletCharacter: "•", runs: [], style: {} } };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidListLevelPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /level outside the supported 0-8 range/.test(error.message),
);

const invalidBodyPropertiesPresentation = Presentation.create();
const invalidBodyPropertiesShape = invalidBodyPropertiesPresentation.slides.add().shapes.add({ text: "invalid body properties" });
invalidBodyPropertiesShape.text.bodyProperties = { insets: { left: -1 } };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBodyPropertiesPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /left inset/.test(error.message),
);
invalidBodyPropertiesShape.text.bodyProperties = { anchor: "middle" };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBodyPropertiesPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /anchor/.test(error.message),
);
invalidBodyPropertiesShape.text.bodyProperties = { rotation: 361 };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBodyPropertiesPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /rotation/.test(error.message),
);
invalidBodyPropertiesShape.text.bodyProperties = { verticalOverflow: "fade" };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBodyPropertiesPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /vertical overflow/.test(error.message),
);
invalidBodyPropertiesShape.text.bodyProperties = { columns: { count: 17 } };
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBodyPropertiesPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_text" && /column count/.test(error.message),
);

const invalidBackgroundPresentation = Presentation.create();
invalidBackgroundPresentation.master.background = { fill: "accent1", mode: "reference", index: -1 };
invalidBackgroundPresentation.slides.add().shapes.add({ text: "invalid background" });
await assert.rejects(
  exportPptxWithOpenChestnut(invalidBackgroundPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_background" && /unsigned 32-bit/.test(error.message),
);

const preservedPresentation = Presentation.create({
  master: { id: "master/preservation", name: "Preservation master", background: "#FFFFFF", placeholders: [] },
});
const preservedSlide = preservedPresentation.slides.add({ name: "Opaque graph" });
preservedSlide.shapes.add({ name: "Editable title", position: { left: 40, top: 32, width: 720, height: 64 }, fill: "#FFFFFF", line: { fill: "#334155", width: 1 }, text: "Before WASM" });
preservedSlide.images.add({
  name: "Preserved image",
  alt: "Opaque image evidence",
  position: { left: 40, top: 140, width: 180, height: 120 },
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=",
});
preservedSlide.charts.add("bar", {
  name: "Preserved chart",
  position: { left: 300, top: 140, width: 560, height: 360 },
  categories: ["Model", "Package"],
  series: [{ name: "Evidence", values: [8, 12] }],
});
const presentationSource = await PresentationFile.exportPptx(preservedPresentation);
const presentationImported = await importPptxWithOpenChestnut(presentationSource);
assert.equal(presentationImported.slides.getItem(0).nativeObjects.items.length, 2);
presentationImported.slides.getItem(0).shapes.items[0].text.set("After WASM");
const presentationPreserved = await exportPptxWithOpenChestnut(presentationImported);
assert.equal(presentationPreserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const presentationRoundTrip = await PresentationFile.importPptx(presentationPreserved);
assert.equal(presentationRoundTrip.slides.getItem(0).shapes.items[0].text.value, "After WASM");
assert.equal(presentationRoundTrip.slides.getItem(0).images.items.length, 1);
assert.equal(presentationRoundTrip.slides.getItem(0).charts.items.length, 1);
presentationImported.slides.getItem(0).nativeObjects.items[0].name = "Unsafe native edit";
await assert.rejects(
  exportPptxWithOpenChestnut(presentationImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);

const unsupportedPresentation = Presentation.create();
unsupportedPresentation.slides.add().images.add({ prompt: "Unsupported direct image authoring" });
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedPresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_features",
);

assert.equal(status.available, true);
assert.equal(status.protocolVersion, 1);
assert.equal(status.assemblyName, "OpenChestnut.Runtime.dll");
assert.ok(status.manifest.totalBytes > 1_000_000);
assert.equal(status.manifest.files.some((file) => file.path.endsWith(".map") || file.path.endsWith(".symbols")), false);

const sofficeStatus = spawnSync("soffice", ["--version"], { encoding: "utf8" });
const popplerStatus = spawnSync("pdftoppm", ["-v"], { encoding: "utf8" });
if (sofficeStatus.status === 0 && popplerStatus.status === 0) {
  const pdf = await createLibreOfficeRenderer({ timeoutMs: 60_000 })({
    input: exported,
    inputType: exported.type,
    outputType: "application/pdf",
    format: "pdf",
    artifactKind: "workbook",
  });
  assert.equal(new TextDecoder().decode(pdf.bytes.slice(0, 4)), "%PDF");
  const png = await createPopplerRenderer({ dpi: 96, timeoutMs: 60_000 })({
    input: pdf,
    inputType: pdf.type,
    outputType: "image/png",
    format: "png",
    artifactKind: "workbook",
    pageIndex: 0,
  });
  assert.deepEqual([...png.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);
}

const noDotnetProbe = `
  import { Workbook } from ${JSON.stringify(new URL("../src/index.mjs", import.meta.url).href)};
  import { exportXlsxWithOpenChestnut } from ${JSON.stringify(new URL("../src/codecs/open-chestnut.mjs", import.meta.url).href)};
  const workbook = Workbook.create();
  workbook.worksheets.add("Sheet1").getRange("A1").values = [["no dotnet on PATH"]];
  const file = await exportXlsxWithOpenChestnut(workbook);
  if (file.bytes[0] !== 0x50 || file.bytes[1] !== 0x4b) process.exit(1);
`;
const child = spawnSync(process.execPath, ["--input-type=module", "-e", noDotnetProbe], {
  cwd: new URL("..", import.meta.url),
  encoding: "utf8",
  env: { ...process.env, PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin" },
});
assert.equal(child.status, 0, `bundled runtime failed without dotnet on PATH\nSTDOUT:\n${child.stdout}\nSTDERR:\n${child.stderr}`);

console.log("OpenChestnut smoke ok");
