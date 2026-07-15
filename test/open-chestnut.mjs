import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import { create, fromBinary, toBinary } from "@bufbuild/protobuf";
import JSZip from "jszip";
import { DocumentFile, DocumentModel, Presentation, PresentationFile, Workbook, SpreadsheetFile } from "../src/index.mjs";
import { createLibreOfficeRenderer } from "../src/renderers/libreoffice.mjs";
import { createPopplerRenderer } from "../src/renderers/poppler.mjs";
import { ArtifactFamily, CellArtifactSchema, CodecOperation, DocumentBlockSchema, DocumentFieldSchema, DocumentHyperlinkSchema, DocumentNumberingSchema, DocumentParagraphSchema, DocumentSourceBindingSchema, DocumentTableCellMarginsSchema, DocumentTableCellSchema, DocumentTableFormattingSchema, DocumentTableSchema, PresentationArtifactSchema, PresentationBackgroundSchema, PresentationLayoutSchema, PresentationLayoutSourceBindingSchema, PresentationMasterSchema, PresentationMasterSourceBindingSchema, PresentationMasterTextStylesSchema, PresentationOpaqueElementSchema, PresentationPlaceholderFrameSchema, PresentationPlaceholderSchema, PresentationSlideSchema, PresentationTextBodyPropertiesSchema, PresentationTextBodySchema, PresentationTextParagraphSchema, PresentationTextRunSchema, SpreadsheetCalculationArtifactSchema, SpreadsheetChartArtifactSchema, SpreadsheetChartAxisArtifactSchema, SpreadsheetChartDataLabelPosition, SpreadsheetChartDataLabelsArtifactSchema, SpreadsheetChartLineDashStyle, SpreadsheetChartLineGrouping, SpreadsheetChartLineOptionsArtifactSchema, SpreadsheetChartLineStyleArtifactSchema, SpreadsheetChartMarkerArtifactSchema, SpreadsheetChartMarkerSymbol, SpreadsheetChartSeriesArtifactSchema, SpreadsheetChartSourceBindingSchema, SpreadsheetChartTextStyleArtifactSchema, SpreadsheetChartType, SpreadsheetConnectionArtifactSchema, SpreadsheetDefinedNameArtifactSchema, SpreadsheetImageArtifactSchema, SpreadsheetImageSourceBindingSchema, SpreadsheetImageTransformArtifactSchema, SpreadsheetOneCellAnchorArtifactSchema, SpreadsheetTableArtifactSchema, SpreadsheetTableColorArtifactSchema, SpreadsheetTableColumnArtifactSchema, SpreadsheetTableFilterArtifactSchema, SpreadsheetTableIconArtifactSchema, SpreadsheetTableQueryArtifactSchema, SpreadsheetTableQueryFieldArtifactSchema, SpreadsheetTableQueryRefreshArtifactSchema, SpreadsheetTableSortConditionArtifactSchema, SpreadsheetTableSortStateArtifactSchema, SpreadsheetTableValueFilterArtifactSchema, SpreadsheetWorkbookViewArtifactSchema, SpreadsheetWorkbookViewSourceBindingSchema, SpreadsheetWorksheetSourceBindingSchema, SpreadsheetWorksheetViewSourceBindingSchema, SpreadsheetWorksheetVisibility, WorkbookArtifactSchema, WorksheetArtifactSchema } from "../src/generated/open_office/artifact/v1/office_artifact_pb.js";
import {
  OpenChestnutCodecError,
  exportDocxWithOpenChestnut,
  exportPptxWithOpenChestnut,
  exportXlsxWithOpenChestnut,
  importDocxWithOpenChestnut,
  importPptxWithOpenChestnut,
  importXlsxWithOpenChestnut,
  invokeOpenChestnut,
  openChestnutStatus,
} from "../src/codecs/open-chestnut.mjs";
import { spreadsheetChartFromWire } from "../src/codecs/open-chestnut-spreadsheet-charts.mjs";
import { materializePresentationNativeGraphs } from "../src/codecs/open-chestnut-presentation-native.mjs";
import { parseSpreadsheetChart } from "../src/spreadsheet/ooxml-drawings.mjs";

function appendComplexColorDifferentialFormat(stylesXml) {
  const collection = /<x:dxfs count="(\d+)">/.exec(stylesXml);
  assert.ok(collection, "fixture requires an existing differential-style collection");
  const id = Number(collection[1]);
  const complexDxf = '<x:dxf><x:font><x:color rgb="FF2563EB" /></x:font><x:fill><x:patternFill patternType="solid"><x:fgColor rgb="FFE11D48" /><x:bgColor indexed="64" /></x:patternFill></x:fill></x:dxf>';
  const xml = stylesXml
    .replace(collection[0], `<x:dxfs count="${id + 1}">`)
    .replace("</x:dxfs>", `${complexDxf}</x:dxfs>`);
  assert.notEqual(xml, stylesXml, "fixture must append a complex differential style");
  return { id, xml };
}

async function addQueryTableGraph(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const contentTypesPath = "[Content_Types].xml";
  const workbookRelationshipsPath = "xl/_rels/workbook.xml.rels";
  const contentTypes = await zip.file(contentTypesPath).async("text");
  zip.file(contentTypesPath, contentTypes.replace(
    "</Types>",
    '<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/><Override PartName="/xl/queryTables/queryTable1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>',
  ));
  const workbookRelationships = await zip.file(workbookRelationshipsPath).async("text");
  zip.file(workbookRelationshipsPath, workbookRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>',
  ));
  zip.file("xl/tables/_rels/table1.xml.rels", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>');
  zip.file("xl/connections.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture"><x:connection id="7" name="Fixture warehouse" description="Read-only warehouse source" type="5" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="0" saveData="1" savePassword="0" credentials="integrated"><x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="SELECT Status, Value FROM Metrics" commandType="2"/><x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="kept"/></x:ext></x:extLst></x:connection><x:connection id="8" name="Opaque companion" type="1" refreshedVersion="8"><x:dbPr connection="Provider=Opaque.Provider;Data Source=opaque.invalid" command="SELECT 1" commandType="2"/></x:connection></x:connections>');
  zip.file("xl/queryTables/queryTable1.xml", '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:query-fixture" name="Warehouse metrics" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" firstBackgroundRefresh="0" refreshOnLoad="0" growShrinkType="insertClear" fillFormulas="0" removeDataOnSave="0" disableEdit="0" preserveFormatting="1" adjustColumnWidth="1" intermediate="0" connectionId="7"><x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="3" unboundColumnsLeft="0" unboundColumnsRight="0"><x:queryTableFields count="2"><x:queryTableField id="1" name="Status" dataBound="1" tableColumnId="1" fillFormulas="0" clipped="0"><x:extLst><x:ext uri="{71C44015-E485-449B-93BE-190C959F820F}"><fixture:fieldOpaque value="kept"/></x:ext></x:extLst></x:queryTableField><x:queryTableField id="2" name="Value" dataBound="1" tableColumnId="2"/></x:queryTableFields><x:queryTableDeletedFields count="2"><x:deletedField name="Legacy Status"/><x:deletedField name="Legacy Value"/></x:queryTableDeletedFields><x:sortState ref="A2:B3" caseSensitive="1" sortMethod="stroke"><x:sortCondition ref="B2:B3" descending="1" customList="ready,pending"/><x:sortCondition ref="A2:A3" sortBy="icon" iconSet="3Arrows" iconId="0"/><x:extLst><x:ext uri="{A1E10EA8-3B88-4BE3-9884-625AB42E9DDC}"><fixture:sortOpaque value="kept"/></x:ext></x:extLst></x:sortState></x:queryTableRefresh><x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:opaque value="kept"/></x:ext></x:extLst></x:queryTable>');
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

async function addOpaqueOpcGraph(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath).async("text");
  const updatedContentTypes = contentTypes.replace(
    "</Types>",
    '<Override PartName="/xl/custom/native.xml" ContentType="application/vnd.open-office-artifact-tool.native+xml"/></Types>',
  );
  assert.notEqual(updatedContentTypes, contentTypes, "fixture must register the custom OPC part content type");
  zip.file(contentTypesPath, updatedContentTypes);
  zip.file("xl/custom/native.xml", '<native xmlns="urn:open-office-artifact-tool:native">preserve me</native>');
  zip.file("xl/custom/_rels/native.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPayload" Type="urn:open-office-artifact-tool:native-payload" Target="https://example.invalid/native-payload" TargetMode="External"/></Relationships>');
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

async function addPresentationNativeGraph(bytes, embeddedWorkbookBytes, { sharedOleOnSecondSlide = false } = {}) {
  const zip = await JSZip.loadAsync(bytes);
  const namespaces = ' xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"';
  const ole = `<p:graphicFrame${namespaces}><p:nvGraphicFramePr><p:cNvPr id="100" name="Embedded workbook"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="rIdNativeOle" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdNativePreview"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="3657600" cy="2286000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
  const diagram = `<p:graphicFrame${namespaces}><p:nvGraphicFramePr><p:cNvPr id="101" name="Preserved diagram"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="457200" y="3657600"/><a:ext cx="5486400" cy="1828800"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" r:dm="rIdNativeDm" r:lo="rIdNativeLo" r:qs="rIdNativeQs" r:cs="rIdNativeCs"/></a:graphicData></a:graphic></p:graphicFrame>`;
  const content = `<p:grpSp${namespaces}><p:nvGrpSpPr><p:cNvPr id="102" name="Native content group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="7000000" y="5000000"/><a:ext cx="952500" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="952500" cy="952500"/></a:xfrm></p:grpSpPr><p:contentPart r:id="rIdNativeContent"/></p:grpSp>`;
  const slidePath = "ppt/slides/slide1.xml";
  const slideXml = await zip.file(slidePath).async("text");
  const updatedSlide = slideXml.replace("</p:spTree>", `${ole}${diagram}${content}</p:spTree>`);
  assert.notEqual(updatedSlide, slideXml, "fixture must append native PresentationML objects");
  zip.file(slidePath, updatedSlide);

  const relationshipsPath = "ppt/slides/_rels/slide1.xml.rels";
  const relationshipsXml = await zip.file(relationshipsPath)?.async("text") || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
  const nativeRelationships = '<Relationship Id="rIdNativeOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/native-workbook.xlsx"/><Relationship Id="rIdNativePreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/native-preview.png"/><Relationship Id="rIdNativeDm" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="../diagrams/native-data.xml"/><Relationship Id="rIdNativeLo" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" Target="../diagrams/native-layout.xml"/><Relationship Id="rIdNativeQs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" Target="../diagrams/native-style.xml"/><Relationship Id="rIdNativeCs" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" Target="../diagrams/native-colors.xml"/><Relationship Id="rIdNativeContent" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/native-content.xml"/>';
  zip.file(relationshipsPath, relationshipsXml.replace("</Relationships>", `${nativeRelationships}</Relationships>`));

  if (sharedOleOnSecondSlide) {
    const secondSlidePath = "ppt/slides/slide2.xml";
    const secondSlideXml = await zip.file(secondSlidePath)?.async("text");
    assert.ok(secondSlideXml?.includes("</p:spTree>"), "shared OLE fixture requires slide2.xml");
    const secondOle = ole
      .replace('id="100" name="Embedded workbook"', 'id="200" name="Shared embedded workbook"')
      .replaceAll("rIdNativeOle", "rIdSharedNativeOle")
      .replaceAll("rIdNativePreview", "rIdSharedNativePreview");
    zip.file(secondSlidePath, secondSlideXml.replace("</p:spTree>", `${secondOle}</p:spTree>`));
    const secondRelationshipsPath = "ppt/slides/_rels/slide2.xml.rels";
    const secondRelationshipsXml = await zip.file(secondRelationshipsPath)?.async("text") || '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    const sharedRelationships = '<Relationship Id="rIdSharedNativeOle" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/native-workbook.xlsx"/><Relationship Id="rIdSharedNativePreview" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/native-preview.png"/>';
    zip.file(secondRelationshipsPath, secondRelationshipsXml.replace("</Relationships>", `${sharedRelationships}</Relationships>`));
  }

  const contentTypesPath = "[Content_Types].xml";
  const contentTypes = await zip.file(contentTypesPath).async("text");
  const overrides = '<Override PartName="/ppt/embeddings/native-workbook.xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/><Override PartName="/ppt/media/native-preview.png" ContentType="image/png"/><Override PartName="/ppt/diagrams/native-data.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramData+xml"/><Override PartName="/ppt/diagrams/native-layout.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramLayout+xml"/><Override PartName="/ppt/diagrams/native-style.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramStyle+xml"/><Override PartName="/ppt/diagrams/native-colors.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.diagramColors+xml"/><Override PartName="/ppt/customXml/native-content.xml" ContentType="application/xml"/><Override PartName="/ppt/customXml/itemProps1.xml" ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"/>';
  zip.file(contentTypesPath, contentTypes.replace("</Types>", `${overrides}</Types>`));
  assert.ok(embeddedWorkbookBytes?.length, "fixture requires a valid embedded XLSX workbook");
  zip.file("ppt/embeddings/native-workbook.xlsx", embeddedWorkbookBytes);
  zip.file("ppt/media/native-preview.png", "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", { base64: true });
  zip.file("ppt/diagrams/native-data.xml", '<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:ptLst/><dgm:cxnLst/><dgm:bg/><dgm:whole/></dgm:dataModel>');
  zip.file("ppt/diagrams/native-layout.xml", '<dgm:layoutDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-layout"><dgm:title val="Native"/><dgm:desc val="Native layout"/><dgm:catLst/><dgm:layoutNode name="root"/></dgm:layoutDef>');
  zip.file("ppt/diagrams/native-style.xml", '<dgm:styleDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-style"><dgm:title val="Native"/><dgm:desc val="Native style"/><dgm:catLst/><dgm:styleLbl name="node0"/></dgm:styleDef>');
  zip.file("ppt/diagrams/native-colors.xml", '<dgm:colorsDef xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" uniqueId="urn:open-office:native-colors"><dgm:title val="Native"/><dgm:desc val="Native colors"/><dgm:catLst/></dgm:colorsDef>');
  zip.file("ppt/customXml/native-content.xml", '<native xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:link="rIdPayload">preserve me</native>');
  zip.file("ppt/customXml/_rels/native-content.xml.rels", '<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdPayload" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" Target="itemProps1.xml"/></Relationships>');
  zip.file("ppt/customXml/itemProps1.xml", '<ds:datastoreItem ds:itemID="{00112233-4455-6677-8899-AABBCCDDEEFF}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs/></ds:datastoreItem>');
  return zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

const legacyTabWire = toBinary(PresentationTextParagraphSchema, create(PresentationTextParagraphSchema, {
  tabStops: [{ positionEmu: 120n, alignment: "left" }],
}));
await assert.rejects(
  materializePresentationNativeGraphs({
    payload: { case: "presentation", value: { slides: [{ elements: [{ content: { case: "opaque", value: { preservedPartPaths: ["ppt/native/../escape.xml"] } } }] }] } },
    opaqueOpc: { parts: [] },
  }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_native_graph",
);
assert.equal(toBinary(CellArtifactSchema, create(CellArtifactSchema, { numberFormatCode: "0.00%" }))[0], 0x22, "Spreadsheet number-format codes must use additive cell field 4.");
assert.equal(toBinary(CellArtifactSchema, create(CellArtifactSchema, { formulaMetadata: { kind: 1, sharedIndex: 7, reference: "C1:C2" } }))[0], 0x2a, "Spreadsheet formula topology must use additive cell field 5.");
assert.equal(toBinary(CellArtifactSchema, create(CellArtifactSchema, { style: { font: { bold: true } } }))[0], 0x32, "Spreadsheet static styles must use additive cell field 6.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { theme: { accent1Rgb: "0F766E" } }))[0], 0x22, "Spreadsheet workbook themes must use additive workbook field 4.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { connections: [{ connectionId: 7, name: "Warehouse", type: 5, refreshedVersion: 8 }] }))[0], 0x2a, "Spreadsheet workbook connections must use additive workbook field 5.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { definedNames: [{ id: "defined-name/1", name: "Data", refersTo: "Sheet1!A1" }] }))[0], 0x32, "Spreadsheet workbook defined names must use additive workbook field 6.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { calculation: { mode: 1 } }))[0], 0x3a, "Spreadsheet workbook calculation policy must use additive workbook field 7.");
assert.equal(toBinary(PresentationOpaqueElementSchema, create(PresentationOpaqueElementSchema, { oleWorkbook: { partPath: "ppt/embeddings/book.xlsx" } }))[0], 0x5a, "Presentation OLE workbooks must use additive opaque-element field 11.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { view: { activeWorksheetId: "worksheet/2" } }))[0], 0x42, "Spreadsheet workbook views must use additive workbook field 8.");
assert.equal(toBinary(WorkbookArtifactSchema, create(WorkbookArtifactSchema, { additionalViews: [{ activeWorksheetId: "worksheet/3" }] }))[0], 0x4a, "Spreadsheet additional workbook windows must use additive workbook field 9.");
assert.equal(toBinary(SpreadsheetWorkbookViewArtifactSchema, create(SpreadsheetWorkbookViewArtifactSchema, { source: { viewXmlSha256: "x" } }))[0], 0x12, "Spreadsheet workbook-view source bindings must use view field 2.");
assert.equal(toBinary(SpreadsheetWorkbookViewArtifactSchema, create(SpreadsheetWorkbookViewArtifactSchema, { selectedWorksheetIds: ["worksheet/1"] }))[0], 0x1a, "Spreadsheet selected worksheet IDs must use view field 3.");
assert.equal(toBinary(SpreadsheetWorkbookViewSourceBindingSchema, create(SpreadsheetWorkbookViewSourceBindingSchema, { semanticSha256: "x" }))[0], 0x22, "Spreadsheet workbook-view semantic hashes must use source-binding field 4.");
assert.equal(toBinary(SpreadsheetWorkbookViewSourceBindingSchema, create(SpreadsheetWorkbookViewSourceBindingSchema, { worksheetViews: [{ worksheetId: "worksheet/1" }] }))[0], 0x32, "Spreadsheet worksheet-view bindings must use workbook-view source field 6.");
assert.deepEqual([...toBinary(SpreadsheetWorksheetViewSourceBindingSchema, create(SpreadsheetWorksheetViewSourceBindingSchema, { tabSelected: false }))], [0x38, 0x00], "Spreadsheet worksheet-view selection bindings must preserve explicit false values at field 7.");
assert.deepEqual([...toBinary(SpreadsheetCalculationArtifactSchema, create(SpreadsheetCalculationArtifactSchema, { calculateOnSave: false }))], [0x10, 0x00], "Spreadsheet calculation booleans must preserve explicit false values.");
assert.deepEqual([...toBinary(SpreadsheetDefinedNameArtifactSchema, create(SpreadsheetDefinedNameArtifactSchema, { hidden: false }))], [0x30, 0x00], "Spreadsheet defined-name hidden state must preserve explicit false values.");
assert.deepEqual([...toBinary(SpreadsheetConnectionArtifactSchema, create(SpreadsheetConnectionArtifactSchema, { keepAlive: false }))], [0x30, 0x00], "Spreadsheet connection booleans must preserve explicit false values.");
assert.equal(toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { tables: [{ name: "Sales" }] }))[0], 0x4a, "Spreadsheet worksheet tables must use additive worksheet field 9.");
assert.equal(toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { sortState: { reference: "A1:B2" } }))[0], 0x52, "Spreadsheet worksheet sort state must use additive worksheet field 10.");
assert.deepEqual([...toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { visibility: SpreadsheetWorksheetVisibility.HIDDEN }))], [0x58, 0x02], "Spreadsheet worksheet visibility must use additive worksheet field 11.");
assert.equal(toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { source: { ordinal: 1 } }))[0], 0x62, "Spreadsheet worksheet source bindings must use additive worksheet field 12.");
assert.equal(toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { images: [{ id: "image/1" }] }))[0], 0x6a, "Spreadsheet worksheet images must use additive worksheet field 13.");
assert.equal(toBinary(WorksheetArtifactSchema, create(WorksheetArtifactSchema, { charts: [{ id: "chart/1" }] }))[0], 0x72, "Spreadsheet worksheet charts must use additive worksheet field 14.");
assert.equal(toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { source: { chartPartPath: "xl/charts/chart1.xml" } }))[0], 0x5a, "Spreadsheet chart source bindings must use chart field 11.");
assert.equal(toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { xAxis: { title: "Quarter" } }))[0], 0x62, "Spreadsheet chart x-axis semantics must use additive chart field 12.");
assert.equal(toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { yAxis: { title: "Revenue" } }))[0], 0x6a, "Spreadsheet chart y-axis semantics must use additive chart field 13.");
assert.equal(toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { titleTextStyle: { fontSizePoints: 12 } }))[0], 0x72, "Spreadsheet chart title text styles must use additive chart field 14.");
assert.equal(toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { lineOptions: { smooth: true } }))[0], 0x7a, "Spreadsheet chart line options must use additive chart field 15.");
assert.deepEqual([...toBinary(SpreadsheetChartArtifactSchema, create(SpreadsheetChartArtifactSchema, { dataLabels: {} }))].slice(0, 2), [0x82, 0x01], "Spreadsheet chart data labels must use additive chart field 16.");
assert.deepEqual([...toBinary(SpreadsheetChartDataLabelsArtifactSchema, create(SpreadsheetChartDataLabelsArtifactSchema, { showValue: true }))], [0x08, 0x01], "Spreadsheet chart show-value labels must use data-label field 1.");
assert.deepEqual([...toBinary(SpreadsheetChartDataLabelsArtifactSchema, create(SpreadsheetChartDataLabelsArtifactSchema, { showCategoryName: true }))], [0x10, 0x01], "Spreadsheet chart show-category labels must use data-label field 2.");
assert.deepEqual([...toBinary(SpreadsheetChartDataLabelsArtifactSchema, create(SpreadsheetChartDataLabelsArtifactSchema, { position: SpreadsheetChartDataLabelPosition.TOP }))], [0x18, 0x09], "Spreadsheet chart label position must preserve optional enum field 3.");
assert.deepEqual([...toBinary(SpreadsheetChartDataLabelsArtifactSchema, create(SpreadsheetChartDataLabelsArtifactSchema, { showSeriesName: false }))], [0x20, 0x00], "Spreadsheet chart show-series-name must preserve explicit false at optional field 4.");
assert.deepEqual([...toBinary(SpreadsheetChartLineOptionsArtifactSchema, create(SpreadsheetChartLineOptionsArtifactSchema, { smooth: false }))], [0x08, 0x00], "Spreadsheet chart smooth options must preserve explicit false at optional field 1.");
assert.deepEqual([...toBinary(SpreadsheetChartLineOptionsArtifactSchema, create(SpreadsheetChartLineOptionsArtifactSchema, { grouping: SpreadsheetChartLineGrouping.STACKED }))], [0x10, 0x02], "Spreadsheet chart grouping must preserve explicit stacked presence at optional field 2.");
assert.deepEqual([...toBinary(SpreadsheetChartLineOptionsArtifactSchema, create(SpreadsheetChartLineOptionsArtifactSchema, { varyColors: true }))], [0x18, 0x01], "Spreadsheet chart vary-colors true must use boolean field 3.");
assert.deepEqual([...toBinary(SpreadsheetChartLineOptionsArtifactSchema, create(SpreadsheetChartLineOptionsArtifactSchema, { varyColors: false }))], [], "Spreadsheet chart vary-colors false must retain the native omitted default.");
assert.equal(toBinary(SpreadsheetChartSeriesArtifactSchema, create(SpreadsheetChartSeriesArtifactSchema, { fill: { source: { case: "rgb", value: "F472B6" } } }))[0], 0x2a, "Spreadsheet chart series fills must use additive series field 5.");
assert.equal(toBinary(SpreadsheetChartSeriesArtifactSchema, create(SpreadsheetChartSeriesArtifactSchema, { line: {} }))[0], 0x32, "Spreadsheet chart series lines must use additive series field 6.");
assert.equal(toBinary(SpreadsheetChartSeriesArtifactSchema, create(SpreadsheetChartSeriesArtifactSchema, { marker: {} }))[0], 0x3a, "Spreadsheet chart series markers must use additive series field 7.");
assert.equal(toBinary(SpreadsheetChartLineStyleArtifactSchema, create(SpreadsheetChartLineStyleArtifactSchema, { dashStyle: SpreadsheetChartLineDashStyle.DASHED }))[0], 0x10, "Spreadsheet chart line dash styles must use line field 2.");
assert.equal(toBinary(SpreadsheetChartLineStyleArtifactSchema, create(SpreadsheetChartLineStyleArtifactSchema, { widthPoints: 0 }))[0], 0x19, "Spreadsheet chart line widths must preserve explicit zero at optional line field 3.");
assert.equal(toBinary(SpreadsheetChartMarkerArtifactSchema, create(SpreadsheetChartMarkerArtifactSchema, { symbol: SpreadsheetChartMarkerSymbol.DIAMOND }))[0], 0x08, "Spreadsheet chart marker symbols must use marker field 1.");
assert.deepEqual([...toBinary(SpreadsheetChartMarkerArtifactSchema, create(SpreadsheetChartMarkerArtifactSchema, { size: 2 }))], [0x10, 0x02], "Spreadsheet chart marker sizes must preserve optional integer presence at marker field 2.");
assert.equal(toBinary(SpreadsheetChartMarkerArtifactSchema, create(SpreadsheetChartMarkerArtifactSchema, { fill: { source: { case: "rgb", value: "E11D48" } } }))[0], 0x1a, "Spreadsheet chart marker fills must use additive marker field 3.");
assert.equal(toBinary(SpreadsheetChartMarkerArtifactSchema, create(SpreadsheetChartMarkerArtifactSchema, { line: {} }))[0], 0x22, "Spreadsheet chart marker outlines must use additive marker field 4.");
assert.equal(toBinary(SpreadsheetChartAxisArtifactSchema, create(SpreadsheetChartAxisArtifactSchema, { textStyle: { fontSizePoints: 10 } }))[0], 0x3a, "Spreadsheet chart axis tick-label styles must use additive axis field 7.");
assert.equal(toBinary(SpreadsheetChartTextStyleArtifactSchema, create(SpreadsheetChartTextStyleArtifactSchema, { fontSizePoints: 10 }))[0], 0x09, "Spreadsheet chart font sizes must preserve optional double presence at text-style field 1.");
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, series: [{ name: "Malformed", fill: { source: { case: "rgb", value: "12345" } } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /#RRGGBB solid color/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, series: [{ name: "Theme", fill: { source: { case: "theme", value: 4 } } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /non-RGB fill source/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, title: "Malformed size", titleTextStyle: { fontSizePoints: 0 }, series: [] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /1 through 4000/i.test(error.message),
  "Malformed wire font sizes must fail before a workbook is mutated.",
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, series: [{ name: "Theme line", line: { color: { source: { case: "theme", value: 4 } } } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /non-RGB color source/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, series: [{ name: "Wide line", line: { widthPoints: 1_585 } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /0 through 1584/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, series: [{ name: "Unknown marker", marker: { symbol: 99 } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /unsupported symbol 99/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, series: [{ name: "Small marker", marker: { size: 1 } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /2 through 72/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, series: [{ name: "Theme marker", marker: { fill: { source: { case: "theme", value: 4 } } } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /non-RGB fill source/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, series: [{ name: "Wide marker", marker: { line: { widthPoints: 1_585 } } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /0 through 1584/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.BAR, series: [{ name: "Bar marker", marker: { symbol: SpreadsheetChartMarkerSymbol.CIRCLE } }] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /require a line chart/i.test(error.message),
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, lineOptions: {}, series: [] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /explicit grouping, smooth, or vary-colors semantics/i.test(error.message),
  "Malformed wire line options must fail before a workbook is mutated.",
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, lineOptions: { grouping: 99 }, series: [] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /unsupported grouping 99/i.test(error.message),
  "Unknown wire grouping values must fail before a workbook is mutated.",
);
assert.throws(
  () => spreadsheetChartFromWire(null, { type: SpreadsheetChartType.LINE, dataLabels: { position: 99 }, series: [] }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /unsupported position 99/i.test(error.message),
  "Unknown wire data-label positions must fail before a workbook is mutated.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:barChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx><c:dPt><c:idx val="0"/><c:spPr><a:solidFill><a:srgbClr val="E11D48"/></a:solidFill></c:spPr></c:dPt><c:cat><c:strLit><c:ptCount val="1"/><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:ptCount val="1"/><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:barChart></c:plotArea></c:chart></c:chartSpace>').series[0].fill,
  undefined,
  "JavaScript chart import must not flatten a data-point fill into a series fill.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:v>Series</c:v></c:tx><c:spPr><a:ln w="25400"><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill><a:prstDash val="dashDot"/></a:ln></c:spPr><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>').series[0],
  { name: "Series", categoryFormula: undefined, formula: undefined, categories: ["A"], values: [1], fill: undefined, line: { width: 2, fill: "#2563EB", style: "dash-dot" }, marker: undefined },
  "JavaScript chart import must keep line fill distinct from the series area fill.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr sz="1250"/><a:t>Styled</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Value</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart><c:catAx><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr/></a:p></c:txPr></c:catAx><c:valAx><c:scaling/></c:valAx></c:plotArea></c:chart></c:chartSpace>'),
  {
    type: "line",
    title: "Styled",
    titleTextStyle: { fontSize: 12.5 },
    hasLegend: false,
    categories: ["A"],
    series: [{ name: "Value", categoryFormula: undefined, formula: undefined, categories: ["A"], values: [1], fill: undefined, line: undefined, marker: undefined }],
    xAxis: { axisType: "textAxis", title: { text: "" }, textStyle: { fontSize: 10 } },
    yAxis: { axisType: "valueAxis", title: { text: "" } },
  },
  "JavaScript fallback import must expose bounded title and tick-label font sizes.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Marker</c:v></c:tx><c:marker><c:symbol val="diamond"/><c:size val="8"/><c:spPr><a:solidFill><a:srgbClr val="E11D48"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill><a:prstDash val="dot"/></a:ln></c:spPr></c:marker><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>').series[0].marker,
  { symbol: "diamond", size: 8, fill: "#E11D48", line: { width: 1.5, fill: "#2563EB", style: "dotted" } },
  "JavaScript fallback import must expose the bounded direct line-series marker fill and outline.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Marker</c:v></c:tx><c:marker><c:symbol val="circle"/><c:spPr><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></c:spPr></c:marker><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser></c:lineChart></c:plotArea></c:chart></c:chartSpace>').series[0].marker,
  undefined,
  "JavaScript fallback import must not flatten theme marker graphs into the bounded marker model.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="stacked"/><c:varyColors val="1"/><c:ser><c:tx><c:v>Smooth</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:smooth val="0"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>').lineOptions,
  { grouping: "stacked", varyColors: true, smooth: false },
  "JavaScript fallback import must preserve chart-level grouping, vary-colors true, and explicit smooth false.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:varyColors val="false"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>').lineOptions,
  { grouping: "standard" },
  "JavaScript fallback import must normalize native vary-colors false to the omitted public default.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:grouping val="standard"/><c:varyColors val="1"/><c:varyColors val="0"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>').lineOptions,
  undefined,
  "JavaScript fallback import must not flatten duplicate vary-colors nodes.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:smooth val="1"/><c:smooth val="0"/></c:lineChart></c:plotArea></c:chart></c:chartSpace>').lineOptions,
  undefined,
  "JavaScript fallback import must not flatten duplicate smooth nodes.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:ser><c:tx><c:v>Labels</c:v></c:tx><c:cat><c:strLit><c:pt idx="0"><c:v>A</c:v></c:pt></c:strLit></c:cat><c:val><c:numLit><c:pt idx="0"><c:v>1</c:v></c:pt></c:numLit></c:val></c:ser><c:dLbls><c:dLblPos val="inEnd"/><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="1"/><c:showSerName val="1"/></c:dLbls></c:lineChart></c:plotArea></c:chart></c:chartSpace>').dataLabels,
  { showValue: true, showCategoryName: true, showSeriesName: true, position: "insideEnd" },
  "JavaScript fallback import must expose the bounded plot-level value/category/series-name/position label profile.",
);
assert.deepEqual(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:dLbls><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/></c:dLbls></c:lineChart></c:plotArea></c:chart></c:chartSpace>').dataLabels,
  { showValue: false, showCategoryName: false, showSeriesName: false },
  "JavaScript fallback import must preserve explicit false series-name visibility.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:dLbls><c:dLblPos val="floating"/><c:showVal val="1"/><c:showCatName val="0"/></c:dLbls></c:lineChart></c:plotArea></c:chart></c:chartSpace>').dataLabels,
  undefined,
  "JavaScript fallback import must not flatten unknown native data-label positions.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:dLbls><c:dLblPos val="t"/><c:dLblPos val="b"/><c:showVal val="1"/><c:showCatName val="0"/></c:dLbls></c:lineChart></c:plotArea></c:chart></c:chartSpace>').dataLabels,
  undefined,
  "JavaScript fallback import must not flatten duplicate native data-label positions.",
);
assert.equal(
  parseSpreadsheetChart('<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart><c:plotArea><c:lineChart><c:dLbls><c:showVal val="1"/><c:showCatName val="0"/><c:showPercent val="1"/></c:dLbls></c:lineChart></c:plotArea></c:chart></c:chartSpace>').dataLabels,
  undefined,
  "JavaScript fallback import must not flatten visible unsupported data-label semantics.",
);
assert.equal(toBinary(SpreadsheetChartAxisArtifactSchema, create(SpreadsheetChartAxisArtifactSchema, { minimum: 0 }))[0], 0x21, "Spreadsheet value-axis minimum must preserve explicit zero at optional field 4.");
assert.equal(toBinary(SpreadsheetChartSourceBindingSchema, create(SpreadsheetChartSourceBindingSchema, { semanticSha256: "x" }))[0], 0x3a, "Spreadsheet chart semantic hashes must use source-binding field 7.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { assetId: "asset/1" }))[0], 0x22, "Spreadsheet image assets must use image field 4.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { twoCellAnchor: { from: {}, to: {} } }))[0], 0x3a, "Spreadsheet two-cell anchors must use additive image field 7.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { absoluteAnchor: { widthEmu: 1n, heightEmu: 1n } }))[0], 0x42, "Spreadsheet absolute anchors must use additive image field 8.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { crop: { leftThousandthPercent: 1 } }))[0], 0x4a, "Spreadsheet image crops must use additive image field 9.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { effects: { grayscale: true } }))[0], 0x52, "Spreadsheet image effects must use additive image field 10.");
assert.equal(toBinary(SpreadsheetImageArtifactSchema, create(SpreadsheetImageArtifactSchema, { transform: { rotationAngle60000: 60_000 } }))[0], 0x5a, "Spreadsheet image transforms must use additive image field 11.");
assert.deepEqual([...toBinary(SpreadsheetImageTransformArtifactSchema, create(SpreadsheetImageTransformArtifactSchema, { flipHorizontal: false }))], [0x10, 0x00], "Spreadsheet picture transforms must preserve explicit false flips.");
assert.deepEqual([...toBinary(SpreadsheetOneCellAnchorArtifactSchema, create(SpreadsheetOneCellAnchorArtifactSchema, { widthEmu: 1n }))], [0x28, 0x01], "Spreadsheet image widths must use signed EMU field 5.");
assert.equal(toBinary(SpreadsheetImageSourceBindingSchema, create(SpreadsheetImageSourceBindingSchema, { semanticSha256: "x" }))[0], 0x2a, "Spreadsheet image semantic hashes must use source-binding field 5.");
assert.equal(toBinary(SpreadsheetWorksheetSourceBindingSchema, create(SpreadsheetWorksheetSourceBindingSchema, { semanticSha256: "x" }))[0], 0x2a, "Spreadsheet worksheet semantic hashes must use source-binding field 5.");
assert.equal(toBinary(SpreadsheetTableArtifactSchema, create(SpreadsheetTableArtifactSchema, { source: { tablePartPath: "xl/tables/table1.xml" } }))[0], 0x6a, "Spreadsheet table source bindings must use additive table field 13.");
assert.equal(toBinary(SpreadsheetTableArtifactSchema, create(SpreadsheetTableArtifactSchema, { columns: [{ name: "Revenue" }] }))[0], 0x72, "Spreadsheet rich table columns must use additive table field 14.");
assert.equal(toBinary(SpreadsheetTableArtifactSchema, create(SpreadsheetTableArtifactSchema, { filters: [{ columnIndex: 1, criteria: { case: "values", value: { values: ["x"] } } }] }))[0], 0x7a, "Spreadsheet table filters must use additive table field 15.");
assert.deepEqual([...toBinary(SpreadsheetTableArtifactSchema, create(SpreadsheetTableArtifactSchema, { sortState: { reference: "A2:B3", conditions: [{ reference: "B2:B3" }] } })).slice(0, 2)], [0x82, 0x01], "Spreadsheet table sort state must use additive table field 16.");
assert.deepEqual([...toBinary(SpreadsheetTableArtifactSchema, create(SpreadsheetTableArtifactSchema, { queryTable: { name: "Query", connectionId: 7 } })).slice(0, 2)], [0x8a, 0x01], "Spreadsheet QueryTable semantics must use additive table field 17.");
assert.deepEqual([...toBinary(SpreadsheetTableQueryArtifactSchema, create(SpreadsheetTableQueryArtifactSchema, { headers: false }))], [0x18, 0x00], "Spreadsheet QueryTable booleans must preserve explicit false values.");
assert.deepEqual([...toBinary(SpreadsheetTableQueryArtifactSchema, create(SpreadsheetTableQueryArtifactSchema, { refresh: {} })).slice(0, 2)], [0xc2, 0x01], "Spreadsheet QueryTable refresh semantics must use additive query field 24.");
assert.deepEqual([...toBinary(SpreadsheetTableQueryRefreshArtifactSchema, create(SpreadsheetTableQueryRefreshArtifactSchema, { preserveSortFilterLayout: false }))], [0x08, 0x00], "Spreadsheet QueryTable refresh booleans must preserve explicit false values.");
assert.deepEqual([...toBinary(SpreadsheetTableQueryRefreshArtifactSchema, create(SpreadsheetTableQueryRefreshArtifactSchema, { deletedFieldNames: ["x"] }))], [0x4a, 0x01, 0x78], "Spreadsheet QueryTable deleted-field history must use additive refresh field 9.");
assert.equal(toBinary(SpreadsheetTableQueryRefreshArtifactSchema, create(SpreadsheetTableQueryRefreshArtifactSchema, { sortState: { reference: "A2:B3" } }))[0], 0x52, "Spreadsheet QueryTable refresh-local sort state must use additive refresh field 10.");
assert.deepEqual([...toBinary(SpreadsheetTableQueryFieldArtifactSchema, create(SpreadsheetTableQueryFieldArtifactSchema, { dataBound: false }))], [0x18, 0x00], "Spreadsheet QueryTable field booleans must preserve explicit false values.");
assert.equal(toBinary(SpreadsheetTableSortStateArtifactSchema, create(SpreadsheetTableSortStateArtifactSchema, { conditions: [{ reference: "B2:B3" }] }))[0], 0x1a, "Spreadsheet sort conditions must use sort-state field 3.");
assert.equal(toBinary(SpreadsheetTableSortStateArtifactSchema, create(SpreadsheetTableSortStateArtifactSchema, { sortMethod: "stroke" }))[0], 0x22, "Spreadsheet locale sort methods must use additive sort-state field 4.");
assert.deepEqual([...toBinary(SpreadsheetTableSortStateArtifactSchema, create(SpreadsheetTableSortStateArtifactSchema, { columnSort: false }))], [0x28, 0x00], "Spreadsheet column-sort direction must use presence-aware sort-state field 5.");
assert.equal(toBinary(SpreadsheetTableColumnArtifactSchema, create(SpreadsheetTableColumnArtifactSchema, { totalsRowFormulaArray: true }))[0], 0x38, "Spreadsheet table totals-formula array state must use column field 7.");
assert.equal(toBinary(SpreadsheetTableFilterArtifactSchema, create(SpreadsheetTableFilterArtifactSchema, { criteria: { case: "values", value: { values: ["x"] } } }))[0], 0x12, "Spreadsheet value-filter criteria must use filter field 2.");
assert.equal(toBinary(SpreadsheetTableFilterArtifactSchema, create(SpreadsheetTableFilterArtifactSchema, { criteria: { case: "dynamic", value: { type: "today" } } }))[0], 0x22, "Spreadsheet dynamic-filter criteria must use additive filter field 4.");
assert.equal(toBinary(SpreadsheetTableFilterArtifactSchema, create(SpreadsheetTableFilterArtifactSchema, { criteria: { case: "top10", value: { top: true, value: 10 } } }))[0], 0x2a, "Spreadsheet Top10-filter criteria must use additive filter field 5.");
assert.equal(toBinary(SpreadsheetTableFilterArtifactSchema, create(SpreadsheetTableFilterArtifactSchema, { criteria: { case: "icon", value: { iconSet: "3Arrows", iconId: 0 } } }))[0], 0x32, "Spreadsheet icon-filter criteria must use additive filter field 6.");
assert.equal(toBinary(SpreadsheetTableFilterArtifactSchema, create(SpreadsheetTableFilterArtifactSchema, { criteria: { case: "color", value: { target: { case: "cellColor", value: true }, color: { source: { case: "rgb", value: "E11D48" } } } } }))[0], 0x3a, "Spreadsheet color-filter criteria must use additive filter field 7.");
assert.equal(toBinary(SpreadsheetTableIconArtifactSchema, create(SpreadsheetTableIconArtifactSchema, { iconId: 0 }))[0], 0x10, "Spreadsheet icon IDs must use presence-aware field 2 even when zero.");
assert.equal(toBinary(SpreadsheetTableColorArtifactSchema, create(SpreadsheetTableColorArtifactSchema, { target: { case: "fontColor", value: true } }))[0], 0x10, "Spreadsheet color targets must use an explicit oneof rather than a scalar default.");
assert.equal(toBinary(SpreadsheetTableSortConditionArtifactSchema, create(SpreadsheetTableSortConditionArtifactSchema, { icon: { iconSet: "3Arrows" } }))[0], 0x1a, "Spreadsheet icon-sort selectors must use additive condition field 3.");
assert.equal(toBinary(SpreadsheetTableSortConditionArtifactSchema, create(SpreadsheetTableSortConditionArtifactSchema, { color: { target: { case: "cellColor", value: true } } }))[0], 0x22, "Spreadsheet color-sort selectors must use additive condition field 4.");
assert.equal(toBinary(SpreadsheetTableSortConditionArtifactSchema, create(SpreadsheetTableSortConditionArtifactSchema, { customList: "high,low" }))[0], 0x2a, "Spreadsheet custom-list sorts must use additive condition field 5.");
assert.equal(toBinary(SpreadsheetTableValueFilterArtifactSchema, create(SpreadsheetTableValueFilterArtifactSchema, { dateGroups: [{ year: 2026, month: 7, grouping: "month" }] }))[0], 0x1a, "Spreadsheet grouped-date criteria must use additive value-filter field 3.");
assert.equal(toBinary(SpreadsheetTableValueFilterArtifactSchema, create(SpreadsheetTableValueFilterArtifactSchema, { calendarType: "gregorian" }))[0], 0x22, "Spreadsheet grouped-date calendar must use additive value-filter field 4.");
assert.equal(toBinary(DocumentBlockSchema, create(DocumentBlockSchema, { content: { case: "hyperlink", value: { text: "x", target: { case: "externalUri", value: "https://example.test" } } } }))[0], 0x6a, "Document hyperlinks must use additive block field 13.");
assert.equal(toBinary(DocumentBlockSchema, create(DocumentBlockSchema, { content: { case: "field", value: { instruction: "PAGE", display: "1" } } }))[0], 0x72, "Document fields must use additive block field 14.");
assert.equal(toBinary(DocumentSourceBindingSchema, create(DocumentSourceBindingSchema, { residualSha256: "x" }))[0], 0x2a, "Document residual hashes must use additive field 5.");
assert.equal(toBinary(DocumentHyperlinkSchema, create(DocumentHyperlinkSchema, { target: { case: "externalUri", value: "https://example.test" } }))[0], 0x12, "Document external hyperlink targets must use field 2.");
assert.equal(toBinary(DocumentFieldSchema, create(DocumentFieldSchema, { instruction: "PAGE" }))[0], 0x0a, "Document field instructions must use field 1.");
assert.equal(toBinary(DocumentParagraphSchema, create(DocumentParagraphSchema, { numbering: { numberingId: 7 } }))[0], 0x1a, "Document paragraph numbering must use additive field 3.");
assert.equal(toBinary(DocumentNumberingSchema, create(DocumentNumberingSchema, { numberingId: 7 }))[0], 0x08, "Document numbering IDs must use field 1.");
assert.deepEqual([...toBinary(DocumentTableSchema, create(DocumentTableSchema, { gridColumns: 3 }))], [0x10, 0x03], "Document table grid width must use additive field 2.");
assert.equal(toBinary(DocumentTableSchema, create(DocumentTableSchema, { formatting: { widthDxa: 1 } }))[0], 0x1a, "Document table formatting must use additive field 3.");
assert.deepEqual([...toBinary(DocumentTableFormattingSchema, create(DocumentTableFormattingSchema, { widthDxa: 1 }))], [0x08, 0x01], "Document table formatting width must use field 1.");
assert.deepEqual([...toBinary(DocumentTableCellMarginsSchema, create(DocumentTableCellMarginsSchema, { start: 1 }))], [0x18, 0x01], "Document table start margin must use field 3.");
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
assert.equal(toBinary(PresentationPlaceholderSchema, create(PresentationPlaceholderSchema, { directFrame: { widthEmu: 1n } }))[0], 0x3a, "Presentation placeholder direct frames must use additive field 7.");
assert.deepEqual([...toBinary(PresentationPlaceholderFrameSchema, create(PresentationPlaceholderFrameSchema, { leftEmu: 1n, topEmu: 2n, widthEmu: 3n, heightEmu: 4n }))], [0x08, 0x01, 0x10, 0x02, 0x18, 0x03, 0x20, 0x04], "Presentation placeholder frames must retain their atomic four-coordinate field contract.");
assert.equal(toBinary(PresentationMasterSourceBindingSchema, create(PresentationMasterSourceBindingSchema, { backgroundSemanticSha256: "x" }))[0], 0x3a, "Presentation master background hashes must use additive field 7.");
assert.deepEqual([...toBinary(PresentationMasterSourceBindingSchema, create(PresentationMasterSourceBindingSchema, { backgroundEditable: true }))], [0x40, 0x01], "Presentation master background editability must use additive field 8.");
assert.equal(toBinary(PresentationLayoutSourceBindingSchema, create(PresentationLayoutSourceBindingSchema, { backgroundSemanticSha256: "x" }))[0], 0x2a, "Presentation layout background hashes must use additive field 5.");
assert.deepEqual([...toBinary(PresentationLayoutSourceBindingSchema, create(PresentationLayoutSourceBindingSchema, { backgroundEditable: true }))], [0x30, 0x01], "Presentation layout background editability must use additive field 6.");
assert.equal(toBinary(PresentationBackgroundSchema, create(PresentationBackgroundSchema, { color: { case: "colorScheme", value: "accent1" }, kind: { case: "styleReferenceIndex", value: 1001 } }))[0], 0x12, "Presentation background theme colors must retain field 2.");

const workbook = Workbook.create({
  dateSystem: "1904",
  calculation: {
    mode: "automaticExceptTables",
    calculateOnSave: false,
    fullCalculationOnLoad: true,
    forceFullCalculation: true,
    iteration: { enabled: true, maxIterations: 100, maxChange: 0.001 },
    fullPrecision: false,
  },
  theme: {
    name: "OpenChestnut Theme",
    colors: {
      dk1: "#101820", lt1: "#F8FAFC", dk2: "#1E3A5F", lt2: "#E2E8F0",
      accent1: "#0F766E", accent2: "#C2410C", accent3: "#4D7C0F", accent4: "#7E22CE",
      accent5: "#0369A1", accent6: "#BE123C", hlink: "#1D4ED8", folHlink: "#7E22CE",
    },
  },
});
const summary = workbook.worksheets.add("Summary");
summary.getRange("A1:B2").values = [["Quarter", 42.5], [true, null]];
summary.getRange("B2").formulas = [["=B1*2"]];
summary.getRange("A1:B1").format = {
  fill: { patternType: "darkGrid", foreground: { theme: 4, tint: 0.4 }, background: "#E2E8F0" },
  font: { bold: true, italic: false, underline: "double", strike: false, color: "#FFFFFF", size: 13, name: "Aptos Display" },
  alignment: { horizontal: "center", vertical: "bottom", wrapText: true, textRotation: 15, indent: 1, shrinkToFit: false, readingOrder: 1 },
  border: {
    left: { style: "thin", color: { indexed: 8 } },
    right: { style: "thin", color: { indexed: 8 } },
    top: { style: "thin", color: { auto: true } },
    bottom: { style: "double", color: "#38BDF8" },
  },
  protection: { locked: false, hidden: true },
};
summary.getRange("B1").format.numberFormat = "0.000 \"units\"";
summary.getRange("B2").format.numberFormat = "0.00%";
summary.freezePanes.freezeRows(1).freezeColumns(1);
summary.showGridLines = false;
summary.columnDimensions.set(0, { width: 18, bestFit: true });
summary.rowDimensions.set(0, { height: 24 });
summary.mergeCells("A3:B3");
const summaryImage = summary.images.add({
  name: "Quarter mark",
  alt: "Quarterly performance",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  anchor: { from: { row: 3, col: 2, rowOffsetPx: 10, colOffsetPx: 5 }, extent: { widthPx: 120, heightPx: 80 } },
});
const summaryChart = summary.charts.add("line", {
  name: "Quarter chart",
  title: "Quarter trend",
  lineOptions: { grouping: "stacked", smooth: true, varyColors: true },
  dataLabels: { showValue: true, showCategoryName: false, showSeriesName: true, position: "outsideEnd" },
  hasLegend: true,
  categories: ["Q1", "Q2"],
  series: [{ name: "Revenue", values: [42.5, 85], fill: "#F472B6", line: { fill: "#0EA5E9", style: "dashed", width: 2 }, marker: { symbol: "diamond", size: 8, fill: "#FDE68A", line: { fill: "#BE123C", style: "dotted", width: 1.5 } } }],
  xAxis: { axisType: "textAxis", title: { text: "Quarter" }, numberFormatCode: "@", tickLabelInterval: 2 },
  yAxis: { axisType: "valueAxis", title: { text: "Revenue" }, numberFormatCode: "$#,##0.0", min: 0, max: 100, majorUnit: 25 },
  position: { left: 420, top: 40, width: 360, height: 220 },
});
summaryChart.series.items[0].categoryFormula = "'Summary'!$A$1:$A$2";
summaryChart.series.items[0].formula = "'Summary'!$B$1:$B$2";
summary.sortState = {
  reference: "A1:B2",
  caseSensitive: true,
  sortMethod: "stroke",
  columnSort: true,
  conditions: [{ reference: "A2:B2", descending: true, customList: "Actual,Plan" }, { reference: "A1:B1", descending: false }],
};
const details = workbook.worksheets.add("Details");
details.visibility = "hidden";
details.getRange("A1:B3").values = [["Status", "Value"], ["ready", 2], ["pending", 1]];
workbook.definedNames.add({ id: "defined-name/summary-data", name: "SummaryData", refersTo: "Summary!$A$1:$B$2", comment: "Summary data body", hidden: false });
workbook.definedNames.add({ id: "defined-name/status-data", name: "StatusData", refersTo: "Details!$A$2:$B$3", scope: "Details", hidden: true });
const detailsTable = details.tables.add({ range: "A1:B3", name: "StatusTable", hasHeaders: true, style: "TableStyleMedium4" });
detailsTable.showFirstColumn = true;
detailsTable.showBandedColumns = true;
detailsTable.columnDefinitions = [
  { name: "Status" },
  { name: "Value", calculatedColumnFormula: "=LEN([@Status])" },
];
detailsTable.filters = [
  { columnIndex: 0, kind: "values", values: ["ready"], includeBlank: true },
  { columnIndex: 1, kind: "custom", matchAll: true, criteria: [{ operator: "greaterThanOrEqual", value: "1" }, { operator: "lessThanOrEqual", value: "2" }] },
];
detailsTable.sortState = {
  reference: "A2:B3",
  caseSensitive: true,
  sortMethod: "stroke",
  conditions: [{ reference: "B2:B3", descending: true, customList: "2,1" }, { reference: "A2:A3", descending: false }],
};
const advancedFilters = workbook.worksheets.add("Advanced Filters");
advancedFilters.visibility = "veryHidden";
advancedFilters.getRange("A1:C3").values = [
  ["Date", "Status", "Score"],
  [45853, "ready", 95],
  [45854, "pending", 80],
];
advancedFilters.getRange("A2:A3").format.numberFormat = "yyyy-mm-dd";
const advancedFilterTable = advancedFilters.tables.add({ range: "A1:C3", name: "AdvancedFilterTable", hasHeaders: true, style: "TableStyleMedium6" });
advancedFilterTable.filters = [
  { columnIndex: 0, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 15 }] },
  { columnIndex: 1, kind: "dynamic", type: "today", value: 45853, maxValue: 45854 },
  { columnIndex: 2, kind: "top10", top: true, percent: true, value: 10, filterValue: 95 },
];
const iconRules = workbook.worksheets.add("Icon Rules");
iconRules.getRange("A1:B3").values = [["Trend", "Rating"], [1, 5], [2, 3]];
const iconTable = iconRules.tables.add({ range: "A1:B3", name: "IconRuleTable", hasHeaders: true, style: "TableStyleMedium7" });
iconTable.filters = [
  { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 0 },
  { columnIndex: 1, kind: "icon", iconSet: "3Flags" },
];
iconTable.sortState = {
  reference: "A2:B3",
  caseSensitive: false,
  conditions: [
    { reference: "B2:B3", descending: true, kind: "icon", iconSet: "5Rating", iconId: 4 },
    { reference: "A2:A3", descending: false, kind: "icon", iconSet: "3Symbols2" },
  ],
};
const colorRules = workbook.worksheets.add("Color Rules");
colorRules.getRange("A1:B3").values = [["Fill", "Font"], [1, 5], [2, 3]];
const colorTable = colorRules.tables.add({ range: "A1:B3", name: "ColorRuleTable", hasHeaders: true, style: "TableStyleMedium8" });
colorTable.filters = [
  { columnIndex: 0, kind: "color", target: "cell", color: "#E11D48" },
  { columnIndex: 1, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
];
colorTable.sortState = {
  reference: "A2:B3",
  caseSensitive: false,
  conditions: [
    { reference: "B2:B3", descending: true, kind: "color", target: "font", color: { theme: 4, tint: -0.25 } },
    { reference: "A2:A3", descending: false, kind: "color", target: "cell", color: "#E11D48" },
  ],
};
workbook.worksheets.setActiveWorksheet(iconRules);
workbook.worksheets.setSelectedWorksheets([summary, iconRules]);
workbook.windows.add({ activeWorksheet: colorRules, selectedWorksheets: [summary, colorRules] });

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
const exportedZip = await JSZip.loadAsync(exported.bytes);
const exportedWorkbookXml = await exportedZip.file("xl/workbook.xml").async("text");
assert.match(exportedWorkbookXml, /<x:workbookView\b[^>]*activeTab="3"/);
assert.match(exportedWorkbookXml, /<x:workbookView\b[^>]*activeTab="4"/);
assert.match(await exportedZip.file("xl/worksheets/sheet1.xml").async("text"), /<x:sheetView\b[^>]*tabSelected="1"/);
assert.match(await exportedZip.file("xl/worksheets/sheet4.xml").async("text"), /<x:sheetView\b[^>]*tabSelected="1"/);
assert.match(await exportedZip.file("xl/worksheets/sheet5.xml").async("text"), /<x:sheetView\b(?=[^>]*workbookViewId="1")(?=[^>]*tabSelected="1")[^>]*>/);
assert.match(exportedWorkbookXml, /<x:sheet\b[^>]*name="Details"[^>]*state="hidden"/);
assert.match(exportedWorkbookXml, /<x:sheet\b[^>]*name="Advanced Filters"[^>]*state="veryHidden"/);
assert.match(exportedWorkbookXml, /<x:definedName name="SummaryData" comment="Summary data body" hidden="0">Summary!\$A\$1:\$B\$2<\/x:definedName>/);
assert.match(exportedWorkbookXml, /<x:definedName name="StatusData" localSheetId="1" hidden="1">Details!\$A\$2:\$B\$3<\/x:definedName>/);
const summaryWorksheetXml = await exportedZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(summaryWorksheetXml, /<x:sortState\b[^>]*ref="A1:B2"/);
assert.match(summaryWorksheetXml, /<x:sortState\b[^>]*columnSort="1"/);
assert.match(summaryWorksheetXml, /<x:sortCondition\b[^>]*ref="A2:B2"[^>]*customList="Actual,Plan"/);
assert.match(summaryWorksheetXml, /<x:drawing\b[^>]*r:id="[^"]+"/);
const summaryDrawingXml = await exportedZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(summaryDrawingXml, /<xdr:oneCellAnchor>/);
assert.match(summaryDrawingXml, /name="Quarter mark"/);
assert.match(summaryDrawingXml, /descr="Quarterly performance"/);
assert.match(summaryDrawingXml, /<xdr:row>3<\/xdr:row>/);
assert.match(summaryDrawingXml, /<xdr:col>2<\/xdr:col>/);
assert.match(summaryDrawingXml, /cx="1143000" cy="762000"/);
const summaryMediaPath = Object.keys(exportedZip.files).find((path) => /^xl\/media\/[^/]+\.png$/i.test(path));
assert.ok(summaryMediaPath, "OpenChestnut must author a worksheet PNG ImagePart");
assert.deepEqual([...(await exportedZip.file(summaryMediaPath).async("uint8array"))], [...Buffer.from(summaryImage.dataUrl.split(",")[1], "base64")]);
assert.match(await exportedZip.file("xl/tables/table1.xml").async("text"), /<x:calculatedColumnFormula>LEN\(\[@Status\]\)<\/x:calculatedColumnFormula>/);
assert.match(await exportedZip.file("xl/tables/table1.xml").async("text"), /<x:filterColumn colId="0"><x:filters blank="1"><x:filter val="ready"\s*\/><\/x:filters><\/x:filterColumn>/);
assert.match(await exportedZip.file("xl/tables/table1.xml").async("text"), /<x:customFilters and="1"><x:customFilter operator="greaterThanOrEqual" val="1"\s*\/><x:customFilter operator="lessThanOrEqual" val="2"\s*\/><\/x:customFilters>/);
assert.match(await exportedZip.file("xl/tables/table1.xml").async("text"), /<x:sortState ref="A2:B3" caseSensitive="1" sortMethod="stroke"><x:sortCondition ref="B2:B3" descending="1" customList="2,1"\s*\/><x:sortCondition ref="A2:A3"\s*\/><\/x:sortState>/);
const advancedFilterXml = await exportedZip.file("xl/tables/table2.xml").async("text");
assert.match(advancedFilterXml, /<x:filters calendarType="gregorian"><x:dateGroupItem year="2026" dateTimeGrouping="day" month="7" day="15"\s*\/><\/x:filters>/);
assert.match(advancedFilterXml, /<x:dynamicFilter type="today" val="45853" maxVal="45854"\s*\/>/);
assert.match(advancedFilterXml, /<x:top10 top="1" percent="1" val="10" filterVal="95"\s*\/>/);
const iconRuleXml = await exportedZip.file("xl/tables/table3.xml").async("text");
assert.match(iconRuleXml, /<x:iconFilter iconSet="3Arrows" iconId="0"\s*\/>/);
assert.match(iconRuleXml, /<x:iconFilter iconSet="3Flags"\s*\/>/);
assert.match(iconRuleXml, /<x:sortCondition ref="B2:B3" descending="1" sortBy="icon" iconSet="5Rating" iconId="4"\s*\/>/);
assert.match(iconRuleXml, /<x:sortCondition ref="A2:A3" sortBy="icon" iconSet="3Symbols2"\s*\/>/);
const colorRuleXml = await exportedZip.file("xl/tables/table4.xml").async("text");
assert.match(colorRuleXml, /<x:colorFilter dxfId="0" cellColor="1"\s*\/>/);
assert.match(colorRuleXml, /<x:colorFilter dxfId="1" cellColor="0"\s*\/>/);
assert.match(colorRuleXml, /<x:sortCondition ref="B2:B3" descending="1" sortBy="fontColor" dxfId="1"\s*\/>/);
assert.match(colorRuleXml, /<x:sortCondition ref="A2:A3" sortBy="cellColor" dxfId="0"\s*\/>/);

const querySourceBytes = await addQueryTableGraph(exported.bytes);
const querySourceZip = await JSZip.loadAsync(querySourceBytes);
const queryRelationshipXml = await querySourceZip.file("xl/tables/_rels/table1.xml.rels").async("uint8array");
const queryImported = await importXlsxWithOpenChestnut(querySourceBytes);
assert.deepEqual(queryImported.connections, [{
  connectionId: 7,
  name: "Fixture warehouse",
  type: 5,
  refreshedVersion: 8,
  description: "Read-only warehouse source",
  keepAlive: false,
  intervalMinutes: 30,
  background: true,
  refreshOnLoad: false,
  saveData: true,
}]);
assert.equal(queryImported.resolve("connection/7"), queryImported.connections[0]);
assert.match(queryImported.inspect({ kind: "connection" }).ndjson, /"name":"Fixture warehouse"/);
const queryTable = queryImported.worksheets.getItem("Details").tables.getItemOrNullObject("StatusTable");
assert.deepEqual(queryTable.queryTable, {
  name: "Warehouse metrics",
  connectionId: 7,
  headers: true,
  rowNumbers: false,
  disableRefresh: false,
  backgroundRefresh: true,
  firstBackgroundRefresh: false,
  refreshOnLoad: false,
  growShrinkType: "insertClear",
  fillFormulas: false,
  removeDataOnSave: false,
  disableEdit: false,
  preserveFormatting: true,
  adjustColumnWidth: true,
  intermediate: false,
  refresh: {
    preserveSortFilterLayout: true,
    fieldIdWrapped: false,
    headersInLastRefresh: true,
    minimumVersion: 0,
    nextId: 3,
    unboundColumnsLeft: 0,
    unboundColumnsRight: 0,
    fields: [
      { id: 1, name: "Status", dataBound: true, fillFormulas: false, clipped: false, tableColumnId: 1 },
      { id: 2, name: "Value", dataBound: true, tableColumnId: 2 },
    ],
    deletedFieldNames: ["Legacy Status", "Legacy Value"],
    sortState: {
      reference: "A2:B3",
      caseSensitive: true,
      sortMethod: "stroke",
      conditions: [
        { reference: "B2:B3", descending: true, customList: "ready,pending" },
        { reference: "A2:A3", descending: false, kind: "icon", iconSet: "3Arrows", iconId: 0 },
      ],
    },
  },
});
assert.match(queryTable.inspectRecord().queryTable.name, /Warehouse metrics/);
queryTable.name = "QueriedStatusTable";
queryTable.queryTable.name = "Warehouse metrics refreshed";
queryTable.queryTable.backgroundRefresh = false;
queryTable.queryTable.refreshOnLoad = true;
queryTable.queryTable.autoFormatId = 3;
queryTable.queryTable.applyFontFormats = true;
queryTable.queryTable.refresh.preserveSortFilterLayout = false;
queryTable.queryTable.refresh.headersInLastRefresh = false;
queryTable.queryTable.refresh.minimumVersion = 1;
queryTable.queryTable.refresh.fields[0].name = "State";
queryTable.queryTable.refresh.fields[0].dataBound = false;
queryTable.queryTable.refresh.fields[0].fillFormulas = true;
queryTable.queryTable.refresh.fields[1].clipped = true;
queryTable.queryTable.refresh.deletedFieldNames[0] = "Legacy State";
queryTable.queryTable.refresh.sortState.caseSensitive = false;
queryTable.queryTable.refresh.sortState.sortMethod = "pinYin";
queryTable.queryTable.refresh.sortState.columnSort = true;
queryTable.queryTable.refresh.sortState.conditions[0].reference = "A3:B3";
queryTable.queryTable.refresh.sortState.conditions[0].descending = false;
queryTable.queryTable.refresh.sortState.conditions[0].customList = "pending,ready";
queryTable.queryTable.refresh.sortState.conditions[1].reference = "A2:B2";
queryTable.queryTable.refresh.sortState.conditions[1].iconId = 1;
Object.assign(queryImported.connections[0], {
  name: "Fixture warehouse curated",
  description: "Curated without executing the source",
  keepAlive: true,
  intervalMinutes: 45,
  background: false,
  refreshOnLoad: true,
  saveData: false,
});
const queryExported = await exportXlsxWithOpenChestnut(queryImported, { recalculate: false });
const queryOutputZip = await JSZip.loadAsync(queryExported.bytes);
assert.deepEqual(await queryOutputZip.file("xl/tables/_rels/table1.xml.rels").async("uint8array"), queryRelationshipXml, "the table/query relationship must retain its source identity");
const queryOutputConnectionXml = await queryOutputZip.file("xl/connections.xml").async("text");
assert.match(queryOutputConnectionXml, /name="Fixture warehouse curated"/);
assert.match(queryOutputConnectionXml, /description="Curated without executing the source"/);
assert.match(queryOutputConnectionXml, /keepAlive="1"/);
assert.match(queryOutputConnectionXml, /interval="45"/);
assert.match(queryOutputConnectionXml, /background="0"/);
assert.match(queryOutputConnectionXml, /refreshOnLoad="1"/);
assert.match(queryOutputConnectionXml, /saveData="0"/);
assert.match(queryOutputConnectionXml, /Provider=Fixture.Provider;Data Source=fixture.invalid/);
assert.match(queryOutputConnectionXml, /SELECT Status, Value FROM Metrics/);
assert.match(queryOutputConnectionXml, /savePassword="0"/);
assert.match(queryOutputConnectionXml, /credentials="integrated"/);
assert.match(queryOutputConnectionXml, /<fixture:connectionOpaque value="kept"/);
assert.match(queryOutputConnectionXml, /id="8" name="Opaque companion" type="1"/);
assert.match(queryOutputConnectionXml, /Provider=Opaque.Provider;Data Source=opaque.invalid/);
const queryOutputXml = await queryOutputZip.file("xl/queryTables/queryTable1.xml").async("text");
assert.match(queryOutputXml, /name="Warehouse metrics refreshed"/);
assert.match(queryOutputXml, /backgroundRefresh="0"/);
assert.match(queryOutputXml, /refreshOnLoad="1"/);
assert.match(queryOutputXml, /autoFormatId="3"/);
assert.match(queryOutputXml, /applyFontFormats="1"/);
assert.match(queryOutputXml, /preserveSortFilterLayout="0"/);
assert.match(queryOutputXml, /headersInLastRefresh="0"/);
assert.match(queryOutputXml, /minimumVersion="1"/);
assert.match(queryOutputXml, /id="1" name="State" dataBound="0" tableColumnId="1" fillFormulas="1" clipped="0"/);
assert.match(queryOutputXml, /id="2" name="Value" dataBound="1" tableColumnId="2" clipped="1"/);
assert.match(queryOutputXml, /<x:queryTableFields count="2">/);
assert.match(queryOutputXml, /<x:deletedField name="Legacy State"/);
assert.match(queryOutputXml, /<x:deletedField name="Legacy Value"/);
assert.match(queryOutputXml, /<x:sortState ref="A2:B3" sortMethod="pinYin" columnSort="1">/);
assert.match(queryOutputXml, /<x:sortCondition ref="A3:B3" customList="pending,ready"/);
assert.match(queryOutputXml, /<x:sortCondition ref="A2:B2" sortBy="icon" iconSet="3Arrows" iconId="1"/);
assert.match(queryOutputXml, /<fixture:fieldOpaque value="kept"/);
assert.match(queryOutputXml, /<fixture:sortOpaque value="kept"/);
assert.match(queryOutputXml, /<fixture:opaque value="kept"/);
const queryReimported = await importXlsxWithOpenChestnut(queryExported);
assert.deepEqual(queryReimported.connections, [{
  connectionId: 7,
  name: "Fixture warehouse curated",
  type: 5,
  refreshedVersion: 8,
  description: "Curated without executing the source",
  keepAlive: true,
  intervalMinutes: 45,
  background: false,
  refreshOnLoad: true,
  saveData: false,
}]);
const queryReimportedTable = queryReimported.worksheets.getItem("Details").tables.getItemOrNullObject("QueriedStatusTable");
assert.equal(queryReimportedTable.queryTable.name, "Warehouse metrics refreshed");
assert.equal(queryReimportedTable.queryTable.backgroundRefresh, false);
assert.equal(queryReimportedTable.queryTable.refreshOnLoad, true);
assert.equal(queryReimportedTable.queryTable.autoFormatId, 3);
assert.equal(queryReimportedTable.queryTable.applyFontFormats, true);
assert.equal(queryReimportedTable.queryTable.refresh.preserveSortFilterLayout, false);
assert.equal(queryReimportedTable.queryTable.refresh.headersInLastRefresh, false);
assert.equal(queryReimportedTable.queryTable.refresh.minimumVersion, 1);
assert.deepEqual(queryReimportedTable.queryTable.refresh.fields, [
  { id: 1, name: "State", dataBound: false, fillFormulas: true, clipped: false, tableColumnId: 1 },
  { id: 2, name: "Value", dataBound: true, clipped: true, tableColumnId: 2 },
]);
assert.deepEqual(queryReimportedTable.queryTable.refresh.deletedFieldNames, ["Legacy State", "Legacy Value"]);
assert.deepEqual(queryReimportedTable.queryTable.refresh.sortState, {
  reference: "A2:B3",
  caseSensitive: false,
  sortMethod: "pinYin",
  columnSort: true,
  conditions: [
    { reference: "A3:B3", descending: false, customList: "pending,ready" },
    { reference: "A2:B2", descending: false, kind: "icon", iconSet: "3Arrows", iconId: 1 },
  ],
});
const queryJavaScriptFallback = await SpreadsheetFile.importXlsx(queryExported);
assert.deepEqual(queryJavaScriptFallback.connections, queryReimported.connections);
assert.match(queryJavaScriptFallback.inspect({ kind: "connection" }).ndjson, /"intervalMinutes":45/);
assert.equal(queryJavaScriptFallback.worksheets.getItem("Details").tables.items[0].queryTable.name, "Warehouse metrics refreshed");
assert.deepEqual(queryJavaScriptFallback.worksheets.getItem("Details").tables.items[0].queryTable.refresh.fields, queryReimportedTable.queryTable.refresh.fields);
assert.deepEqual(queryJavaScriptFallback.worksheets.getItem("Details").tables.items[0].queryTable.refresh.deletedFieldNames, queryReimportedTable.queryTable.refresh.deletedFieldNames);
assert.deepEqual(queryJavaScriptFallback.worksheets.getItem("Details").tables.items[0].queryTable.refresh.sortState, queryReimportedTable.queryTable.refresh.sortState);
await assert.rejects(
  SpreadsheetFile.exportXlsx(queryJavaScriptFallback),
  /JavaScript XLSX codec cannot author or source-preserve QueryTable\/external-connection graphs/,
);

const invalidQueryConnection = await importXlsxWithOpenChestnut(querySourceBytes);
invalidQueryConnection.worksheets.getItem("Details").tables.items[0].queryTable.connectionId = 999;
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidQueryConnection, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /does not identify a connection/i.test(error.message),
);
const reboundWorkbookConnection = await importXlsxWithOpenChestnut(querySourceBytes);
reboundWorkbookConnection.connections[0].connectionId = 8;
await assert.rejects(
  exportXlsxWithOpenChestnut(reboundWorkbookConnection, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_connection" && /id\/type\/version identity/i.test(error.message),
);
const removedWorkbookConnection = await importXlsxWithOpenChestnut(querySourceBytes);
removedWorkbookConnection.connections.pop();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedWorkbookConnection, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_connection" && /cannot remove imported connection/i.test(error.message),
);
const invalidWorkbookConnectionInterval = await importXlsxWithOpenChestnut(querySourceBytes);
invalidWorkbookConnectionInterval.connections[0].intervalMinutes = 32_768;
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidWorkbookConnectionInterval, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_connection" && /bounded source-editable profile/i.test(error.message),
);
const removedQuery = await importXlsxWithOpenChestnut(querySourceBytes);
removedQuery.worksheets.getItem("Details").tables.items[0].queryTable = undefined;
await assert.rejects(
  exportXlsxWithOpenChestnut(removedQuery, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /add or remove a worksheet QueryTable graph/i.test(error.message),
);
const reboundQueryField = await importXlsxWithOpenChestnut(querySourceBytes);
reboundQueryField.worksheets.getItem("Details").tables.items[0].queryTable.refresh.fields[0].id = 99;
await assert.rejects(
  exportXlsxWithOpenChestnut(reboundQueryField, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /reorder or rebind/i.test(error.message),
);
const removedQueryField = await importXlsxWithOpenChestnut(querySourceBytes);
removedQueryField.worksheets.getItem("Details").tables.items[0].queryTable.refresh.fields.pop();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedQueryField, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /add or remove query refresh fields/i.test(error.message),
);
const invalidQueryNextId = await importXlsxWithOpenChestnut(querySourceBytes);
invalidQueryNextId.worksheets.getItem("Details").tables.items[0].queryTable.refresh.nextId = 2;
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidQueryNextId, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /unused positive field ID/i.test(error.message),
);
const removedDeletedQueryField = await importXlsxWithOpenChestnut(querySourceBytes);
removedDeletedQueryField.worksheets.getItem("Details").tables.items[0].queryTable.refresh.deletedFieldNames.pop();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedDeletedQueryField, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /add or remove query refresh deleted fields/i.test(error.message),
);
const removedRefreshSortCondition = await importXlsxWithOpenChestnut(querySourceBytes);
removedRefreshSortCondition.worksheets.getItem("Details").tables.items[0].queryTable.refresh.sortState.conditions.pop();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedRefreshSortCondition, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /add or remove query refresh sort conditions/i.test(error.message),
);
const invalidRefreshSortRange = await importXlsxWithOpenChestnut(querySourceBytes);
invalidRefreshSortRange.worksheets.getItem("Details").tables.items[0].queryTable.refresh.sortState.reference = "A2:C3";
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidRefreshSortRange, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /contained in the source table range/i.test(error.message),
);
const fabricatedQueryWorkbook = Workbook.create();
const fabricatedQuerySheet = fabricatedQueryWorkbook.worksheets.add("Query");
fabricatedQuerySheet.getRange("A1:B2").values = [["Key", "Value"], ["x", 1]];
fabricatedQuerySheet.tables.add({ range: "A1:B2", name: "FabricatedQuery", queryTable: { name: "Unsafe", connectionId: 1 } });
await assert.rejects(
  exportXlsxWithOpenChestnut(fabricatedQueryWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /cannot fabricate/i.test(error.message),
);
const fabricatedConnectionWorkbook = Workbook.create({ connections: [{ connectionId: 7, name: "Unsafe", type: 5, refreshedVersion: 8 }] });
fabricatedConnectionWorkbook.worksheets.add("Connection").getRange("A1").values = [["x"]];
await assert.rejects(
  exportXlsxWithOpenChestnut(fabricatedConnectionWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_connection" && /cannot fabricate/i.test(error.message),
);

const imported = await importXlsxWithOpenChestnut(exported);
assert.equal(imported.dateSystem, "1904");
assert.equal(imported.theme.name, "OpenChestnut Theme");
assert.equal(imported.theme.colors.accent1, "#0F766E");
assert.deepEqual(imported.calculation, workbook.calculation);
assert.equal(imported.worksheets.items.length, 5);
assert.deepEqual(imported.worksheets.items.map((item) => item.visibility), ["visible", "hidden", "veryHidden", "visible", "visible"]);
assert.equal(imported.worksheets.getActiveWorksheet().name, "Icon Rules");
assert.deepEqual(imported.worksheets.getSelectedWorksheets().map((sheet) => sheet.name), ["Summary", "Icon Rules"]);
assert.equal(imported.windows.count, 2);
assert.equal(imported.windows.getItemAt(1).getActiveWorksheet().name, "Color Rules");
assert.deepEqual(imported.windows.getItemAt(1).getSelectedWorksheets().map((sheet) => sheet.name), ["Summary", "Color Rules"]);
assert.throws(() => { imported.worksheets.getActiveWorksheet().visibility = "hidden"; }, /select another active worksheet first/i);
assert.throws(() => { imported.worksheets.getItem("Summary").visibility = "hidden"; }, /selected worksheet/i);
imported.worksheets.setSelectedWorksheets(["Summary", "Icon Rules", "Color Rules"]);
const groupEdited = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
const groupEditedZip = await JSZip.loadAsync(groupEdited.bytes);
assert.match(await groupEditedZip.file("xl/worksheets/sheet1.xml").async("text"), /<x:sheetView\b[^>]*tabSelected="1"/);
assert.match(await groupEditedZip.file("xl/worksheets/sheet4.xml").async("text"), /<x:sheetView\b[^>]*tabSelected="1"/);
assert.match(await groupEditedZip.file("xl/worksheets/sheet5.xml").async("text"), /<x:sheetView\b[^>]*tabSelected="1"/);
assert.deepEqual((await importXlsxWithOpenChestnut(groupEdited)).worksheets.getSelectedWorksheets().map((sheet) => sheet.name), ["Summary", "Icon Rules", "Color Rules"]);
imported.worksheets.setActiveWorksheet("Color Rules");
const activeEdited = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
const activeEditedZip = await JSZip.loadAsync(activeEdited.bytes);
assert.match(await activeEditedZip.file("xl/workbook.xml").async("text"), /<x:workbookView\b[^>]*activeTab="4"/);
const activeEditedImport = await importXlsxWithOpenChestnut(activeEdited);
assert.equal(activeEditedImport.worksheets.getActiveWorksheet().name, "Color Rules");
assert.deepEqual(activeEditedImport.worksheets.getSelectedWorksheets().map((sheet) => sheet.name), ["Color Rules"]);
assert.equal(activeEditedImport.windows.getItemAt(1).getActiveWorksheet().name, "Color Rules");
imported.windows.getItemAt(1).setActiveWorksheet("Summary");
const secondaryEdited = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
const secondaryEditedZip = await JSZip.loadAsync(secondaryEdited.bytes);
assert.match(await secondaryEditedZip.file("xl/workbook.xml").async("text"), /<x:workbookView\b[^>]*activeTab="0"/);
const secondaryEditedImport = await importXlsxWithOpenChestnut(secondaryEdited);
assert.equal(secondaryEditedImport.windows.getItemAt(1).getActiveWorksheet().name, "Summary");
assert.deepEqual(secondaryEditedImport.windows.getItemAt(1).getSelectedWorksheets().map((sheet) => sheet.name), ["Summary"]);
assert.deepEqual(imported.definedNames.toJSON(), [
  { id: "defined-name/1", name: "SummaryData", refersTo: "Summary!$A$1:$B$2", scope: undefined, comment: "Summary data body", hidden: false },
  { id: "defined-name/2", name: "StatusData", refersTo: "Details!$A$2:$B$3", scope: "Details", comment: undefined, hidden: true },
]);
assert.match(imported.inspect({ kind: "definedName", target: "SummaryData" }).ndjson, /"hidden":false/);
const importedTable = imported.worksheets.getItem("Details").tables.getItemOrNullObject("StatusTable");
assert.equal(importedTable.isNullObject, undefined);
assert.equal(importedTable.range, "A1:B3");
assert.deepEqual(importedTable.columnNames, ["Status", "Value"]);
assert.equal(importedTable.columnDefinitions[1].calculatedColumnFormula, "=LEN([@Status])");
assert.deepEqual(importedTable.filters, [
  { columnIndex: 0, kind: "values", values: ["ready"], includeBlank: true },
  { columnIndex: 1, kind: "custom", matchAll: true, criteria: [{ operator: "greaterThanOrEqual", value: "1" }, { operator: "lessThanOrEqual", value: "2" }] },
]);
assert.deepEqual(importedTable.sortState, {
  reference: "A2:B3",
  caseSensitive: true,
  sortMethod: "stroke",
  conditions: [{ reference: "B2:B3", descending: true, customList: "2,1" }, { reference: "A2:A3", descending: false }],
});
const importedAdvancedFilterTable = imported.worksheets.getItem("Advanced Filters").tables.getItemOrNullObject("AdvancedFilterTable");
assert.deepEqual(importedAdvancedFilterTable.filters, advancedFilterTable.filters);
const importedIconTable = imported.worksheets.getItem("Icon Rules").tables.getItemOrNullObject("IconRuleTable");
assert.deepEqual(importedIconTable.filters, iconTable.filters);
assert.deepEqual(importedIconTable.sortState, iconTable.sortState);
const importedColorTable = imported.worksheets.getItem("Color Rules").tables.getItemOrNullObject("ColorRuleTable");
assert.deepEqual(importedColorTable.filters, colorTable.filters);
assert.deepEqual(importedColorTable.sortState, colorTable.sortState);
assert.equal(importedTable.style, "TableStyleMedium4");
assert.equal(importedTable.showFirstColumn, true);
assert.equal(importedTable.showBandedColumns, true);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(imported.worksheets.getItem("Summary").freezePanes.toJSON(), { rows: 1, columns: 1, frozen: true, topLeftCell: "B2", activePane: "bottomRight" });
assert.equal(imported.worksheets.getItem("Summary").showGridLines, false);
assert.equal(imported.worksheets.getItem("Summary").columnDimensions.get(0).width, 18);
assert.deepEqual(imported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);
assert.deepEqual(imported.worksheets.getItem("Summary").sortState, summary.sortState);
assert.equal(imported.worksheets.getItem("Summary").getRange("B1").format.numberFormat, "0.000 \"units\"");
assert.equal(imported.worksheets.getItem("Summary").getRange("B2").format.numberFormat, "0.00%");
assert.equal(imported.worksheets.getItem("Summary").getRange("A1").format.font.bold, true);
assert.equal(imported.worksheets.getItem("Summary").getRange("A1").format.font.underline, "double");
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1").format.fill, { patternType: "darkGrid", foreground: { theme: 4, tint: 0.4 }, background: "#E2E8F0" });
assert.equal(imported.worksheets.getItem("Summary").getRange("A1").format.border.bottom.style, "double");
assert.deepEqual(imported.worksheets.getItem("Summary").getRange("A1").format.protection, { locked: false, hidden: true });
const importedSummaryImage = imported.worksheets.getItem("Summary").images.items[0];
assert.equal(importedSummaryImage.name, "Quarter mark");
assert.equal(importedSummaryImage.alt, "Quarterly performance");
assert.equal(importedSummaryImage.dataUrl, summaryImage.dataUrl);
assert.deepEqual(importedSummaryImage.anchor, { from: { row: 3, col: 2, rowOffsetPx: 10, colOffsetPx: 5 }, extent: { widthPx: 120, heightPx: 80 } });
assert.equal(imported.resolve(importedSummaryImage.id), importedSummaryImage);
const importedSummaryChart = imported.worksheets.getItem("Summary").charts.items[0];
assert.equal(importedSummaryChart.type, "line");
assert.equal(importedSummaryChart.name, "Quarter chart");
assert.equal(importedSummaryChart.title, "Quarter trend");
assert.deepEqual(importedSummaryChart.lineOptions, { grouping: "stacked", smooth: true, varyColors: true });
assert.deepEqual(importedSummaryChart.dataLabels, { showValue: true, showCategoryName: false, showSeriesName: true, position: "outsideEnd" });
assert.equal(importedSummaryChart.hasLegend, true);
assert.deepEqual(importedSummaryChart.categories, ["Q1", "Q2"]);
assert.deepEqual(importedSummaryChart.series.items[0], {
  name: "Revenue",
  values: [42.5, 85],
  categoryFormula: "'Summary'!$A$1:$A$2",
  formula: "'Summary'!$B$1:$B$2",
  fill: "#F472B6",
  line: { fill: "#0EA5E9", style: "dashed", width: 2 },
  marker: { symbol: "diamond", size: 8, fill: "#FDE68A", line: { fill: "#BE123C", style: "dotted", width: 1.5 } },
});
assert.deepEqual(importedSummaryChart.position, { left: 420, top: 40, width: 360, height: 220 });
assert.deepEqual(importedSummaryChart.xAxis, { axisType: "textAxis", title: { text: "Quarter" }, numberFormatCode: "@", tickLabelInterval: 2 });
assert.deepEqual(importedSummaryChart.yAxis, { axisType: "valueAxis", title: { text: "Revenue" }, numberFormatCode: "$#,##0.0", min: 0, max: 100, majorUnit: 25 });
assert.equal(imported.resolve(importedSummaryChart.id), importedSummaryChart);
assert.match(imported.inspect({ kind: "workbook,sheet,formula" }).ndjson, /"dateSystem":"1904"/);
assert.equal(imported.verify().ok, true);
assert.equal(imported.resolve(imported.worksheets.getItem("Summary").id).name, "Summary");

// Open XML SDK serializes SpreadsheetML with a legal namespace prefix. Keep the
// JavaScript migration oracle able to read the same package while both codecs
// coexist, so cross-codec fixtures compare semantics instead of XML spelling.
const javascriptImported = await SpreadsheetFile.importXlsx(exported);
assert.equal(javascriptImported.dateSystem, "1904");
assert.equal(javascriptImported.theme.name, "OpenChestnut Theme");
assert.equal(javascriptImported.theme.colors.accent1, "#0F766E");
assert.deepEqual(javascriptImported.calculation, workbook.calculation);
assert.equal(javascriptImported.worksheets.items.length, 5);
assert.deepEqual(javascriptImported.worksheets.items.map((item) => item.visibility), ["visible", "hidden", "veryHidden", "visible", "visible"]);
assert.deepEqual(javascriptImported.definedNames.items.map((item) => item.toJSON()), [
  { id: javascriptImported.definedNames.items[0].id, name: "SummaryData", refersTo: "Summary!$A$1:$B$2", scope: undefined, comment: "Summary data body", hidden: false },
  { id: javascriptImported.definedNames.items[1].id, name: "StatusData", refersTo: "Details!$A$2:$B$3", scope: "Details", comment: undefined, hidden: true },
]);
assert.equal(javascriptImported.worksheets.getItem("Details").tables.items[0].name, "StatusTable");
assert.deepEqual(javascriptImported.worksheets.getItem("Details").tables.items[0].columnNames, ["Status", "Value"]);
assert.equal(javascriptImported.worksheets.getItem("Details").tables.items[0].columnDefinitions[1].calculatedColumnFormula, "=LEN([@Status])");
assert.deepEqual(javascriptImported.worksheets.getItem("Details").tables.items[0].filters, detailsTable.filters);
assert.deepEqual(javascriptImported.worksheets.getItem("Details").tables.items[0].sortState, detailsTable.sortState);
assert.deepEqual(javascriptImported.worksheets.getItem("Advanced Filters").tables.items[0].filters, advancedFilterTable.filters);
assert.deepEqual(javascriptImported.worksheets.getItem("Icon Rules").tables.items[0].filters, iconTable.filters);
assert.deepEqual(javascriptImported.worksheets.getItem("Icon Rules").tables.items[0].sortState, iconTable.sortState);
assert.deepEqual(javascriptImported.worksheets.getItem("Color Rules").tables.items[0].filters, colorTable.filters);
assert.deepEqual(javascriptImported.worksheets.getItem("Color Rules").tables.items[0].sortState, colorTable.sortState);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").values, [["Quarter", 42.5], [true, 85]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").getRange("A1:B2").formulas, [[null, null], [null, "=B1*2"]]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").mergedRanges, ["A3:B3"]);
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").sortState, summary.sortState);
assert.equal(javascriptImported.worksheets.getItem("Summary").getRange("B1").format.numberFormat, "0.000 \"units\"");
assert.equal(javascriptImported.worksheets.getItem("Summary").getRange("B2").format.numberFormat, "0.00%");
assert.equal(javascriptImported.worksheets.getItem("Summary").getRange("A1").format.font.bold, true);
assert.equal(javascriptImported.worksheets.getItem("Summary").getRange("A1").format.fill.patternType, "darkGrid");
assert.equal(javascriptImported.worksheets.getItem("Summary").getRange("A1").format.border.bottom.style, "double");
assert.equal(javascriptImported.worksheets.getItem("Summary").images.items[0].name, "Quarter mark");
assert.equal(javascriptImported.worksheets.getItem("Summary").images.items[0].dataUrl, summaryImage.dataUrl);
assert.equal(javascriptImported.worksheets.getItem("Summary").charts.items[0].title, "Quarter trend");
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].dataLabels, { showValue: true, showCategoryName: false, showSeriesName: true, position: "outsideEnd" });
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].categories, ["Q1", "Q2"]);
assert.equal(javascriptImported.worksheets.getItem("Summary").charts.items[0].series.items[0].fill, "#F472B6");
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].series.items[0].line, { width: 2, fill: "#0EA5E9", style: "dashed" });
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].series.items[0].marker, { symbol: "diamond", size: 8, fill: "#FDE68A", line: { width: 1.5, fill: "#BE123C", style: "dotted" } });
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].xAxis, { axisType: "textAxis", title: { text: "Quarter" }, numberFormatCode: "@", tickLabelInterval: 2 });
assert.deepEqual(javascriptImported.worksheets.getItem("Summary").charts.items[0].yAxis, { axisType: "valueAxis", title: { text: "Revenue" }, numberFormatCode: "$#,##0.0", min: 0, max: 100, majorUnit: 25 });

imported.worksheets.getItem("Summary").getRange("B1").format.numberFormat = "$#,##0.00";
imported.worksheets.getItem("Details").visibility = "veryHidden";
Object.assign(imported.definedNames.getItem("SummaryData"), { name: "SummaryRange", refersTo: "Summary!$B$1:$B$2", comment: "Updated summary range", hidden: true });
imported.setTheme({ name: "OpenChestnut Edited", colors: { ...imported.theme.colors, accent2: "#22C55E" } });
imported.setCalculation({ ...imported.calculation, mode: "manual", forceFullCalculation: false, iteration: { ...imported.calculation.iteration, maxIterations: 250, maxChange: 0.0001 } });
imported.worksheets.getItem("Summary").getRange("A1").format = {
  ...imported.worksheets.getItem("Summary").getRange("A1").format,
  fill: "#22C55E",
  font: { ...imported.worksheets.getItem("Summary").getRange("A1").format.font, bold: false },
};
importedTable.name = "EditedStatusTable";
importedTable.style = "TableStyleMedium9";
importedTable.showLastColumn = true;
importedTable.showBandedColumns = false;
importedTable.columnNames[1] = "Score";
importedTable.columnDefinitions[1].calculatedColumnFormula = "=LEN([@Status])+1";
importedTable.filters[0].values = ["pending"];
importedTable.filters[0].includeBlank = false;
importedTable.filters[1].matchAll = false;
importedTable.filters[1].criteria[0].value = "0";
importedTable.sortState.caseSensitive = false;
importedTable.sortState.sortMethod = "pinYin";
importedTable.sortState.conditions[0].descending = false;
importedTable.sortState.conditions[0].customList = "1,2";
importedTable.sortState.conditions[1].descending = true;
importedAdvancedFilterTable.filters[0].dateGroups[0].day = 16;
importedAdvancedFilterTable.filters[1].type = "yesterday";
importedAdvancedFilterTable.filters[1].value = 45852;
importedAdvancedFilterTable.filters[1].maxValue = 45853;
importedAdvancedFilterTable.filters[2].top = false;
importedAdvancedFilterTable.filters[2].percent = false;
importedAdvancedFilterTable.filters[2].value = 5;
importedIconTable.filters[0].iconId = 2;
importedIconTable.filters[1].iconId = 1;
importedIconTable.sortState.conditions[0].iconSet = "4Rating";
importedIconTable.sortState.conditions[0].iconId = 3;
importedIconTable.sortState.conditions[1].iconId = 2;
importedColorTable.filters[0].color = "#22C55E";
importedColorTable.sortState.conditions[1].color = "#22C55E";
importedColorTable.filters[1].color = "#2563EB";
importedColorTable.sortState.conditions[0].color = "#2563EB";
importedSummaryImage.name = "Updated quarter mark";
importedSummaryImage.alt = "Updated quarterly performance";
importedSummaryImage.anchor = { from: { row: 5, col: 4, rowOffsetPx: 12, colOffsetPx: 6 }, extent: { widthPx: 160, heightPx: 100 } };
importedSummaryChart.name = "Updated quarter chart";
importedSummaryChart.title = "Updated quarter trend";
importedSummaryChart.lineOptions = { grouping: "percentStacked", smooth: false, varyColors: false };
importedSummaryChart.dataLabels = { showValue: false, showCategoryName: true, showSeriesName: false, position: "top" };
importedSummaryChart.hasLegend = false;
importedSummaryChart.categories[1] = "Q2 actual";
importedSummaryChart.series.items[0].name = "Actual revenue";
importedSummaryChart.series.items[0].values[1] = 90;
importedSummaryChart.series.items[0].fill = "#2563EB";
importedSummaryChart.series.items[0].line = { fill: "#7C3AED", style: "dash-dot", width: 2.5 };
importedSummaryChart.series.items[0].marker = { symbol: "triangle", size: 10, fill: "#DCFCE7", line: { fill: "#166534", style: "dashed", width: 2 } };
importedSummaryChart.xAxis.title.text = "Fiscal quarter";
importedSummaryChart.xAxis.numberFormatCode = "mmm";
importedSummaryChart.xAxis.tickLabelInterval = 1;
importedSummaryChart.yAxis.title.text = "Revenue USD";
importedSummaryChart.yAxis.numberFormatCode = "$0";
importedSummaryChart.yAxis.min = -10;
importedSummaryChart.yAxis.max = 120;
importedSummaryChart.yAxis.majorUnit = 10;
const secondExport = await exportXlsxWithOpenChestnut(imported, { recalculate: false });
assert.deepEqual([...secondExport.bytes.slice(0, 2)], [0x50, 0x4b]);
const secondImported = await importXlsxWithOpenChestnut(secondExport);
assert.equal(secondImported.worksheets.getItem("Details").visibility, "veryHidden");
assert.equal(secondImported.worksheets.getItem("Summary").getRange("B1").format.numberFormat, "$#,##0.00");
assert.equal(secondImported.theme.name, "OpenChestnut Edited");
assert.equal(secondImported.theme.colors.accent2, "#22C55E");
assert.deepEqual(secondImported.calculation, {
  mode: "manual",
  calculateOnSave: false,
  fullCalculationOnLoad: true,
  forceFullCalculation: false,
  iteration: { enabled: true, maxIterations: 250, maxChange: 0.0001 },
  fullPrecision: false,
});
assert.deepEqual(secondImported.definedNames.getItem("SummaryRange").toJSON(), {
  id: "defined-name/1",
  name: "SummaryRange",
  refersTo: "Summary!$B$1:$B$2",
  scope: undefined,
  comment: "Updated summary range",
  hidden: true,
});
const secondImage = secondImported.worksheets.getItem("Summary").images.items[0];
assert.equal(secondImage.name, "Updated quarter mark");
assert.equal(secondImage.alt, "Updated quarterly performance");
assert.deepEqual(secondImage.anchor, { from: { row: 5, col: 4, rowOffsetPx: 12, colOffsetPx: 6 }, extent: { widthPx: 160, heightPx: 100 } });
assert.equal(secondImage.dataUrl, summaryImage.dataUrl);
const secondChart = secondImported.worksheets.getItem("Summary").charts.items[0];
assert.equal(secondChart.name, "Updated quarter chart");
assert.equal(secondChart.title, "Updated quarter trend");
assert.deepEqual(secondChart.lineOptions, { grouping: "percentStacked", smooth: false });
assert.deepEqual(secondChart.dataLabels, { showValue: false, showCategoryName: true, showSeriesName: false, position: "top" });
assert.equal(secondChart.hasLegend, false);
assert.deepEqual(secondChart.categories, ["Q1", "Q2 actual"]);
assert.deepEqual(secondChart.series.items[0].values, [42.5, 90]);
assert.equal(secondChart.series.items[0].formula, "'Summary'!$B$1:$B$2");
assert.equal(secondChart.series.items[0].fill, "#2563EB");
assert.deepEqual(secondChart.series.items[0].line, { width: 2.5, fill: "#7C3AED", style: "dash-dot" });
assert.deepEqual(secondChart.series.items[0].marker, { symbol: "triangle", size: 10, fill: "#DCFCE7", line: { width: 2, fill: "#166534", style: "dashed" } });
assert.deepEqual(secondChart.xAxis, { axisType: "textAxis", title: { text: "Fiscal quarter" }, numberFormatCode: "mmm", tickLabelInterval: 1 });
assert.deepEqual(secondChart.yAxis, { axisType: "valueAxis", title: { text: "Revenue USD" }, numberFormatCode: "$0", min: -10, max: 120, majorUnit: 10 });
const removedSeriesFill = await importXlsxWithOpenChestnut(secondExport);
removedSeriesFill.worksheets.getItem("Summary").charts.items[0].series.items[0].fill = undefined;
const withoutSeriesFill = await exportXlsxWithOpenChestnut(removedSeriesFill, { recalculate: false });
const withoutSeriesFillRoundTrip = await importXlsxWithOpenChestnut(withoutSeriesFill);
assert.equal(withoutSeriesFillRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].fill, undefined);
withoutSeriesFillRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].fill = "#22C55E";
const readdedSeriesFill = await exportXlsxWithOpenChestnut(withoutSeriesFillRoundTrip, { recalculate: false });
assert.equal((await importXlsxWithOpenChestnut(readdedSeriesFill)).worksheets.getItem("Summary").charts.items[0].series.items[0].fill, "#22C55E");
const removedSeriesLine = await importXlsxWithOpenChestnut(secondExport);
delete removedSeriesLine.worksheets.getItem("Summary").charts.items[0].series.items[0].line;
const withoutSeriesLine = await exportXlsxWithOpenChestnut(removedSeriesLine, { recalculate: false });
const withoutSeriesLineRoundTrip = await importXlsxWithOpenChestnut(withoutSeriesLine);
assert.equal(withoutSeriesLineRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].line, undefined);
withoutSeriesLineRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].stroke = { color: "#22C55E", style: "dotted", weight: 1.25 };
const readdedSeriesLine = await exportXlsxWithOpenChestnut(withoutSeriesLineRoundTrip, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(readdedSeriesLine)).worksheets.getItem("Summary").charts.items[0].series.items[0].line, { width: 1.25, fill: "#22C55E", style: "dotted" });
const removedSeriesMarker = await importXlsxWithOpenChestnut(secondExport);
delete removedSeriesMarker.worksheets.getItem("Summary").charts.items[0].series.items[0].marker;
const withoutSeriesMarker = await exportXlsxWithOpenChestnut(removedSeriesMarker, { recalculate: false });
const withoutSeriesMarkerRoundTrip = await importXlsxWithOpenChestnut(withoutSeriesMarker);
assert.equal(withoutSeriesMarkerRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].marker, undefined);
withoutSeriesMarkerRoundTrip.worksheets.getItem("Summary").charts.items[0].series.items[0].marker = { symbol: "plus", size: 12, fill: "#FEF3C7", line: { fill: "#92400E", style: "solid", width: 1 } };
const readdedSeriesMarker = await exportXlsxWithOpenChestnut(withoutSeriesMarkerRoundTrip, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(readdedSeriesMarker)).worksheets.getItem("Summary").charts.items[0].series.items[0].marker, { symbol: "plus", size: 12, fill: "#FEF3C7", line: { width: 1, fill: "#92400E", style: "solid" } });
const removedDataLabels = await importXlsxWithOpenChestnut(secondExport);
delete removedDataLabels.worksheets.getItem("Summary").charts.items[0].dataLabels;
const withoutDataLabels = await exportXlsxWithOpenChestnut(removedDataLabels, { recalculate: false });
const withoutDataLabelsRoundTrip = await importXlsxWithOpenChestnut(withoutDataLabels);
assert.equal(withoutDataLabelsRoundTrip.worksheets.getItem("Summary").charts.items[0].dataLabels, undefined);
withoutDataLabelsRoundTrip.worksheets.getItem("Summary").charts.items[0].dataLabels = { showValue: true, showCategoryName: true, showSeriesName: false, position: "center" };
const readdedDataLabels = await exportXlsxWithOpenChestnut(withoutDataLabelsRoundTrip, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(readdedDataLabels)).worksheets.getItem("Summary").charts.items[0].dataLabels, { showValue: true, showCategoryName: true, showSeriesName: false, position: "center" });
const removedDefinedName = await importXlsxWithOpenChestnut(exported);
removedDefinedName.definedNames.delete("SummaryData");
await assert.rejects(
  exportXlsxWithOpenChestnut(removedDefinedName, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_defined_name" && /cannot remove imported defined name/i.test(error.message),
);
const removedCalculation = await importXlsxWithOpenChestnut(exported);
removedCalculation.setCalculation(undefined);
await assert.rejects(
  exportXlsxWithOpenChestnut(removedCalculation, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_calculation" && /cannot remove imported calculation/i.test(error.message),
);
const removedImage = await importXlsxWithOpenChestnut(exported);
removedImage.worksheets.getItem("Summary").images.deleteAll();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedImage, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image_topology" && /cannot remove imported image/i.test(error.message),
);
const removedChart = await importXlsxWithOpenChestnut(exported);
removedChart.worksheets.getItem("Summary").charts.deleteAll();
await assert.rejects(
  exportXlsxWithOpenChestnut(removedChart, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart_topology" && /cannot remove imported chart/i.test(error.message),
);
const addedChart = await importXlsxWithOpenChestnut(exported);
addedChart.worksheets.getItem("Summary").charts.add("bar", { name: "Unexpected chart", categories: ["A"], series: [{ name: "Value", values: [1] }] });
await assert.rejects(
  exportXlsxWithOpenChestnut(addedChart, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart_topology" && /cannot be added/i.test(error.message),
);
const movedChart = await importXlsxWithOpenChestnut(exported);
movedChart.worksheets.getItem("Summary").charts.items[0].position.left += 1;
await assert.rejects(
  exportXlsxWithOpenChestnut(movedChart, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart_edit" && /anchor geometry/i.test(error.message),
);
const chartTypesWorkbook = Workbook.create();
const chartTypesSheet = chartTypesWorkbook.worksheets.add("Chart types");
for (const [index, chartType] of ["bar", "line", "pie"].entries()) {
  chartTypesSheet.charts.add(chartType, {
    name: `${chartType} chart`,
    title: `${chartType} title`,
    hasLegend: index !== 2,
    categories: ["A", "B"],
    series: [{ name: "Score", values: [index + 1, index + 2], fill: ["#E11D48", "#2563EB", "#22C55E"][index] }],
    ...(chartType === "line" ? {
      xAxis: { title: { text: "Category" }, numberFormatCode: "@", tickLabelInterval: 2 },
      yAxis: { title: { text: "Score" }, numberFormatCode: "0.0", min: 0, max: 10, tickLabelInterval: 2 },
    } : {}),
    position: { left: 40 + index * 240, top: 40, width: 220, height: 160 },
  });
}
const chartTypesExport = await exportXlsxWithOpenChestnut(chartTypesWorkbook);
const chartTypesRoundTrip = await importXlsxWithOpenChestnut(chartTypesExport);
assert.deepEqual(chartTypesRoundTrip.worksheets.getItem("Chart types").charts.items.map((chart) => chart.type), ["bar", "line", "pie"]);
assert.deepEqual(chartTypesRoundTrip.worksheets.getItem("Chart types").charts.items.map((chart) => chart.hasLegend), [true, true, false]);
assert.deepEqual(chartTypesRoundTrip.worksheets.getItem("Chart types").charts.items.map((chart) => chart.series.items[0].fill), ["#E11D48", "#2563EB", "#22C55E"]);
assert.deepEqual((await SpreadsheetFile.importXlsx(chartTypesExport)).worksheets.getItem("Chart types").charts.items.map((chart) => chart.type), ["bar", "line", "pie"]);
const javascriptChartTypesExport = await SpreadsheetFile.exportXlsx(chartTypesWorkbook);
const javascriptChartTypesRoundTrip = await importXlsxWithOpenChestnut(javascriptChartTypesExport);
assert.deepEqual(javascriptChartTypesRoundTrip.worksheets.getItem("Chart types").charts.items.map((chart) => chart.type), ["bar", "line", "pie"]);
assert.deepEqual(javascriptChartTypesRoundTrip.worksheets.getItem("Chart types").charts.items[1].xAxis, { axisType: "textAxis", title: { text: "Category" }, numberFormatCode: "@", tickLabelInterval: 2 });
assert.deepEqual(javascriptChartTypesRoundTrip.worksheets.getItem("Chart types").charts.items[1].yAxis, { axisType: "valueAxis", title: { text: "Score" }, numberFormatCode: "0.0", min: 0, max: 10, majorUnit: 2 });
assert.deepEqual(javascriptChartTypesRoundTrip.worksheets.getItem("Chart types").charts.items.map((chart) => chart.series.items[0].fill), ["#E11D48", "#2563EB", "#22C55E"]);
const mismatchedChartWorkbook = Workbook.create();
mismatchedChartWorkbook.worksheets.add("Invalid chart").charts.add("bar", { name: "Mismatch", categories: ["A", "B"], series: [{ name: "Value", values: [1] }] });
await assert.rejects(
  exportXlsxWithOpenChestnut(mismatchedChartWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /1 values for 2 categories/i.test(error.message),
);
const unsupportedChartWorkbook = Workbook.create();
unsupportedChartWorkbook.worksheets.add("Unsupported chart").charts.add("combo", { name: "Combo", categories: ["A"], series: [{ name: "Value", values: [1] }] });
await assert.rejects(
  exportXlsxWithOpenChestnut(unsupportedChartWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /bar, line, or pie/i.test(error.message),
);
const styledChartWorkbook = Workbook.create();
const styledChart = styledChartWorkbook.worksheets.add("Styled chart").charts.add("bar", { name: "Styled", categories: ["A"], series: [{ name: "Value", values: [1] }] });
styledChart.series.items[0].fill = "blue";
assert.ok(styledChartWorkbook.verify().issues.some((issue) => issue.type === "invalidChartSeriesFill"));
await assert.rejects(
  exportXlsxWithOpenChestnut(styledChartWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /#RRGGBB solid color/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(styledChartWorkbook), /#RRGGBB solid color/i);
const invalidLineWorkbook = Workbook.create();
const invalidLineChart = invalidLineWorkbook.worksheets.add("Invalid line").charts.add("line", { name: "Invalid line", categories: ["A"], series: [{ name: "Value", values: [1], line: { fill: "blue", style: "dashed", width: 2 } }] });
assert.ok(invalidLineWorkbook.verify().issues.some((issue) => issue.type === "invalidChartSeriesLine"));
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidLineWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /#RRGGBB color/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidLineWorkbook), /#RRGGBB color/i);
invalidLineChart.series.items[0].line = { fill: "#2563EB", style: "double", width: 2 };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidLineWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /solid, dashed, dotted, dash-dot, dash-dot-dot/i.test(error.message),
);
invalidLineChart.series.items[0].line = { fill: "#2563EB", style: "solid", width: 2 };
invalidLineChart.series.items[0].stroke = { color: "#E11D48", style: "solid", weight: 2 };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidLineWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /aliases must describe the same/i.test(error.message),
);
const invalidMarkerWorkbook = Workbook.create();
const invalidMarkerChart = invalidMarkerWorkbook.worksheets.add("Invalid marker").charts.add("line", { name: "Invalid marker", categories: ["A"], series: [{ name: "Value", values: [1], marker: { symbol: "picture", size: 8 } }] });
assert.ok(invalidMarkerWorkbook.verify().issues.some((issue) => issue.type === "invalidChartSeriesMarker"));
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidMarkerWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /none, dot, circle/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidMarkerWorkbook), /none, dot, circle/i);
invalidMarkerChart.series.items[0].marker = { symbol: "circle", size: 1 };
await assert.rejects(exportXlsxWithOpenChestnut(invalidMarkerWorkbook), /2 through 72/i);
invalidMarkerChart.series.items[0].marker = { symbol: "circle", size: 8, fill: "red" };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidMarkerWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /fill must be a #RRGGBB color/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidMarkerWorkbook), /fill must be a #RRGGBB color/i);
invalidMarkerChart.series.items[0].marker = { symbol: "circle", size: 8, line: { fill: "#E11D48", style: "double" } };
await assert.rejects(exportXlsxWithOpenChestnut(invalidMarkerWorkbook), /solid, dashed, dotted, dash-dot, dash-dot-dot/i);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidMarkerWorkbook), /solid, dashed, dotted, dash-dot, dash-dot-dot/i);
const barMarkerWorkbook = Workbook.create();
barMarkerWorkbook.worksheets.add("Bar marker").charts.add("bar", { name: "Bar marker", categories: ["A"], series: [{ name: "Value", values: [1], marker: { symbol: "circle", size: 8 } }] });
assert.ok(barMarkerWorkbook.verify().issues.some((issue) => issue.type === "invalidChartSeriesMarker"));
await assert.rejects(exportXlsxWithOpenChestnut(barMarkerWorkbook), /markers require a line chart/i);
await assert.rejects(SpreadsheetFile.exportXlsx(barMarkerWorkbook), /markers require a line chart/i);
const invalidLineOptionsWorkbook = Workbook.create();
const invalidLineOptionsSheet = invalidLineOptionsWorkbook.worksheets.add("Invalid line options");
assert.throws(
  () => invalidLineOptionsSheet.charts.add("line", { name: "Rejected options", categories: ["A"], series: [{ name: "Value", values: [1] }], lineOptions: { grouping: "clustered" } }),
  /grouping must be one of standard, stacked, percentStacked/i,
);
const invalidLineOptionsChart = invalidLineOptionsSheet.charts.add("line", { name: "Invalid options", categories: ["A"], series: [{ name: "Value", values: [1] }] });
invalidLineOptionsChart.lineOptions = { smooth: true, gapWidth: 10 };
assert.ok(invalidLineOptionsWorkbook.verify().issues.some((issue) => issue.type === "invalidChartLineOptions"));
await assert.rejects(exportXlsxWithOpenChestnut(invalidLineOptionsWorkbook), /supports only grouping, smooth, and varyColors/i);
invalidLineOptionsChart.lineOptions = { smooth: "true" };
await assert.rejects(exportXlsxWithOpenChestnut(invalidLineOptionsWorkbook), /smooth must be a boolean/i);
invalidLineOptionsChart.lineOptions = { varyColors: 1 };
await assert.rejects(exportXlsxWithOpenChestnut(invalidLineOptionsWorkbook), /varyColors must be a boolean/i);
invalidLineOptionsChart.lineOptions = { smooth: true };
const barLineOptionsWorkbook = Workbook.create();
barLineOptionsWorkbook.worksheets.add("Bar options").charts.add("bar", { name: "Bar options", categories: ["A"], series: [{ name: "Value", values: [1] }], lineOptions: { smooth: true } });
await assert.rejects(exportXlsxWithOpenChestnut(barLineOptionsWorkbook), /lineOptions require a line chart/i);
const invalidDataLabelsWorkbook = Workbook.create();
const invalidDataLabelsSheet = invalidDataLabelsWorkbook.worksheets.add("Invalid data labels");
assert.throws(
  () => invalidDataLabelsSheet.charts.add("line", { name: "Rejected labels", categories: ["A"], series: [{ name: "Value", values: [1] }], dataLabels: { position: "floating" } }),
  /position must be one of/i,
);
const invalidDataLabelsChart = invalidDataLabelsSheet.charts.add("line", { name: "Invalid labels", categories: ["A"], series: [{ name: "Value", values: [1] }] });
invalidDataLabelsChart.dataLabels = {};
assert.ok(invalidDataLabelsWorkbook.verify().issues.some((issue) => issue.type === "invalidChartDataLabels"));
await assert.rejects(exportXlsxWithOpenChestnut(invalidDataLabelsWorkbook), /must define showValue, showCategoryName, showSeriesName, or position/i);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidDataLabelsWorkbook), /must define showValue, showCategoryName, showSeriesName, or position/i);
invalidDataLabelsChart.dataLabels = { showValue: "yes" };
await assert.rejects(exportXlsxWithOpenChestnut(invalidDataLabelsWorkbook), /showValue must be a boolean/i);
invalidDataLabelsChart.dataLabels = { showSeriesName: "yes" };
await assert.rejects(exportXlsxWithOpenChestnut(invalidDataLabelsWorkbook), /showSeriesName must be a boolean/i);
const reversedAxisWorkbook = Workbook.create();
reversedAxisWorkbook.worksheets.add("Invalid axis").charts.add("line", { name: "Reversed axis", categories: ["A"], series: [{ name: "Value", values: [1] }], yAxis: { min: 10, max: 0 } });
await assert.rejects(
  exportXlsxWithOpenChestnut(reversedAxisWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /min must be less than/i.test(error.message),
);
const unsupportedAxisWorkbook = Workbook.create();
unsupportedAxisWorkbook.worksheets.add("Unsupported axis").charts.add("line", { name: "Date axis", categories: ["A"], series: [{ name: "Value", values: [1] }], xAxis: { axisType: "dateAxis" } });
await assert.rejects(
  exportXlsxWithOpenChestnut(unsupportedAxisWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /textAxis/i.test(error.message),
);
const styledAxisWorkbook = Workbook.create();
const styledAxisChart = styledAxisWorkbook.worksheets.add("Styled axis").charts.add("line", {
  name: "Styled axis",
  title: "Styled sizes",
  titleTextStyle: { fontSize: 12.5 },
  lineOptions: { grouping: "stacked", smooth: true, varyColors: true },
  dataLabels: { showValue: true, showCategoryName: true, showSeriesName: true, position: "r" },
  categories: ["A", "B"],
  series: [{ name: "Value", values: [1, 2], line: { fill: "#2563EB", style: "dash-dot-dot", width: 2.25 }, marker: { symbol: "star", size: 10, fill: "#FACC15", line: { fill: "#7C2D12", style: "dotted", width: 1.25 } } }],
  xAxis: { textStyle: { fontSize: 10 } },
  yAxis: { textStyle: { fontSize: 9 } },
});
assert.match(styledAxisChart.toSvg(), /stroke="#2563EB" stroke-width="2.25" stroke-dasharray="8 4 2 4 2 4"[\s\S]*font-size="10"/);
assert.match(styledAxisChart.toSvg(), /<polygon points="[^"]+" fill="#FACC15" stroke="#7C2D12" stroke-width="1.25" stroke-dasharray="2 4"/);
assert.match(styledAxisChart.toSvg(), /<path d="M [^"]+ C [^"]+" fill="none"/);
assert.match(styledAxisChart.toSvg(), /data-chart-label-position="right"[\s\S]*data-chart-label-series="0"[\s\S]*>Value: A: 1<\/text>/);
const groupedPreviewChart = styledAxisWorkbook.worksheets.getItem("Styled axis").charts.add("line", {
  name: "Grouped preview",
  lineOptions: { grouping: "standard" },
  categories: ["A", "B"],
  series: [{ name: "One", values: [1, 2] }, { name: "Two", values: [3, 4] }],
});
const standardPreview = groupedPreviewChart.toSvg();
groupedPreviewChart.lineOptions = { grouping: "stacked" };
const stackedPreview = groupedPreviewChart.toSvg();
groupedPreviewChart.lineOptions = { grouping: "percentStacked" };
const percentPreview = groupedPreviewChart.toSvg();
assert.equal([...stackedPreview.matchAll(/data-series-index=/g)].length, 2, "Worksheet SVG preview must render every grouped line series.");
assert.notEqual(stackedPreview, standardPreview, "Stacked grouping must change deterministic worksheet SVG coordinates.");
assert.notEqual(percentPreview, stackedPreview, "Percent-stacked grouping must normalize deterministic worksheet SVG coordinates.");
styledAxisWorkbook.worksheets.getItem("Styled axis").charts.items.pop();
const styledAxisNative = await exportXlsxWithOpenChestnut(styledAxisWorkbook);
const styledAxisNativeZip = await JSZip.loadAsync(styledAxisNative.bytes);
const styledAxisNativeChartPath = Object.keys(styledAxisNativeZip.files).find((name) => /\/charts\/chart\d+\.xml$/.test(name));
assert.ok(styledAxisNativeChartPath, "OpenChestnut styled chart export must contain one ChartPart.");
const styledAxisNativeXml = await styledAxisNativeZip.file(styledAxisNativeChartPath).async("text");
assert.match(styledAxisNativeXml, /<a:rPr sz="1250"\s*\/>/);
assert.match(styledAxisNativeXml, /<c:catAx>[\s\S]*?<a:defRPr sz="1000"\s*\/>/);
assert.match(styledAxisNativeXml, /<c:valAx>[\s\S]*?<a:defRPr sz="900"\s*\/>/);
assert.match(styledAxisNativeXml, /<a:ln w="28575"><a:solidFill><a:srgbClr val="2563EB"\s*\/><\/a:solidFill><a:prstDash val="lgDashDotDot"\s*\/><\/a:ln>/);
assert.match(styledAxisNativeXml, /<c:marker><c:symbol val="star"\s*\/><c:size val="10"\s*\/><c:spPr><a:solidFill><a:srgbClr val="FACC15"\s*\/><\/a:solidFill><a:ln w="15875"><a:solidFill><a:srgbClr val="7C2D12"\s*\/><\/a:solidFill><a:prstDash val="dot"\s*\/><\/a:ln><\/c:spPr><\/c:marker>/);
assert.match(styledAxisNativeXml, /<c:grouping val="stacked"\s*\/>/);
assert.match(styledAxisNativeXml, /<c:varyColors val="1"\s*\/>/);
assert.match(styledAxisNativeXml, /<c:smooth val="1"\s*\/>/);
assert.match(styledAxisNativeXml, /<c:dLbls><c:dLblPos val="r"\s*\/><c:showVal val="1"\s*\/><c:showCatName val="1"\s*\/><c:showSerName val="1"\s*\/><\/c:dLbls>/);
const styledAxisImported = await importXlsxWithOpenChestnut(styledAxisNative);
const importedStyledAxisChart = styledAxisImported.worksheets.getItem("Styled axis").charts.items[0];
assert.deepEqual(importedStyledAxisChart.titleTextStyle, { fontSize: 12.5 });
assert.deepEqual(importedStyledAxisChart.xAxis.textStyle, { fontSize: 10 });
assert.deepEqual(importedStyledAxisChart.yAxis.textStyle, { fontSize: 9 });
assert.deepEqual(importedStyledAxisChart.series.items[0].line, { fill: "#2563EB", style: "dash-dot-dot", width: 2.25 });
assert.deepEqual(importedStyledAxisChart.series.items[0].marker, { symbol: "star", size: 10, fill: "#FACC15", line: { width: 1.25, fill: "#7C2D12", style: "dotted" } });
assert.deepEqual(importedStyledAxisChart.lineOptions, { grouping: "stacked", smooth: true, varyColors: true });
assert.deepEqual(importedStyledAxisChart.dataLabels, { showValue: true, showCategoryName: true, showSeriesName: true, position: "right" });
importedStyledAxisChart.titleTextStyle.fontSize = 14;
importedStyledAxisChart.xAxis.textStyle.fontSize = 11;
delete importedStyledAxisChart.yAxis.textStyle;
delete importedStyledAxisChart.series.items[0].line;
importedStyledAxisChart.series.items[0].stroke = { color: "#7C3AED", style: "dotted", weight: 1.5 };
importedStyledAxisChart.series.items[0].marker = { symbol: "plus", size: 12, fill: "#E0E7FF", line: { fill: "#4338CA", style: "dashed", width: 2 } };
const styledAxisEdited = await exportXlsxWithOpenChestnut(styledAxisImported, { recalculate: false });
const styledAxisEditedRoundTrip = await importXlsxWithOpenChestnut(styledAxisEdited);
const editedStyledAxisChart = styledAxisEditedRoundTrip.worksheets.getItem("Styled axis").charts.items[0];
assert.deepEqual(editedStyledAxisChart.titleTextStyle, { fontSize: 14 });
assert.deepEqual(editedStyledAxisChart.xAxis.textStyle, { fontSize: 11 });
assert.equal(editedStyledAxisChart.yAxis.textStyle, undefined);
assert.deepEqual(editedStyledAxisChart.series.items[0].line, { fill: "#7C3AED", style: "dotted", width: 1.5 });
assert.deepEqual(editedStyledAxisChart.series.items[0].marker, { symbol: "plus", size: 12, fill: "#E0E7FF", line: { width: 2, fill: "#4338CA", style: "dashed" } });
assert.deepEqual(editedStyledAxisChart.dataLabels, { showValue: true, showCategoryName: true, showSeriesName: true, position: "right" });
const styledAxisFallback = await SpreadsheetFile.exportXlsx(styledAxisWorkbook);
const styledAxisFallbackRoundTrip = await importXlsxWithOpenChestnut(styledAxisFallback);
assert.deepEqual(styledAxisFallbackRoundTrip.worksheets.getItem("Styled axis").charts.items[0].titleTextStyle, { fontSize: 12.5 });
assert.deepEqual(styledAxisFallbackRoundTrip.worksheets.getItem("Styled axis").charts.items[0].series.items[0].line, { width: 2.25, fill: "#2563EB", style: "dash-dot-dot" });
assert.deepEqual(styledAxisFallbackRoundTrip.worksheets.getItem("Styled axis").charts.items[0].series.items[0].marker, { symbol: "star", size: 10, fill: "#FACC15", line: { width: 1.25, fill: "#7C2D12", style: "dotted" } });
assert.deepEqual(styledAxisFallbackRoundTrip.worksheets.getItem("Styled axis").charts.items[0].lineOptions, { grouping: "stacked", smooth: true, varyColors: true });
assert.deepEqual(styledAxisFallbackRoundTrip.worksheets.getItem("Styled axis").charts.items[0].dataLabels, { showValue: true, showCategoryName: true, showSeriesName: true, position: "right" });

const invalidTextStyleWorkbook = Workbook.create();
invalidTextStyleWorkbook.worksheets.add("Invalid text style").charts.add("line", { name: "Invalid style", title: "Invalid", titleTextStyle: { fontSize: 0 }, categories: ["A"], series: [{ name: "Value", values: [1] }] });
assert.ok(invalidTextStyleWorkbook.verify().issues.some((issue) => issue.type === "invalidChartTextStyle"));
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidTextStyleWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_chart" && /1 through 4000/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidTextStyleWorkbook), /1 through 4000/i);
const unsupportedAxisTitleStyleWorkbook = Workbook.create();
unsupportedAxisTitleStyleWorkbook.worksheets.add("Axis title style").charts.add("line", { name: "Axis title style", categories: ["A"], series: [{ name: "Value", values: [1] }], xAxis: { title: { text: "Category", textStyle: { fontSize: 10 } } } });
await assert.rejects(
  exportXlsxWithOpenChestnut(unsupportedAxisTitleStyleWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /axis-title text styling/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(unsupportedAxisTitleStyleWorkbook), /axis-title text styling/i);
const pieAxisWorkbook = Workbook.create();
pieAxisWorkbook.worksheets.add("Pie axis").charts.add("pie", { name: "Pie axis", categories: ["A"], series: [{ name: "Value", values: [1] }], xAxis: { title: { text: "Forbidden" } } });
await assert.rejects(
  exportXlsxWithOpenChestnut(pieAxisWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_chart" && /cannot carry category\/value axes/i.test(error.message),
);
const addedImage = await importXlsxWithOpenChestnut(exported);
addedImage.worksheets.getItem("Summary").images.add({ name: "Unexpected image", dataUrl: summaryImage.dataUrl });
await assert.rejects(
  exportXlsxWithOpenChestnut(addedImage, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image_topology" && /cannot add image/i.test(error.message),
);
const replacedImage = await importXlsxWithOpenChestnut(exported);
const replacementImageDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAIAAAACCAQAAABFaP0WAAAADUlEQVR42mNk+M/wHwAF/gL+3c5GAAAAAElFTkSuQmCC";
replacedImage.worksheets.getItem("Summary").images.items[0].dataUrl = replacementImageDataUrl;
const replacedImageExport = await exportXlsxWithOpenChestnut(replacedImage, { recalculate: false });
const replacedImageRoundTrip = await importXlsxWithOpenChestnut(replacedImageExport);
assert.equal(replacedImageRoundTrip.worksheets.getItem("Summary").images.items[0].dataUrl, replacementImageDataUrl);
const crossFormatImage = await importXlsxWithOpenChestnut(exported);
crossFormatImage.worksheets.getItem("Summary").images.items[0].dataUrl = "data:image/jpeg;base64,/9j/2Q==";
await assert.rejects(
  exportXlsxWithOpenChestnut(crossFormatImage, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_image_edit" && /content type/i.test(error.message),
);
const externalImageWorkbook = Workbook.create();
externalImageWorkbook.worksheets.add("External").images.add({ name: "External image", uri: "https://example.test/image.png" });
await assert.rejects(
  exportXlsxWithOpenChestnut(externalImageWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_image" && /embedded dataUrl bytes/i.test(error.message),
);
const jpegImageWorkbook = Workbook.create();
jpegImageWorkbook.worksheets.add("JPEG").images.add({
  name: "JPEG marker",
  alt: "JPEG marker",
  dataUrl: "data:image/jpeg;base64,/9j/2Q==",
  anchor: { from: { row: 0, col: 0 }, extent: { widthPx: 32, heightPx: 24 } },
});
const jpegImageRoundTrip = await importXlsxWithOpenChestnut(await exportXlsxWithOpenChestnut(jpegImageWorkbook));
assert.equal(jpegImageRoundTrip.worksheets.getItem("JPEG").images.items[0].dataUrl, "data:image/jpeg;base64,/9j/2Q==");
const twoCellImageWorkbook = Workbook.create();
const twoCellImageSheet = twoCellImageWorkbook.worksheets.add("Two-cell image");
twoCellImageSheet.getRange("A1:F8").values = Array.from({ length: 8 }, (_, row) => Array.from({ length: 6 }, (_, column) => `${row + 1}:${column + 1}`));
const twoCellImageDataUrl = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=";
twoCellImageSheet.images.add({
  name: "Two-cell mark",
  alt: "Move without resize",
  dataUrl: twoCellImageDataUrl,
  anchor: {
    type: "twoCell",
    from: { row: 1, col: 1, rowOffsetPx: 2, colOffsetPx: 3 },
    to: { row: 4, col: 5, rowOffsetPx: 6, colOffsetPx: 7 },
    editAs: "oneCell",
  },
});
const twoCellImageExport = await exportXlsxWithOpenChestnut(twoCellImageWorkbook);
const twoCellImageZip = await JSZip.loadAsync(twoCellImageExport.bytes);
const twoCellDrawingXml = await twoCellImageZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(twoCellDrawingXml, /<xdr:twoCellAnchor editAs="oneCell">/);
assert.match(twoCellDrawingXml, /<xdr:from><xdr:col>1<\/xdr:col><xdr:colOff>28575<\/xdr:colOff><xdr:row>1<\/xdr:row><xdr:rowOff>19050<\/xdr:rowOff><\/xdr:from>/);
assert.match(twoCellDrawingXml, /<xdr:to><xdr:col>5<\/xdr:col><xdr:colOff>66675<\/xdr:colOff><xdr:row>4<\/xdr:row><xdr:rowOff>57150<\/xdr:rowOff><\/xdr:to>/);
const twoCellImageImported = await importXlsxWithOpenChestnut(twoCellImageExport);
const importedTwoCellImage = twoCellImageImported.worksheets.getItem("Two-cell image").images.items[0];
assert.deepEqual(importedTwoCellImage.anchor, {
  type: "twoCell",
  from: { row: 1, col: 1, rowOffsetPx: 2, colOffsetPx: 3 },
  to: { row: 4, col: 5, rowOffsetPx: 6, colOffsetPx: 7 },
  editAs: "oneCell",
});
importedTwoCellImage.name = "Edited two-cell mark";
importedTwoCellImage.anchor = {
  type: "twoCell",
  from: { row: 2, col: 2, rowOffsetPx: 4, colOffsetPx: 5 },
  to: { row: 6, col: 7, rowOffsetPx: 8, colOffsetPx: 9 },
  editAs: "absolute",
};
const editedTwoCellExport = await exportXlsxWithOpenChestnut(twoCellImageImported, { recalculate: false });
const editedTwoCellRoundTrip = await importXlsxWithOpenChestnut(editedTwoCellExport);
assert.equal(editedTwoCellRoundTrip.worksheets.getItem("Two-cell image").images.items[0].name, "Edited two-cell mark");
assert.deepEqual(editedTwoCellRoundTrip.worksheets.getItem("Two-cell image").images.items[0].anchor, importedTwoCellImage.anchor);
const javascriptTwoCellExport = await SpreadsheetFile.exportXlsx(twoCellImageWorkbook);
const javascriptTwoCellZip = await JSZip.loadAsync(new Uint8Array(await javascriptTwoCellExport.arrayBuffer()));
assert.match(await javascriptTwoCellZip.file("xl/drawings/drawing1.xml").async("text"), /<xdr:twoCellAnchor editAs="oneCell">/);
const javascriptTwoCellImport = await SpreadsheetFile.importXlsx(javascriptTwoCellExport);
assert.deepEqual(javascriptTwoCellImport.worksheets.getItem("Two-cell image").images.items[0].anchor, {
  type: "twoCell",
  from: { row: 1, col: 1, rowOffsetPx: 2, colOffsetPx: 3 },
  to: { row: 4, col: 5, rowOffsetPx: 6, colOffsetPx: 7 },
  editAs: "oneCell",
});
const changedTwoCellKind = await importXlsxWithOpenChestnut(twoCellImageExport);
changedTwoCellKind.worksheets.getItem("Two-cell image").images.items[0].anchor = { from: { row: 1, col: 1 }, extent: { widthPx: 120, heightPx: 80 } };
await assert.rejects(
  exportXlsxWithOpenChestnut(changedTwoCellKind, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_image_edit" && /anchor kind/i.test(error.message),
);
const invalidTwoCellWorkbook = Workbook.create();
invalidTwoCellWorkbook.worksheets.add("Invalid").images.add({
  name: "Invalid two-cell mark",
  dataUrl: twoCellImageDataUrl,
  anchor: { type: "twoCell", from: { row: 4, col: 4 }, to: { row: 3, col: 5 } },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidTwoCellWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /strictly after/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidTwoCellWorkbook), /strictly after/i);
const absoluteImageWorkbook = Workbook.create();
const absoluteImageSheet = absoluteImageWorkbook.worksheets.add("Absolute image");
absoluteImageSheet.images.add({
  name: "Absolute mark",
  alt: "Page-relative worksheet image",
  dataUrl: twoCellImageDataUrl,
  anchor: {
    type: "absolute",
    position: { leftPx: -20, topPx: 30 },
    extent: { widthPx: 120, heightPx: 80 },
  },
});
const absoluteImageExport = await exportXlsxWithOpenChestnut(absoluteImageWorkbook);
const absoluteImageZip = await JSZip.loadAsync(absoluteImageExport.bytes);
const absoluteDrawingXml = await absoluteImageZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(absoluteDrawingXml, /<xdr:absoluteAnchor>/);
assert.match(absoluteDrawingXml, /<xdr:pos x="-190500" y="285750"\s*\/>/);
assert.match(absoluteDrawingXml, /<xdr:ext cx="1143000" cy="762000"\s*\/>/);
const absoluteImageImported = await importXlsxWithOpenChestnut(absoluteImageExport);
const importedAbsoluteImage = absoluteImageImported.worksheets.getItem("Absolute image").images.items[0];
assert.deepEqual(importedAbsoluteImage.anchor, {
  type: "absolute",
  position: { leftPx: -20, topPx: 30 },
  extent: { widthPx: 120, heightPx: 80 },
});
importedAbsoluteImage.name = "Edited absolute mark";
importedAbsoluteImage.anchor = {
  type: "absolute",
  position: { leftPx: 40, topPx: -10 },
  extent: { widthPx: 160, heightPx: 100 },
};
const editedAbsoluteExport = await exportXlsxWithOpenChestnut(absoluteImageImported, { recalculate: false });
const editedAbsoluteRoundTrip = await importXlsxWithOpenChestnut(editedAbsoluteExport);
assert.equal(editedAbsoluteRoundTrip.worksheets.getItem("Absolute image").images.items[0].name, "Edited absolute mark");
assert.deepEqual(editedAbsoluteRoundTrip.worksheets.getItem("Absolute image").images.items[0].anchor, importedAbsoluteImage.anchor);
const javascriptAbsoluteExport = await SpreadsheetFile.exportXlsx(absoluteImageWorkbook);
const javascriptAbsoluteZip = await JSZip.loadAsync(new Uint8Array(await javascriptAbsoluteExport.arrayBuffer()));
assert.match(await javascriptAbsoluteZip.file("xl/drawings/drawing1.xml").async("text"), /<xdr:absoluteAnchor><xdr:pos x="-190500" y="285750"\/>/);
const javascriptAbsoluteImport = await SpreadsheetFile.importXlsx(javascriptAbsoluteExport);
assert.deepEqual(javascriptAbsoluteImport.worksheets.getItem("Absolute image").images.items[0].anchor, {
  type: "absolute",
  position: { leftPx: -20, topPx: 30 },
  extent: { widthPx: 120, heightPx: 80 },
});
const changedAbsoluteKind = await importXlsxWithOpenChestnut(absoluteImageExport);
changedAbsoluteKind.worksheets.getItem("Absolute image").images.items[0].anchor = { from: { row: 1, col: 1 }, extent: { widthPx: 120, heightPx: 80 } };
await assert.rejects(
  exportXlsxWithOpenChestnut(changedAbsoluteKind, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_image_edit" && /anchor kind/i.test(error.message),
);
const invalidAbsoluteWorkbook = Workbook.create();
invalidAbsoluteWorkbook.worksheets.add("Invalid absolute").images.add({
  name: "Invalid absolute mark",
  dataUrl: twoCellImageDataUrl,
  anchor: { type: "absolute", position: { leftPx: 0, topPx: 0 }, extent: { widthPx: 0, heightPx: 80 } },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidAbsoluteWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /positive/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidAbsoluteWorkbook), /positive extent geometry/i);
const cropImageWorkbook = Workbook.create();
const cropImageSheet = cropImageWorkbook.worksheets.add("Crop image");
cropImageSheet.images.add({
  name: "Cropped mark",
  alt: "Inset and outset source rectangle",
  dataUrl: twoCellImageDataUrl,
  crop: { leftPercent: 10, topPercent: -5, rightPercent: 15, bottomPercent: 20 },
  anchor: { from: { row: 1, col: 2 }, extent: { widthPx: 120, heightPx: 80 } },
});
const cropImageExport = await exportXlsxWithOpenChestnut(cropImageWorkbook);
const cropImageZip = await JSZip.loadAsync(cropImageExport.bytes);
const cropDrawingXml = await cropImageZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(cropDrawingXml, /<a:srcRect l="10000" t="-5000" r="15000" b="20000"[^>]*\/>/);
assert.ok(cropDrawingXml.indexOf("<a:blip") < cropDrawingXml.indexOf("<a:srcRect"));
assert.ok(cropDrawingXml.indexOf("<a:srcRect") < cropDrawingXml.indexOf("<a:stretch"));
const cropImageImported = await importXlsxWithOpenChestnut(cropImageExport);
const importedCropImage = cropImageImported.worksheets.getItem("Crop image").images.items[0];
assert.deepEqual(importedCropImage.crop, { leftPercent: 10, topPercent: -5, rightPercent: 15, bottomPercent: 20 });
importedCropImage.crop = { leftPercent: 12, topPercent: 7, rightPercent: -3, bottomPercent: 18 };
const editedCropExport = await exportXlsxWithOpenChestnut(cropImageImported, { recalculate: false });
const editedCropRoundTrip = await importXlsxWithOpenChestnut(editedCropExport);
assert.deepEqual(editedCropRoundTrip.worksheets.getItem("Crop image").images.items[0].crop, importedCropImage.crop);
editedCropRoundTrip.worksheets.getItem("Crop image").images.items[0].crop = undefined;
const removedCropExport = await exportXlsxWithOpenChestnut(editedCropRoundTrip, { recalculate: false });
const removedCropZip = await JSZip.loadAsync(removedCropExport.bytes);
assert.doesNotMatch(await removedCropZip.file("xl/drawings/drawing1.xml").async("text"), /srcRect/);
assert.equal((await importXlsxWithOpenChestnut(removedCropExport)).worksheets.getItem("Crop image").images.items[0].crop, undefined);
const javascriptCropExport = await SpreadsheetFile.exportXlsx(cropImageWorkbook);
const javascriptCropZip = await JSZip.loadAsync(new Uint8Array(await javascriptCropExport.arrayBuffer()));
assert.match(await javascriptCropZip.file("xl/drawings/drawing1.xml").async("text"), /<a:srcRect l="10000" t="-5000" r="15000" b="20000"\/>/);
assert.deepEqual((await SpreadsheetFile.importXlsx(javascriptCropExport)).worksheets.getItem("Crop image").images.items[0].crop, {
  leftPercent: 10,
  topPercent: -5,
  rightPercent: 15,
  bottomPercent: 20,
});
javascriptCropZip.file("xl/drawings/drawing1.xml", (await javascriptCropZip.file("xl/drawings/drawing1.xml").async("text")).replace('l="10000"', 'l="100001"'));
javascriptCropZip.remove("customXml/open-office-artifact.json");
assert.equal((await SpreadsheetFile.importXlsx(await javascriptCropZip.generateAsync({ type: "uint8array" }))).worksheets.getItem("Crop image").images.items[0].crop, undefined);
const invalidCropPairWorkbook = Workbook.create();
invalidCropPairWorkbook.worksheets.add("Invalid crop").images.add({
  name: "Invalid crop",
  dataUrl: twoCellImageDataUrl,
  crop: { leftPercent: 60, rightPercent: 40 },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidCropPairWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /positive source rectangle/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidCropPairWorkbook), /positive source rectangle/i);
const invalidCropBoundWorkbook = Workbook.create();
invalidCropBoundWorkbook.worksheets.add("Invalid crop bound").images.add({
  name: "Invalid crop bound",
  dataUrl: twoCellImageDataUrl,
  crop: { leftPercent: 100.001 },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidCropBoundWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /between -100 and 100/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidCropBoundWorkbook), /-100 through 100/i);
const effectsImageWorkbook = Workbook.create();
const effectsImageSheet = effectsImageWorkbook.worksheets.add("Picture effects");
effectsImageSheet.images.add({
  name: "Effect mark",
  alt: "Bounded grayscale luminance and opacity",
  dataUrl: twoCellImageDataUrl,
  effects: { grayscale: true, brightnessPercent: 15, contrastPercent: -10, opacityPercent: 65 },
  anchor: { from: { row: 1, col: 2 }, extent: { widthPx: 120, heightPx: 80 } },
});
const effectsImageExport = await exportXlsxWithOpenChestnut(effectsImageWorkbook);
const effectsImageZip = await JSZip.loadAsync(effectsImageExport.bytes);
const effectsDrawingXml = await effectsImageZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(effectsDrawingXml, /<a:alphaModFix amt="65000"[^>]*\/>/);
assert.match(effectsDrawingXml, /<a:grayscl[^>]*\/>/);
assert.match(effectsDrawingXml, /<a:lum bright="15000" contrast="-10000"[^>]*\/>/);
assert.ok(effectsDrawingXml.indexOf("alphaModFix") < effectsDrawingXml.indexOf("grayscl"));
assert.ok(effectsDrawingXml.indexOf("grayscl") < effectsDrawingXml.indexOf("<a:lum"));
const effectsImageImported = await importXlsxWithOpenChestnut(effectsImageExport);
const importedEffectsImage = effectsImageImported.worksheets.getItem("Picture effects").images.items[0];
assert.deepEqual(importedEffectsImage.effects, { grayscale: true, brightnessPercent: 15, contrastPercent: -10, opacityPercent: 65 });
assert.match(importedEffectsImage.toSvg(), /filter:grayscale\(1\) brightness\(1\.15\) contrast\(0\.9\)/);
assert.match(importedEffectsImage.toSvg(), /opacity="0\.65"/);
importedEffectsImage.effects = { brightnessPercent: -20, contrastPercent: 25, opacityPercent: 0 };
const editedEffectsExport = await exportXlsxWithOpenChestnut(effectsImageImported, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(editedEffectsExport)).worksheets.getItem("Picture effects").images.items[0].effects, importedEffectsImage.effects);
const removedEffectsWorkbook = await importXlsxWithOpenChestnut(editedEffectsExport);
removedEffectsWorkbook.worksheets.getItem("Picture effects").images.items[0].effects = undefined;
const removedEffectsExport = await exportXlsxWithOpenChestnut(removedEffectsWorkbook, { recalculate: false });
const removedEffectsZip = await JSZip.loadAsync(removedEffectsExport.bytes);
assert.doesNotMatch(await removedEffectsZip.file("xl/drawings/drawing1.xml").async("text"), /alphaModFix|grayscl|<a:lum\b/);
assert.equal((await importXlsxWithOpenChestnut(removedEffectsExport)).worksheets.getItem("Picture effects").images.items[0].effects, undefined);
const javascriptEffectsExport = await SpreadsheetFile.exportXlsx(effectsImageWorkbook);
const javascriptEffectsZip = await JSZip.loadAsync(new Uint8Array(await javascriptEffectsExport.arrayBuffer()));
const javascriptEffectsXml = await javascriptEffectsZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(javascriptEffectsXml, /<a:blip[^>]*><a:alphaModFix amt="65000"\/><a:grayscl\/><a:lum bright="15000" contrast="-10000"\/><\/a:blip>/);
assert.deepEqual((await SpreadsheetFile.importXlsx(javascriptEffectsExport)).worksheets.getItem("Picture effects").images.items[0].effects, {
  grayscale: true,
  brightnessPercent: 15,
  contrastPercent: -10,
  opacityPercent: 65,
});
javascriptEffectsZip.file("xl/drawings/drawing1.xml", javascriptEffectsXml.replace("</a:blip>", "<a:grayscl/></a:blip>"));
javascriptEffectsZip.remove("customXml/open-office-artifact.json");
assert.equal((await SpreadsheetFile.importXlsx(await javascriptEffectsZip.generateAsync({ type: "uint8array" }))).worksheets.getItem("Picture effects").images.items[0].effects, undefined);
const invalidEffectWorkbook = Workbook.create();
invalidEffectWorkbook.worksheets.add("Invalid effect").images.add({
  name: "Invalid effect",
  dataUrl: twoCellImageDataUrl,
  effects: { brightnessPercent: 100.001 },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidEffectWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /between -100 and 100/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidEffectWorkbook), /between -100 and 100/i);
const invalidOpacityWorkbook = Workbook.create();
invalidOpacityWorkbook.worksheets.add("Invalid opacity").images.add({
  name: "Invalid opacity",
  dataUrl: twoCellImageDataUrl,
  effects: { opacityPercent: -0.001 },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidOpacityWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /between 0 and 100/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidOpacityWorkbook), /between 0 and 100/i);
const transformImageWorkbook = Workbook.create();
transformImageWorkbook.worksheets.add("Picture transform").images.add({
  name: "Transformed mark",
  alt: "Rotated and mirrored worksheet picture",
  dataUrl: twoCellImageDataUrl,
  transform: { rotationDegrees: 30.5, flipHorizontal: true, flipVertical: false },
  anchor: { from: { row: 1, col: 2 }, extent: { widthPx: 120, heightPx: 80 } },
});
const transformImageExport = await exportXlsxWithOpenChestnut(transformImageWorkbook);
const transformImageZip = await JSZip.loadAsync(transformImageExport.bytes);
const transformDrawingXml = await transformImageZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(transformDrawingXml, /<a:xfrm\b[^>]*rot="1830000"[^>]*flipH="1"[^>]*flipV="0"[^>]*\/>/);
const transformImageImported = await importXlsxWithOpenChestnut(transformImageExport);
const importedTransformImage = transformImageImported.worksheets.getItem("Picture transform").images.items[0];
assert.deepEqual(importedTransformImage.transform, { rotationDegrees: 30.5, flipHorizontal: true, flipVertical: false });
assert.match(importedTransformImage.toSvg(), /transform="translate\([^\"]+\) rotate\(30\.5\) scale\(-1 1\) translate\([^\"]+\)"/);
importedTransformImage.transform = { rotationDegrees: -45, flipHorizontal: false, flipVertical: true };
const editedTransformExport = await exportXlsxWithOpenChestnut(transformImageImported, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(editedTransformExport)).worksheets.getItem("Picture transform").images.items[0].transform, importedTransformImage.transform);
const removedTransformWorkbook = await importXlsxWithOpenChestnut(editedTransformExport);
removedTransformWorkbook.worksheets.getItem("Picture transform").images.items[0].transform = undefined;
const removedTransformExport = await exportXlsxWithOpenChestnut(removedTransformWorkbook, { recalculate: false });
const removedTransformZip = await JSZip.loadAsync(removedTransformExport.bytes);
assert.doesNotMatch(await removedTransformZip.file("xl/drawings/drawing1.xml").async("text"), /<a:xfrm\b/);
assert.equal((await importXlsxWithOpenChestnut(removedTransformExport)).worksheets.getItem("Picture transform").images.items[0].transform, undefined);
const javascriptTransformExport = await SpreadsheetFile.exportXlsx(transformImageWorkbook);
const javascriptTransformZip = await JSZip.loadAsync(new Uint8Array(await javascriptTransformExport.arrayBuffer()));
const javascriptTransformXml = await javascriptTransformZip.file("xl/drawings/drawing1.xml").async("text");
assert.match(javascriptTransformXml, /<a:xfrm rot="1830000" flipH="1" flipV="0"\/>/);
assert.deepEqual((await SpreadsheetFile.importXlsx(javascriptTransformExport)).worksheets.getItem("Picture transform").images.items[0].transform, {
  rotationDegrees: 30.5,
  flipHorizontal: true,
  flipVertical: false,
});
javascriptTransformZip.file("xl/drawings/drawing1.xml", javascriptTransformXml.replace('rot="1830000"', 'rot="21600001"'));
javascriptTransformZip.remove("customXml/open-office-artifact.json");
assert.equal((await SpreadsheetFile.importXlsx(await javascriptTransformZip.generateAsync({ type: "uint8array" }))).worksheets.getItem("Picture transform").images.items[0].transform, undefined);
const opaqueTransformZip = await JSZip.loadAsync(transformImageExport.bytes);
opaqueTransformZip.file("xl/drawings/drawing1.xml", transformDrawingXml.replace('rot="1830000"', 'rot="21600001"'));
const opaqueTransformWorkbook = await importXlsxWithOpenChestnut(await opaqueTransformZip.generateAsync({ type: "uint8array" }));
const opaqueTransformImage = opaqueTransformWorkbook.worksheets.getItem("Picture transform").images.items[0];
assert.equal(opaqueTransformImage.transform, undefined);
opaqueTransformImage.name = "Opaque transform retained";
const opaqueMetadataExport = await exportXlsxWithOpenChestnut(opaqueTransformWorkbook, { recalculate: false });
const opaqueMetadataZip = await JSZip.loadAsync(opaqueMetadataExport.bytes);
assert.match(await opaqueMetadataZip.file("xl/drawings/drawing1.xml").async("text"), /<a:xfrm\b[^>]*rot="21600001"/);
const opaqueTransformEdit = await importXlsxWithOpenChestnut(opaqueMetadataExport);
opaqueTransformEdit.worksheets.getItem("Picture transform").images.items[0].transform = { rotationDegrees: 10 };
await assert.rejects(
  exportXlsxWithOpenChestnut(opaqueTransformEdit, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_spreadsheet_image_edit" && /transform/i.test(error.message),
);
const invalidTransformWorkbook = Workbook.create();
invalidTransformWorkbook.worksheets.add("Invalid transform").images.add({
  name: "Invalid transform",
  dataUrl: twoCellImageDataUrl,
  transform: { rotationDegrees: 360.001 },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidTransformWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /between -360 and 360/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(invalidTransformWorkbook), /between -360 and 360/i);
const emptyTransformWorkbook = Workbook.create();
emptyTransformWorkbook.worksheets.add("Empty transform").images.add({
  name: "Empty transform",
  dataUrl: twoCellImageDataUrl,
  transform: {},
});
await assert.rejects(
  exportXlsxWithOpenChestnut(emptyTransformWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_spreadsheet_image" && /must define/i.test(error.message),
);
await assert.rejects(SpreadsheetFile.exportXlsx(emptyTransformWorkbook), /must define/i);
const activeVisibilityEdit = await importXlsxWithOpenChestnut(exported);
assert.throws(
  () => { activeVisibilityEdit.worksheets.getItem("Icon Rules").visibility = "hidden"; },
  /select another active worksheet first/i,
);
const noVisibleWorkbook = Workbook.create();
noVisibleWorkbook.worksheets.add("Hidden", { visibility: "hidden" }).getRange("A1").values = [[1]];
await assert.rejects(
  exportXlsxWithOpenChestnut(noVisibleWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "missing_visible_worksheet" && /at least one visible worksheet/i.test(error.message),
);
assert.equal(secondImported.worksheets.getItem("Summary").getRange("A1").format.fill, "#22C55E");
assert.equal(secondImported.worksheets.getItem("Summary").getRange("A1").format.font.bold, false);
const secondTable = secondImported.worksheets.getItem("Details").tables.getItemOrNullObject("EditedStatusTable");
assert.equal(secondTable.style, "TableStyleMedium9");
assert.equal(secondTable.showLastColumn, true);
assert.equal(secondTable.showBandedColumns, false);
assert.deepEqual(secondTable.columnNames, ["Status", "Score"]);
assert.equal(secondTable.columnDefinitions[1].calculatedColumnFormula, "=LEN([@Status])+1");
assert.deepEqual(secondTable.filters, [
  { columnIndex: 0, kind: "values", values: ["pending"], includeBlank: false },
  { columnIndex: 1, kind: "custom", matchAll: false, criteria: [{ operator: "greaterThanOrEqual", value: "0" }, { operator: "lessThanOrEqual", value: "2" }] },
]);
assert.deepEqual(secondTable.sortState, {
  reference: "A2:B3",
  caseSensitive: false,
  sortMethod: "pinYin",
  conditions: [{ reference: "B2:B3", descending: false, customList: "1,2" }, { reference: "A2:A3", descending: true }],
});
assert.deepEqual(secondImported.worksheets.getItem("Advanced Filters").tables.getItemOrNullObject("AdvancedFilterTable").filters, [
  { columnIndex: 0, kind: "values", values: [], includeBlank: false, calendarType: "gregorian", dateGroups: [{ grouping: "day", year: 2026, month: 7, day: 16 }] },
  { columnIndex: 1, kind: "dynamic", type: "yesterday", value: 45852, maxValue: 45853 },
  { columnIndex: 2, kind: "top10", top: false, percent: false, value: 5, filterValue: 95 },
]);
assert.deepEqual(secondImported.worksheets.getItem("Icon Rules").tables.getItemOrNullObject("IconRuleTable").filters, [
  { columnIndex: 0, kind: "icon", iconSet: "3Arrows", iconId: 2 },
  { columnIndex: 1, kind: "icon", iconSet: "3Flags", iconId: 1 },
]);
assert.deepEqual(secondImported.worksheets.getItem("Icon Rules").tables.getItemOrNullObject("IconRuleTable").sortState, {
  reference: "A2:B3",
  caseSensitive: false,
  conditions: [
    { reference: "B2:B3", descending: true, kind: "icon", iconSet: "4Rating", iconId: 3 },
    { reference: "A2:A3", descending: false, kind: "icon", iconSet: "3Symbols2", iconId: 2 },
  ],
});
assert.deepEqual(secondImported.worksheets.getItem("Color Rules").tables.getItemOrNullObject("ColorRuleTable").filters, [
  { columnIndex: 0, kind: "color", target: "cell", color: "#22C55E" },
  { columnIndex: 1, kind: "color", target: "font", color: "#2563EB" },
]);
assert.deepEqual(secondImported.worksheets.getItem("Color Rules").tables.getItemOrNullObject("ColorRuleTable").sortState.conditions, [
  { reference: "B2:B3", descending: true, kind: "color", target: "font", color: "#2563EB" },
  { reference: "A2:A3", descending: false, kind: "color", target: "cell", color: "#22C55E" },
]);
secondTable.delete();
await assert.rejects(
  exportXlsxWithOpenChestnut(secondImported, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /cannot remove imported table/i.test(error.message),
);

const complexTableZip = await JSZip.loadAsync(exported.bytes);
const complexTablePath = "xl/tables/table1.xml";
const complexTableSourceXml = await complexTableZip.file(complexTablePath).async("text");
const complexTableXml = complexTableSourceXml.replace(/(<(?:[A-Za-z_][\w.-]*:)?table\b)/, '$1 published="0"');
assert.notEqual(complexTableXml, complexTableSourceXml, "fixture must add a legal unmodeled table attribute");
complexTableZip.file(complexTablePath, complexTableXml);
const complexTableImported = await importXlsxWithOpenChestnut(await complexTableZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
const complexTableSheet = complexTableImported.worksheets.getItem("Details");
assert.equal(complexTableSheet.tables.items.length, 0, "unmodeled table profiles must not appear as editable public tables");
complexTableSheet.getRange("B2").values = [[7]];
const complexTablePreserved = await exportXlsxWithOpenChestnut(complexTableImported, { recalculate: false });
assert.equal(await (await JSZip.loadAsync(complexTablePreserved.bytes)).file(complexTablePath).async("text"), complexTableXml, "unmodeled table parts must remain byte-exact across unrelated edits");
complexTableSheet.getRange("D1:E2").values = [["Key", "Metric"], ["x", 1]];
complexTableSheet.tables.add({ range: "D1:E2", name: "ReplacementTable" });
await assert.rejects(
  exportXlsxWithOpenChestnut(complexTableImported, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /add or remove worksheet tables/i.test(error.message),
);

const colorSortZip = await JSZip.loadAsync(exported.bytes);
const colorSortSourceXml = await colorSortZip.file(complexTablePath).async("text");
const colorSortStylesPath = "xl/styles.xml";
const colorSortStyles = appendComplexColorDifferentialFormat(await colorSortZip.file(colorSortStylesPath).async("text"));
colorSortZip.file(colorSortStylesPath, colorSortStyles.xml);
const colorSortXml = colorSortSourceXml.replace(/(<x:sortCondition ref="B2:B3" descending="1")/, `$1 sortBy="cellColor" dxfId="${colorSortStyles.id}"`);
assert.notEqual(colorSortXml, colorSortSourceXml, "fixture must add a complex color-sort profile");
colorSortZip.file(complexTablePath, colorSortXml);
const colorSortBytes = await colorSortZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const colorSortImported = await importXlsxWithOpenChestnut(colorSortBytes);
const colorSortSheet = colorSortImported.worksheets.getItem("Details");
assert.equal(colorSortSheet.tables.items.length, 0, "complex color sorts must remain hidden read-only table slots");
colorSortSheet.getRange("B2").values = [[9]];
const colorSortPreserved = await exportXlsxWithOpenChestnut(colorSortImported, { recalculate: false });
assert.equal(await (await JSZip.loadAsync(colorSortPreserved.bytes)).file(complexTablePath).async("text"), colorSortXml, "complex color-sort table XML must remain byte-exact");

const colorFilterZip = await JSZip.loadAsync(exported.bytes);
const colorFilterSourceXml = await colorFilterZip.file(complexTablePath).async("text");
const colorFilterStylesPath = "xl/styles.xml";
const colorFilterStyles = appendComplexColorDifferentialFormat(await colorFilterZip.file(colorFilterStylesPath).async("text"));
colorFilterZip.file(colorFilterStylesPath, colorFilterStyles.xml);
const colorFilterXml = colorFilterSourceXml.replace(/<x:filters blank="1">[\s\S]*?<\/x:filters>/, `<x:colorFilter dxfId="${colorFilterStyles.id}" cellColor="1"/>`);
assert.notEqual(colorFilterXml, colorFilterSourceXml, "fixture must add a complex color-filter profile");
colorFilterZip.file(complexTablePath, colorFilterXml);
const colorFilterImported = await importXlsxWithOpenChestnut(await colorFilterZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
const colorFilterSheet = colorFilterImported.worksheets.getItem("Details");
assert.equal(colorFilterSheet.tables.items.length, 0, "complex color filters must remain hidden read-only table slots");
colorFilterSheet.getRange("B2").values = [[11]];
const colorFilterPreserved = await exportXlsxWithOpenChestnut(colorFilterImported, { recalculate: false });
assert.equal(await (await JSZip.loadAsync(colorFilterPreserved.bytes)).file(complexTablePath).async("text"), colorFilterXml, "complex color-filter table XML must remain byte-exact");

const complexThemeZip = await JSZip.loadAsync(exported.bytes);
const complexThemePath = "xl/theme/theme1.xml";
const complexThemeSourceXml = await complexThemeZip.file(complexThemePath).async("text");
const complexThemeXml = complexThemeSourceXml.replace(
  '<a:accent1><a:srgbClr val="0F766E" /></a:accent1>',
  '<a:accent1><a:hslClr hue="5400000" sat="100000" lum="50000" /></a:accent1>',
);
assert.notEqual(complexThemeXml, complexThemeSourceXml, "fixture must replace the authored RGB theme slot");
complexThemeZip.file(complexThemePath, complexThemeXml);
const complexThemeBytes = await complexThemeZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const complexThemeImported = await importXlsxWithOpenChestnut(complexThemeBytes);
const complexThemePreserved = await exportXlsxWithOpenChestnut(complexThemeImported, { recalculate: false });
const complexThemePreservedXml = await (await JSZip.loadAsync(complexThemePreserved.bytes)).file(complexThemePath).async("text");
assert.equal(complexThemePreservedXml, complexThemeXml, "unmodeled workbook themes must remain byte-exact when unchanged");
complexThemeImported.setTheme({ name: "Lossy replacement", colors: { ...complexThemeImported.theme.colors, accent1: "#2563EB" } });
await assert.rejects(
  exportXlsxWithOpenChestnut(complexThemeImported, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_theme" && /cannot be replaced losslessly/i.test(error.message),
);

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

const opaqueGraphResponse = await invokeOpenChestnut({
  protocolVersion: 1,
  operation: CodecOperation.IMPORT_XLSX,
  family: ArtifactFamily.WORKBOOK,
  file: await addOpaqueOpcGraph(exported.bytes),
});
const nativeOpaquePart = opaqueGraphResponse.artifact.opaqueOpc.parts.find((part) => part.path === "xl/custom/native.xml");
assert.ok(nativeOpaquePart, "bundled WASM import must expose the custom OPC part");
assert.equal(nativeOpaquePart.contentType, "application/vnd.open-office-artifact-tool.native+xml");
assert.equal(nativeOpaquePart.data.length, 0, "source-backed opaque parts must not duplicate package bytes");
const nativeOutgoingRelationship = nativeOpaquePart.relationships.find((relationship) => relationship.id === "rIdPayload");
assert.ok(nativeOutgoingRelationship, "bundled WASM import must expose source-local OPC relationship adjacency");
assert.equal(nativeOutgoingRelationship.sourcePath, nativeOpaquePart.path);
assert.equal(nativeOutgoingRelationship.target, "https://example.invalid/native-payload");
assert.ok(opaqueGraphResponse.artifact.opaqueOpc.packageRelationships.some((relationship) =>
  relationship.sourcePath === nativeOpaquePart.path && relationship.id === nativeOutgoingRelationship.id));
nativeOutgoingRelationship.target = "https://example.invalid/tampered";
await assert.rejects(
  invokeOpenChestnut({
    protocolVersion: 1,
    operation: CodecOperation.EXPORT_XLSX,
    family: ArtifactFamily.WORKBOOK,
    artifact: opaqueGraphResponse.artifact,
  }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "source_package_graph_mismatch",
  "the bundled runtime must bind per-part relationship metadata to the source snapshot",
);

await assert.rejects(
  importXlsxWithOpenChestnut(exported, { limits: { maxInputBytes: 16 } }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "input_budget_exceeded",
);

const styled = Workbook.create();
const styledSheet = styled.worksheets.add("Sheet1");
styledSheet.getRange("A1:B2").values = [["Label", "Value"], ["styled", 1]];
styledSheet.getRange("A1:B1").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" } };
styledSheet.tables.add("A1:B2", true, "StyledTable").style = "TableStyleMedium4";
styledSheet.getRange("B2").dataValidation = { rule: { type: "list", values: ["1", "2"] } };
await assert.rejects(
  exportXlsxWithOpenChestnut(styled),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_workbook_features",
);
const styledSource = await SpreadsheetFile.exportXlsx(styled);
const styledImported = await importXlsxWithOpenChestnut(styledSource);
styledImported.worksheets.getItem("Sheet1").getRange("B2").values = [[2]];
styledImported.worksheets.getItem("Sheet1").getRange("B2").format.numberFormat = "$#,##0.00";
const styledPreserved = await exportXlsxWithOpenChestnut(styledImported, { recalculate: false });
const styledRoundTrip = await SpreadsheetFile.importXlsx(styledPreserved);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("B2").values[0][0], 2);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("A1").format.font.bold, true);
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").getRange("B2").format.numberFormat, "$#,##0.00");
assert.equal(styledRoundTrip.worksheets.getItem("Sheet1").tables.items[0].name, "StyledTable");

const invalidTableWorkbook = Workbook.create();
const invalidTableSheet = invalidTableWorkbook.worksheets.add("Sheet1");
invalidTableSheet.getRange("A1:B2").values = [["Duplicate", "Duplicate"], [1, 2]];
invalidTableSheet.tables.add({ range: "A1:B2", name: "InvalidColumns", columnNames: ["Duplicate", "Duplicate"] });
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /duplicate column name/i.test(error.message),
);
const invalidFormulaTableWorkbook = Workbook.create();
const invalidFormulaTableSheet = invalidFormulaTableWorkbook.worksheets.add("Sheet1");
invalidFormulaTableSheet.getRange("A1:B2").values = [["Name", "Score"], ["x", 1]];
invalidFormulaTableSheet.tables.add({
  range: "A1:B2",
  name: "InvalidFormulaTable",
  columnDefinitions: [{ name: "Name" }, { name: "Score", calculatedColumnFormula: "LEN([@Name])" }],
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidFormulaTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /invalid calculated-column formula/i.test(error.message),
);
const invalidFilterTableWorkbook = Workbook.create();
const invalidFilterTableSheet = invalidFilterTableWorkbook.worksheets.add("Sheet1");
invalidFilterTableSheet.getRange("A1:B2").values = [["Name", "Score"], ["x", 1]];
invalidFilterTableSheet.tables.add({ range: "A1:B2", name: "InvalidFilterTable", filters: [{ columnIndex: 2, kind: "values", values: ["x"] }] });
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidFilterTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /filter column index/i.test(error.message),
);
const invalidSortTableWorkbook = Workbook.create();
const invalidSortTableSheet = invalidSortTableWorkbook.worksheets.add("Sheet1");
invalidSortTableSheet.getRange("A1:B3").values = [["Name", "Score"], ["x", 1], ["y", 2]];
invalidSortTableSheet.tables.add({
  range: "A1:B3",
  name: "InvalidSortTable",
  sortState: { reference: "A2:B3", conditions: [{ reference: "A2:B3", descending: true }] },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidSortTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /sort condition/i.test(error.message),
);
const invalidColumnSortTableWorkbook = Workbook.create();
const invalidColumnSortTableSheet = invalidColumnSortTableWorkbook.worksheets.add("Sheet1");
invalidColumnSortTableSheet.getRange("A1:B3").values = [["Name", "Score"], ["x", 1], ["y", 2]];
invalidColumnSortTableSheet.tables.add({
  range: "A1:B3",
  name: "InvalidColumnSortTable",
  sortState: { reference: "A2:B3", columnSort: false, conditions: [{ reference: "A2:A3" }] },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidColumnSortTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /inside an AutoFilter/i.test(error.message),
);
const invalidLocaleSortWorkbook = Workbook.create();
const invalidLocaleSortSheet = invalidLocaleSortWorkbook.worksheets.add("Sheet1");
invalidLocaleSortSheet.getRange("A1:B3").values = [["Name", "Score"], ["x", 1], ["y", 2]];
invalidLocaleSortSheet.tables.add({
  range: "A1:B3",
  name: "InvalidLocaleSortTable",
  sortState: { reference: "A2:B3", sortMethod: "radical", conditions: [{ reference: "A2:A3" }] },
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidLocaleSortWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /locale-specific sort method/i.test(error.message),
);
const invalidIconTableWorkbook = Workbook.create();
const invalidIconTableSheet = invalidIconTableWorkbook.worksheets.add("Sheet1");
invalidIconTableSheet.getRange("A1:B2").values = [["Name", "Score"], ["x", 1]];
invalidIconTableSheet.tables.add({
  range: "A1:B2",
  name: "InvalidIconTable",
  filters: [{ columnIndex: 1, kind: "icon", iconSet: "3Arrows", iconId: 3 }],
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidIconTableWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /invalid icon filter/i.test(error.message),
);
const invalidColorTargetWorkbook = Workbook.create();
const invalidColorTargetSheet = invalidColorTargetWorkbook.worksheets.add("Sheet1");
invalidColorTargetSheet.getRange("A1:A2").values = [["Value"], [1]];
invalidColorTargetSheet.tables.add({
  range: "A1:A2",
  name: "InvalidColorTargetTable",
  filters: [{ columnIndex: 0, kind: "color", target: "background", color: "#E11D48" }],
});
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidColorTargetWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /color target/i.test(error.message),
);
const missingTableColorWorkbook = Workbook.create();
const missingTableColorSheet = missingTableColorWorkbook.worksheets.add("Sheet1");
missingTableColorSheet.getRange("A1:A2").values = [["Value"], [1]];
missingTableColorSheet.tables.add({
  range: "A1:A2",
  name: "MissingTableColor",
  filters: [{ columnIndex: 0, kind: "color", target: "cell" }],
});
await assert.rejects(
  exportXlsxWithOpenChestnut(missingTableColorWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_worksheet_table" && /provide a color/i.test(error.message),
);

const invalidNumberFormat = Workbook.create();
invalidNumberFormat.worksheets.add("Sheet1").getRange("A1").values = [[1]];
invalidNumberFormat.worksheets.getItem("Sheet1").getRange("A1").format.numberFormat = "x".repeat(4097);
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidNumberFormat),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_number_format",
);

const invalidFontSize = Workbook.create();
invalidFontSize.worksheets.add("Sheet1").getRange("A1").values = [[1]];
invalidFontSize.worksheets.getItem("Sheet1").getRange("A1").format = { font: { size: 500 } };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidFontSize),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_style" && /font size/i.test(error.message),
);

const invalidThemeColor = Workbook.create();
invalidThemeColor.worksheets.add("Sheet1").getRange("A1").values = [[1]];
invalidThemeColor.worksheets.getItem("Sheet1").getRange("A1").format = { fill: { patternType: "solid", foreground: { theme: 12 } } };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidThemeColor),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_style" && /theme color index/i.test(error.message),
);

const invalidBorderStyle = Workbook.create();
invalidBorderStyle.worksheets.add("Sheet1").getRange("A1").values = [[1]];
invalidBorderStyle.worksheets.getItem("Sheet1").getRange("A1").format = { border: { bottom: { style: "triple", color: "#000000" } } };
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidBorderStyle),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_style" && /border style/i.test(error.message),
);

const invalidWorkbookTheme = Workbook.create();
invalidWorkbookTheme.worksheets.add("Sheet1").getRange("A1").values = [[1]];
invalidWorkbookTheme.theme.colors.accent1 = "not-a-color";
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidWorkbookTheme),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_workbook_theme" && /color/i.test(error.message),
);

const formulaWorkbook = Workbook.create();
const formulaSheet = formulaWorkbook.worksheets.add("FormulaTopology");
formulaSheet.getRange("A1:B2").values = [[2, 3], [4, 5]];
formulaSheet.getRange("C1:C2").formulas = [["=A1+B1"], ["=A2+B2"]];
formulaSheet.getRange("E1").formulas = [["=SUM(A1:A2)"]];
formulaSheet.getRange("E1").values = [[6]];
for (const address of ["C1", "C2"]) {
  const cell = formulaSheet.store.get(address);
  cell.formulaType = "shared";
  cell.sharedIndex = 7;
  cell.sharedRef = "C1:C2";
}
const arrayCell = formulaSheet.store.get("E1");
arrayCell.formulaType = "array";
arrayCell.arrayRef = "E1:E2";
const formulaExported = await exportXlsxWithOpenChestnut(formulaWorkbook, { recalculate: false });
const formulaZip = await JSZip.loadAsync(formulaExported.bytes);
const formulaXml = await formulaZip.file("xl/worksheets/sheet1.xml").async("text");
assert.match(formulaXml, /<x:f t="shared" ref="C1:C2" si="7">A1\+B1<\/x:f>/);
assert.match(formulaXml, /<x:f t="shared" si="7"\s*\/>/);
assert.match(formulaXml, /<x:f t="array" ref="E1:E2">SUM\(A1:A2\)<\/x:f>/);

const formulaImported = await importXlsxWithOpenChestnut(formulaExported);
const importedFormulaSheet = formulaImported.worksheets.getItem("FormulaTopology");
assert.deepEqual(importedFormulaSheet.getRange("C1:C2").formulas, [["=A1+B1"], ["=A2+B2"]]);
assert.deepEqual(
  ["C1", "C2"].map((address) => {
    const cell = importedFormulaSheet.store.get(address);
    return { formulaType: cell.formulaType, sharedIndex: cell.sharedIndex, sharedRef: cell.sharedRef };
  }),
  [
    { formulaType: "shared", sharedIndex: 7, sharedRef: "C1:C2" },
    { formulaType: "shared", sharedIndex: 7, sharedRef: "C1:C2" },
  ],
);
assert.equal(importedFormulaSheet.store.get("E1").formulaType, "array");
assert.equal(importedFormulaSheet.store.get("E1").arrayRef, "E1:E2");

const dynamicWorkbook = Workbook.create();
const dynamicSheet = dynamicWorkbook.worksheets.add("DynamicArray");
dynamicSheet.getRange("A1").formulas = [["=SEQUENCE(2,2,101,1)"]];
const dynamicExported = await exportXlsxWithOpenChestnut(dynamicWorkbook);
const dynamicZip = await JSZip.loadAsync(dynamicExported.bytes);
const dynamicMetadataPath = Object.keys(dynamicZip.files).find((path) => /^xl\/metadata\d*\.xml$/.test(path));
assert.ok(dynamicMetadataPath, "OpenChestnut must emit a workbook cell-metadata part for dynamic arrays.");
const dynamicXml = await dynamicZip.file("xl/worksheets/sheet1.xml").async("text");
const dynamicMetadataXml = await dynamicZip.file(dynamicMetadataPath).async("text");
assert.match(dynamicXml, /<x:c\b(?=[^>]*\br="A1")(?=[^>]*\bcm="1")[^>]*><x:f t="array" ref="A1:B2">SEQUENCE\(2,2,101,1\)<\/x:f>/);
assert.match(dynamicMetadataXml, /name="XLDAPR"/);
assert.match(dynamicMetadataXml, /dynamicArrayProperties\b[^>]*fDynamic="1"[^>]*fCollapsed="0"/);
const dynamicImported = await importXlsxWithOpenChestnut(dynamicExported);
assert.equal(dynamicImported.worksheets.getItem("DynamicArray").store.get("A1").formulaType, "dynamicArray");
assert.equal(dynamicImported.worksheets.getItem("DynamicArray").store.get("A1").dynamicArrayRef, "A1:B2");
assert.deepEqual(dynamicImported.worksheets.getItem("DynamicArray").getRange("A1:B2").values, [[101, 102], [103, 104]]);
const dynamicJavaScriptImported = await SpreadsheetFile.importXlsx(dynamicExported);
assert.equal(dynamicJavaScriptImported.worksheets.getItem("DynamicArray").store.get("A1").formulaType, "dynamicArray");
assert.deepEqual(dynamicJavaScriptImported.worksheets.getItem("DynamicArray").getRange("A1:B2").values, [[101, 102], [103, 104]]);
dynamicImported.worksheets.getItem("DynamicArray").getRange("A1").formulas = [["=SEQUENCE(2,2,201,1)"]];
const dynamicEdited = await exportXlsxWithOpenChestnut(dynamicImported);
const dynamicEditedImport = await importXlsxWithOpenChestnut(dynamicEdited);
assert.equal(dynamicEditedImport.worksheets.getItem("DynamicArray").store.get("A1").formulaType, "dynamicArray");
assert.deepEqual(dynamicEditedImport.worksheets.getItem("DynamicArray").getRange("A1:B2").values, [[201, 202], [203, 204]]);

const blockedDynamicWorkbook = Workbook.create();
const blockedDynamicSheet = blockedDynamicWorkbook.worksheets.add("BlockedDynamic");
blockedDynamicSheet.getRange("C1").values = [["occupied"]];
blockedDynamicSheet.getRange("A1").formulas = [["=SEQUENCE(1,3)"]];
await assert.rejects(
  exportXlsxWithOpenChestnut(blockedDynamicWorkbook),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_dynamic_array_edit" && /blocked dynamic array/.test(error.message),
);

const sourceWithoutDynamicMetadata = Workbook.create();
sourceWithoutDynamicMetadata.worksheets.add("Plain").getRange("A1").values = [["plain"]];
const importedSourceWithoutDynamicMetadata = await importXlsxWithOpenChestnut(await exportXlsxWithOpenChestnut(sourceWithoutDynamicMetadata));
importedSourceWithoutDynamicMetadata.worksheets.getItem("Plain").getRange("C1").formulas = [["=SEQUENCE(1,2)"]];
await assert.rejects(
  exportXlsxWithOpenChestnut(importedSourceWithoutDynamicMetadata),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_dynamic_array_edit" && /recognized XLDAPR/.test(error.message),
);

const formulaJavaScriptImported = await SpreadsheetFile.importXlsx(formulaExported);
assert.deepEqual(formulaJavaScriptImported.worksheets.getItem("FormulaTopology").getRange("C1:C2").formulas, [["=A1+B1"], ["=A2+B2"]]);
assert.equal(formulaJavaScriptImported.worksheets.getItem("FormulaTopology").store.get("C2").sharedRef, "C1:C2");

importedFormulaSheet.getRange("C2").values = [[99]];
importedFormulaSheet.getRange("C1").format.numberFormat = "0.00";
const cachedFormulaEdit = await exportXlsxWithOpenChestnut(formulaImported, { recalculate: false });
const cachedFormulaXml = await (await JSZip.loadAsync(cachedFormulaEdit.bytes)).file("xl/worksheets/sheet1.xml").async("text");
assert.match(cachedFormulaXml, /<x:f t="shared" ref="C1:C2" si="7">A1\+B1<\/x:f>/);
assert.match(cachedFormulaXml, /<x:f t="shared" si="7"\s*\/>/);
assert.equal((await importXlsxWithOpenChestnut(cachedFormulaEdit)).worksheets.getItem("FormulaTopology").store.get("C2").sharedIndex, 7);

importedFormulaSheet.getRange("C2").formulas = [["=A2-B2"]];
assert.equal(importedFormulaSheet.store.get("C1").formulaType, undefined, "Editing one shared member must detach the complete native group.");
assert.equal(importedFormulaSheet.store.get("C2").formulaType, undefined);
const detachedFormulaExport = await exportXlsxWithOpenChestnut(formulaImported, { recalculate: false });
const detachedFormulaXml = await (await JSZip.loadAsync(detachedFormulaExport.bytes)).file("xl/worksheets/sheet1.xml").async("text");
assert.doesNotMatch(detachedFormulaXml, /t="shared"/);
const detachedFormulaRoundTrip = await importXlsxWithOpenChestnut(detachedFormulaExport);
assert.deepEqual(detachedFormulaRoundTrip.worksheets.getItem("FormulaTopology").getRange("C1:C2").formulas, [["=A1+B1"], ["=A2-B2"]]);
assert.equal(detachedFormulaRoundTrip.worksheets.getItem("FormulaTopology").store.get("C1").formulaType, undefined);

const invalidSharedWorkbook = Workbook.create();
const invalidSharedSheet = invalidSharedWorkbook.worksheets.add("InvalidShared");
invalidSharedSheet.getRange("C1").formulas = [["=A1+B1"]];
Object.assign(invalidSharedSheet.store.get("C1"), { formulaType: "shared", sharedIndex: 1, sharedRef: "C1:C2" });
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidSharedWorkbook, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_formula" && /contains 1 members/.test(error.message),
);

const invalidArrayWorkbook = Workbook.create();
const invalidArraySheet = invalidArrayWorkbook.worksheets.add("InvalidArray");
invalidArraySheet.getRange("D1:D2").formulas = [["=1+1"], ["=2+2"]];
Object.assign(invalidArraySheet.store.get("D1"), { formulaType: "array", arrayRef: "D1:D2" });
await assert.rejects(
  exportXlsxWithOpenChestnut(invalidArrayWorkbook, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_formula" && /another formula inside legacy array range/.test(error.message),
);

const oversizedArrayWorkbook = Workbook.create();
const oversizedArraySheet = oversizedArrayWorkbook.worksheets.add("OversizedArray");
oversizedArraySheet.getRange("A1").formulas = [["=1+1"]];
Object.assign(oversizedArraySheet.store.get("A1"), { formulaType: "array", arrayRef: "A1:XFD1048576" });
await assert.rejects(
  exportXlsxWithOpenChestnut(oversizedArrayWorkbook, { recalculate: false }),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_cell_formula" && /topology exceeds 1048576 cells/.test(error.message),
);

const digitFunctionWorkbook = Workbook.create();
const digitFunctionSheet = digitFunctionWorkbook.worksheets.add("DigitFunction");
digitFunctionSheet.getRange("A1:A2").values = [[10], [100]];
digitFunctionSheet.getRange("C1:C2").formulas = [["=LOG10(A1)"], ["=LOG10(A2)"]];
for (const address of ["C1", "C2"]) Object.assign(digitFunctionSheet.store.get(address), { formulaType: "shared", sharedIndex: 9, sharedRef: "C1:C2" });
const digitFunctionExport = await exportXlsxWithOpenChestnut(digitFunctionWorkbook, { recalculate: false });
assert.deepEqual((await importXlsxWithOpenChestnut(digitFunctionExport)).worksheets.getItem("DigitFunction").getRange("C1:C2").formulas, [["=LOG10(A1)"], ["=LOG10(A2)"]]);

const minimalDocument = DocumentModel.create({
  name: "OpenChestnut brief",
  blocks: [
    { kind: "paragraph", text: "Quarterly brief", styleId: "Normal", runs: [{ text: "Quarterly ", style: { bold: true } }, { text: "brief", style: { italic: true } }] },
    {
      kind: "table",
      styleId: "TableGrid",
      values: [["Revenue", "42"], ["Status", "Ready"]],
      widthDxa: 9000,
      indentDxa: 240,
      columnWidthsDxa: [3600, 5400],
      cellMarginsDxa: { top: 60, bottom: 80, start: 100, end: 140 },
      borderColor: "445566",
      borderSize: 8,
      headerFill: "E2E8F0",
    },
  ],
});
const docxExported = await exportDocxWithOpenChestnut(minimalDocument);
assert.deepEqual([...docxExported.bytes.slice(0, 4)], [0x50, 0x4b, 0x03, 0x04]);
assert.equal(docxExported.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
assert.equal(docxExported.metadata.codec, "open-chestnut");
assert.equal((await DocumentFile.inspectDocx(docxExported)).ok, true);
const directFormattedXml = await (await JSZip.loadAsync(docxExported.bytes)).file("word/document.xml").async("text");
assert.match(directFormattedXml, /<w:tblW w:w="9000" w:type="dxa"\s*\/>/);
assert.match(directFormattedXml, /<w:tblInd w:w="240" w:type="dxa"\s*\/>/);
assert.match(directFormattedXml, /<w:tblGrid><w:gridCol w:w="3600"\s*\/><w:gridCol w:w="5400"\s*\/><\/w:tblGrid>/);
assert.match(directFormattedXml, /<w:tblCellMar><w:top w:w="60" w:type="dxa"\s*\/><w:start w:w="100" w:type="dxa"\s*\/><w:bottom w:w="80" w:type="dxa"\s*\/><w:end w:w="140" w:type="dxa"\s*\/><\/w:tblCellMar>/);
assert.match(directFormattedXml, /<w:top w:val="single" w:color="445566" w:sz="8" w:space="0"\s*\/>/);
assert.match(directFormattedXml, /<w:tcW w:w="3600" w:type="dxa"\s*\/>[\s\S]*?<w:shd w:val="clear" w:color="auto" w:fill="E2E8F0"\s*\/>[\s\S]*?<w:rPr><w:b\s*\/><\/w:rPr>/);
const docxImported = await importDocxWithOpenChestnut(docxExported);
assert.equal(docxImported.name, "Imported document");
assert.equal(docxImported.blocks[0].text, "Quarterly brief");
assert.equal(docxImported.blocks[0].runs[0].style.bold, true);
assert.deepEqual(docxImported.blocks[1].values, [["Revenue", "42"], ["Status", "Ready"]]);
assert.equal(docxImported.blocks[1].widthDxa, 9000);
assert.equal(docxImported.blocks[1].indentDxa, 240);
assert.deepEqual(docxImported.blocks[1].columnWidthsDxa, [3600, 5400]);
assert.deepEqual(docxImported.blocks[1].cellMarginsDxa, { top: 60, bottom: 80, start: 100, end: 140 });
assert.equal(docxImported.blocks[1].borderColor, "445566");
assert.equal(docxImported.blocks[1].borderSize, 8);
assert.equal(docxImported.blocks[1].headerFill, "E2E8F0");
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
const tableFormattingImported = await importDocxWithOpenChestnut(docxExported);
tableFormattingImported.blocks[1].widthDxa = 9600;
tableFormattingImported.blocks[1].indentDxa = 360;
tableFormattingImported.blocks[1].columnWidthsDxa = [3200, 6400];
tableFormattingImported.blocks[1].cellMarginsDxa = { top: 40, bottom: 60, start: 80, end: 100 };
tableFormattingImported.blocks[1].borderColor = "AA3300";
tableFormattingImported.blocks[1].borderSize = 12;
tableFormattingImported.blocks[1].headerFill = "FFF2CC";
const tableFormattingEdited = await exportDocxWithOpenChestnut(tableFormattingImported);
const tableFormattingRoundTrip = await importDocxWithOpenChestnut(tableFormattingEdited);
assert.equal(tableFormattingRoundTrip.blocks[1].widthDxa, 9600);
assert.equal(tableFormattingRoundTrip.blocks[1].indentDxa, 360);
assert.deepEqual(tableFormattingRoundTrip.blocks[1].columnWidthsDxa, [3200, 6400]);
assert.deepEqual(tableFormattingRoundTrip.blocks[1].cellMarginsDxa, { top: 40, bottom: 60, start: 80, end: 100 });
assert.equal(tableFormattingRoundTrip.blocks[1].borderColor, "AA3300");
assert.equal(tableFormattingRoundTrip.blocks[1].borderSize, 12);
assert.equal(tableFormattingRoundTrip.blocks[1].headerFill, "FFF2CC");
const unrecognizedFormattingZip = await JSZip.loadAsync(docxExported.bytes);
const unrecognizedFormattingXml = await unrecognizedFormattingZip.file("word/document.xml").async("text");
unrecognizedFormattingZip.file("word/document.xml", unrecognizedFormattingXml.replace("</w:tblPr>", '<w:tblLook w:val="04A0"/></w:tblPr>'));
const unrecognizedFormattingDocx = await unrecognizedFormattingZip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
const unrecognizedFormattingImported = await importDocxWithOpenChestnut(unrecognizedFormattingDocx);
unrecognizedFormattingImported.blocks[1].headerFill = "FFF2CC";
await assert.rejects(
  exportDocxWithOpenChestnut(unrecognizedFormattingImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit" && /recognized the complete bounded profile/.test(error.message),
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
assert.deepEqual(authoredMergedDocument.blocks[0].columnWidthsDxa, [3120, 3120, 3120]);
const authoredMergedDocx = await exportDocxWithOpenChestnut(authoredMergedDocument);
const authoredMergedXml = await (await JSZip.loadAsync(authoredMergedDocx.bytes)).file("word/document.xml").async("text");
assert.match(authoredMergedXml, /<w:tblGrid><w:gridCol w:w="3120"\s*\/><w:gridCol w:w="3120"\s*\/><w:gridCol w:w="3120"\s*\/><\/w:tblGrid>/);
assert.match(authoredMergedXml, /<w:gridSpan w:val="2"\s*\/>/);
assert.match(authoredMergedXml, /<w:vMerge w:val="restart"\s*\/>/);
assert.match(authoredMergedXml, /<w:vMerge w:val="continue"\s*\/>/);
const authoredMergedRoundTrip = await importDocxWithOpenChestnut(authoredMergedDocx);
assert.equal(authoredMergedRoundTrip.blocks[0].gridColumns, 3);
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(0, 0).rowSpan, 2);
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(1, 0).verticalMerge, "continue");
assert.equal(authoredMergedRoundTrip.blocks[0].getCell(2, 1).columnSpan, 2);
assert.deepEqual(authoredMergedRoundTrip.blocks[0].columnWidthsDxa, [3120, 3120, 3120]);

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

const invalidTableFormatting = DocumentModel.create({ blocks: [{ kind: "table", values: [["A"]] }] });
invalidTableFormatting.blocks[0].borderSize = 1;
await assert.rejects(
  exportDocxWithOpenChestnut(invalidTableFormatting),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_document_table" && /borderSize/.test(error.message),
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
  listType: "number", numberFormat: "upperLetter", start: 3, levelText: "%1)", numberingId: 88, abstractNumberingId: 9,
});
groupedNumberingDocument.addListItem("Second grouped item", {
  listType: "number", numberFormat: "upperLetter", start: 3, levelText: "%1)", numberingId: 88, abstractNumberingId: 9,
});
groupedNumberingDocument.addListItem("Nested grouped item", {
  listType: "number", numberFormat: "lowerRoman", level: 2, start: 4, levelText: "%1.%2.%3.", numberingId: 88, abstractNumberingId: 9,
});
const groupedNumberingSource = await exportDocxWithOpenChestnut(groupedNumberingDocument);
const groupedNumberingSourceZip = await JSZip.loadAsync(groupedNumberingSource.bytes);
const groupedNumberingSourceXml = await groupedNumberingSourceZip.file("word/numbering.xml").async("text");
assert.match(groupedNumberingSourceXml, /<w:abstractNum w:abstractNumId="9">[\s\S]*?<w:lvl w:ilvl="0">[\s\S]*?<w:numFmt w:val="upperLetter"\s*\/>[\s\S]*?<w:lvl w:ilvl="2">[\s\S]*?<w:numFmt w:val="lowerRoman"\s*\/>/);
assert.match(groupedNumberingSourceXml, /<w:num w:numId="88"><w:abstractNumId w:val="9"\s*\/><\/w:num>/);
const groupedNumberingImported = await importDocxWithOpenChestnut(groupedNumberingSource);
for (const block of groupedNumberingImported.blocks.slice(0, 2)) {
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
assert.equal(groupedNumberingRoundTrip.blocks.slice(0, 2).every((block) => block.numberFormat === "lowerRoman" && block.start === 5 && block.levelText === "%1."), true);
assert.equal(groupedNumberingRoundTrip.blocks[2].level, 2);
assert.equal(groupedNumberingRoundTrip.blocks[2].start, 4);

const partialNumberingImported = await importDocxWithOpenChestnut(groupedNumberingSource);
partialNumberingImported.blocks[0].start = 9;
await assert.rejects(
  exportDocxWithOpenChestnut(partialNumberingImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_edit" && /coherently/.test(error.message),
);

const conflictingDirectNumbering = DocumentModel.create({ blocks: [] });
conflictingDirectNumbering.addListItem("Decimal", { listType: "number", numberFormat: "decimal", numberingId: 91, abstractNumberingId: 11 });
conflictingDirectNumbering.addListItem("Roman", { listType: "number", numberFormat: "upperRoman", numberingId: 91, abstractNumberingId: 11 });
await assert.rejects(
  exportDocxWithOpenChestnut(conflictingDirectNumbering),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_document_numbering" && /conflicting definitions/.test(error.message),
);

const styleLinkedDirectNumbering = DocumentModel.create({ blocks: [] });
styleLinkedDirectNumbering.addListItem("Style-linked", { listType: "number", numberingId: 92, numberingStyleId: "AgentNumbering" });
await assert.rejects(
  exportDocxWithOpenChestnut(styleLinkedDirectNumbering),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_document_features" && /style-linked/.test(error.message),
);

const automaticDirectNumbering = DocumentModel.create({ blocks: [] });
automaticDirectNumbering.addListItem("First bullet", { listType: "bullet" });
automaticDirectNumbering.addListItem("Second bullet", { listType: "bullet" });
automaticDirectNumbering.addListItem("First number", { listType: "number" });
const automaticDirectRoundTrip = await importDocxWithOpenChestnut(await exportDocxWithOpenChestnut(automaticDirectNumbering));
assert.equal(automaticDirectRoundTrip.blocks[0].numberingId, automaticDirectRoundTrip.blocks[1].numberingId);
assert.notEqual(automaticDirectRoundTrip.blocks[0].numberingId, automaticDirectRoundTrip.blocks[2].numberingId);
assert.ok(automaticDirectRoundTrip.blocks.every((block) => block.numberingId > 0 && block.abstractNumberingId > 0));

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
assert.deepEqual(importedMasterPlaceholder.position, { left: 80, top: 60, width: 720, height: 120 });
assert.equal(importedLayoutPlaceholder.type, "body");
assert.equal(importedLayoutPlaceholder.idx, 2);
assert.equal(importedLayoutPlaceholder.text[0].runs[0].text, "Layout prompt");
assert.deepEqual(importedLayoutPlaceholder.position, { left: 80, top: 200, width: 720, height: 120 });
importedMasterPlaceholder.text[0].runs[0].text = "Edited master prompt";
importedMasterPlaceholder.position = { left: 96, top: 70, width: 700, height: 110 };
importedLayoutPlaceholder.text[0].runs[0] = { text: "Edited layout prompt", style: {}, link: { uri: "https://example.com/layout-help", tooltip: "Layout help" } };
importedLayoutPlaceholder.textBodyProperties.anchor = "bottom";
importedLayoutPlaceholder.position = { left: 88, top: 220, width: 680, height: 100 };
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
assert.match(masterPlaceholderEditedXml, /<a:xfrm\b[^>]*rot="60000"[^>]*><a:off x="914400" y="666750"\s*\/><a:ext cx="6667500" cy="1047750"\s*\/><\/a:xfrm>/);
assert.match(masterPlaceholderEditedXml, /<a:t>Edited master prompt<\/a:t>/);
const layoutEditedXml = await masterStyleEditedZip.file(layoutPartPath).async("text");
assert.match(layoutEditedXml, /<p:bg><p:bgRef idx="1002"><a:schemeClr\b[^>]*val="accent2"[^>]*\/><\/p:bgRef><\/p:bg>/);
assert.match(layoutEditedXml, /<a:bodyPr\b[^>]*anchor="b"/);
assert.match(layoutEditedXml, /<a:xfrm><a:off x="838200" y="2095500"\s*\/><a:ext cx="6477000" cy="952500"\s*\/><\/a:xfrm>/);
assert.match(layoutEditedXml, /<a:t>Edited layout prompt<\/a:t>/);
assert.match(await masterStyleEditedZip.file("ppt/slideLayouts/_rels/slideLayout1.xml.rels").async("text"), /Type="[^"]+\/hyperlink" Target="https:\/\/example\.com\/layout-help" TargetMode="External"/);
const masterStyleRoundTrip = await importPptxWithOpenChestnut(masterStyleEdited);
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.title[0].alignment, "right");
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.body[1], undefined);
assert.equal(masterStyleRoundTrip.master.textParagraphStyles.other[2].bulletImage.uri, "https://example.com/master-marker.png");
assert.deepEqual(masterStyleRoundTrip.master.background, { fill: "#112233", mode: "solid" });
assert.deepEqual(masterStyleRoundTrip.layouts.items[0].background, { fill: "accent2", mode: "reference", index: 1002 });
assert.equal(masterStyleRoundTrip.master.placeholders[0].text[0].runs[0].text, "Edited master prompt");
assert.deepEqual(masterStyleRoundTrip.master.placeholders[0].position, { left: 96, top: 70, width: 700, height: 110 });
assert.equal(masterStyleRoundTrip.layouts.items[0].placeholders[0].text[0].runs[0].link.uri, "https://example.com/layout-help");
assert.equal(masterStyleRoundTrip.layouts.items[0].placeholders[0].textBodyProperties.anchor, "bottom");
assert.deepEqual(masterStyleRoundTrip.layouts.items[0].placeholders[0].position, { left: 88, top: 220, width: 680, height: 100 });
const invalidPlaceholderFrame = await importPptxWithOpenChestnut(masterStyleSource);
invalidPlaceholderFrame.layouts.items[0].placeholders[0].position.width = 0;
await assert.rejects(
  exportPptxWithOpenChestnut(invalidPlaceholderFrame),
  (error) => error instanceof OpenChestnutCodecError && error.code === "invalid_presentation_frame",
);
const inheritedPlaceholderZip = await JSZip.loadAsync(masterStyleSource);
inheritedPlaceholderZip.file(masterPartPath, (await inheritedPlaceholderZip.file(masterPartPath).async("text")).replace(/<a:xfrm rot="60000"><a:off x="762000" y="571500"\/><a:ext cx="6858000" cy="1143000"\/><\/a:xfrm>/, ""));
const inheritedPlaceholderPresentation = await importPptxWithOpenChestnut(await inheritedPlaceholderZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
assert.equal(inheritedPlaceholderPresentation.master.placeholders[0].position, undefined, "Missing direct placeholder geometry must remain inherited rather than gaining a synthetic frame.");
inheritedPlaceholderPresentation.master.placeholders[0].text[0].runs[0].text = "Inherited frame text edit";
const inheritedPlaceholderTextEdit = await exportPptxWithOpenChestnut(inheritedPlaceholderPresentation);
const inheritedPlaceholderTextEditZip = await JSZip.loadAsync(inheritedPlaceholderTextEdit.bytes);
assert.match(await inheritedPlaceholderTextEditZip.file(masterPartPath).async("text"), /<a:t>Inherited frame text edit<\/a:t>/);
const inheritedPlaceholderFrameEdit = await importPptxWithOpenChestnut(await inheritedPlaceholderZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
inheritedPlaceholderFrameEdit.master.placeholders[0].position = { left: 80, top: 60, width: 720, height: 120 };
await assert.rejects(
  exportPptxWithOpenChestnut(inheritedPlaceholderFrameEdit),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);
const complexPlaceholderFrameZip = await JSZip.loadAsync(masterStyleSource);
complexPlaceholderFrameZip.file(layoutPartPath, (await complexPlaceholderFrameZip.file(layoutPartPath).async("text")).replace(
  /(<a:xfrm><a:off x="762000" y="1905000"\/><a:ext cx="6858000" cy="1143000"\/>)/,
  '$1<a:chOff xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" x="0" y="0"/>',
));
const complexPlaceholderFramePresentation = await importPptxWithOpenChestnut(await complexPlaceholderFrameZip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
assert.equal(complexPlaceholderFramePresentation.layouts.items[0].placeholders[0].position, undefined);
complexPlaceholderFramePresentation.layouts.items[0].placeholders[0].position = { left: 80, top: 200, width: 720, height: 120 };
await assert.rejects(
  exportPptxWithOpenChestnut(complexPlaceholderFramePresentation),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);
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
const retainedLayoutBackground = masterStyleImported.layouts.items[0].background;
masterStyleImported.master.clearBackground();
masterStyleImported.layouts.items[0].clearBackground();
const removedBackgrounds = await exportPptxWithOpenChestnut(masterStyleImported);
const removedBackgroundsZip = await JSZip.loadAsync(removedBackgrounds.bytes);
assert.doesNotMatch(await removedBackgroundsZip.file(masterPartPath).async("text"), /<p:bg>/);
assert.doesNotMatch(await removedBackgroundsZip.file(layoutPartPath).async("text"), /<p:bg>/);
const removedBackgroundsRoundTrip = await importPptxWithOpenChestnut(removedBackgrounds);
assert.equal(removedBackgroundsRoundTrip.master.background, undefined);
assert.equal(removedBackgroundsRoundTrip.layouts.items[0].background, undefined);
assert.deepEqual(removedBackgroundsRoundTrip.master.effectiveBackground(), { fill: "#ffffff", mode: "solid" });
assert.deepEqual(removedBackgroundsRoundTrip.layouts.items[0].effectiveBackground(), { fill: "#ffffff", mode: "solid" });
masterStyleImported.master.setBackground(retainedMasterBackground);
masterStyleImported.layouts.items[0].setBackground(retainedLayoutBackground);
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
unsupportedBackgroundImported.layouts.items[0].clearBackground();
await assert.rejects(
  exportPptxWithOpenChestnut(unsupportedBackgroundImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit" && /not safely removable/.test(error.message),
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
const originalEmbeddedWorkbook = Workbook.create();
originalEmbeddedWorkbook.worksheets.add("Embedded").getRange("A1").values = [["Original embedded workbook"]];
const originalEmbeddedWorkbookFile = await SpreadsheetFile.exportXlsx(originalEmbeddedWorkbook);
const sharedOlePresentation = Presentation.create();
sharedOlePresentation.slides.add({ name: "Shared OLE one" }).shapes.add({ name: "One", position: { left: 40, top: 40, width: 300, height: 80 }, text: "One" });
sharedOlePresentation.slides.add({ name: "Shared OLE two" }).shapes.add({ name: "Two", position: { left: 40, top: 40, width: 300, height: 80 }, text: "Two" });
const sharedOleSource = await PresentationFile.exportPptx(sharedOlePresentation);
const sharedOleBytes = await addPresentationNativeGraph(sharedOleSource.bytes, originalEmbeddedWorkbookFile.bytes, { sharedOleOnSecondSlide: true });
const sharedOleFallback = await PresentationFile.importPptx(sharedOleBytes);
const sharedFallbackObjects = sharedOleFallback.slides.items.flatMap((slide) => slide.nativeObjects.items).filter((object) => object.nativeKind === "oleObject");
assert.equal(sharedFallbackObjects.length, 2);
assert.equal(sharedFallbackObjects.every((object) => object.oleWorkbook === undefined), true);
assert.throws(() => sharedFallbackObjects[0].getEmbeddedWorkbook(), /no editable embedded XLSX workbook/);
const sharedOleOpenChestnut = await importPptxWithOpenChestnut(sharedOleBytes);
const sharedOpenChestnutObjects = sharedOleOpenChestnut.slides.items.flatMap((slide) => slide.nativeObjects.items).filter((object) => object.nativeKind === "oleObject");
assert.equal(sharedOpenChestnutObjects.length, 2);
assert.equal(sharedOpenChestnutObjects.every((object) => object.oleWorkbook === undefined), true);
const presentationNativeSource = await addPresentationNativeGraph(new Uint8Array(await presentationSource.arrayBuffer()), originalEmbeddedWorkbookFile.bytes);
const presentationImported = await importPptxWithOpenChestnut(presentationNativeSource);
const importedNativeObjects = presentationImported.slides.getItem(0).nativeObjects.items;
assert.equal(importedNativeObjects.length, 5);
const importedOle = importedNativeObjects.find((object) => object.nativeKind === "oleObject");
const importedDiagram = importedNativeObjects.find((object) => object.nativeKind === "diagram");
const importedContentPart = importedNativeObjects.find((object) => object.nativeKind === "contentPart");
assert.ok(importedOle && importedDiagram && importedContentPart);
assert.equal(importedOle.rootRelationships.length, 2);
assert.equal(importedOle.parts.length, 2);
assert.deepEqual(importedOle.oleWorkbook, {
  partPath: "ppt/embeddings/native-workbook.xlsx",
  contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  sourceSha256: importedOle.parts.find((part) => part.path === "ppt/embeddings/native-workbook.xlsx").sourceSha256,
  relationshipId: "rIdNativeOle",
});
const readOriginalEmbeddedWorkbook = await SpreadsheetFile.importXlsx(importedOle.getEmbeddedWorkbook());
assert.equal(readOriginalEmbeddedWorkbook.worksheets.getItem(0).getRange("A1").values[0][0], "Original embedded workbook");
assert.equal(importedDiagram.rootRelationships.length, 4);
assert.equal(importedDiagram.parts.length, 4);
assert.equal(importedContentPart.rootRelationships.length, 1);
assert.equal(importedContentPart.parts.length, 2);
assert.deepEqual(importedContentPart.parts.map((part) => part.path), ["ppt/customXml/native-content.xml", "ppt/customXml/itemProps1.xml"]);
assert.equal(importedOle.editable, true);
assert.equal(importedDiagram.editable, true);
assert.equal(importedContentPart.editable, true);
assert.equal(importedNativeObjects.filter((object) => !new Set(["oleObject", "diagram", "contentPart"]).has(object.nativeKind)).every((object) => object.editable === false), true);
const nativeInspect = presentationImported.inspect({ kind: "nativeObject", maxChars: 50_000 }).ndjson;
assert.match(nativeInspect, /"nativeKind":"oleObject"/);
assert.match(nativeInspect, /"nativeKind":"diagram"/);
assert.match(nativeInspect, /"nativeKind":"contentPart"/);
assert.match(nativeInspect, /"nativeParts":\[\{"path":"ppt\/customXml\/native-content.xml"/);
assert.match(nativeInspect, /"editableFields":\["name","position","embeddedWorkbook"\]/);
assert.equal(presentationImported.slides.getItem(0).resolve(importedDiagram.id), importedDiagram);
presentationImported.slides.getItem(0).shapes.items[0].text.set("After WASM");
const replacementEmbeddedWorkbook = Workbook.create();
replacementEmbeddedWorkbook.worksheets.add("Embedded").getRange("A1").values = [["Replacement workbook marker"]];
const replacementEmbeddedWorkbookFile = await SpreadsheetFile.exportXlsx(replacementEmbeddedWorkbook);
importedOle.replaceEmbeddedWorkbook(replacementEmbeddedWorkbookFile);
importedOle.setName("Edited embedded workbook").setPosition({ left: 160, top: 120, width: 420, height: 260 });
importedDiagram.setName("Edited SmartArt").setPosition({ left: 80, top: 410, width: 540, height: 150 });
importedContentPart.setName("Edited content part").setPosition({ left: 730, top: 480, width: 130, height: 110 });
const presentationPreserved = await exportPptxWithOpenChestnut(presentationImported);
assert.equal(presentationPreserved.metadata.diagnostics.some((item) => item.code === "opaque_content_preserved"), true);
const presentationPreservedZip = await JSZip.loadAsync(presentationPreserved.bytes);
assert.deepEqual(await presentationPreservedZip.file("ppt/embeddings/native-workbook.xlsx").async("uint8array"), replacementEmbeddedWorkbookFile.bytes);
assert.match(await presentationPreservedZip.file("ppt/customXml/native-content.xml").async("text"), /preserve me/);
const presentationPreservedSlideXml = await presentationPreservedZip.file("ppt/slides/slide1.xml").async("text");
assert.match(presentationPreservedSlideXml, /name="Edited embedded workbook"/);
assert.match(presentationPreservedSlideXml, /<a:off x="1524000" y="1143000"\s*\/>/);
assert.match(presentationPreservedSlideXml, /name="Edited SmartArt"/);
assert.match(presentationPreservedSlideXml, /name="Edited content part"/);
const presentationNativeRoundTrip = await importPptxWithOpenChestnut(presentationPreserved);
const roundTripOle = presentationNativeRoundTrip.slides.getItem(0).nativeObjects.items.find((object) => object.nativeKind === "oleObject");
const readReplacementEmbeddedWorkbook = await SpreadsheetFile.importXlsx(roundTripOle.getEmbeddedWorkbook());
assert.equal(readReplacementEmbeddedWorkbook.worksheets.getItem(0).getRange("A1").values[0][0], "Replacement workbook marker");
const editedNativeRoundTrip = presentationNativeRoundTrip.slides.getItem(0).nativeObjects.items.filter((object) => object.editable);
assert.deepEqual(editedNativeRoundTrip.map((object) => object.name), ["Edited embedded workbook", "Edited SmartArt", "Edited content part"]);
assert.deepEqual(editedNativeRoundTrip.map((object) => object.position), [
  { left: 160, top: 120, width: 420, height: 260 },
  { left: 80, top: 410, width: 540, height: 150 },
  { left: 730, top: 480, width: 130, height: 110 },
]);
const presentationRoundTrip = await PresentationFile.importPptx(presentationPreserved);
assert.equal(presentationRoundTrip.slides.getItem(0).shapes.items[0].text.value, "After WASM");
assert.equal(presentationRoundTrip.slides.getItem(0).images.items.length, 1);
assert.equal(presentationRoundTrip.slides.getItem(0).charts.items.length, 1);
const presentationFallback = await PresentationFile.exportPptx(presentationImported);
const presentationFallbackRoundTrip = await PresentationFile.importPptx(presentationFallback);
const presentationFallbackSlide = presentationFallbackRoundTrip.slides.getItem(0);
assert.deepEqual(presentationFallbackSlide.nativeObjects.items.map((object) => object.nativeKind).sort(), ["diagram", "oleObject"].sort());
assert.equal(presentationFallbackSlide.groups.items[0].nativeObjects.items[0].nativeKind, "contentPart");
assert.equal(presentationFallbackSlide.nativeObjects.items.find((object) => object.nativeKind === "oleObject").name, "Edited embedded workbook");
const fallbackOle = presentationFallbackSlide.nativeObjects.items.find((object) => object.nativeKind === "oleObject");
assert.ok(fallbackOle.oleWorkbook);
const fallbackEmbeddedWorkbook = await SpreadsheetFile.importXlsx(fallbackOle.getEmbeddedWorkbook());
assert.equal(fallbackEmbeddedWorkbook.worksheets.getItem(0).getRange("A1").values[0][0], "Replacement workbook marker");
const fallbackOnlyReplacement = Workbook.create();
fallbackOnlyReplacement.worksheets.add("Embedded").getRange("A1").values = [["JavaScript fallback replacement"]];
fallbackOle.replaceEmbeddedWorkbook(await SpreadsheetFile.exportXlsx(fallbackOnlyReplacement));
const fallbackOnlyExport = await PresentationFile.exportPptx(presentationFallbackRoundTrip);
const fallbackOnlyRoundTrip = await PresentationFile.importPptx(fallbackOnlyExport);
const fallbackOnlyWorkbook = await SpreadsheetFile.importXlsx(fallbackOnlyRoundTrip.slides.getItem(0).nativeObjects.items.find((object) => object.nativeKind === "oleObject").getEmbeddedWorkbook());
assert.equal(fallbackOnlyWorkbook.worksheets.getItem(0).getRange("A1").values[0][0], "JavaScript fallback replacement");
assert.deepEqual(presentationFallbackSlide.nativeObjects.items.find((object) => object.nativeKind === "diagram").position, { left: 80, top: 410, width: 540, height: 150 });
assert.equal(presentationFallbackSlide.groups.items[0].name, "Edited content part");
assert.deepEqual(presentationFallbackSlide.groups.items[0].position, { left: 730, top: 480, width: 130, height: 110 });
const readOnlyNativeObject = presentationImported.slides.getItem(0).nativeObjects.items.find((object) => !object.editable);
assert.ok(readOnlyNativeObject);
assert.throws(() => readOnlyNativeObject.setPosition({ left: 1 }), /read-only/);
assert.throws(() => importedDiagram.replaceEmbeddedWorkbook(replacementEmbeddedWorkbookFile), /no editable embedded XLSX workbook/);
const readOnlyNativeName = readOnlyNativeObject.name;
readOnlyNativeObject.name = "Unsafe native edit";
await assert.rejects(
  exportPptxWithOpenChestnut(presentationImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);
readOnlyNativeObject.name = readOnlyNativeName;
const originalNativeRawXml = importedOle.rawXml;
importedOle.rawXml += " ";
await assert.rejects(
  exportPptxWithOpenChestnut(presentationImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);
importedOle.rawXml = originalNativeRawXml;
const originalReplacementBytes = Uint8Array.from(importedOle.embeddedWorkbookPart().bytes);
importedOle.replaceEmbeddedWorkbook(Uint8Array.of(0x50, 0x4b, 0x03, 0x04, 1, 2, 3, 4));
await assert.rejects(
  exportPptxWithOpenChestnut(presentationImported),
  (error) => error instanceof OpenChestnutCodecError && new Set(["invalid_opc_package", "invalid_presentation_ole_workbook"]).has(error.code),
);
await assert.rejects(PresentationFile.exportPptx(presentationImported), /not a valid XLSX package/);
importedOle.replaceEmbeddedWorkbook(originalReplacementBytes);
const previewPart = importedOle.parts.find((part) => part.path === "ppt/media/native-preview.png");
const originalPreviewBytes = Uint8Array.from(previewPart.bytes);
previewPart.bytes[previewPart.bytes.length - 1] ^= 0x01;
await assert.rejects(
  exportPptxWithOpenChestnut(presentationImported),
  (error) => error instanceof OpenChestnutCodecError && error.code === "unsupported_presentation_edit",
);
await assert.rejects(PresentationFile.exportPptx(presentationImported), /changed read-only part/);
previewPart.bytes = originalPreviewBytes;

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
