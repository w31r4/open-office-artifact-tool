import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

export const XLSX_THREADED_REVIEW_FIXTURE = Object.freeze({
  workbookName: "reviewed-budget.xlsx",
  sheetName: "Forecast",
  address: "F19",
  root: Object.freeze({
    id: "{11111111-1111-4111-8111-111111111111}",
    personId: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
    author: "Scenario owner",
    userId: "scenario.owner@example.com",
    date: "2026-07-17T09:00:00.000Z",
    text: "Please confirm the downside cash buffer before board circulation.",
  }),
  priorReply: Object.freeze({
    id: "{22222222-2222-4222-8222-222222222222}",
    personId: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    author: "Risk reviewer",
    userId: "risk.reviewer@example.com",
    date: "2026-07-17T09:30:00.000Z",
    text: "Sensitivity analysis is attached to the approved planning case.",
  }),
  requestedReply: "Approved after sensitivity review",
});

export const XLSX_GROWTH_UPDATE_FIXTURE = Object.freeze({
  workbookName: "operating-plan.xlsx",
  targetSheetName: "Forecast",
  canarySheetName: "Approved Baseline",
  growthAddress: "B9",
  marginAddress: "B10",
  originalGrowth: 0.08,
  replacementGrowth: 0.1,
  grossMargin: 0.6,
  revenueFormulas: Object.freeze([
    "=B4*(1+$B$9)",
    "=B5*(1+$B$9)",
    "=B6*(1+$B$9)",
  ]),
  revisedRevenue: Object.freeze([110, 121, 133.1]),
  canaryText: "Approved Baseline — do not modify",
});

// This fixture contains one recognized imported workbook connection plus the
// QueryTable that consumes it. The PromptBench task is deliberately narrower
// than general external-data editing: it may turn only the connection's
// explicit refresh-on-open bit off, leaving the QueryTable and every other
// connection property intact.
export const XLSX_CONNECTION_REFRESH_FIXTURE = Object.freeze({
  workbookName: "external-sales-refresh-on-open.xlsx",
  sheetName: "External Data",
  tableName: "ExternalSales",
  connectionId: 7,
  connectionName: "Fixture warehouse",
  connectionCommand: "SELECT Region, Revenue FROM Sales",
  connectionOpaqueValue: "kept",
});

// This is deliberately a native, source-bound PivotTable fixture rather than
// a hand-written OOXML package. The corresponding PromptBench slice may turn
// off only this uniquely-owned cache's explicit refresh-on-load request.
export const XLSX_PIVOT_REFRESH_FIXTURE = Object.freeze({
  workbookName: "regional-revenue-refresh-on-open.xlsx",
  sheetName: "Pivot Summary",
  pivotName: "Revenue by region",
  sourceSheetName: "Data",
  sourceRange: "Data!A1:B5",
  targetRange: "A1",
});

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

function xmlEscape(value) {
  return String(value).replaceAll("&", "&amp;").replaceAll('"', "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

export const DOCX_CLASSIC_COMMENT_FIXTURE = Object.freeze({
  documentName: "legal-review.docx",
  title: "Controlled rollout legal review",
  anchorText: "Decision: proceed with controlled rollout.",
  supportingText: "The control plan, owner, and retention schedule remain unchanged.",
  comment: Object.freeze({
    author: "Legal reviewer",
    initials: "LR",
    date: "2026-07-18T09:00:00Z",
    originalText: "Please confirm the final retention wording.",
    replacementText: "Approved after legal review.",
  }),
});

// This fixture is intentionally narrow: two ordinary paragraphs share one
// uniquely used default HeaderPart, while a PAGE footer is a canary for the
// source-owned field boundary. The ready PromptBench task may change only the
// first header paragraph and must leave every other package part byte-stable.
export const DOCX_HEADER_TEXT_FIXTURE = Object.freeze({
  documentName: "board-brief-header.docx",
  title: "Board brief — controlled rollout",
  body: Object.freeze([
    "Decision: proceed with the approved controls and named accountable owner.",
    "The review record, retention schedule, and approval evidence remain unchanged.",
    "This document's header is the only requested source-bound edit.",
  ]),
  header: Object.freeze({
    sectionIndex: 0,
    referenceType: "default",
    originalText: "Northwind | Internal",
    replacementText: "Northwind | Reviewed",
    companionText: "Retain the body and footer exactly.",
  }),
  footer: Object.freeze({
    text: "1",
    fieldInstruction: "PAGE",
  }),
});

export const PPTX_TITLE_NOTES_FIXTURE = Object.freeze({
  presentationName: "launch-review.pptx",
  targetSlideName: "Go-no-go decision",
  untouchedSlideName: "Unchanged appendix",
  titleShapeName: "approval-title",
  originalTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  supportingText: "The scope, owner, and retained controls remain unchanged.",
  originalNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
  targetBackground: "#F1F5F9",
  untouchedBackground: "#FFF7ED",
});

// PromptBench keeps the ordinary title/notes example as a small plain-text
// workflow, while its ready presentation slice exercises the more important
// imported rich-notes boundary: a source-bound NotesSlide can change one
// existing ordinary run without flattening paragraphs, bullets, or sibling
// run formatting.
export const PPTX_RICH_NOTES_FIXTURE = Object.freeze({
  presentationName: "rich-notes-review.pptx",
  targetSlideName: "Go-no-go decision",
  untouchedSlideName: "Unchanged appendix",
  titleShapeName: "approval-title",
  originalTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  supportingText: "The speaker-note topology, visible layout, and appendix remain unchanged.",
  originalNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved control set.\nClose with the accountable owner.",
  originalNotesParagraphs: Object.freeze([
    Object.freeze({
      bulletCharacter: "•",
      runs: Object.freeze([
        Object.freeze({ text: "Lead with ", style: Object.freeze({ bold: true, fontSize: 18, fontFamily: "Aptos", color: "#0F172A" }) }),
        Object.freeze({ text: "the pending legal condition.", style: Object.freeze({ italic: true, fontSize: 18, color: "#7C2D12" }) }),
      ]),
    }),
    Object.freeze({
      autoNumber: Object.freeze({ type: "arabicPeriod", startAt: 2 }),
      runs: Object.freeze([
        Object.freeze({ text: "Close with the accountable owner.", style: Object.freeze({ fontSize: 16 }) }),
      ]),
    }),
  ]),
  targetRun: Object.freeze({
    paragraphIndex: 0,
    runIndex: 1,
    expectedText: "the pending legal condition.",
    replacementText: "the approved control set.",
    expectedStyle: Object.freeze({ italic: true, fontSize: 18, color: "#7c2d12" }),
    replacementStyle: Object.freeze({ bold: true, italic: false, fontSize: 18, color: "#0f766e" }),
  }),
  targetBackground: "#F1F5F9",
  untouchedBackground: "#FFF7ED",
});

// Reuse the same two-slide package for the non-visual source-bound rename
// case. The title, notes, direct backgrounds, and appendix are deliberate
// semantic and render canaries: the requested edit is only p:cSld/@name.
export const PPTX_SLIDE_NAME_FIXTURE = Object.freeze({
  presentationName: PPTX_TITLE_NOTES_FIXTURE.presentationName,
  expectedName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
  replacementName: "Go decision: controlled rollout",
  untouchedSlideName: PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
});

// This fixture exercises the narrow imported-slide clone profile rather than
// treating a presentation relationship graph as generally editable. Its
// source slide owns three accepted closed leaves: one canonical notes slide,
// one legacy comments XML leaf with a presentation-wide author catalog, and
// one literal-data chart whose ChartPart has no relationship graph. The
// appendix is a visible/package canary.
export const PPTX_CLOSED_LEAF_CLONE_FIXTURE = Object.freeze({
  presentationName: "release-review.pptx",
  sourceSlideName: "Release decision",
  appendixSlideName: "Appendix canary",
  sourceTitle: "Decision: approve controlled rollout",
  sourceSupportingText: "The original slide, notes, legacy comment, and appendix must remain unchanged.",
  sourceNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
  sourceComment: "Confirm the original evidence before delivery.",
  commentAuthor: "Presentation Reviewer",
  commentCreated: "2026-07-18T03:05:00Z",
  chartTitle: "Control evidence by stage",
  chartCategories: Object.freeze(["Ready", "Watch", "Blocked"]),
  chartSeriesName: "Controls",
  chartValues: Object.freeze([68, 24, 8]),
  customShowName: "Board review route",
  customShowNativeId: 31,
  customShowText: "Open board review route",
  oleObjectName: "Embedded control evidence",
  oleWorkbookPart: "ppt/embeddings/release-control-evidence.xlsx",
  oleWorkbookRelationshipId: "rIdReleaseControlWorkbook",
  olePreviewPart: "ppt/media/release-control-evidence-preview.png",
  olePreviewRelationshipId: "rIdReleaseControlPreview",
  oleWorkbookMarker: "Release control evidence",
  sourceBackground: "#E0F2FE",
  appendixBackground: "#FEF3C7",
  appendixText: "Appendix: immutable evidence",
});

function commentConfig(comment) {
  return {
    id: comment.id,
    personId: comment.personId,
    author: comment.author,
    date: comment.date,
    person: {
      id: comment.personId,
      displayName: comment.author,
      userId: comment.userId,
      providerId: "None",
    },
  };
}

export async function generateXlsxThreadedReview(target) {
  const fixture = XLSX_THREADED_REVIEW_FIXTURE;
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(fixture.sheetName);
  sheet.getRange("A1:F19").values = [
    ["FY27 downside cash review", null, null, null, null, null],
    ["Scenario", "Revenue", "Gross margin", "EBITDA", "Cash buffer", "Board status"],
    ["Base", 1480, 0.58, 255, 122, "Approved"],
    ["Upside", 1640, 0.6, 302, 168, "Approved"],
    ["Downside", 1210, 0.51, 142, 45, "Pending review"],
    [null, null, null, null, null, null],
    ["Control", "Value", null, null, null, null],
    ["Minimum required cash buffer", 40, null, null, null, null],
    ["Downside buffer check", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Board review notes", null, null, null, null, null],
    ["The threaded review target is the final board-status cell below.", null, null, null, null, null],
    [null, null, null, null, null, null],
    ["Forecast", "Value", null, null, null, null],
    ["Opening cash", 210, null, null, null, null],
    ["Operating cash flow", -118, null, null, null, null],
    ["Committed spend", -47, null, null, null, null],
    ["Minimum buffer", -40, null, null, null, null],
    ["Downside cash buffer", null, null, null, null, null],
  ];
  sheet.getRange("B9").formulas = [["=IF(F5=\"Pending review\",\"REVIEW\",\"PASS\")"]];
  sheet.getRange("F19").formulas = [["=SUM(B15:B18)"]];
  sheet.getRange("A1:F1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  sheet.getRange("A2:F2").format = { fill: "#E2E8F0", font: { bold: true } };
  sheet.getRange("A14:B14").format = { fill: "#E2E8F0", font: { bold: true } };
  sheet.getRange("A1:F19").format.columnWidthPx = 130;
  sheet.getRange("A1:A19").format.columnWidthPx = 220;
  sheet.freezePanes.freezeRows(2);

  workbook.comments.setSelf({ displayName: "Finance workflow" });
  const thread = workbook.comments.addThread(
    { cell: sheet.getRange(fixture.address) },
    fixture.root.text,
    { id: "downside-cash-review", author: fixture.root.author, resolved: false, comment: commentConfig(fixture.root) },
  );
  thread.addReply(fixture.priorReply.text, commentConfig(fixture.priorReply));

  workbook.recalculate();
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxGrowthUpdate(target) {
  const fixture = XLSX_GROWTH_UPDATE_FIXTURE;
  const workbook = Workbook.create();
  const forecast = workbook.worksheets.add(fixture.targetSheetName);
  forecast.getRange("A1:D10").values = [
    ["FY27 Operating Plan", null, null, null],
    ["Update only the monthly growth assumption; preserve every formula and the approved baseline.", null, null, null],
    ["Month", "Revenue", "Gross Profit", "Growth"],
    ["Jan", 100, null, null],
    ["Feb", null, null, null],
    ["Mar", null, null, null],
    ["Apr", null, null, null],
    [null, null, null, null],
    ["Monthly growth", fixture.originalGrowth, null, null],
    ["Gross margin", fixture.grossMargin, null, null],
  ];
  forecast.getRange("B5:B7").formulas = fixture.revenueFormulas.map((formula) => [formula]);
  forecast.getRange("C4:C7").formulas = [
    ["=B4*$B$10"],
    ["=B5*$B$10"],
    ["=B6*$B$10"],
    ["=B7*$B$10"],
  ];
  forecast.getRange("D4").formulas = [["=0"]];
  forecast.getRange("D5:D7").formulas = [
    ["=B5/B4-1"],
    ["=B6/B5-1"],
    ["=B7/B6-1"],
  ];
  forecast.getRange("A1:D1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF", size: 14 } };
  forecast.getRange("A3:D3").format = { fill: "#E2E8F0", font: { bold: true } };
  forecast.getRange("A9:B10").format = { fill: "#FEF3C7", font: { bold: true } };
  forecast.getRange("B4:C7").setNumberFormat("$#,##0.00");
  forecast.getRange("B9:B10").setNumberFormat("0.0%");
  forecast.getRange("D4:D7").setNumberFormat("0.0%");
  forecast.getRange("A1:D10").format.columnWidthPx = 150;
  forecast.getRange("A1:A10").format.columnWidthPx = 280;
  forecast.freezePanes.freezeRows(3);

  const baseline = workbook.worksheets.add(fixture.canarySheetName);
  baseline.getRange("A1:C5").values = [
    [fixture.canaryText, null, null],
    ["Metric", "Approved value", "Status"],
    ["Monthly growth", fixture.originalGrowth, "Board approved"],
    ["Gross margin", fixture.grossMargin, "Board approved"],
    ["Scope", "No changes authorized", "Canary"],
  ];
  baseline.getRange("A1:C1").format = { fill: "#14532D", font: { bold: true, color: "#FFFFFF", size: 14 } };
  baseline.getRange("A2:C2").format = { fill: "#DCFCE7", font: { bold: true } };
  baseline.getRange("B3:B4").setNumberFormat("0.0%");
  baseline.getRange("A1:C5").format.columnWidthPx = 170;
  baseline.getRange("A1:A5").format.columnWidthPx = 260;
  baseline.freezePanes.freezeRows(2);

  workbook.recalculate();
  const verification = workbook.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated XLSX growth-update fixture failed model verification: " + verification.ndjson);
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: XLSX_MIME };
}

async function attachConnectionRefreshFixture(file, fixture) {
  const bytes = new Uint8Array(await file.arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const tablePartPath = Object.keys(zip.files).find((name) => /^xl\/tables\/table1\.xml$/i.test(name));
  if (!tablePartPath || zip.file("xl/connections.xml")) throw new Error("XLSX connection-refresh fixture could not find an unbound TablePart.");
  const tableRelationshipPath = `${path.posix.dirname(tablePartPath)}/_rels/${path.posix.basename(tablePartPath)}.rels`;
  const [contentTypes, workbookRelationships] = await Promise.all([
    zip.file("[Content_Types].xml")?.async("text"),
    zip.file("xl/_rels/workbook.xml.rels")?.async("text"),
  ]);
  if (!contentTypes?.includes("</Types>") || !workbookRelationships?.includes("</Relationships>") || zip.file(tableRelationshipPath)) {
    throw new Error("XLSX connection-refresh fixture could not safely add the required OPC relationships.");
  }
  const queryPartPath = "xl/queryTables/queryTable1.xml";
  zip.file("[Content_Types].xml", contentTypes.replace(
    "</Types>",
    `<Override PartName="/xl/connections.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"/><Override PartName="/${queryPartPath}" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"/></Types>`,
  ));
  zip.file("xl/_rels/workbook.xml.rels", workbookRelationships.replace(
    "</Relationships>",
    '<Relationship Id="rIdConnections" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" Target="connections.xml"/></Relationships>',
  ));
  zip.file(tableRelationshipPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdQueryTable" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" Target="../queryTables/queryTable1.xml"/></Relationships>');
  zip.file("xl/connections.xml", `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:connections xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:promptbench"><x:connection id="${fixture.connectionId}" name="${xmlEscape(fixture.connectionName)}" description="Read-only warehouse source" type="5" refreshedVersion="8" keepAlive="0" interval="30" background="1" refreshOnLoad="1" saveData="1" savePassword="0" credentials="integrated"><x:dbPr connection="Provider=Fixture.Provider;Data Source=fixture.invalid" command="${xmlEscape(fixture.connectionCommand)}" commandType="2"/><x:extLst><x:ext uri="{E5A74D42-D212-4CC7-9D5B-A7393F4D8A61}"><fixture:connectionOpaque value="${xmlEscape(fixture.connectionOpaqueValue)}"/></x:ext></x:extLst></x:connection></x:connections>`);
  zip.file(queryPartPath, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><x:queryTable xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:fixture="urn:open-office-artifact-tool:promptbench" name="Warehouse sales" headers="1" rowNumbers="0" disableRefresh="0" backgroundRefresh="1" firstBackgroundRefresh="0" refreshOnLoad="0" growShrinkType="insertClear" fillFormulas="0" removeDataOnSave="0" disableEdit="0" preserveFormatting="1" adjustColumnWidth="1" intermediate="0" connectionId="${fixture.connectionId}"><x:queryTableRefresh preserveSortFilterLayout="1" fieldIdWrapped="0" headersInLastRefresh="1" minimumVersion="0" nextId="3" unboundColumnsLeft="0" unboundColumnsRight="0"><x:queryTableFields count="2"><x:queryTableField id="1" name="Region" dataBound="1" tableColumnId="1" fillFormulas="0" clipped="0"/><x:queryTableField id="2" name="Revenue" dataBound="1" tableColumnId="2"/></x:queryTableFields></x:queryTableRefresh><x:extLst><x:ext uri="{A1D56E5F-35B8-4C51-9C80-779E6A39D52B}"><fixture:queryOpaque value="kept"/></x:ext></x:extLst></x:queryTable>`);
  return new Uint8Array(await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }));
}

export async function generateXlsxConnectionRefresh(target) {
  const fixture = XLSX_CONNECTION_REFRESH_FIXTURE;
  const workbook = Workbook.create();
  const sheet = workbook.worksheets.add(fixture.sheetName);
  sheet.getRange("A1:B3").values = [
    ["Region", "Revenue"],
    ["North", 120],
    ["South", 90],
  ];
  sheet.tables.add("A1:B3", true, fixture.tableName);
  sheet.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  sheet.getRange("A1:B3").format.columnWidthPx = 150;
  sheet.freezePanes.freezeRows(1);
  const exported = await SpreadsheetFile.exportXlsx(workbook);
  const patched = await attachConnectionRefreshFixture(exported, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, patched);
  return { path: target, type: XLSX_MIME };
}

export async function generateXlsxPivotRefresh(target) {
  const fixture = XLSX_PIVOT_REFRESH_FIXTURE;
  const workbook = Workbook.create();
  const data = workbook.worksheets.add(fixture.sourceSheetName);
  data.getRange("A1:B5").values = [
    ["Region", "Revenue"],
    ["East", 120],
    ["West", 90],
    ["East", 30],
    ["North", 60],
  ];
  data.getRange("A1:B1").format = { fill: "#0F172A", font: { bold: true, color: "#FFFFFF" } };
  data.getRange("A1:B5").format.columnWidthPx = 128;
  data.freezePanes.freezeRows(1);
  const summary = workbook.worksheets.add(fixture.sheetName);
  summary.getRange("A1:B5").format = { border: { bottom: { style: "thin", color: "#CBD5E1" } } };
  summary.getRange("A1:B1").format = { fill: "#DBEAFE", font: { bold: true, color: "#1E3A8A" } };
  summary.getRange("A1:B5").format.columnWidthPx = 144;
  summary.pivotTables.add({
    name: fixture.pivotName,
    sourceRange: fixture.sourceRange,
    targetRange: fixture.targetRange,
    rowFields: ["Region"],
    valueFields: [{ field: "Revenue", summarizeBy: "sum", name: "Revenue" }],
    rowGrandTotals: true,
    columnGrandTotals: true,
    refreshPolicy: { refreshOnLoad: true, saveData: true, enableRefresh: true },
  });
  const exported = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const bytes = new Uint8Array(await exported.arrayBuffer());
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir);
  const cacheDefinitions = paths.filter((name) => /^pivotCache\/pivotCacheDefinition\d*\.xml$/i.test(name));
  const pivotTables = paths.filter((name) => /^xl\/pivotTables\/pivotTable\d*\.xml$/i.test(name));
  if (cacheDefinitions.length !== 1 || pivotTables.length !== 1) {
    throw new Error("XLSX Pivot refresh fixture must contain exactly one native PivotTable and cache definition.");
  }
  const cacheXml = await zip.file(cacheDefinitions[0]).async("text");
  if (!/refreshOnLoad="(?:1|true)"/i.test(cacheXml)) {
    throw new Error("XLSX Pivot refresh fixture did not author an explicit cache refreshOnLoad=true request.");
  }
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, bytes);
  return { path: target, type: XLSX_MIME };
}

export async function generateDocxClassicCommentReview(target) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  const decision = document.addParagraph(fixture.anchorText, {
    paragraphFormat: { spaceAfterTwips: 120 },
    runs: [{ text: fixture.anchorText, style: { bold: true } }],
  });
  document.addParagraph(fixture.supportingText, {
    paragraphFormat: { spaceAfterTwips: 120 },
  });
  document.addParagraph("Reviewer instruction: preserve the decision text and update only the attached classic comment.");
  document.addComment(decision, fixture.comment.originalText, {
    author: fixture.comment.author,
    initials: fixture.comment.initials,
    date: fixture.comment.date,
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX classic-comment fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: DOCX_MIME };
}

export async function generateDocxHeaderTextReview(target) {
  const fixture = DOCX_HEADER_TEXT_FIXTURE;
  const document = DocumentModel.create({
    name: fixture.title,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 11, color: "#172033" },
    blocks: [],
  });
  document.addParagraph(fixture.title, {
    paragraphFormat: { spaceAfterTwips: 160 },
    runs: [{ text: fixture.title, style: { bold: true, fontSize: 16, color: "#123B5D" } }],
  });
  for (const text of fixture.body) document.addParagraph(text, { paragraphFormat: { spaceAfterTwips: 120 } });
  document.addHeader(fixture.header.originalText, {
    id: "header/review-target",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
  });
  document.addHeader(fixture.header.companionText, {
    id: "header/companion",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
  });
  document.addFooter(fixture.footer.text, {
    id: "footer/page",
    sectionIndex: fixture.header.sectionIndex,
    referenceType: fixture.header.referenceType,
    fieldInstruction: fixture.footer.fieldInstruction,
  });
  const verification = document.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated DOCX header-text fixture failed model verification: " + verification.ndjson);
  const exported = await DocumentFile.exportDocx(document);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: DOCX_MIME };
}

export async function generatePptxTitleNotesReview(target) {
  const fixture = PPTX_TITLE_NOTES_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const decision = presentation.slides.add({ name: fixture.targetSlideName });
  decision.setBackground({ fill: fixture.targetBackground, mode: "solid" });
  const title = decision.shapes.add({
    name: fixture.titleShapeName,
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.originalTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
  const supporting = decision.shapes.add({
    name: "supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 880, height: 80 },
    text: fixture.supportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  decision.addNotes(fixture.originalNotes);

  const appendix = presentation.slides.add({ name: fixture.untouchedSlideName });
  appendix.setBackground({ fill: fixture.untouchedBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: "Appendix: unchanged evidence",
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#7C2D12" };

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX title/notes fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generatePptxRichNotesReview(target) {
  const fixture = PPTX_RICH_NOTES_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const decision = presentation.slides.add({ name: fixture.targetSlideName });
  decision.setBackground({ fill: fixture.targetBackground, mode: "solid" });
  const title = decision.shapes.add({
    name: fixture.titleShapeName,
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.originalTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0F172A" };
  const supporting = decision.shapes.add({
    name: "supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 880, height: 80 },
    text: fixture.supportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  decision.addNotes(fixture.originalNotesParagraphs);

  const appendix = presentation.slides.add({ name: fixture.untouchedSlideName });
  appendix.setBackground({ fill: fixture.untouchedBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: "Appendix: unchanged evidence",
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#7C2D12" };

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX rich-notes fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await exported.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generatePptxSlideNameReview(target) {
  return generatePptxTitleNotesReview(target);
}

async function addClosedCloneOleWorkbook(exported, fixture) {
  const embeddedWorkbook = Workbook.create();
  embeddedWorkbook.worksheets.add("Evidence").getRange("A1:B3").values = [
    [fixture.oleWorkbookMarker, null],
    ["Control", "Status"],
    ["Release gate", "Approved"],
  ];
  const embeddedWorkbookFile = await SpreadsheetFile.exportXlsx(embeddedWorkbook);
  const zip = await JSZip.loadAsync(exported.bytes);
  const [slideXml, slideRelationships] = await Promise.all([
    zip.file("ppt/slides/slide1.xml").async("text"),
    zip.file("ppt/slides/_rels/slide1.xml.rels").async("text"),
  ]);
  const oleFrame = `<p:graphicFrame xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><p:nvGraphicFramePr><p:cNvPr id="100" name="${fixture.oleObjectName}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="10191750" y="3238500"/><a:ext cx="1524000" cy="1143000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj showAsIcon="1" r:id="${fixture.oleWorkbookRelationshipId}" imgW="965200" imgH="609600" progId="Excel.Sheet.12"><p:embed/><p:pic><p:nvPicPr><p:cNvPr id="0" name=""/><p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${fixture.olePreviewRelationshipId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="10191750" y="3238500"/><a:ext cx="1524000" cy="1143000"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>`;
  return PresentationFile.patchPptx(exported, [
    { path: "ppt/slides/slide1.xml", xml: slideXml.replace("</p:spTree>", `${oleFrame}</p:spTree>`) },
    { path: "ppt/slides/_rels/slide1.xml.rels", xml: slideRelationships.replace("</Relationships>", `<Relationship Id="${fixture.oleWorkbookRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/release-control-evidence.xlsx"/><Relationship Id="${fixture.olePreviewRelationshipId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/release-control-evidence-preview.png"/></Relationships>`) },
    { path: fixture.oleWorkbookPart, bytes: embeddedWorkbookFile.bytes, contentType: XLSX_MIME },
    { path: fixture.olePreviewPart, bytes: Buffer.from("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=", "base64"), contentType: "image/png" },
  ]);
}

export async function generatePptxClosedLeafClone(target) {
  const fixture = PPTX_CLOSED_LEAF_CLONE_FIXTURE;
  const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
  const source = presentation.slides.add({ name: fixture.sourceSlideName });
  source.setBackground({ fill: fixture.sourceBackground, mode: "solid" });
  const title = source.shapes.add({
    name: "release-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 1040, height: 96 },
    text: fixture.sourceTitle,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text.style = { fontSize: 34, bold: true, color: "#0C4A6E" };
  const supporting = source.shapes.add({
    name: "release-supporting-copy",
    geometry: "textbox",
    position: { left: 72, top: 194, width: 920, height: 88 },
    text: fixture.sourceSupportingText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  supporting.text.style = { fontSize: 18, color: "#334155" };
  const route = source.shapes.add({
    name: "board-route-link",
    geometry: "textbox",
    position: { left: 1048, top: 194, width: 184, height: 108 },
    text: [{ runs: [{
      text: fixture.customShowText,
      link: { customShow: fixture.customShowName, returnToSlide: true, tooltip: "Open the board route" },
    }] }],
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  route.text.style = { fontSize: 15, bold: true, color: "#0369A1" };
  source.addNotes(fixture.sourceNotes);
  source.comments.addThread(undefined, fixture.sourceComment, {
    author: fixture.commentAuthor,
    created: fixture.commentCreated,
    position: { x: 360, y: 240 },
  });
  source.charts.add("bar", {
    name: "release-evidence-chart",
    position: { left: 72, top: 318, width: 980, height: 320 },
    title: fixture.chartTitle,
    categories: [...fixture.chartCategories],
    series: [{ name: fixture.chartSeriesName, values: [...fixture.chartValues], fill: "#0284C7" }],
    axes: {
      category: { title: "Evidence stage" },
      value: { title: "Share", min: 0, max: 80, majorUnit: 20 },
    },
    legend: false,
    dataLabels: { showValue: true, position: "top" },
  });

  const appendix = presentation.slides.add({ name: fixture.appendixSlideName });
  appendix.setBackground({ fill: fixture.appendixBackground, mode: "solid" });
  const appendixTitle = appendix.shapes.add({
    name: "appendix-title",
    geometry: "textbox",
    position: { left: 72, top: 72, width: 900, height: 96 },
    text: fixture.appendixText,
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  appendixTitle.text.style = { fontSize: 30, bold: true, color: "#92400E" };
  presentation.customShows.add({
    name: fixture.customShowName,
    nativeId: fixture.customShowNativeId,
    slides: [source, appendix],
  });

  const verification = presentation.verify({ visualQa: true });
  if (!verification.ok) throw new Error("Generated PPTX closed-leaf clone fixture failed model verification: " + verification.ndjson);
  const exported = await PresentationFile.exportPptx(presentation);
  const patchedSource = await addClosedCloneOleWorkbook(exported, fixture);
  await fs.mkdir(path.dirname(target), { recursive: true });
  await fs.writeFile(target, new Uint8Array(await patchedSource.arrayBuffer()));
  return { path: target, type: PPTX_MIME };
}

export async function generateOfficeInput(generator, target) {
  if (generator === "xlsx-threaded-review") return generateXlsxThreadedReview(target);
  if (generator === "xlsx-growth-update") return generateXlsxGrowthUpdate(target);
  if (generator === "xlsx-connection-refresh") return generateXlsxConnectionRefresh(target);
  if (generator === "xlsx-pivot-refresh") return generateXlsxPivotRefresh(target);
  if (generator === "docx-classic-comment-review") return generateDocxClassicCommentReview(target);
  if (generator === "docx-header-text-review") return generateDocxHeaderTextReview(target);
  if (generator === "pptx-title-notes-review") return generatePptxTitleNotesReview(target);
  if (generator === "pptx-rich-notes-review") return generatePptxRichNotesReview(target);
  if (generator === "pptx-slide-name-review") return generatePptxSlideNameReview(target);
  if (generator === "pptx-closed-leaf-clone") return generatePptxClosedLeafClone(target);
  return null;
}
