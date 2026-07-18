import fs from "node:fs/promises";
import path from "node:path";

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

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

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

// Reuse the same two-slide package for the non-visual source-bound rename
// case. The title, notes, direct backgrounds, and appendix are deliberate
// semantic and render canaries: the requested edit is only p:cSld/@name.
export const PPTX_SLIDE_NAME_FIXTURE = Object.freeze({
  presentationName: PPTX_TITLE_NOTES_FIXTURE.presentationName,
  expectedName: PPTX_TITLE_NOTES_FIXTURE.targetSlideName,
  replacementName: "Go decision: controlled rollout",
  untouchedSlideName: PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName,
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

export async function generatePptxSlideNameReview(target) {
  return generatePptxTitleNotesReview(target);
}

export async function generateOfficeInput(generator, target) {
  if (generator === "xlsx-threaded-review") return generateXlsxThreadedReview(target);
  if (generator === "xlsx-growth-update") return generateXlsxGrowthUpdate(target);
  if (generator === "docx-classic-comment-review") return generateDocxClassicCommentReview(target);
  if (generator === "pptx-title-notes-review") return generatePptxTitleNotesReview(target);
  if (generator === "pptx-slide-name-review") return generatePptxSlideNameReview(target);
  return null;
}
