import fs from "node:fs/promises";
import path from "node:path";

import { SpreadsheetFile, Workbook } from "../src/index.mjs";

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

const XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

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

export async function generateOfficeInput(generator, target) {
  if (generator === "xlsx-threaded-review") return generateXlsxThreadedReview(target);
  return null;
}
