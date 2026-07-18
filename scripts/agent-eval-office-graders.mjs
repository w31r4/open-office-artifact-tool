import { docxGradedCaseIds, gradeDocxCase } from "./agent-eval-docx-graders.mjs";
import { gradePptxCase, pptxGradedCaseIds } from "./agent-eval-presentation-graders.mjs";
import { gradeSpreadsheetCase, spreadsheetGradedCaseIds } from "./agent-eval-spreadsheet-graders.mjs";

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };

/**
 * Cross-format dispatch only. Each Office family owns its independent semantic
 * grader; keeping this module small prevents package-level orchestration from
 * becoming a second XLSX/DOCX/PPTX parser.
 */
export async function gradeOfficeCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const spreadsheet = await gradeSpreadsheetCase({ item, workspace, finalMessage, trace, weights });
  if (spreadsheet.supported) return spreadsheet;
  const document = await gradeDocxCase({ item, workspace, finalMessage, trace, weights });
  return document.supported ? document : gradePptxCase({ item, workspace, finalMessage, trace, weights });
}

export const officeGradedCaseIds = new Set([
  ...spreadsheetGradedCaseIds,
  ...docxGradedCaseIds,
  ...pptxGradedCaseIds,
]);
