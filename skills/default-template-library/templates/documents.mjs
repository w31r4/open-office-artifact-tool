import { DocumentModel } from "open-office-artifact-tool";

import { addDocumentStyles } from "./common.mjs";

function createDocument(name, preset, header) {
  const document = DocumentModel.create({
    name,
    designPreset: preset,
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 10, color: "#18232D" },
    blocks: [],
  });
  addDocumentStyles(document);
  document.addHeader(header, { referenceType: "default", sectionIndex: 0 });
  document.addFooter("Confidential working draft — Page ", {
    referenceType: "default",
    sectionIndex: 0,
    fieldInstruction: "PAGE",
  });
  return document;
}

function numberedActions(document, prefix, actions) {
  actions.forEach((text, index) => document.addListItem(text, {
    name: `${prefix}-${index + 1}`,
    styleId: "Normal",
    listType: "number",
    start: 1,
    level: 0,
    numberFormat: "decimal",
    levelText: "%1.",
    numberingId: 81,
    abstractNumberingId: 8,
  }));
}

export function buildStrategyMemorandum() {
  const document = createDocument("Strategy Memorandum", "source_free_strategy_memorandum", "STRATEGY MEMORANDUM");
  document.addParagraph("Strategy Memorandum", { name: "memo-title", styleId: "TemplateTitle" });
  document.addParagraph("A source-free decision frame for a consequential choice", {
    name: "memo-subtitle",
    styleId: "TemplateSubtitle",
  });
  document.addParagraph("Decision", { name: "decision-heading", styleId: "TemplateHeading" });
  document.addParagraph("Recommendation: commit to the smallest reversible move that produces decision-quality evidence.", {
    name: "recommendation",
    styleId: "TemplateCallout",
    paragraphFormat: { leftIndentTwips: 240, rightIndentTwips: 240 },
  });
  document.addParagraph("Context", { name: "context-heading", styleId: "TemplateHeading" });
  document.addParagraph("State the outcome, current evidence, non-negotiable constraints, and the decision owner before describing a solution.", {
    name: "context",
    styleId: "Normal",
  });
  document.addParagraph("Options", { name: "options-heading", styleId: "TemplateHeading" });
  document.addTable({
    name: "option-comparison",
    styleId: "TableGrid",
    widthDxa: 9000,
    columnWidthsDxa: [2200, 6800],
    headerFill: "D9F0EE",
    borderColor: "9DBCB7",
    borderSize: 6,
    values: [
      ["Option", "Decision frame"],
      ["Hold", "Preserve current capacity. Risk: delays learning. Signal: new evidence changes the constraint."],
      ["Pilot", "Learn with bounded exposure. Risk: requires focused ownership. Signal: success criteria are met."],
      ["Commit", "Capture full value sooner. Risk: locks in an assumption. Signal: pilot evidence is sufficient."],
    ],
  });
  document.addParagraph("Constraints", {
    name: "constraints-heading",
    styleId: "TemplateHeading",
    // Keep the complete checklist and its follow-up actions on one page.
    paragraphFormat: { pageBreakBefore: true },
  });
  [
    "Name the constraint that cannot be traded away.",
    "Separate known facts from assumptions that require validation.",
    "Define the evidence that changes the recommendation.",
  ].forEach((text, index) => document.addListItem(text, {
    name: `constraint-${index + 1}`,
    styleId: "Normal",
    listType: "bullet",
    level: 0,
  }));
  document.addParagraph("Next actions", { name: "next-actions-heading", styleId: "TemplateHeading" });
  numberedActions(document, "next-action", [
    "Assign one accountable owner for the pilot.",
    "Set the earliest decision-review date.",
    "Record the evidence and the final decision in the delivery artifact.",
  ]);
  return document;
}

export function buildDesignReport() {
  const document = createDocument("Design Report", "source_free_design_report", "DESIGN REPORT");
  document.addParagraph("Design Report", { name: "design-report-title", styleId: "TemplateTitle" });
  document.addParagraph("A source-free record of a bounded design decision, evidence, rollout, and verification plan", {
    name: "design-report-subtitle",
    styleId: "TemplateSubtitle",
  });
  document.addParagraph("Executive summary", { name: "executive-summary-heading", styleId: "TemplateHeading" });
  document.addParagraph("Recommendation: approve the smallest implementation that proves the critical user outcome, preserves a rollback path, and produces measurable evidence for the next decision.", {
    name: "design-recommendation",
    styleId: "TemplateCallout",
    paragraphFormat: { leftIndentTwips: 240, rightIndentTwips: 240 },
  });
  document.addParagraph("Problem and success criteria", { name: "problem-heading", styleId: "TemplateHeading" });
  document.addParagraph("Describe the observed user or operator problem, the boundary of this change, and the measurable evidence that will make the decision reversible or irreversible.", {
    name: "problem-context",
    styleId: "Normal",
  });
  document.addTable({
    name: "success-criteria",
    styleId: "TableGrid",
    widthDxa: 9000,
    columnWidthsDxa: [2800, 6200],
    headerFill: "D9F0EE",
    borderColor: "9DBCB7",
    borderSize: 6,
    values: [
      ["Criterion", "Measure and decision use"],
      ["User outcome", "Observed completion confirms the proposed flow."],
      ["Operational safety", "A rollback drill confirms safe operation."],
      ["Delivery cost", "A named constraint confirms the scope."],
    ],
  });
  document.addParagraph("Proposed approach", {
    name: "approach-heading",
    styleId: "TemplateHeading",
    // Keep the heading and its first table together in native pagination.
    paragraphFormat: { pageBreakBefore: true },
  });
  document.addParagraph("State the interface, data boundary, and the one deliberate simplification. Keep unrelated redesigns outside this report so reviewers can validate the requested decision independently.", {
    name: "approach-description",
    styleId: "Normal",
  });
  document.addTable({
    name: "design-boundary",
    styleId: "TableGrid",
    widthDxa: 9000,
    columnWidthsDxa: [2800, 6200],
    headerFill: "E8F5F3",
    borderColor: "91C7C0",
    borderSize: 6,
    values: [
      ["Area", "Decision boundary"],
      ["User flow", "One clear action and result. No silent destructive action. Owner: Product."],
      ["Service boundary", "Validate before side effects. Keep input and output auditable. Owner: Engineering."],
      ["Evidence", "Record outcome and failure reason. Keep the result queryable. Owner: Operations."],
    ],
  });
  document.addParagraph("Risks and mitigations", {
    name: "risks-heading",
    styleId: "TemplateHeading",
    // Keep the complete risk register off the tail of the design-boundary page.
    paragraphFormat: { pageBreakBefore: true },
  });
  document.addTable({
    name: "risk-register",
    styleId: "TableGrid",
    widthDxa: 9000,
    columnWidthsDxa: [3000, 6000],
    headerFill: "FEF3C7",
    borderColor: "E8B854",
    borderSize: 6,
    values: [
      ["Risk", "Mitigation and owner"],
      ["Assumption fails", "Run a pilot. Owner: decision owner."],
      ["Integration change", "Fail closed. Owner: engineering."],
      ["Rollout impact", "Stop and revert. Owner: operations."],
    ],
  });
  document.addParagraph("Rollout and verification", { name: "rollout-heading", styleId: "TemplateHeading" });
  numberedActions(document, "rollout-action", [
    "Confirm named owners, preconditions, and success criteria before enabling the change.",
    "Run the bounded rollout and review the evidence at the agreed checkpoint.",
    "Record the decision, follow-up work, and any changed assumptions in the delivery artifact.",
  ]);
  document.addParagraph("Open decisions", { name: "open-decisions-heading", styleId: "TemplateHeading" });
  document.addParagraph("List only questions that change scope, risk, ownership, or the acceptance criteria. Assign each one an owner and a deadline before implementation begins.", {
    name: "open-decisions",
    styleId: "Normal",
  });
  return document;
}
