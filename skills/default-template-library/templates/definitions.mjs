import { buildDesignReport, buildStrategyMemorandum } from "./documents.mjs";
import { buildOperatingReview, buildProjectKickoff } from "./presentations.mjs";
import { buildFinancialBudget, buildProjectTracker } from "./workbooks.mjs";

function definition({ artifactKind, extension, build, title, slideNames, sheetNames, previewSheet, previewRange }) {
  return Object.freeze({
    artifactKind,
    extension,
    build,
    validation: Object.freeze({
      title,
      slideNames: slideNames && Object.freeze([...slideNames]),
      sheetNames: sheetNames && Object.freeze([...sheetNames]),
      previewSheet,
      previewRange,
    }),
  });
}

export const TEMPLATE_DEFINITIONS = Object.freeze({
  "artifact-template-design-report": definition({
    artifactKind: "document",
    extension: ".docx",
    build: buildDesignReport,
    title: "Design Report",
  }),
  "artifact-template-strategy-memorandum": definition({
    artifactKind: "document",
    extension: ".docx",
    build: buildStrategyMemorandum,
    title: "Strategy Memorandum",
  }),
  "artifact-template-operating-review": definition({
    artifactKind: "presentation",
    extension: ".pptx",
    build: buildOperatingReview,
    title: "Operating Review",
    slideNames: ["Operating scorecard", "Delivery and risks", "Decisions and owners"],
  }),
  "artifact-template-project-kickoff": definition({
    artifactKind: "presentation",
    extension: ".pptx",
    build: buildProjectKickoff,
    title: "Project Kickoff",
    slideNames: ["Kickoff overview", "Scope and plan", "Owners and decisions"],
  }),
  "artifact-template-financial-budget": definition({
    artifactKind: "workbook",
    extension: ".xlsx",
    build: buildFinancialBudget,
    title: "Financial Budget",
    sheetNames: ["Assumptions", "Monthly Plan", "Budget Summary"],
    previewSheet: "Budget Summary",
    previewRange: "A1:D7",
  }),
  "artifact-template-project-tracker": definition({
    artifactKind: "workbook",
    extension: ".xlsx",
    build: buildProjectTracker,
    title: "Project Tracker",
    sheetNames: ["Project Summary", "Work Plan", "Risks and Decisions"],
    previewSheet: "Project Summary",
    previewRange: "A1:D8",
  }),
});
