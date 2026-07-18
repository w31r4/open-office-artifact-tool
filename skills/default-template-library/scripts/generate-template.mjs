import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "open-office-artifact-tool";

const TEMPLATE_DEFINITIONS = Object.freeze({
  "artifact-template-strategy-memorandum": Object.freeze({
    artifactKind: "document",
    extension: ".docx",
    build: buildStrategyMemorandum,
  }),
  "artifact-template-project-kickoff": Object.freeze({
    artifactKind: "presentation",
    extension: ".pptx",
    build: buildProjectKickoff,
  }),
  "artifact-template-financial-budget": Object.freeze({
    artifactKind: "workbook",
    extension: ".xlsx",
    build: buildFinancialBudget,
  }),
});

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value.trim();
}

function assertVerified(report, label) {
  if (!report?.ok) throw new Error(`${label} verification failed: ${report?.ndjson || JSON.stringify(report?.issues || [])}`);
}

function noLine() {
  return { fill: "transparent", width: 0 };
}

function addDocumentStyles(document) {
  for (const [id, style] of Object.entries({
    TemplateTitle: {
      name: "Template Title",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 24,
      bold: true,
      color: "#0F3D4C",
      spaceAfterTwips: 90,
      keepNext: true,
    },
    TemplateSubtitle: {
      name: "Template Subtitle",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10.5,
      color: "#4B6470",
      spaceAfterTwips: 150,
      keepNext: true,
    },
    TemplateHeading: {
      name: "Template Heading",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos Display",
      fontSize: 13,
      bold: true,
      color: "#0F3D4C",
      spaceBeforeTwips: 160,
      spaceAfterTwips: 65,
      keepNext: true,
    },
    TemplateCallout: {
      name: "Template Callout",
      type: "paragraph",
      basedOn: "Normal",
      fontFamily: "Aptos",
      fontSize: 10,
      bold: true,
      color: "#0F5F61",
      spaceBeforeTwips: 70,
      spaceAfterTwips: 110,
      keepNext: true,
    },
    TableGrid: {
      name: "Table Grid",
      type: "table",
      fontFamily: "Aptos",
      fontSize: 9,
    },
  })) document.styles.add(id, style);
}

export function buildStrategyMemorandum() {
  const document = DocumentModel.create({
    name: "Strategy Memorandum",
    designPreset: "source_free_strategy_memorandum",
    defaultRunStyle: { fontFamily: "Aptos", fontSize: 10, color: "#18232D" },
    blocks: [],
  });
  addDocumentStyles(document);
  document.addHeader("STRATEGY MEMORANDUM", { referenceType: "default", sectionIndex: 0 });
  document.addFooter("Confidential working draft — Page ", {
    referenceType: "default",
    sectionIndex: 0,
    fieldInstruction: "PAGE",
  });
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
    widthDxa: 9000,
    columnWidthsDxa: [2200, 2400, 2200, 2200],
    headerFill: "D9F0EE",
    borderColor: "9DBCB7",
    borderSize: 6,
    values: [
      ["Option", "Expected outcome", "Main risk", "Decision signal"],
      ["Hold", "Preserve current capacity", "Delays learning", "New evidence changes the constraint"],
      ["Pilot", "Learn with bounded exposure", "Requires focused ownership", "Success criteria are met"],
      ["Commit", "Capture full value sooner", "Locks in an assumption", "Pilot evidence is sufficient"],
    ],
  });
  document.addParagraph("Constraints", { name: "constraints-heading", styleId: "TemplateHeading" });
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
  [
    "Assign one accountable owner for the pilot.",
    "Set the earliest decision-review date.",
    "Record the evidence and the final decision in the delivery artifact.",
  ].forEach((text, index) => document.addListItem(text, {
    name: `next-action-${index + 1}`,
    styleId: "Normal",
    listType: "number",
    start: 1,
    level: 0,
    numberFormat: "decimal",
    levelText: "%1.",
    numberingId: 81,
    abstractNumberingId: 8,
  }));
  return document;
}

function addKickoffText(slide, config) {
  return slide.shapes.add({
    geometry: "rect",
    fill: "transparent",
    line: noLine(),
    ...config,
  });
}

export function buildProjectKickoff() {
  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });
  const title = presentation.slides.add({
    name: "Kickoff overview",
    notes: "Introduce the outcome first, then the commitment requested from the group.",
    background: { fill: "#F4FBFA" },
  });
  title.shapes.add({
    name: "title-accent",
    geometry: "rect",
    position: { left: 0, top: 0, width: 1280, height: 20 },
    fill: "#0F766E",
    line: noLine(),
  });
  addKickoffText(title, {
    name: "kickoff-title",
    text: "Project Kickoff",
    position: { left: 84, top: 108, width: 860, height: 92 },
    textStyle: { fontFamily: "Aptos Display", fontSize: 48, bold: true, color: "#0F3D4C" },
  });
  addKickoffText(title, {
    name: "kickoff-subtitle",
    text: "A source-free working deck for outcome, scope, ownership, and decisions",
    position: { left: 88, top: 220, width: 760, height: 76 },
    textStyle: { fontFamily: "Aptos", fontSize: 21, color: "#48656A" },
  });
  title.shapes.add({
    name: "outcome-panel",
    geometry: "roundRect",
    position: { left: 86, top: 346, width: 1090, height: 196 },
    fill: "#D9F0EE",
    line: { fill: "#88B9B3", width: 1 },
    text: "OUTCOME\nDeliver a measurable first release while preserving a clear rollback path and a shared record of decisions.",
    textStyle: { fontFamily: "Aptos", fontSize: 22, bold: true, color: "#173B3D" },
  });

  const scope = presentation.slides.add({
    name: "Scope and plan",
    notes: "Confirm the boundary before discussing detailed implementation tasks.",
    background: { fill: "#FFFFFF" },
  });
  addKickoffText(scope, {
    name: "scope-title",
    text: "Scope and plan",
    position: { left: 80, top: 64, width: 700, height: 60 },
    textStyle: { fontFamily: "Aptos Display", fontSize: 34, bold: true, color: "#0F3D4C" },
  });
  const planCards = [
    ["In scope", "Outcome, decision record, bounded delivery, and verification evidence."],
    ["Not in scope", "Unbounded redesigns, unspecified integrations, and unowned operational work."],
    ["First checkpoint", "Review evidence, risks, and the next irreversible decision."],
  ];
  planCards.forEach(([heading, body], index) => {
    const left = 80 + index * 385;
    scope.shapes.add({
      name: `scope-card-${index + 1}`,
      geometry: "roundRect",
      position: { left, top: 210, width: 330, height: 282 },
      fill: index === 1 ? "#FEF3C7" : "#E8F5F3",
      line: { fill: index === 1 ? "#E8B854" : "#91C7C0", width: 1 },
      text: `${heading}\n\n${body}`,
      textStyle: { fontFamily: "Aptos", fontSize: 18, bold: true, color: "#0F3D4C" },
    });
  });

  const operating = presentation.slides.add({
    name: "Owners and decisions",
    notes: "Close by making ownership and the decision cadence explicit.",
    background: { fill: "#0F3D4C" },
  });
  addKickoffText(operating, {
    name: "operating-title",
    text: "Owners and decisions",
    position: { left: 80, top: 66, width: 790, height: 60 },
    textStyle: { fontFamily: "Aptos Display", fontSize: 34, bold: true, color: "#FFFFFF" },
  });
  const rows = [
    ["Outcome owner", "Owns priorities, success criteria, and the decision record."],
    ["Delivery owner", "Owns the plan, risk surfacing, and evidence collection."],
    ["Decision cadence", "Review weekly; escalate only a named blocking decision."],
  ];
  rows.forEach(([heading, body], index) => {
    const top = 194 + index * 140;
    operating.shapes.add({
      name: `operating-row-${index + 1}`,
      geometry: "roundRect",
      position: { left: 80, top, width: 1110, height: 102 },
      fill: "#174F5D",
      line: { fill: "#4D8991", width: 1 },
      text: `${heading}: ${body}`,
      textStyle: { fontFamily: "Aptos", fontSize: 17, bold: true, color: "#FFFFFF" },
    });
  });
  return presentation;
}

function titleRange(sheet, range, text) {
  sheet.getRange(range).values = [[text, null, null, null, null]];
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: "#0F3D4C",
    font: { bold: true, color: "#FFFFFF" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 30,
  };
}

function sectionRange(sheet, range, text) {
  sheet.getRange(range).values = [[text, null, null, null, null]];
  sheet.getRange(range).merge();
  sheet.getRange(range).format = {
    fill: "#D9F0EE",
    font: { bold: true, color: "#0F3D4C" },
    alignment: { horizontal: "left", vertical: "center" },
    rowHeightPx: 22,
  };
}

export function buildFinancialBudget() {
  const workbook = Workbook.create({ calculation: { mode: "automatic", fullCalculationOnLoad: true } });
  const assumptions = workbook.worksheets.add("Assumptions");
  const plan = workbook.worksheets.add("Monthly Plan");
  const summary = workbook.worksheets.add("Budget Summary");
  for (const sheet of [assumptions, plan, summary]) sheet.showGridLines = false;

  titleRange(assumptions, "A1:E1", "Financial Budget — Assumptions");
  sectionRange(assumptions, "A3:E3", "Planning inputs");
  assumptions.getRange("A4:E8").values = [
    ["Input", "Value", "Unit", "Owner", "Purpose"],
    ["Starting monthly revenue", 120000, "USD", "Commercial", "Baseline for monthly plan"],
    ["Monthly growth", 0.04, "%", "Commercial", "Revenue change per month"],
    ["Gross margin", 0.62, "%", "Finance", "Gross profit expectation"],
    ["Fixed operating cost", 41000, "USD", "Operations", "Monthly operating cost"],
  ];
  assumptions.getRange("A4:E4").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  assumptions.getRange("B5:B8").format = { fill: "#FFF2CC", font: { bold: true, color: "#1D4ED8" } };
  assumptions.getRange("B6:B7").format.numberFormat = "0.0%";
  assumptions.getRange("B5:B5").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  assumptions.getRange("B8:B8").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  assumptions.getRange("A1:A12").format.columnWidthPx = 210;
  assumptions.getRange("B1:B12").format.columnWidthPx = 120;
  assumptions.getRange("C1:C12").format.columnWidthPx = 95;
  assumptions.getRange("D1:D12").format.columnWidthPx = 125;
  assumptions.getRange("E1:E12").format.columnWidthPx = 245;
  assumptions.freezePanes.freezeRows(4);

  titleRange(plan, "A1:F1", "Financial Budget — Monthly Plan");
  plan.getRange("A3:F3").values = [["Month", "Revenue", "Gross profit", "Operating cost", "Operating result", "Variance to target"]];
  plan.getRange("A3:F3").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  plan.getRange("A4:A9").values = [["Month 1"], ["Month 2"], ["Month 3"], ["Month 4"], ["Month 5"], ["Month 6"]];
  plan.getRange("B4").formulas = [["='Assumptions'!$B$5"]];
  plan.getRange("B5").formulas = [["=B4*(1+'Assumptions'!$B$6)"]];
  plan.getRange("B5:B9").fillDown();
  plan.getRange("C4:C9").formulas = [["=B4*'Assumptions'!$B$7"], ["=B5*'Assumptions'!$B$7"], ["=B6*'Assumptions'!$B$7"], ["=B7*'Assumptions'!$B$7"], ["=B8*'Assumptions'!$B$7"], ["=B9*'Assumptions'!$B$7"]];
  plan.getRange("D4:D9").formulas = [["='Assumptions'!$B$8"], ["='Assumptions'!$B$8"], ["='Assumptions'!$B$8"], ["='Assumptions'!$B$8"], ["='Assumptions'!$B$8"], ["='Assumptions'!$B$8"]];
  plan.getRange("E4:E9").formulas = [["=C4-D4"], ["=C5-D5"], ["=C6-D6"], ["=C7-D7"], ["=C8-D8"], ["=C9-D9"]];
  plan.getRange("F4:F9").formulas = [["=E4-$B$12"], ["=E5-$B$12"], ["=E6-$B$12"], ["=E7-$B$12"], ["=E8-$B$12"], ["=E9-$B$12"]];
  plan.getRange("A11:F11").values = [["Total / target", null, null, null, null, null]];
  plan.getRange("A11:F11").format = { fill: "#D9F0EE", font: { bold: true, color: "#0F3D4C" } };
  plan.getRange("B11:E11").formulas = [["=SUM(B4:B9)", "=SUM(C4:C9)", "=SUM(D4:D9)", "=SUM(E4:E9)"]];
  plan.getRange("B12").formulas = [["=E11/6"]];
  plan.getRange("A12").values = [["Monthly operating-result target"]];
  plan.getRange("A12:B12").format = { fill: "#F8FAFC", font: { bold: true, color: "#0F3D4C" } };
  plan.getRange("B4:F12").format.numberFormat = "$#,##0;[Red]($#,##0);-";
  plan.getRange("A1:A14").format.columnWidthPx = 150;
  plan.getRange("B1:F14").format.columnWidthPx = 135;
  plan.freezePanes.freezeRows(3);

  titleRange(summary, "A1:E1", "Financial Budget — Summary");
  summary.getRange("A3:E3").values = [["Metric", "Value", "Check", "Status", "Owner action"]];
  summary.getRange("A3:E3").format = { fill: "#0F766E", font: { bold: true, color: "#FFFFFF" }, alignment: { horizontal: "center" } };
  summary.getRange("A4:E7").values = [
    ["Six-month revenue", null, "Positive total", null, "Confirm commercial plan"],
    ["Six-month operating result", null, "Positive total", null, "Review cost assumptions"],
    ["Average monthly result", null, "Meets target", null, "Escalate a material shortfall"],
    ["Model status", null, "No failed checks", null, "Approve or revise plan"],
  ];
  summary.getRange("B4:B7").formulas = [["='Monthly Plan'!$B$11"], ["='Monthly Plan'!$E$11"], ["='Monthly Plan'!$B$12"], ["=COUNTIF(D4:D6,\"CHECK\")"]];
  summary.getRange("D4:D6").formulas = [["=IF(B4>0,\"OK\",\"CHECK\")"], ["=IF(B5>0,\"OK\",\"CHECK\")"], ["=IF(B6>0,\"OK\",\"CHECK\")"]];
  summary.getRange("D7").formulas = [["=IF(B7=0,\"OK\",\"CHECK\")"]];
  summary.getRange("B4:B6").format = { fill: "#F0FDF4", font: { bold: true, color: "#166534" }, numberFormat: "$#,##0;[Red]($#,##0);-" };
  summary.getRange("B7:D7").format = { fill: "#F8FAFC", font: { bold: true, color: "#0F3D4C" } };
  summary.getRange("D4:D7").format = { fill: "#DCFCE7", font: { bold: true, color: "#166534" }, alignment: { horizontal: "center" } };
  summary.getRange("D4:D7").conditionalFormats.add("containsText", { text: "CHECK", format: { fill: "#FEE2E2", font: { bold: true, color: "#B91C1C" } } });
  summary.getRange("A1:A10").format.columnWidthPx = 210;
  summary.getRange("B1:B10").format.columnWidthPx = 145;
  summary.getRange("C1:C10").format.columnWidthPx = 145;
  summary.getRange("D1:D10").format.columnWidthPx = 90;
  summary.getRange("E1:E10").format.columnWidthPx = 220;
  summary.freezePanes.freezeRows(3);
  workbook.worksheets.setActiveWorksheet("Budget Summary");
  workbook.recalculate();
  return workbook;
}

async function exportDocumentTemplate(document) {
  assertVerified(document.verify({ visualQa: true }), "Strategy Memorandum model");
  const first = await DocumentFile.exportDocx(document);
  const imported = await DocumentFile.importDocx(first);
  assert.equal(imported.blocks.some((block) => block.text === "Strategy Memorandum"), true, "DOCX import must retain the template title.");
  assertVerified(imported.verify({ visualQa: true }), "Strategy Memorandum first import");
  const final = await DocumentFile.exportDocx(imported);
  const reimported = await DocumentFile.importDocx(final);
  assert.equal(reimported.blocks.some((block) => block.text === "Strategy Memorandum"), true, "DOCX second import must retain the template title.");
  const preview = await reimported.render({ format: "svg" });
  assert.equal(preview.type, "image/svg+xml");
  assert.match(await preview.text(), /Strategy Memorandum/);
  return { file: final, preview, validation: { verify: true, secondImport: true } };
}

async function exportPresentationTemplate(presentation) {
  assertVerified(presentation.verify({ visualQa: true }), "Project Kickoff model");
  const first = await PresentationFile.exportPptx(presentation);
  const imported = await PresentationFile.importPptx(first);
  assert.equal(imported.slides.count, 3, "PPTX import must retain all source-free kickoff slides.");
  assert.equal(imported.slides.getItem(0).name, "Kickoff overview");
  assert.equal(imported.slides.getItem(0).shapes.items.some((shape) => shape.text.value === "Project Kickoff"), true);
  assertVerified(imported.verify({ visualQa: true }), "Project Kickoff first import");
  const final = await PresentationFile.exportPptx(imported);
  const reimported = await PresentationFile.importPptx(final);
  assert.deepEqual(reimported.slides.items.map((slide) => slide.name), ["Kickoff overview", "Scope and plan", "Owners and decisions"]);
  const preview = await reimported.export({ format: "montage" });
  assert.equal(preview.type, "image/svg+xml");
  assert.match(await preview.text(), /Project Kickoff/);
  return { file: final, preview, validation: { verify: true, secondImport: true, slides: 3 } };
}

async function exportWorkbookTemplate(workbook) {
  assertVerified(workbook.verify({ visualQa: true }), "Financial Budget model");
  const first = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
  const imported = await SpreadsheetFile.importXlsx(first);
  imported.recalculate();
  assert.equal(imported.worksheets.getItem("Monthly Plan").getRange("B5").formulas[0][0], "=B4*(1+'Assumptions'!$B$6)");
  assert.deepEqual(imported.worksheets.getItem("Budget Summary").getRange("D4:D7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);
  assertVerified(imported.verify({ visualQa: true }), "Financial Budget first import");
  const final = await SpreadsheetFile.exportXlsx(imported, { recalculate: false });
  const reimported = await SpreadsheetFile.importXlsx(final);
  reimported.recalculate();
  assert.deepEqual(reimported.worksheets.items.map((sheet) => sheet.name), ["Assumptions", "Monthly Plan", "Budget Summary"]);
  assert.deepEqual(reimported.worksheets.getItem("Budget Summary").getRange("D4:D7").values, [["OK"], ["OK"], ["OK"], ["OK"]]);
  const preview = await reimported.render({ sheetName: "Budget Summary", range: "A1:E7", autoCrop: "all", format: "svg" });
  assert.equal(preview.type, "image/svg+xml");
  assert.match(await preview.text(), /Financial Budget/);
  return { file: final, preview, validation: { verify: true, secondImport: true, sheets: 3 } };
}

async function exportSourceFreeTemplate(definition, artifact) {
  if (definition.artifactKind === "document") return exportDocumentTemplate(artifact);
  if (definition.artifactKind === "presentation") return exportPresentationTemplate(artifact);
  if (definition.artifactKind === "workbook") return exportWorkbookTemplate(artifact);
  throw new Error(`Unsupported template artifact kind: ${definition.artifactKind}`);
}

async function assertNewFile(target, label) {
  try {
    await fs.access(target);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}. Choose a new output path instead of overwriting it.`);
}

function normalizedOutputPath(value, definition) {
  const outputPath = path.resolve(requiredText(value, "outputPath"));
  if (path.extname(outputPath).toLowerCase() !== definition.extension) {
    throw new Error(`Output path must end in ${definition.extension}: ${outputPath}`);
  }
  return outputPath;
}

export async function generateTemplate({ templateId, outputPath, auditPath } = {}) {
  const id = requiredText(templateId, "templateId");
  const definition = TEMPLATE_DEFINITIONS[id];
  if (!definition) throw new RangeError(`Unknown or unavailable source-free template: ${id}`);
  const finalPath = normalizedOutputPath(outputPath, definition);
  const finalAuditPath = path.resolve(auditPath ? requiredText(auditPath, "auditPath") : `${finalPath}.audit.json`);
  if (finalAuditPath === finalPath) throw new Error("auditPath must be distinct from outputPath.");
  await Promise.all([
    assertNewFile(finalPath, "Template output"),
    assertNewFile(finalAuditPath, "Template audit"),
  ]);
  await Promise.all([
    fs.mkdir(path.dirname(finalPath), { recursive: true }),
    fs.mkdir(path.dirname(finalAuditPath), { recursive: true }),
  ]);

  const artifact = definition.build();
  const generated = await exportSourceFreeTemplate(definition, artifact);
  const temporaryOutput = `${finalPath}.tmp-${process.pid}-${crypto.randomUUID()}`;
  const temporaryAudit = `${finalAuditPath}.tmp-${process.pid}-${crypto.randomUUID()}`;
  let outputPromoted = false;
  try {
    await generated.file.save(temporaryOutput);
    const bytes = await fs.readFile(temporaryOutput);
    const audit = {
      schema: "open-office-artifact-tool.template-library.v1",
      status: "succeeded",
      template: {
        id,
        artifactKind: definition.artifactKind,
        provenance: "project-authored-source-free",
        retainedReference: false,
        retainedPreview: false,
      },
      source: null,
      output: { path: finalPath, bytes: bytes.length, sha256: sha256(bytes) },
      provider: { actual: "open-chestnut", silentFallback: false },
      savePolicy: { strategy: "create-new" },
      validation: {
        ...generated.validation,
        modelRender: { type: generated.preview.type, bytes: generated.preview.bytes.length },
      },
    };
    await fs.writeFile(temporaryAudit, `${JSON.stringify(audit, null, 2)}\n`);
    await fs.rename(temporaryOutput, finalPath);
    outputPromoted = true;
    await fs.rename(temporaryAudit, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([
      fs.rm(temporaryOutput, { force: true }),
      fs.rm(temporaryAudit, { force: true }),
      outputPromoted ? fs.rm(finalPath, { force: true }) : Promise.resolve(),
    ]);
    throw error;
  }
}

function usage() {
  return [
    "Usage:",
    "  node generate-template.mjs --template-id <id> --output <path> [--audit <path>]",
    "",
    "Ready template IDs:",
    ...Object.entries(TEMPLATE_DEFINITIONS).map(([id, definition]) => `  ${id} (${definition.extension})`),
  ].join("\n");
}

function parseCli(argv) {
  const result = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (token === "--help" || token === "-h") return { help: true };
    if (!token.startsWith("--")) throw new Error(`Unexpected argument: ${token}`);
    const key = token.slice(2);
    if (!new Set(["template-id", "output", "audit"]).has(key)) throw new Error(`Unknown option: ${token}`);
    const value = argv[index + 1];
    if (value === undefined || value.startsWith("--")) throw new Error(`${token} requires a value.`);
    result[key] = value;
    index += 1;
  }
  return {
    templateId: result["template-id"],
    outputPath: result.output,
    auditPath: result.audit,
  };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  try {
    const request = parseCli(process.argv.slice(2));
    if (request.help) {
      console.log(usage());
    } else {
      const result = await generateTemplate(request);
      console.log(JSON.stringify({
        templateId: result.audit.template.id,
        artifactKind: result.audit.template.artifactKind,
        outputPath: result.outputPath,
        auditPath: result.auditPath,
        outputSha256: result.audit.output.sha256,
      }));
    }
  } catch (error) {
    console.error(error instanceof Error ? error.message : String(error));
    process.exitCode = 1;
  }
}
