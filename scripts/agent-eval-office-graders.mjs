import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { XLSX_THREADED_REVIEW_FIXTURE } from "./agent-eval-office-fixtures.mjs";
import { docxGradedCaseIds, gradeDocxCase } from "./agent-eval-docx-graders.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";
import { gradePptxCase, pptxGradedCaseIds } from "./agent-eval-presentation-graders.mjs";

const xlsxGradedCaseIds = new Set(["xlsx-threaded-reply-resolve"]);
const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;
const SHIPPED_THREADED_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/spreadsheets|node_modules\/open-office-artifact-tool\/skills\/spreadsheets\/skills\/spreadsheets)\/examples\/openchestnut-threaded-comment-reply-workflow\.mjs(?:$|[\s"'`])/i;

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function decodeXml(value = "") {
  return String(value)
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", "\"")
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function xmlAttributes(opening = "") {
  const result = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    const name = match[1].split(":").at(-1);
    result[name] = decodeXml(match[2]);
  }
  return result;
}

function innerText(xml = "") {
  return decodeXml(String(xml).replace(/<[^>]+>/g, ""));
}

function parseThreadedComments(xml = "") {
  const comments = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?threadedComment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?threadedComment>/g)) {
    const opening = /^<(?:[\w.-]+:)?threadedComment\b[^>]*>/.exec(match[0])?.[0] || "";
    const attributes = xmlAttributes(opening);
    comments.push({
      id: String(attributes.id || "").toUpperCase(),
      parentId: attributes.parentId ? String(attributes.parentId).toUpperCase() : null,
      personId: String(attributes.personId || "").toUpperCase(),
      ref: String(attributes.ref || "").toUpperCase(),
      date: attributes.dT || null,
      done: new Set(["1", "true", "on"]).has(String(attributes.done || "0").toLowerCase()),
      text: innerText(match[1]),
    });
  }
  return comments;
}

function parsePersons(xml = "") {
  const people = new Map();
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?person\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    const id = String(attributes.id || "").toUpperCase();
    if (!id) continue;
    people.set(id, {
      id,
      displayName: attributes.displayName || "",
      userId: attributes.userId || "",
      providerId: attributes.providerId || "",
    });
  }
  return people;
}

export async function inspectThreadedWorkbook(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const threadedPaths = paths.filter((name) => /^xl\/threadedcomments\/[^/]+\.xml$/i.test(name));
  const personPaths = paths.filter((name) => /^xl\/persons\/[^/]+\.xml$/i.test(name));
  const classicCommentPaths = paths.filter((name) => /^xl\/comments\d*\.xml$/i.test(name));
  const threadedXml = await Promise.all(threadedPaths.map((name) => zip.file(name).async("text")));
  const personXml = await Promise.all(personPaths.map((name) => zip.file(name).async("text")));
  const workbookXml = await zip.file("xl/workbook.xml")?.async("text") || "";
  const comments = threadedXml.flatMap(parseThreadedComments);
  const people = new Map();
  for (const xml of personXml) {
    for (const [id, person] of parsePersons(xml)) people.set(id, person);
  }
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    threadedPaths,
    personPaths,
    classicCommentPaths,
    comments,
    people: [...people.values()],
    forecastSheetPresent: /<[^>]*sheet\b[^>]*name="Forecast"/i.test(workbookXml),
  };
}

function auditProvider(audit) {
  const provider = audit?.provider;
  return String(typeof provider === "string" ? provider : provider?.actual || provider?.selected || provider?.name || "");
}

function auditVersion(audit) {
  const provider = audit?.provider;
  return String(provider?.version || audit?.providerVersion || "");
}

function auditFallbackIsFalse(audit) {
  const provider = audit?.provider || {};
  return provider.silentFallback === false || provider.fallbackUsed === false || audit?.silentFallback === false || audit?.fallbackUsed === false;
}

function auditStrategy(audit) {
  const policy = audit?.savePolicy || audit?.savePolicy || audit?.save_strategy;
  return String(typeof policy === "string" ? policy : policy?.strategy || audit?.strategy || "");
}

function auditOperation(audit) {
  const operation = audit?.operation;
  return String(typeof operation === "string" ? operation : operation?.type || operation?.name || "");
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[`${side}Sha256`] || "");
}

function usedTypedXlsxRoundTrip(commandText) {
  const directPublicApi = /(?:SpreadsheetFile\.)?importXlsx/i.test(commandText) && /(?:SpreadsheetFile\.)?exportXlsx/i.test(commandText);
  // The runner fingerprints both the copied .agents Skill and installed package tree
  // before execution. Accepting this exact published entrypoint therefore preserves
  // the typed-route requirement without rewarding an arbitrary local wrapper.
  return directPublicApi || SHIPPED_THREADED_WORKFLOW.test(commandText);
}

function originalCommentEquivalent(left, right) {
  return left && right
    && left.id === right.id
    && left.parentId === right.parentId
    && left.personId === right.personId
    && left.ref === right.ref
    && left.date === right.date
    && left.text === right.text;
}

export function gradeXlsxThreadedReplyEvidence({ evidence, audit, commands, item }) {
  const fixture = XLSX_THREADED_REVIEW_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceRoot = source.comments.find((comment) => comment.id === fixture.root.id);
  const sourcePriorReply = source.comments.find((comment) => comment.id === fixture.priorReply.id);
  const outputRoot = output.comments.find((comment) => comment.id === fixture.root.id);
  const outputPriorReply = output.comments.find((comment) => comment.id === fixture.priorReply.id);
  const requested = output.comments.filter((comment) => comment.text === fixture.requestedReply);
  const added = requested.length === 1 ? requested[0] : null;
  const outputPeople = new Map(output.people.map((person) => [person.id, person]));
  const commandText = commands.join("\n");
  const visual = evidence.visual;
  const visualAvailable = Boolean(visual?.source?.available && visual?.output?.available);
  const outputRendered = visual?.output?.ok === true && visual.output.pages.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = visual?.source?.pageCount === visual?.output?.pageCount;
  const directReply = {
    exactlyOneRequestedText: requested.length === 1,
    expectedCommentCount: output.comments.length === source.comments.length + 1,
    targetCell: added?.ref === fixture.address,
    rootParent: added?.parentId === fixture.root.id,
    newCommentId: Boolean(added && added.id !== fixture.root.id && added.id !== fixture.priorReply.id && GUID.test(added.id)),
    newPersonId: Boolean(added && GUID.test(added.personId) && outputPeople.has(added.personId)),
  };
  return [
    check("xlsx-machine:forecast-sheet", "machine", source.forecastSheetPresent && output.forecastSheetPresent),
    check("xlsx-machine:canonical-threaded-parts", "machine", source.threadedPaths.length === 1 && output.threadedPaths.length === 1 && source.personPaths.length === 1 && output.personPaths.length === 1, {
      source: { threaded: source.threadedPaths, persons: source.personPaths },
      output: { threaded: output.threadedPaths, persons: output.personPaths },
    }),
    check("xlsx-machine:original-thread-preserved", "machine", originalCommentEquivalent(sourceRoot, outputRoot) && originalCommentEquivalent(sourcePriorReply, outputPriorReply), {
      sourceRoot,
      outputRoot,
      sourcePriorReply,
      outputPriorReply,
    }),
    check("xlsx-machine:direct-reply", "machine", Object.values(directReply).every(Boolean), { added, directReply, outputComments: output.comments }),
    check("xlsx-machine:reply-order", "machine", output.comments.map((comment) => comment.id).slice(0, 3).join(",") === [fixture.root.id, fixture.priorReply.id, added?.id].join(","), {
      ids: output.comments.map((comment) => comment.id),
    }),
    check("xlsx-machine:thread-resolved", "machine", output.comments.length === 3 && output.comments.every((comment) => comment.done), {
      done: output.comments.map((comment) => ({ id: comment.id, done: comment.done })),
    }),
    check("xlsx-visual:native-render", "visual", visualAvailable && outputRendered && pageCountsMatch, { visual }),
    gate("xlsx-security:no-classic-note-downgrade", "security", output.classicCommentPaths.length === 0 && output.threadedPaths.length === 1, {
      classicCommentPaths: output.classicCommentPaths,
      threadedPaths: output.threadedPaths,
    }),
    gate("xlsx-security:audit-byte-provenance", "security", auditHash(audit, "source") === source.sha256 && auditHash(audit, "output") === output.sha256 && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("xlsx-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("xlsx-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("xlsx-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), { strategy: auditStrategy(audit) }),
    check("xlsx-trace:threaded-operation", "trace", /thread|comment/i.test(auditOperation(audit)), { operation: auditOperation(audit) }),
    check("xlsx-trace:typed-roundtrip", "trace", usedTypedXlsxRoundTrip(commandText), {
      expected: "public SpreadsheetFile importXlsx/exportXlsx calls or the integrity-protected published threaded-comment workflow",
    }),
    check("xlsx-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

async function readAudit(workspace) {
  try {
    return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8"));
  } catch {
    return null;
  }
}

export async function gradeOfficeCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  if (!xlsxGradedCaseIds.has(item.id)) {
    const docx = await gradeDocxCase({ item, workspace, finalMessage, trace, weights });
    return docx.supported ? docx : gradePptxCase({ item, workspace, finalMessage, trace, weights });
  }
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  let source;
  let output;
  try {
    source = await inspectThreadedWorkbook(path.join(workspace, "inputs", XLSX_THREADED_REVIEW_FIXTURE.workbookName));
    output = await inspectThreadedWorkbook(path.join(workspace, "outputs", "reviewed-budget-resolved.xlsx"));
  } catch (error) {
    const checks = [
      gate("xlsx-machine:readable-output", "machine", false, { error: error.message }),
      gate("xlsx-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(path.join(workspace, "inputs", XLSX_THREADED_REVIEW_FIXTURE.workbookName), "xlsx-source"),
    renderOfficeFile(path.join(workspace, "outputs", "reviewed-budget-resolved.xlsx"), "xlsx-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler spreadsheet rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }
  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeXlsxThreadedReplyEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

export const officeGradedCaseIds = new Set([...xlsxGradedCaseIds, ...docxGradedCaseIds, ...pptxGradedCaseIds]);
