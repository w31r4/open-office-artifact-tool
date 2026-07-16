import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const oracleScript = path.join(scriptDirectory, "agent-eval-pdf-oracle.py");
const supportedCases = new Set([
  "pdf-bounded-contract-id-replace",
  "pdf-overflow-replace-refusal",
  "pdf-active-content-public-sanitize",
]);
const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };

function check(id, category, passed, details = {}) {
  return { id, category, gate: false, passed: Boolean(passed), ...details };
}

function gate(id, category, passed, details = {}) {
  return { id, category, gate: true, passed: Boolean(passed), ...details };
}

function sameNumber(left, right, tolerance = 0.001) {
  return Number.isFinite(left) && Number.isFinite(right) && Math.abs(left - right) <= tolerance;
}

function samePageBoxes(sourcePages = [], outputPages = []) {
  return sourcePages.length === outputPages.length && sourcePages.every((source, index) => {
    const output = outputPages[index];
    return output
      && sameNumber(source.width, output.width)
      && sameNumber(source.height, output.height)
      && source.rotation === output.rotation;
  });
}

function sameBoundingBox(left = [], right = [], tolerance = 0.01) {
  return left.length === 4 && right.length === 4 && left.every((value, index) => sameNumber(value, right[index], tolerance));
}

function auditProvider(audit) {
  const provider = audit?.provider;
  if (typeof provider === "string") return provider;
  return provider?.actual || provider?.selected || provider?.name || provider?.provider || audit?.actualProvider || "";
}

function auditFallback(audit) {
  const provider = audit?.provider;
  const values = [
    provider?.silentFallback,
    provider?.silent_fallback,
    provider?.fallbackUsed,
    provider?.fallback_used,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length ? values.every((value) => value === false || value === "false") : null;
}

function auditProviderVersion(audit) {
  const provider = audit?.provider;
  return provider?.version || provider?.providerVersion || audit?.providerVersion || audit?.provider_version || "";
}

function auditSourceHash(audit) {
  return audit?.source?.sha256 || audit?.sourceSha256 || audit?.source_sha256 || "";
}

function auditOutputHash(audit) {
  return audit?.output?.sha256 || audit?.outputSha256 || audit?.output_sha256 || "";
}

function auditSaveStrategy(audit) {
  const policy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  if (typeof policy === "string") return policy;
  return policy?.strategy || policy?.selected || audit?.strategy || "";
}

function auditOperation(audit) {
  const operation = audit?.operation;
  if (typeof operation === "string") return operation;
  if (Array.isArray(operation)) return operation.map((value) => typeof value === "string" ? value : value?.type || value?.operation || "").join(" ");
  return operation?.type || operation?.operation || operation?.name || operation?.performed || "";
}

export function extractCompletedCommands(trace) {
  const commands = new Map();
  for (const line of String(trace || "").split(/\r?\n/)) {
    if (!line.trim()) continue;
    let event;
    try { event = JSON.parse(line); } catch { continue; }
    const item = event?.item;
    if (event?.type === "item.completed" && item?.type === "command_execution" && typeof item.command === "string") {
      commands.set(item.id || `${commands.size}`, item.command);
    }
  }
  return [...commands.values()];
}

function boundedTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const probeIndex = commandText.search(/pymupdf_edit\.py\s+probe\b/i);
  const editIndex = commandText.search(/pymupdf_edit\.py\s+edit\b/i);
  const bypassPatterns = [
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bxref_set_(?:key|stream)\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^sanitize$/i.test(String(auditSaveStrategy(audit))), { expected: "sanitize", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:capability-probe", "trace", probeIndex >= 0 && editIndex >= 0 && probeIndex < editIndex, { expected: "shipped PyMuPDF capability probe before mutation", actual: { probeObserved: probeIndex >= 0, editObserved: editIndex >= 0, preflightOrder: probeIndex >= 0 && editIndex >= 0 && probeIndex < editIndex } }),
    check("pdf-trace:typed-edit-primitive", "trace", /pymupdf_edit\.py\s+edit\b/i.test(commandText) && /replace[_ -]?text/i.test(operation), { expected: "pymupdf_edit.py edit + replace_text", actual: operation || "unreported" }),
    check("pdf-trace:no-content-stream-bypass", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function overflowTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const providerProbe = /pymupdf_edit\.py\s+probe\b/i.test(commandText)
    || /\b(?:import|from)\s+(?:fitz|pymupdf)\b/i.test(commandText)
    || /importlib\.metadata[\s\S]{0,200}pymupdf/i.test(commandText);
  const mutationPatterns = [
    /pymupdf_edit\.py\s+edit\b/i,
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\badd_redact_annot\s*\(/i,
    /\bapply_redactions\s*\(/i,
    /\binsert_text(?:box)?\s*\(/i,
    /\bdoc(?:ument)?\.save\s*\(/i,
    /\bwriter\.write\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:capability-probe", "trace", providerProbe, { expected: "PyMuPDF capability/version evidence" }),
    gate("pdf-trace:no-mutation-after-failed-preflight", "trace", !mutationPatterns.some((pattern) => pattern.test(commandText)), { forbidden: mutationPatterns.map(String) }),
  ];
}

function activeContentTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const invocationIndex = (pattern) => commands.findIndex((command) => {
    const match = String(command).match(pattern);
    if (!match) return false;
    const tail = String(command).slice((match.index || 0) + match[0].length);
    return !/^\s+(?:--help|-h)\b/i.test(tail);
  });
  const probeIndex = invocationIndex(/pymupdf_edit\.py\s+probe\b/i);
  const planIndex = invocationIndex(/pdf_provider\.py\s+plan\b/i);
  const editIndex = invocationIndex(/pymupdf_edit\.py\s+edit\b/i);
  const afterEdit = editIndex >= 0 ? commands.slice(editIndex).join("\n") : "";
  const residueAfterEdit = /residue_scan\.py\b/i.test(afterEdit);
  const renderAfterEdit = /\bpdftoppm\b/i.test(afterEdit);
  const auditAfterEdit = /pdf_audit\.py\s+validate\b/i.test(afterEdit);
  const bypassPatterns = [
    /\bupdate_stream\s*\(/i,
    /\bset_contents\s*\(/i,
    /\bxref_set_(?:key|stream)\s*\(/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /pymupdf/i.test(String(auditProvider(audit))), { expected: "pymupdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^sanitize$/i.test(String(auditSaveStrategy(audit))), { expected: "sanitize", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:probe-plan-before-mutation", "trace", probeIndex >= 0 && planIndex >= 0 && editIndex > probeIndex && editIndex > planIndex, {
      actual: {
        probeObserved: probeIndex >= 0,
        planObserved: planIndex >= 0,
        editObserved: editIndex >= 0,
        probeCommandIndex: probeIndex,
        planCommandIndex: planIndex,
        editCommandIndex: editIndex,
      },
    }),
    check("pdf-trace:typed-scrub-primitive", "trace", /pymupdf_edit\.py\s+edit\b/i.test(commandText) && /scrub|active[_ -]?content/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:post-mutation-residue-scan", "trace", residueAfterEdit, { actual: { editObserved: editIndex >= 0, postMutationResidueScanObserved: residueAfterEdit } }),
    check("pdf-trace:post-mutation-poppler-render", "trace", renderAfterEdit, { actual: { editObserved: editIndex >= 0, postMutationRenderObserved: renderAfterEdit } }),
    check("pdf-trace:audit-byte-validation", "trace", auditAfterEdit, { actual: { postMutationAuditValidationObserved: auditAfterEdit } }),
    gate("pdf-trace:no-content-stream-bypass", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

export function gradeBoundedReplaceEvidence({ evidence, audit, commands, item }) {
  const expectedPageCount = item.grade.machine.pageCount;
  const oldTerms = item.grade.machine.notContains;
  const [newTerm, expectedNewCount] = Object.entries(item.grade.machine.contains)[0];
  const oldTerm = oldTerms[0];
  const targetPageNumber = item.grade.visual.allowedDiff[0].page;
  const output = evidence.output;
  const source = evidence.source;
  const renderPages = evidence.visual.pages || [];
  const nonTargetPages = renderPages.filter((page) => page.page !== targetPageNumber);
  const targetPage = renderPages.find((page) => page.page === targetPageNumber);
  const residue = oldTerms.every((term) => (
    output.termCounts?.[term] === 0
    && output.rawTermCounts?.[term] === 0
    && output.decodedStreamTermCounts?.[term] === 0
    && output.metadataTermCounts?.[term] === 0
  ));
  const auditStatus = String(audit?.status || "");
  return [
    check("pdf-machine:source-target-unique", "machine", source.termCounts?.[oldTerm] === 1 && source.pages?.[targetPageNumber - 1]?.termCounts?.[oldTerm] === 1, { expected: { term: oldTerm, count: 1, page: targetPageNumber }, actual: source.termCounts?.[oldTerm] }),
    check("pdf-machine:page-count", "machine", output.pageCount === expectedPageCount, { expected: expectedPageCount, actual: output.pageCount }),
    check("pdf-machine:replacement-count", "machine", output.termCounts?.[newTerm] === expectedNewCount && output.pages?.[targetPageNumber - 1]?.termCounts?.[newTerm] === expectedNewCount, { expected: { term: newTerm, count: expectedNewCount, page: targetPageNumber }, actual: output.termCounts?.[newTerm] }),
    check("pdf-machine:old-text-absent", "machine", output.termCounts?.[oldTerm] === 0, { expected: 0, actual: output.termCounts?.[oldTerm] }),
    check("pdf-machine:page-boxes-unchanged", "machine", samePageBoxes(source.pages, output.pages)),
    check("pdf-machine:font-geometry-unchanged", "machine", evidence.sourceStyle?.found && evidence.outputStyle?.found && JSON.stringify(evidence.sourceStyle.fonts) === JSON.stringify(evidence.outputStyle.fonts) && JSON.stringify(evidence.sourceStyle.sizes) === JSON.stringify(evidence.outputStyle.sizes) && sameBoundingBox(evidence.sourceStyle.bbox, evidence.outputStyle.bbox), { expected: { fonts: evidence.sourceStyle?.fonts, sizes: evidence.sourceStyle?.sizes, bbox: evidence.sourceStyle?.bbox }, actual: { fonts: evidence.outputStyle?.fonts, sizes: evidence.outputStyle?.sizes, bbox: evidence.outputStyle?.bbox } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === expectedPageCount && evidence.visual.outputPageCount === expectedPageCount && renderPages.length === expectedPageCount && renderPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: renderPages.length }),
    check("pdf-visual:non-target-pages-identical", "visual", nonTargetPages.length === expectedPageCount - 1 && nonTargetPages.every((page) => page.changedPixelsBBox === null), { actual: nonTargetPages.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    check("pdf-visual:target-diff-contained", "visual", Boolean(targetPage?.changedPixelsBBox) && targetPage?.changedWithinAllowedMask === true, { allowedMask: evidence.visual.allowedMask, actual: targetPage?.changedPixelsBBox || null }),
    gate("pdf-security:no-old-term-residue", "security", residue, { term: oldTerm, evidence: { text: output.termCounts?.[oldTerm], raw: output.rawTermCounts?.[oldTerm], decodedStreams: output.decodedStreamTermCounts?.[oldTerm], metadata: output.metadataTermCounts?.[oldTerm] } }),
    gate("pdf-security:all-streams-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0, { actual: { source: source.decodedStreamErrors || [], output: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-revision", "security", output.startxrefCount === 1 && output.eofCount === 1, { expected: { startxref: 1, eof: 1 }, actual: { startxref: output.startxrefCount, eof: output.eofCount } }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...boundedTraceChecks(audit, commands),
  ];
}

export function gradeOverflowRefusalEvidence({ evidence, audit, commands, finalMessage, item }) {
  const geometry = evidence.geometry || {};
  const auditGeometry = audit?.geometry_evidence || audit?.geometryEvidence || {};
  const operation = String(auditOperation(audit));
  const diagnostics = item.grade.machine.diagnosticTerms || [];
  const diagnosticText = `${finalMessage}\n${audit?.reason || ""}`.toLowerCase();
  const auditWidth = Number(auditGeometry.replacement_text_width_pt ?? auditGeometry.replacementWidth);
  const auditAvailable = Number(auditGeometry.available_width_from_text_origin_pt ?? auditGeometry.availableWidth);
  const auditFits = auditGeometry.fits_outer_cell_without_wrap ?? auditGeometry.fitsObservedTextArea ?? auditGeometry.fits;
  return [
    check("pdf-machine:source-target-unique", "machine", evidence.source.termCounts?.Approved === 1 && evidence.source.pages?.[0]?.termCounts?.Approved === 1, { expected: 1, actual: evidence.source.termCounts?.Approved }),
    check("pdf-machine:independent-overflow-proof", "machine", sameNumber(geometry.targetRectangle?.width, 70) && geometry.fits === false && geometry.replacementWidth > geometry.availableWidth, { actual: geometry }),
    check("pdf-machine:audit-overflow-proof", "machine", Number.isFinite(auditWidth) && Number.isFinite(auditAvailable) && auditWidth > auditAvailable && auditFits === false, { actual: { replacementWidth: auditWidth, availableWidth: auditAvailable, fits: auditFits } }),
    check("pdf-machine:safe-refusal-status", "machine", audit?.status === "failed_closed" && audit?.delivered_modified_pdf !== true, { actual: audit?.status || "unreported" }),
    check("pdf-machine:diagnostic", "machine", diagnostics.some((term) => diagnosticText.includes(String(term).toLowerCase())), { expectedAny: diagnostics }),
    gate("pdf-security:no-partial-output", "security", audit?.validation?.no_partial_modified_pdf_in_outputs === true || audit?.noPartialOutput === true, { actual: audit?.validation?.no_partial_modified_pdf_in_outputs ?? audit?.noPartialOutput ?? "unreported" }),
    gate("pdf-security:no-mutation-claimed", "security", /(?:no (?:pdf )?mutation|not performed|未(?:执行|修改)|没有修改)/i.test(operation), { actual: operation || "unreported" }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source.sha256, { expected: evidence.source.sha256, actual: auditSourceHash(audit) || "unreported" }),
    ...overflowTraceChecks(audit, commands),
  ];
}

function evidenceTermCount(evidence, term) {
  return Number(evidence?.termCounts?.[term] || 0)
    + Number(evidence?.rawTermCounts?.[term] || 0)
    + Number(evidence?.decodedStreamTermCounts?.[term] || 0)
    + Number(evidence?.metadataTermCounts?.[term] || 0);
}

function structureTermCount(structure, term) {
  return Number(structure?.attachmentTermCounts?.[term] || 0)
    + Number(structure?.structureTermCounts?.[term] || 0);
}

export function gradeActiveContentSanitizeEvidence({ evidence, audit, commands, item }) {
  const source = evidence.source;
  const output = evidence.output;
  const sourceStructure = evidence.sourceStructure;
  const outputStructure = evidence.outputStructure;
  const forbiddenNames = item.grade.machine.forbiddenNames;
  const forbiddenActions = item.grade.machine.forbiddenActionTypes;
  const residueTerms = item.grade.machine.residueTerms;
  const sourceNamesPresent = forbiddenNames.every((name) => (
    Number(sourceStructure.structuralNameCounts?.[name] || 0)
    + Number(sourceStructure.actionTypeCounts?.[name] || 0)
  ) > 0);
  const outputNamesAbsent = forbiddenNames.every((name) => (
    Number(outputStructure.structuralNameCounts?.[name] || 0)
    + Number(outputStructure.actionTypeCounts?.[name] || 0)
  ) === 0);
  const sourceActionsPresent = forbiddenActions.every((name) => Number(sourceStructure.actionTypeCounts?.[name] || 0) > 0);
  const outputActionsAbsent = forbiddenActions.every((name) => Number(outputStructure.actionTypeCounts?.[name] || 0) === 0);
  const sourceTermsPresent = residueTerms.every((term) => evidenceTermCount(source, term) + structureTermCount(sourceStructure, term) > 0);
  const outputTermsAbsent = residueTerms.every((term) => evidenceTermCount(output, term) + structureTermCount(outputStructure, term) === 0);
  const visualPages = evidence.visual.pages || [];
  const auditStatus = String(audit?.status || "");
  const validationText = JSON.stringify(audit?.validation || {});
  return [
    check("pdf-machine:source-risk-fixture-complete", "machine", sourceNamesPresent && sourceActionsPresent && sourceStructure.attachments.length > 0 && sourceStructure.commentAnnotations.length > 0 && sourceStructure.populatedWidgets.length > 0 && Object.keys(sourceStructure.personalMetadata).length > 0 && sourceTermsPresent, { actual: { names: sourceStructure.structuralNameCounts, actions: sourceStructure.actionTypeCounts, attachments: sourceStructure.attachments.length, comments: sourceStructure.commentAnnotations.length, populatedWidgets: sourceStructure.populatedWidgets.length, personalMetadata: sourceStructure.personalMetadata } }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", output.pageCount === source.pageCount && samePageBoxes(source.pages, output.pages), { actual: { source: source.pageCount, output: output.pageCount } }),
    check("pdf-machine:attachments-removed", "machine", outputStructure.attachments.length === item.grade.machine.attachments, { expected: item.grade.machine.attachments, actual: outputStructure.attachments.length }),
    check("pdf-machine:comments-removed", "machine", outputStructure.commentAnnotations.length === item.grade.machine.commentAnnotations, { expected: item.grade.machine.commentAnnotations, actual: outputStructure.commentAnnotations.length }),
    check("pdf-machine:form-values-cleared", "machine", outputStructure.populatedWidgets.length === 0, { actual: outputStructure.populatedWidgets }),
    check("pdf-machine:personal-metadata-absent", "machine", Object.keys(outputStructure.personalMetadata).length === 0, { actual: outputStructure.personalMetadata }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === source.pageCount && evidence.visual.outputPageCount === source.pageCount && visualPages.length === source.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: visualPages.length }),
    check("pdf-visual:ordinary-content-stable", "visual", visualPages.every((page) => page.changedOnlyWithinAllowedMasks), { allowedMasks: evidence.visual.allowedMasks, actual: visualPages.map(({ page, changedOutsideAllowedMasksBBox }) => ({ page, changedOutsideAllowedMasksBBox })) }),
    check("pdf-visual:active-appearance-removed", "visual", visualPages.some((page) => page.changedPixelsBBox !== null), { actual: visualPages.map(({ page, changedPixelsBBox }) => ({ page, changedPixelsBBox })) }),
    gate("pdf-security:active-names-absent", "security", outputNamesAbsent && outputActionsAbsent, { actual: { names: outputStructure.structuralNameCounts, actions: outputStructure.actionTypeCounts } }),
    gate("pdf-security:all-risk-canaries-absent", "security", outputTermsAbsent, { terms: residueTerms }),
    gate("pdf-security:all-streams-decodable", "security", source.decodedStreamErrors?.length === 0 && output.decodedStreamErrors?.length === 0, { actual: { source: source.decodedStreamErrors || [], output: output.decodedStreamErrors || [] } }),
    gate("pdf-security:single-revision", "security", output.startxrefCount === 1 && output.eofCount === 1, { expected: { startxref: 1, eof: 1 }, actual: { startxref: output.startxrefCount, eof: output.eofCount } }),
    gate("pdf-security:original-prefix-absent", "security", evidence.originalPrefixPreserved === false, { actual: evidence.originalPrefixPreserved }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === source.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: source.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    check("pdf-security:audit-residue-and-render-evidence", "security", /residue/i.test(validationText) && /render|visual/i.test(validationText), { actual: audit?.validation || "unreported" }),
    ...activeContentTraceChecks(audit, commands),
  ];
}

export function summarizeCaseScore(checks, grade, weights = defaultWeights, hardGatesPassed = true) {
  const categories = ["machine", "visual", "security", "trace"];
  const categoryScores = {};
  let availableWeight = 0;
  let earnedWeight = 0;
  for (const category of categories) {
    const applicable = !(category === "visual" && grade?.visual?.notApplicable === true);
    const categoryChecks = checks.filter((entry) => entry.category === category);
    const passed = applicable ? categoryChecks.length > 0 && categoryChecks.every((entry) => entry.passed) : true;
    const weight = Number(weights?.[category] ?? defaultWeights[category]);
    if (applicable) {
      availableWeight += weight;
      if (passed) earnedWeight += weight;
    }
    categoryScores[category] = { applicable, weight, passed, checks: categoryChecks.length };
  }
  const rawScorePercent = availableWeight ? Math.round(earnedWeight / availableWeight * 10_000) / 100 : 0;
  return {
    categoryScores,
    rawScorePercent,
    scorePercent: hardGatesPassed ? rawScorePercent : 0,
    caseSpecificPassed: categories.every((category) => !categoryScores[category].applicable || categoryScores[category].passed),
  };
}

function oraclePython() {
  return process.env.OPEN_OFFICE_AGENT_EVAL_PYTHON || process.env.OPEN_OFFICE_PDF_PROVIDER_PYTHON || "python3";
}

function invokeOracle(payload, needsPoppler) {
  const python = oraclePython();
  const dependencyProbe = spawnSync(python, ["-c", "import PIL, pdfplumber, pypdf, reportlab"], { encoding: "utf8", env: process.env });
  if (dependencyProbe.status !== 0) {
    return { infrastructureError: `PDF grader requires Pillow, pdfplumber, pypdf, and ReportLab in ${python}: ${String(dependencyProbe.stderr || dependencyProbe.error?.message || "probe failed").trim()}` };
  }
  const poppler = process.env.OPEN_OFFICE_AGENT_EVAL_PDFTOPPM || "pdftoppm";
  if (needsPoppler) {
    const popplerProbe = spawnSync(poppler, ["-v"], { encoding: "utf8", env: process.env });
    if (popplerProbe.status !== 0) return { infrastructureError: `PDF visual grader requires pdftoppm: ${String(popplerProbe.stderr || popplerProbe.error?.message || "probe failed").trim()}` };
  }
  const result = spawnSync(python, [oracleScript], {
    encoding: "utf8",
    input: JSON.stringify({ ...payload, poppler }),
    env: process.env,
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.status !== 0) return { oracleError: `PDF oracle rejected the candidate artifact (${result.status}): ${String(result.stderr || result.stdout || result.error?.message || "unknown error").trim()}` };
  try { return { evidence: JSON.parse(result.stdout) }; } catch (error) { return { infrastructureError: `PDF oracle returned invalid JSON: ${error.message}` }; }
}

async function readAudit(workspace) {
  try { return JSON.parse(await fs.readFile(path.join(workspace, "outputs", "audit.json"), "utf8")); } catch { return null; }
}

function missingArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...boundedTraceChecks(audit, commands),
  ];
}

function unreadableArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...boundedTraceChecks(audit, commands),
  ];
}

function missingActiveContentArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...activeContentTraceChecks(audit, commands),
  ];
}

function unreadableActiveContentArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...activeContentTraceChecks(audit, commands),
  ];
}

export async function gradePdfCase({ item, workspace, evaluator, finalMessage, trace, weights = defaultWeights }) {
  if (!supportedCases.has(item.id)) return { supported: false };
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  let oracle;
  let checks;
  if (item.id === "pdf-bounded-contract-id-replace") {
    const output = path.join(workspace, "outputs", "contract-updated.pdf");
    try { await fs.access(output); } catch {
      checks = missingArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "bounded-replace",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      old: item.grade.machine.notContains[0],
      new: Object.keys(item.grade.machine.contains)[0],
      targetPage: item.grade.visual.allowedDiff[0].page,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeBoundedReplaceEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-overflow-replace-refusal") {
    oracle = invokeOracle({
      kind: "overflow-refusal",
      source: path.join(workspace, "inputs", "source.pdf"),
      old: "Approved",
      new: "Approved subject to twelve additional contractual conditions",
    }, false);
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeOverflowRefusalEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, item });
  } else {
    const output = path.join(workspace, "outputs", "public-safe.pdf");
    try { await fs.access(output); } catch {
      checks = missingActiveContentArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "active-content-sanitize",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      terms: item.grade.machine.residueTerms,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableActiveContentArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeActiveContentSanitizeEvidence({ evidence: oracle.evidence, audit, commands, item });
  }
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence: oracle.evidence, pending: [], ...score };
}

export { supportedCases as pdfGradedCaseIds };
