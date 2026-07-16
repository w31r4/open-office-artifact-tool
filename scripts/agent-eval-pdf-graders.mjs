import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const oracleScript = path.join(scriptDirectory, "agent-eval-pdf-oracle.py");
const supportedCases = new Set([
  "pdf-bounded-contract-id-replace",
  "pdf-overflow-replace-refusal",
  "pdf-acroform-visible-preserved",
  "pdf-attachment-quarantine-inventory",
  "pdf-active-content-public-sanitize",
  "pdf-greenfield-accessible-report",
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

function completedInvocation(commands, pattern) {
  for (let commandIndex = 0; commandIndex < commands.length; commandIndex += 1) {
    const command = String(commands[commandIndex]);
    const expression = new RegExp(pattern.source, pattern.flags.replace("g", "") + "g");
    for (const match of command.matchAll(expression)) {
      const tail = command.slice((match.index || 0) + match[0].length);
      if (/^\s+(?:--help|-h)\b/i.test(tail)) continue;
      return { commandIndex, offset: match.index || 0 };
    }
  }
  return null;
}

function invocationBefore(left, right) {
  return Boolean(left && right && (
    left.commandIndex < right.commandIndex
    || left.commandIndex === right.commandIndex && left.offset < right.offset
  ));
}

function commandTextAfter(commands, position) {
  if (!position) return "";
  return [String(commands[position.commandIndex]).slice(position.offset), ...commands.slice(position.commandIndex + 1)].join("\n");
}

function acroformTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const inspect = completedInvocation(commands, /pypdf_edit\.py\s+inspect\b/i);
  const checkProvider = completedInvocation(commands, /pdf_provider\.py\s+check\b[^\n]*--provider\s+pypdf\b/i);
  const plan = completedInvocation(commands, /pdf_provider\.py\s+plan\b/i);
  const fill = completedInvocation(commands, /pypdf_edit\.py\s+fill-form\b/i);
  const afterFill = commandTextAfter(commands, fill);
  const bypassPatterns = [
    /\bPdfWriter\s*\(/,
    /\bupdate_page_form_field_values\s*\(/,
    /\bclone_document_from_reader\s*\(/,
  ];
  const positions = Object.fromEntries(Object.entries({ inspect, checkProvider, plan, fill }).map(([name, value]) => [name, value?.commandIndex ?? -1]));
  return [
    check("pdf-trace:provider", "trace", /pypdf/i.test(String(auditProvider(audit))), { expected: "pypdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^incremental$/i.test(String(auditSaveStrategy(audit))), { expected: "incremental", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:inspect-check-plan-before-mutation", "trace", invocationBefore(inspect, fill) && invocationBefore(checkProvider, fill) && invocationBefore(plan, fill), { actual: positions }),
    check("pdf-trace:typed-fill-form-primitive", "trace", Boolean(fill) && /fill[-_ ]?form/i.test(operation), { actual: operation || "unreported" }),
    check("pdf-trace:post-mutation-poppler-render", "trace", /\bpdftoppm\b/i.test(afterFill), { actual: { fillObserved: Boolean(fill), postMutationRenderObserved: /\bpdftoppm\b/i.test(afterFill) } }),
    check("pdf-trace:audit-byte-validation", "trace", /pdf_audit\.py\s+validate\b/i.test(afterFill), { actual: { postMutationAuditValidationObserved: /pdf_audit\.py\s+validate\b/i.test(afterFill) } }),
    gate("pdf-trace:no-ad-hoc-pypdf-writer", "trace", !bypassPatterns.some((pattern) => pattern.test(commandText)), { forbidden: bypassPatterns.map(String) }),
  ];
}

function attachmentTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  // Shell-safe invocations commonly quote a variable-expanded script path,
  // for example `"$SCRIPTS/pypdf_edit.py" inspect`. Accept that closing quote
  // without weakening the requirement that the shipped typed primitive ran.
  const inspect = completedInvocation(commands, /pypdf_edit\.py["']?\s+inspect\b/i);
  const checkProvider = completedInvocation(commands, /pdf_provider\.py["']?\s+check\b[^\n]*--provider\s+pypdf\b/i);
  const plan = completedInvocation(commands, /pdf_provider\.py["']?\s+plan\b[^\n]*--task\s+extract-attachments\b/i);
  const extract = completedInvocation(commands, /pypdf_edit\.py["']?\s+extract-attachments\b/i);
  const afterExtract = commandTextAfter(commands, extract);
  const auditValidatedAfterExtraction = /pdf_audit\.py["']?\s+validate\b/i.test(afterExtract);
  const manualPatterns = [
    /\bPdfReader\s*\(/,
    /\.attachment_list\b/,
    /\.attachments\b/,
    /\[\s*["']\/EmbeddedFiles["']\s*\]/,
    /\bget_data\s*\(/,
  ];
  const payloadOpenPatterns = [
    /\b(?:cat|less|more|head|tail|strings|unzip|7z|tar)\b[^\n]*outputs\/quarantine/i,
    // Codex command traces are normally wrapped in `/bin/zsh -lc`; that
    // launcher alone is not evidence that a quarantined payload was executed.
    // Match only an actual payload interpreter command at the beginning of a
    // command or after a shell separator.
    /(?:^|[;&|]\s*)(?:"?\$PYTHON_BIN"?|(?:\/[\w./-]+\/)?(?:python\d*|node))\s+[^\n]*outputs\/quarantine\//im,
    /(?:^|[;&|]\s*)(?:\/[\w./-]+\/)?(?:bash|sh|zsh)\s+(?!-lc\b)[^\n]*outputs\/quarantine\//im,
    /\bchmod\b[^\n]*outputs\/quarantine\//i,
  ];
  const positions = Object.fromEntries(Object.entries({ inspect, checkProvider, plan, extract }).map(([name, value]) => [name, value?.commandIndex ?? -1]));
  return [
    check("pdf-trace:provider", "trace", /pypdf/i.test(String(auditProvider(audit))), { expected: "pypdf", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^read-only$/i.test(String(auditSaveStrategy(audit))), { expected: "read-only", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:preflight-audit", "trace", audit?.preflight?.probeCompleted === true && audit?.preflight?.planCompleted === true, { actual: audit?.preflight || "unreported" }),
    check("pdf-trace:inspect-check-plan-before-extraction", "trace", invocationBefore(inspect, extract) && invocationBefore(checkProvider, extract) && invocationBefore(plan, extract), { actual: positions }),
    check("pdf-trace:typed-attachment-primitive", "trace", Boolean(extract) && /extract[-_ ]?attachments/i.test(auditOperation(audit)), { actual: auditOperation(audit) || "unreported" }),
    check("pdf-trace:audit-byte-validation", "trace", auditValidatedAfterExtraction, { actual: { extractObserved: Boolean(extract), postExtractionAuditValidationObserved: auditValidatedAfterExtraction } }),
    gate("pdf-trace:no-ad-hoc-pypdf-extraction", "trace", !manualPatterns.some((pattern) => pattern.test(commandText)), { forbidden: manualPatterns.map(String) }),
    gate("pdf-trace:no-payload-open-or-execution", "trace", !payloadOpenPatterns.some((pattern) => pattern.test(commandText)), { forbidden: payloadOpenPatterns.map(String) }),
  ];
}

function accessibleReportTraceChecks(audit, commands) {
  const commandText = commands.join("\n");
  const operation = auditOperation(audit);
  const manualWriterPatterns = [
    /%PDF-\d/i,
    /\bPdfWriter\s*\(/,
    /\breportlab\b/i,
    /\bpymupdf\b|\bfitz\b/i,
  ];
  return [
    check("pdf-trace:provider", "trace", /^artifact-tool$/i.test(String(auditProvider(audit))), { expected: "artifact-tool", actual: auditProvider(audit) }),
    check("pdf-trace:provider-version", "trace", Boolean(String(auditProviderVersion(audit)).trim()), { actual: auditProviderVersion(audit) || "unreported" }),
    check("pdf-trace:save-policy", "trace", /^rewrite$/i.test(String(auditSaveStrategy(audit))), { expected: "rewrite", actual: auditSaveStrategy(audit) }),
    gate("pdf-trace:no-silent-fallback", "trace", auditFallback(audit) === true, { expected: false, actual: auditFallback(audit) === null ? "unreported" : !auditFallback(audit) }),
    check("pdf-trace:typed-greenfield-primitive", "trace", /accessible-board-report\.mjs/i.test(commandText) && /create[-_ ]?accessible[-_ ]?report/i.test(operation), { expected: "shipped accessible-board-report.mjs + create-accessible-report", actual: { operation, exampleObserved: /accessible-board-report\.mjs/i.test(commandText) } }),
    check("pdf-trace:poppler-review", "trace", audit?.validation?.poppler?.status === "passed" && audit?.validation?.poppler?.pages?.length === 6, { actual: audit?.validation?.poppler || "unreported" }),
    gate("pdf-trace:no-ad-hoc-pdf-writer", "trace", !manualWriterPatterns.some((pattern) => pattern.test(commandText)), { forbidden: manualWriterPatterns.map(String) }),
  ];
}

function pdfUaOverclaim(text) {
  return /PDF\/?UA\s+(?:certified|compliant|conformant|validated|passed)\b|(?<!不)(?:已通过|符合)\s*PDF\/?UA|PDF\/?UA\s*(?:已认证|已合规|认证通过)/i.test(String(text || ""));
}

export function gradeAccessibleReportEvidence({ evidence, audit, commands, finalMessage, item }) {
  const structure = evidence.structure || {};
  const output = evidence.output || {};
  const visualPages = evidence.visual?.pages || [];
  const table = structure.tables?.find((candidate) => candidate.pages?.length >= 2);
  const validation = audit?.validation || {};
  const headingLevels = item.grade.machine.headingLevels || [1, 2, 3];
  const claimText = `${JSON.stringify(audit || {})}\n${finalMessage || ""}`;
  const link = structure.links?.[0];
  const pageText = structure.pageText || [];
  return [
    check("pdf-machine:page-count", "machine", output.pageCount === item.grade.machine.pageCount, { expected: item.grade.machine.pageCount, actual: output.pageCount }),
    check("pdf-machine:catalog-title-language-tagging", "machine", structure.tagged === true && structure.language === item.grade.machine.language && Boolean(String(structure.title || "").trim()), { expected: { tagged: true, language: item.grade.machine.language, title: "non-empty" }, actual: { tagged: structure.tagged, language: structure.language, title: structure.title } }),
    check("pdf-machine:h1-h3-structure", "machine", headingLevels.every((level) => Number(structure.roles?.[`H${level}`] || 0) > 0), { expected: headingLevels, actual: structure.roles }),
    check("pdf-machine:cross-page-semantic-table", "machine", Boolean(table && table.pages.length >= 2 && table.rows >= 2 && table.headers >= 1 && table.dataCells >= 1), { actual: structure.tables || [] }),
    check("pdf-machine:figure-alt", "machine", Number(structure.roles?.Figure || 0) >= 1 && structure.figuresWithAlt === Number(structure.roles?.Figure || 0), { actual: { figures: structure.roles?.Figure || 0, withAlt: structure.figuresWithAlt } }),
    check("pdf-machine:meaningful-tagged-link", "machine", Number(structure.roles?.Link || 0) >= 1 && structure.linkObjrAssociations >= 1 && /^https?:\/\//i.test(String(link?.uri || "")) && Number.isInteger(link?.structParent), { actual: { roles: structure.roles?.Link || 0, objr: structure.linkObjrAssociations, links: structure.links } }),
    check("pdf-machine:reading-order-ids", "machine", structure.rootIds?.length >= 12 && new Set(structure.rootIds).size === structure.rootIds.length && structure.rootIds.includes("risk-register"), { actual: structure.rootIds || [] }),
    check("pdf-machine:running-artifacts", "machine", structure.artifactMarkers >= item.grade.machine.pageCount * 2, { expectedAtLeast: item.grade.machine.pageCount * 2, actual: structure.artifactMarkers }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), { actual: audit?.status || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual?.pageCount === item.grade.machine.pageCount && visualPages.length === item.grade.machine.pageCount && visualPages.every((page) => page.nonBlank && page.bytes > 1_000), { renderer: evidence.visual?.renderer, pages: visualPages }),
    check("pdf-visual:no-edge-clipping", "visual", visualPages.length > 0 && visualPages.every((page) => page.touchesEdge === false), { actual: visualPages.map(({ page, inkBBox, touchesEdge }) => ({ page, inkBBox, touchesEdge })) }),
    check("pdf-visual:consistent-page-geometry", "visual", visualPages.length > 0 && visualPages.every((page) => page.width === visualPages[0].width && page.height === visualPages[0].height), { actual: visualPages.map(({ page, width, height }) => ({ page, width, height })) }),
    check("pdf-visual:cross-page-table-readable", "visual", Boolean(table?.pages?.every((pageNumber) => /风险/.test(String(pageText[pageNumber - 1] || "")))), { tablePages: table?.pages || [], pageText: table?.pages?.map((pageNumber) => String(pageText[pageNumber - 1] || "").slice(0, 240)) || [] }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source?.sha256 && auditOutputHash(audit) === output.sha256, { expected: { source: evidence.source?.sha256, output: output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:no-pdfua-overclaim", "security", !pdfUaOverclaim(claimText), { actual: pdfUaOverclaim(claimText) ? "positive PDF/UA certification claim detected" : "no overclaim detected" }),
    check("pdf-security:modeled-scope-explicit", "security", validation.modeledVerify?.status === "passed" && /modeled|PdfArtifact/i.test(String(validation.modeledVerify?.scope || "")), { actual: validation.modeledVerify || "unreported" }),
    check("pdf-security:verapdf-machine-layer-separate", "security", Boolean(validation.veraPdfMachine) && ["not-run", "completed", "completed-with-findings", "probe-failed"].includes(validation.veraPdfMachine.status) && /machine|not.*certification|No veraPDF/i.test(JSON.stringify(validation.veraPdfMachine)), { actual: validation.veraPdfMachine || "unreported" }),
    check("pdf-security:human-pdfua-required", "security", validation.humanPdfUa?.status === "required" && /No complete PDF\/UA certification|不.*完整.*PDF\/UA|人工/i.test(JSON.stringify(validation.humanPdfUa)), { actual: validation.humanPdfUa || "unreported" }),
    ...accessibleReportTraceChecks(audit, commands),
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

export function gradeAcroFormEvidence({ evidence, audit, commands, item }) {
  const expectedFields = item.grade.machine.fields;
  const sourceFields = evidence.sourceForm.fields || {};
  const outputFields = evidence.outputForm.fields || {};
  const sourceWidgets = evidence.sourceForm.widgets || [];
  const outputWidgets = evidence.outputForm.widgets || [];
  const visualPages = evidence.visual.pages || [];
  const widgetChanges = evidence.visual.widgetChanges || [];
  const expectedOutputValue = (name, value) => outputFields[name]?.fieldType === "/Btn" && value
    ? `/${String(value).replace(/^\//, "")}`
    : String(value);
  const sourceFixtureComplete = Object.keys(expectedFields).every((name) => (
    sourceFields[name]
    && sourceFields[name].value === ""
    && sourceFields[name].readOnly === false
  ))
    && sourceFields.company_type?.states?.includes("/LLC")
    && sourceFields.company_type?.states?.includes("/Corporation")
    && sourceFields.terms_ack?.fieldType === "/Btn"
    && sourceFields.terms_ack?.value === "/Yes"
    && sourceFields.terms_ack?.readOnly === false;
  const valuesExact = Object.entries(expectedFields).every(([name, value]) => (
    outputFields[name]?.value === expectedOutputValue(name, value)
  ));
  const topologyPreserved = sourceWidgets.length === outputWidgets.length && sourceWidgets.every((source, index) => {
    const output = outputWidgets[index];
    return output
      && source.page === output.page
      && source.name === output.name
      && source.fieldType === output.fieldType
      && sameBoundingBox(source.rect, output.rect)
      && output.readOnly === false;
  });
  const appearancesPresent = outputWidgets.length > 0 && outputWidgets.every((widget) => widget.appearancePresent === true);
  const companyWidgets = outputWidgets.filter((widget) => widget.name === "company_type");
  const radioStateCorrect = companyWidgets.length === 2
    && companyWidgets.filter((widget) => widget.selectedState === "/LLC").length === 1
    && companyWidgets.filter((widget) => widget.selectedState === "/Off").length === 1;
  const sourceCheckbox = sourceWidgets.find((widget) => widget.name === "terms_ack");
  const outputCheckbox = outputWidgets.find((widget) => widget.name === "terms_ack");
  const checkboxPreserved = outputFields.terms_ack?.value === sourceFields.terms_ack?.value
    && sourceCheckbox?.selectedState === "/Yes"
    && outputCheckbox?.selectedState === sourceCheckbox.selectedState;
  const changedWidgets = widgetChanges.filter((widget) => widget.expectedChange);
  const untouchedWidgets = widgetChanges.filter((widget) => !widget.expectedChange);
  const expectedChangedWidgetCount = Object.entries(expectedFields).filter(([, value]) => String(value).length > 0).length;
  const expectedUntouchedWidgetCount = sourceWidgets.length - expectedChangedWidgetCount;
  const sensitiveFieldsBlank = ["tin", "signature"].every((name) => (
    outputFields[name]?.value === ""
    && outputFields[name]?.defaultValue === ""
    && widgetChanges.filter((widget) => widget.name === name).every((widget) => widget.changedPixelsBBox === null)
  ));
  const auditStatus = String(audit?.status || "");
  return [
    check("pdf-machine:source-form-fixture-complete", "machine", sourceFixtureComplete, { actual: sourceFields }),
    check("pdf-machine:page-count-and-boxes-unchanged", "machine", evidence.output.pageCount === evidence.source.pageCount && samePageBoxes(evidence.source.pages, evidence.output.pages), { actual: { source: evidence.source.pageCount, output: evidence.output.pageCount } }),
    check("pdf-machine:field-values-exact", "machine", valuesExact, { expected: expectedFields, actual: Object.fromEntries(Object.entries(outputFields).map(([name, field]) => [name, field.value])) }),
    check("pdf-machine:widget-topology-preserved", "machine", topologyPreserved, { actual: { sourceWidgets: sourceWidgets.length, outputWidgets: outputWidgets.length } }),
    check("pdf-machine:appearances-present", "machine", appearancesPresent && evidence.outputForm.needAppearances === false, { actual: { appearancesPresent, needAppearances: evidence.outputForm.needAppearances } }),
    check("pdf-machine:radio-appearance-state", "machine", radioStateCorrect, { actual: companyWidgets.map(({ appearanceStates, selectedState }) => ({ appearanceStates, selectedState })) }),
    check("pdf-machine:checkbox-state-preserved", "machine", checkboxPreserved, { actual: { source: sourceCheckbox?.selectedState, output: outputCheckbox?.selectedState, fieldValue: outputFields.terms_ack?.value } }),
    check("pdf-machine:interactive-form-preserved", "machine", evidence.outputForm.acroFormPresent === true && evidence.outputForm.fieldTreeRoots === evidence.sourceForm.fieldTreeRoots && evidence.outputForm.fieldTreeRoots > 0, { actual: { acroFormPresent: evidence.outputForm.acroFormPresent, fieldTreeRoots: evidence.outputForm.fieldTreeRoots } }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:all-pages-rendered", "visual", evidence.visual.sourcePageCount === evidence.source.pageCount && evidence.visual.outputPageCount === evidence.source.pageCount && visualPages.length === evidence.source.pageCount && visualPages.every((page) => page.sameDimensions && page.nonBlank), { renderer: evidence.visual.renderer, pages: visualPages.length }),
    check("pdf-visual:changes-contained-to-target-widgets", "visual", visualPages.every((page) => page.changedOnlyWithinAllowedMasks), { actual: visualPages.map(({ page, changedOutsideAllowedMasksBBox }) => ({ page, changedOutsideAllowedMasksBBox })) }),
    check("pdf-visual:filled-values-visible", "visual", changedWidgets.length === expectedChangedWidgetCount && changedWidgets.every((widget) => widget.changedPixelsBBox !== null && widget.changedInteriorPixelsBBox !== null), { expected: expectedChangedWidgetCount, actual: changedWidgets }),
    check("pdf-visual:untouched-widgets-stable", "visual", untouchedWidgets.length === expectedUntouchedWidgetCount && untouchedWidgets.every((widget) => widget.changedPixelsBBox === null), { expected: expectedUntouchedWidgetCount, actual: untouchedWidgets }),
    gate("pdf-security:incremental-prefix-preserved", "security", evidence.originalPrefixPreserved === true, { actual: evidence.originalPrefixPreserved }),
    gate("pdf-security:single-appended-revision", "security", evidence.output.startxrefCount === evidence.source.startxrefCount + 1 && evidence.output.eofCount === evidence.source.eofCount + 1, { expected: { startxref: evidence.source.startxrefCount + 1, eof: evidence.source.eofCount + 1 }, actual: { startxref: evidence.output.startxrefCount, eof: evidence.output.eofCount } }),
    gate("pdf-security:sensitive-fields-blank", "security", sensitiveFieldsBlank, { actual: { tin: outputFields.tin, signature: outputFields.signature } }),
    gate("pdf-security:all-streams-decodable", "security", evidence.source.decodedStreamErrors?.length === 0 && evidence.output.decodedStreamErrors?.length === 0, { actual: { source: evidence.source.decodedStreamErrors || [], output: evidence.output.decodedStreamErrors || [] } }),
    gate("pdf-security:audit-provenance", "security", auditSourceHash(audit) === evidence.source.sha256 && auditOutputHash(audit) === evidence.output.sha256, { expected: { source: evidence.source.sha256, output: evidence.output.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    ...acroformTraceChecks(audit, commands),
  ];
}

export function gradeAttachmentQuarantineEvidence({ evidence, audit, commands, item }) {
  const expected = evidence.expectedAttachments || [];
  const manifest = evidence.manifest || {};
  const manifestAttachments = manifest.attachments || [];
  const files = evidence.quarantine?.files || [];
  const expectedCount = item.grade.machine.attachmentCount;
  const expectedHashes = expected.map((entry) => entry.sha256).sort();
  const manifestHashes = manifestAttachments.map((entry) => entry.sha256).sort();
  const fileHashes = files.map((entry) => entry.sha256).sort();
  const relativeManifestPaths = manifestAttachments.map((entry) => String(entry.savedPath || "").replaceAll("\\", "/"));
  const manifestSavedNames = manifestAttachments.map((entry) => String(entry.savedName || ""));
  const filePaths = files.map((entry) => entry.path);
  const sourceFixtureComplete = expected.length === expectedCount
    && expected.filter((entry) => entry.scope === "document").length === item.grade.machine.documentAttachments
    && expected.filter((entry) => entry.scope === "page").length === item.grade.machine.pageAttachments
    && expected.filter((entry) => entry.displayName === "report.txt").length === 3
    && evidence.unsafeRawPaths?.some((entry) => entry.displayName === "../escape.exe");
  const manifestFieldsExact = expected.every((entry) => manifestAttachments.some((actual) => (
    actual.scope === entry.scope
    && actual.page === entry.page
    && actual.annotationIndex === entry.annotationIndex
    && actual.internalKey === entry.internalKey
    && actual.displayName === entry.displayName
    && actual.mime === entry.mime
    && actual.bytes === entry.bytes
    && actual.sha256 === entry.sha256
  )));
  const pathsUnique = new Set(relativeManifestPaths.map((value) => value.toLowerCase())).size === expectedCount
    && new Set(manifestSavedNames.map((value) => value.toLowerCase())).size === expectedCount
    && new Set(filePaths.map((value) => value.toLowerCase())).size === expectedCount;
  const pathsContained = relativeManifestPaths.every((value) => (
    value.startsWith("quarantine/")
    && !value.split("/").includes("..")
    && !path.isAbsolute(value)
  ))
    && files.every((entry) => entry.flat === true && !entry.path.split("/").includes(".."))
    && manifestAttachments.some((entry) => entry.displayName === "../escape.exe" && entry.savedName === "escape.exe" && entry.nameSanitized === true);
  const auditStatus = String(audit?.status || "");
  const validation = audit?.validation || {};
  return [
    check("pdf-machine:source-attachment-fixture-complete", "machine", sourceFixtureComplete, { actual: { count: expected.length, scopes: expected.map((entry) => entry.scope), unsafeRawPaths: evidence.unsafeRawPaths } }),
    check("pdf-machine:manifest-schema-and-count", "machine", manifest.schema === "open-office-artifact-tool.pdf-attachments.v1" && manifestAttachments.length === expectedCount, { expected: expectedCount, actual: { schema: manifest.schema, count: manifestAttachments.length } }),
    check("pdf-machine:manifest-fields-exact", "machine", manifestFieldsExact, { expected, actual: manifestAttachments }),
    check("pdf-machine:attachment-bytes-and-hashes", "machine", JSON.stringify(expectedHashes) === JSON.stringify(manifestHashes) && JSON.stringify(expectedHashes) === JSON.stringify(fileHashes), { expected: expectedHashes, manifest: manifestHashes, files: fileHashes }),
    check("pdf-machine:unique-output-paths", "machine", pathsUnique, { actual: relativeManifestPaths }),
    check("pdf-machine:audit-success", "machine", /^(?:success|succeeded|completed)$/i.test(auditStatus), { actual: auditStatus || "unreported" }),
    check("pdf-visual:source-remains-original-pdf", "visual", manifest.source?.sha256 === evidence.source.sha256 && manifest.validation?.sourceUnchanged === true && evidence.source.pageCount === 1, { actual: { manifestSource: manifest.source?.sha256, oracleSource: evidence.source.sha256, sourceUnchanged: manifest.validation?.sourceUnchanged } }),
    check("pdf-visual:no-derived-pdf-or-page-transform", "visual", files.length === expectedCount && files.every((entry) => !entry.path.toLowerCase().endsWith(".pdf")), { actual: filePaths }),
    gate("pdf-security:path-traversal-contained", "security", pathsContained, { actual: { manifestPaths: relativeManifestPaths, files: filePaths } }),
    gate("pdf-security:regular-quarantine-files-only", "security", evidence.quarantine?.invalid?.length === 0 && files.length === expectedCount, { actual: evidence.quarantine }),
    gate("pdf-security:no-missing-or-extra-payloads", "security", JSON.stringify(expectedHashes) === JSON.stringify(fileHashes), { expected: expectedHashes, actual: fileHashes }),
    gate("pdf-security:source-and-manifest-provenance", "security", auditSourceHash(audit) === evidence.source.sha256 && auditOutputHash(audit) === evidence.manifestFile.sha256, { expected: { source: evidence.source.sha256, output: evidence.manifestFile.sha256 }, actual: { source: auditSourceHash(audit) || "unreported", output: auditOutputHash(audit) || "unreported" } }),
    gate("pdf-security:nothing-opened-or-executed", "security", manifest.validation?.attachmentsOpenedOrExecuted === false && validation.attachmentsOpenedOrExecuted === false, { actual: { manifest: manifest.validation?.attachmentsOpenedOrExecuted, audit: validation.attachmentsOpenedOrExecuted } }),
    check("pdf-security:audit-validation-evidence", "security", validation.sourceUnchanged === true && (validation.allHashesVerified === true || validation.allAttachmentHashesVerified === true) && validation.allPathsContained === true && validation.duplicateNamesSeparated === true, { actual: validation }),
    ...attachmentTraceChecks(audit, commands),
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

function missingAcroFormArtifactChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...acroformTraceChecks(audit, commands),
  ];
}

function unreadableAcroFormArtifactChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...acroformTraceChecks(audit, commands),
  ];
}

function missingAttachmentQuarantineChecks(audit, commands) {
  return [
    check("pdf-machine:attachment-manifest-and-quarantine-available", "machine", false),
    check("pdf-visual:source-read-only-evidence-available", "visual", false),
    gate("pdf-security:attachment-manifest-and-quarantine-available", "security", false),
    ...attachmentTraceChecks(audit, commands),
  ];
}

function unreadableAttachmentQuarantineChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:attachment-evidence-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:source-read-only-evidence-readable", "visual", false, { actual: oracleError }),
    gate("pdf-security:attachment-evidence-readable-by-oracle", "security", false, { actual: oracleError }),
    ...attachmentTraceChecks(audit, commands),
  ];
}

function missingAccessibleReportChecks(audit, commands) {
  return [
    check("pdf-machine:artifact-available-for-oracle", "machine", false),
    check("pdf-visual:artifact-available-for-oracle", "visual", false),
    gate("pdf-security:artifact-available-for-oracle", "security", false),
    ...accessibleReportTraceChecks(audit, commands),
  ];
}

function unreadableAccessibleReportChecks(audit, commands, oracleError) {
  return [
    check("pdf-machine:artifact-readable-by-oracle", "machine", false, { actual: oracleError }),
    check("pdf-visual:artifact-renderable-by-oracle", "visual", false, { actual: oracleError }),
    gate("pdf-security:artifact-readable-by-oracle", "security", false, { actual: oracleError }),
    ...accessibleReportTraceChecks(audit, commands),
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
  } else if (item.id === "pdf-acroform-visible-preserved") {
    const output = path.join(workspace, "outputs", "form-filled.pdf");
    try { await fs.access(output); } catch {
      checks = missingAcroFormArtifactChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "acroform-visible",
      source: path.join(workspace, "inputs", "source.pdf"),
      output,
      fields: item.grade.machine.fields,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAcroFormArtifactChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAcroFormEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-attachment-quarantine-inventory") {
    const manifest = path.join(workspace, "outputs", "attachments.json");
    const quarantine = path.join(workspace, "outputs", "quarantine");
    try {
      await fs.access(manifest);
      if (!(await fs.stat(quarantine)).isDirectory()) throw new Error("quarantine is not a directory");
    } catch {
      checks = missingAttachmentQuarantineChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "attachment-quarantine",
      source: path.join(workspace, "inputs", "source.pdf"),
      manifest,
      quarantine,
    }, false);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAttachmentQuarantineChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAttachmentQuarantineEvidence({ evidence: oracle.evidence, audit, commands, item });
  } else if (item.id === "pdf-greenfield-accessible-report") {
    const output = path.join(workspace, "outputs", "readiness-report.pdf");
    try { await fs.access(output); } catch {
      checks = missingAccessibleReportChecks(audit, commands);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: null, pending: [], ...score };
    }
    oracle = invokeOracle({
      kind: "accessible-report",
      source: path.join(workspace, "inputs", "report-data.json"),
      output,
      renderRoot: path.join(evaluator, "pdf-oracle-render"),
    }, true);
    if (!oracle.evidence && oracle.oracleError) {
      checks = unreadableAccessibleReportChecks(audit, commands, oracle.oracleError);
      const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
      return { supported: true, graded: true, checks, evidence: { oracleError: oracle.oracleError }, pending: [], ...score };
    }
    if (!oracle.evidence) return { supported: true, graded: false, checks: [], pending: ["PDF case grader infrastructure"], infrastructureErrors: [oracle.infrastructureError] };
    checks = gradeAccessibleReportEvidence({ evidence: oracle.evidence, audit, commands, finalMessage, item });
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
