import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import { DOCX_CLASSIC_COMMENT_FIXTURE, DOCX_HEADER_TEXT_FIXTURE } from "./agent-eval-office-fixtures.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

export const docxGradedCaseIds = new Set(["docx-classic-comment-text-edit", "docx-header-text-edit"]);

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const SHIPPED_CLASSIC_COMMENT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/open-office-artifact-tool\/skills\/documents\/skills\/documents)\/examples\/openchestnut-classic-comment-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_HEADER_TEXT_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/documents|node_modules\/open-office-artifact-tool\/skills\/documents\/skills\/documents)\/examples\/openchestnut-header-text-edit-workflow\.mjs(?:$|[\s"'`])/i;

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
    result[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return result;
}

function wordText(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")))
    .join("");
}

function escapeRegex(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXmlText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function tagIds(xml = "", name) {
  const pattern = new RegExp("<(?:[\\w.-]+:)?"+ name +"\\b[^>]*>", "gi");
  return [...String(xml).matchAll(pattern)]
    .map((match) => xmlAttributes(match[0]).id)
    .filter((id) => id !== undefined)
    .map(String);
}

function parseClassicComments(xml = "") {
  const comments = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?comment\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?comment>/g)) {
    const opening = /^<(?:[\w.-]+:)?comment\b[^>]*>/.exec(match[0])?.[0] || "";
    const attributes = xmlAttributes(opening);
    comments.push({
      id: String(attributes.id || ""),
      author: attributes.author || "",
      initials: attributes.initials || "",
      date: attributes.date || "",
      text: wordText(match[1]),
    });
  }
  return comments;
}

function parseParagraphs(xml = "") {
  const paragraphs = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?p\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?p>/g)) {
    const body = match[1];
    paragraphs.push({
      text: wordText(body),
      commentRangeStarts: tagIds(body, "commentRangeStart"),
      commentRangeEnds: tagIds(body, "commentRangeEnd"),
      commentReferences: tagIds(body, "commentReference"),
    });
  }
  return paragraphs;
}

export async function inspectClassicCommentDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const commentsPath = paths.find((name) => name.toLowerCase() === "word/comments.xml");
  const documentPath = paths.find((name) => name.toLowerCase() === "word/document.xml");
  const commentsXml = commentsPath ? await zip.file(commentsPath)?.async("text") || "" : "";
  const documentXml = documentPath ? await zip.file(documentPath)?.async("text") || "" : "";
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    commentsPath: commentsPath || null,
    documentPath: documentPath || null,
    commentsExtendedPaths: paths.filter((name) => /^word\/commentsExtended\.xml$/i.test(name)),
    peoplePaths: paths.filter((name) => /^word\/people\.xml$/i.test(name)),
    comments: parseClassicComments(commentsXml),
    paragraphs: parseParagraphs(documentXml),
  };
}

function normalizeTargetText(xml, expectedText) {
  const pattern = new RegExp(`(<w:t(?:\\s[^>]*)?>)${escapeRegex(escapeXmlText(expectedText))}(</w:t>)`, "g");
  let matches = 0;
  const normalized = xml.replace(pattern, (_whole, open, close) => {
    matches += 1;
    return `${open}__OPEN_CHESTNUT_TARGET_TEXT__${close}`;
  });
  return { normalized, matches };
}

function headerFooterPartRecord(partPath, xml) {
  return {
    path: partPath,
    xml,
    paragraphs: parseParagraphs(xml).map((paragraph) => paragraph.text),
  };
}

/**
 * Independent package-level reader for the source-bound header-edit case.
 * It intentionally does not use DocumentFile so a wrapper cannot satisfy the
 * black-box evaluator by merely repeating its own model projection.
 */
export async function inspectHeaderTextDocx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const entries = await Promise.all(paths.map(async (partPath) => [partPath, Buffer.from(await zip.file(partPath)?.async("uint8array") || [])]));
  const partHashes = Object.fromEntries(entries.map(([partPath, value]) => [partPath, sha256(value)]));
  const content = new Map(entries.map(([partPath, value]) => [partPath, value.toString("utf8")]));
  const documentPath = paths.find((name) => name.toLowerCase() === "word/document.xml") || null;
  const headerParts = paths.filter((name) => /^word\/header\d+\.xml$/i.test(name))
    .map((partPath) => headerFooterPartRecord(partPath, content.get(partPath) || ""));
  const footerParts = paths.filter((name) => /^word\/footer\d+\.xml$/i.test(name))
    .map((partPath) => headerFooterPartRecord(partPath, content.get(partPath) || ""));
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes,
    documentPath,
    bodyParagraphs: parseParagraphs(documentPath ? content.get(documentPath) || "" : "").map((paragraph) => paragraph.text),
    headerParts,
    footerParts,
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
  const values = [
    provider.silentFallback,
    provider.silent_fallback,
    provider.fallbackUsed,
    provider.fallback_used,
    audit?.silentFallback,
    audit?.silent_fallback,
    audit?.fallbackUsed,
    audit?.fallback_used,
  ].filter((value) => value !== undefined);
  return values.length > 0 && values.every((value) => value === false || value === "false");
}

function auditStrategy(audit) {
  const policy = audit?.savePolicy || audit?.save_policy || audit?.saveStrategy || audit?.save_strategy;
  return String(typeof policy === "string" ? policy : policy?.strategy || policy?.selected || audit?.strategy || "");
}

function auditOperation(audit) {
  const operation = audit?.operation;
  return String(typeof operation === "string" ? operation : operation?.type || operation?.name || operation?.operation || "");
}

function auditHash(audit, side) {
  const record = audit?.[side] || {};
  return String(record.sha256 || audit?.[side + "Sha256"] || audit?.[side + "_sha256"] || "");
}

function sameCommentMetadata(left, right) {
  return left && right
    && left.id === right.id
    && left.author === right.author
    && left.initials === right.initials
    && left.date === right.date;
}

function anchoredParagraph(document, commentId, anchorText) {
  const matches = document.paragraphs.filter((paragraph) => paragraph.text === anchorText);
  if (matches.length !== 1) return null;
  const paragraph = matches[0];
  const id = String(commentId);
  return {
    paragraph,
    allMarkersPresent: paragraph.commentRangeStarts.includes(id)
      && paragraph.commentRangeEnds.includes(id)
      && paragraph.commentReferences.includes(id),
  };
}

function paragraphsPreserved(source, output) {
  return JSON.stringify(source.paragraphs.map((paragraph) => paragraph.text))
    === JSON.stringify(output.paragraphs.map((paragraph) => paragraph.text));
}

function visualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount;
  const pixelsStable = pageCountsMatch
    && source?.pages?.length === output?.pages?.length
    && source?.pages?.every((page, index) => {
      const next = output.pages[index];
      return next
        && page.width === next.width
        && page.height === next.height
        && page.pixelSha256 === next.pixelSha256;
    });
  return { available, rendered, pageCountsMatch, pixelsStable };
}

function usedTypedDocxRoundTrip(commandText) {
  const directPublicApi = /(?:DocumentFile\.)?importDocx/i.test(commandText)
    && /(?:DocumentFile\.)?exportDocx/i.test(commandText);
  return directPublicApi || SHIPPED_CLASSIC_COMMENT_WORKFLOW.test(commandText);
}

export function gradeDocxClassicCommentEvidence({ evidence, audit, commands }) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceComment = source.comments[0];
  const outputComment = output.comments[0];
  const sourceAnchor = anchoredParagraph(source, sourceComment?.id, fixture.anchorText);
  const outputAnchor = anchoredParagraph(output, outputComment?.id, fixture.anchorText);
  const visual = visualEvidence(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const classicTopology = source.commentsPath === "word/comments.xml"
    && output.commentsPath === "word/comments.xml"
    && source.comments.length === 1
    && output.comments.length === 1
    && source.commentsExtendedPaths.length === 0
    && output.commentsExtendedPaths.length === 0
    && source.peoplePaths.length === 0
    && output.peoplePaths.length === 0;
  return [
    check("docx-machine:fixture-classic-comment", "machine", Boolean(sourceComment)
      && sourceComment.author === fixture.comment.author
      && sourceComment.initials === fixture.comment.initials
      && sourceComment.date === fixture.comment.date
      && sourceComment.text === fixture.comment.originalText, { sourceComment }),
    check("docx-machine:comment-text-edited", "machine", sameCommentMetadata(sourceComment, outputComment)
      && outputComment?.text === fixture.comment.replacementText, { sourceComment, outputComment }),
    check("docx-machine:comment-anchor-preserved", "machine", sourceAnchor?.allMarkersPresent === true
      && outputAnchor?.allMarkersPresent === true
      && sourceComment?.id === outputComment?.id, {
      sourceAnchor,
      outputAnchor,
      sourceCommentId: sourceComment?.id,
      outputCommentId: outputComment?.id,
    }),
    check("docx-machine:visible-document-text-preserved", "machine", paragraphsPreserved(source, output), {
      source: source.paragraphs.map((paragraph) => paragraph.text),
      output: output.paragraphs.map((paragraph) => paragraph.text),
    }),
    check("docx-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("docx-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("docx-visual:comment-only-body-stable", "visual", visual.pixelsStable, {
      visual: evidence.visual,
      note: "Headless LibreOffice does not render review balloons reliably; classic-comment structure is checked independently.",
    }),
    gate("docx-security:classic-topology-and-modern-graph-absence", "security", classicTopology, {
      commentsPath: { source: source.commentsPath, output: output.commentsPath },
      commentsExtendedPaths: { source: source.commentsExtendedPaths, output: output.commentsExtendedPaths },
      peoplePaths: { source: source.peoplePaths, output: output.peoplePaths },
      counts: { source: source.comments.length, output: output.comments.length },
    }),
    gate("docx-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("docx-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("docx-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("docx-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("docx-trace:classic-comment-operation", "trace", /classic|comment/i.test(auditOperation(audit)), {
      operation: auditOperation(audit),
    }),
    check("docx-trace:typed-roundtrip", "trace", usedTypedDocxRoundTrip(commandText), {
      expected: "public DocumentFile importDocx/exportDocx calls or the integrity-protected published classic-comment workflow",
    }),
    check("docx-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

function headerPartWithText(document, text) {
  const matches = document.headerParts.filter((part) => part.paragraphs.filter((paragraph) => paragraph === text).length === 1);
  return matches.length === 1 ? matches[0] : null;
}

function changedPackageParts(source, output) {
  if (JSON.stringify(source.paths) !== JSON.stringify(output.paths)) return null;
  return source.paths.filter((partPath) => source.partHashes[partPath] !== output.partHashes[partPath]);
}

function normalizedHeaderResidual(sourcePart, outputPart) {
  if (!sourcePart || !outputPart || sourcePart.path !== outputPart.path) return { ok: false, source: null, output: null };
  const source = normalizeTargetText(sourcePart.xml, DOCX_HEADER_TEXT_FIXTURE.header.originalText);
  const output = normalizeTargetText(outputPart.xml, DOCX_HEADER_TEXT_FIXTURE.header.replacementText);
  return {
    ok: source.matches === 1 && output.matches === 1 && source.normalized === output.normalized,
    source,
    output,
  };
}

function headerVisibleChange(source, output) {
  const visual = visualEvidence(source, output);
  const pixelsChanged = visual.pageCountsMatch
    && source?.pages?.length === output?.pages?.length
    && source.pages.some((page, index) => page.pixelSha256 !== output.pages[index]?.pixelSha256);
  return { ...visual, pixelsChanged };
}

function usedTypedHeaderRoundTrip(commandText) {
  const directPublicApi = /(?:DocumentFile\.)?importDocx/i.test(commandText)
    && /(?:DocumentFile\.)?exportDocx/i.test(commandText);
  return directPublicApi || SHIPPED_HEADER_TEXT_WORKFLOW.test(commandText);
}

export function gradeDocxHeaderTextEvidence({ evidence, audit, commands }) {
  const fixture = DOCX_HEADER_TEXT_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceTarget = headerPartWithText(source, fixture.header.originalText);
  const outputTarget = headerPartWithText(output, fixture.header.replacementText);
  const residual = normalizedHeaderResidual(sourceTarget, outputTarget);
  const changed = changedPackageParts(source, output);
  const companionStable = Boolean(sourceTarget && outputTarget)
    && JSON.stringify(sourceTarget.paragraphs.map((paragraph) => paragraph === fixture.header.originalText ? "__target__" : paragraph))
      === JSON.stringify(outputTarget.paragraphs.map((paragraph) => paragraph === fixture.header.replacementText ? "__target__" : paragraph));
  const footerStable = source.footerParts.length === output.footerParts.length
    && source.footerParts.every((part, index) => part.path === output.footerParts[index]?.path
      && source.partHashes[part.path] === output.partHashes[part.path]);
  const visual = headerVisibleChange(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const auditTarget = audit?.operation?.target || {};
  return [
    check("docx-header-machine:fixture-source-profile", "machine", Boolean(sourceTarget)
      && sourceTarget.paragraphs.includes(fixture.header.companionText)
      && source.footerParts.length === 1
      && source.bodyParagraphs.includes(fixture.title)
      && fixture.body.every((paragraph) => source.bodyParagraphs.includes(paragraph)), {
      sourceTarget,
      footerParts: source.footerParts.map((part) => ({ path: part.path, paragraphs: part.paragraphs })),
      bodyParagraphs: source.bodyParagraphs,
    }),
    check("docx-header-machine:requested-text-edited", "machine", Boolean(sourceTarget && outputTarget)
      && outputTarget.paragraphs.includes(fixture.header.replacementText)
      && !outputTarget.paragraphs.includes(fixture.header.originalText)
      && companionStable, {
      sourceTarget,
      outputTarget,
      companionStable,
    }),
    check("docx-header-machine:body-and-footer-stable", "machine", JSON.stringify(source.bodyParagraphs) === JSON.stringify(output.bodyParagraphs)
      && footerStable, {
      body: { source: source.bodyParagraphs, output: output.bodyParagraphs },
      footerStable,
    }),
    check("docx-header-machine:header-residual-stable", "machine", residual.ok, {
      sourceMatches: residual.source?.matches ?? 0,
      outputMatches: residual.output?.matches ?? 0,
    }),
    check("docx-header-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("docx-header-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("docx-header-visual:header-change-visible", "visual", visual.pixelsChanged, {
      visual: evidence.visual,
      note: "The package gate proves the changed part; native rendering confirms that the requested page-furniture edit remains visible.",
    }),
    gate("docx-header-security:only-target-header-part-changed", "security", Boolean(sourceTarget)
      && JSON.stringify(changed) === JSON.stringify([sourceTarget.path]), {
      changed,
      targetPart: sourceTarget?.path || null,
    }),
    gate("docx-header-security:footer-field-and-package-inventory-preserved", "security", footerStable
      && changed !== null
      && source.footerParts.some((part) => part.paragraphs.includes(fixture.footer.text)), {
      footerStable,
      sourcePaths: source.paths,
      outputPaths: output.paths,
    }),
    gate("docx-header-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("docx-header-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("docx-header-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("docx-header-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("docx-header-trace:source-bound-header-operation", "trace", /header/i.test(auditOperation(audit))
      && auditTarget.sectionIndex === fixture.header.sectionIndex
      && auditTarget.referenceType === fixture.header.referenceType
      && auditTarget.partPath === sourceTarget?.path, {
      operation: audit?.operation || null,
      expected: { sectionIndex: fixture.header.sectionIndex, referenceType: fixture.header.referenceType, partPath: sourceTarget?.path || null },
    }),
    check("docx-header-trace:typed-roundtrip", "trace", usedTypedHeaderRoundTrip(commandText), {
      expected: "public DocumentFile importDocx/exportDocx calls or the integrity-protected published header-text workflow",
    }),
    check("docx-header-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
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

async function gradeDocxClassicCommentCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_CLASSIC_COMMENT_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  const outputPath = path.join(workspace, "outputs", "legal-review-updated.docx");
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectClassicCommentDocx(sourcePath),
      inspectClassicCommentDocx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("docx-machine:readable-output", "machine", false, { error: error.message }),
      gate("docx-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "docx-source"),
    renderOfficeFile(outputPath, "docx-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler document rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }

  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeDocxClassicCommentEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeDocxHeaderTextCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  const fixture = DOCX_HEADER_TEXT_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.documentName);
  const outputPath = path.join(workspace, "outputs", "board-brief-header-reviewed.docx");
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectHeaderTextDocx(sourcePath),
      inspectHeaderTextDocx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("docx-header-machine:readable-output", "machine", false, { error: error.message }),
      gate("docx-header-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "docx-header-source"),
    renderOfficeFile(outputPath, "docx-header-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler document rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }

  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = gradeDocxHeaderTextEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

export async function gradeDocxCase(options) {
  if (options.item.id === "docx-classic-comment-text-edit") return gradeDocxClassicCommentCase(options);
  if (options.item.id === "docx-header-text-edit") return gradeDocxHeaderTextCase(options);
  return { supported: false };
}
