import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  PPTX_SLIDE_NAME_FIXTURE,
  PPTX_TITLE_NOTES_FIXTURE,
} from "./agent-eval-office-fixtures.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

export const pptxGradedCaseIds = new Set([
  "pptx-title-and-notes-edit",
  "pptx-source-bound-slide-name-edit",
]);

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const SHIPPED_TITLE_NOTES_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/presentations|node_modules\/open-office-artifact-tool\/skills\/presentations\/skills\/presentations)\/examples\/openchestnut-title-notes-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_SLIDE_NAME_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/presentations|node_modules\/open-office-artifact-tool\/skills\/presentations\/skills\/presentations)\/examples\/openchestnut-slide-name-edit-workflow\.mjs(?:$|[\s"'`])/i;

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
  const attributes = {};
  for (const match of String(opening).matchAll(/([:\w.-]+)="([^"]*)"/g)) {
    attributes[match[1].split(":").at(-1)] = decodeXml(match[2]);
  }
  return attributes;
}

function drawingTexts(xml = "") {
  return [...String(xml).matchAll(/<(?:[\w.-]+:)?t\b[^>]*>([\s\S]*?)<\/(?:[\w.-]+:)?t>/g)]
    .map((match) => decodeXml(match[1].replace(/<[^>]+>/g, "")));
}

function shapeTextByName(xml = "", name) {
  for (const shape of String(xml).matchAll(/<(?:[\w.-]+:)?sp\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?sp>/g)) {
    const properties = /<(?:[\w.-]+:)?cNvPr\b[^>]*>/.exec(shape[0])?.[0] || "";
    if (xmlAttributes(properties).name === name) return drawingTexts(shape[0]).join("\n");
  }
  return null;
}

function slideName(xml = "") {
  const opening = /<(?:[\w.-]+:)?cSld\b[^>]*>/.exec(String(xml))?.[0] || "";
  return xmlAttributes(opening).name || null;
}

function directBackground(xml = "") {
  const background = /<(?:[\w.-]+:)?bg\b[^>]*>[\s\S]*?<\/(?:[\w.-]+:)?bg>/.exec(String(xml))?.[0] || "";
  const color = /<(?:[\w.-]+:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(background)?.[1];
  return color ? "#" + color.toUpperCase() : null;
}

function numericPptxOrder(left, right) {
  return Number(/\d+/.exec(left)?.[0]) - Number(/\d+/.exec(right)?.[0]);
}

async function partHashes(zip, paths) {
  const hashes = {};
  for (const filePath of paths) hashes[filePath] = sha256(await zip.file(filePath).async("uint8array"));
  return hashes;
}

export async function inspectTitleNotesPptx(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const slidePaths = paths.filter((name) => /^ppt\/slides\/slide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const notesPaths = paths.filter((name) => /^ppt\/notesSlides\/notesSlide\d+\.xml$/i.test(name)).sort(numericPptxOrder);
  const slides = [];
  for (const slidePath of slidePaths) {
    const xml = await zip.file(slidePath).async("text");
    slides.push({
      path: slidePath,
      name: slideName(xml),
      texts: drawingTexts(xml),
      title: shapeTextByName(xml, PPTX_TITLE_NOTES_FIXTURE.titleShapeName),
      background: directBackground(xml),
    });
  }
  const notes = {};
  for (const notesPath of notesPaths) notes[notesPath] = drawingTexts(await zip.file(notesPath).async("text")).join("\n");
  const target = slides.find((slide) => slide.name === PPTX_TITLE_NOTES_FIXTURE.targetSlideName) || null;
  const untouched = slides.find((slide) => slide.name === PPTX_TITLE_NOTES_FIXTURE.untouchedSlideName) || null;
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await partHashes(zip, paths),
    slidePaths,
    notesPaths,
    slides,
    target,
    untouched,
    targetNotesPath: "ppt/notesSlides/notesSlide1.xml",
    targetNotes: notes["ppt/notesSlides/notesSlide1.xml"] || null,
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
  const policy = audit?.savePolicy || audit?.save_strategy;
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

function sameArray(left, right) {
  return left.length === right.length && left.every((value, index) => value === right[index]);
}

function packageChanges(source, output) {
  const paths = [...new Set([...source.paths, ...output.paths])].sort();
  return paths.filter((filePath) => source.partHashes[filePath] !== output.partHashes[filePath]);
}

function visualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount === 2;
  const targetChanged = pageCountsMatch && source.pages?.[0]?.pixelSha256 !== output.pages?.[0]?.pixelSha256;
  const untouchedStable = pageCountsMatch
    && source.pages?.[1]?.width === output.pages?.[1]?.width
    && source.pages?.[1]?.height === output.pages?.[1]?.height
    && source.pages?.[1]?.pixelSha256 === output.pages?.[1]?.pixelSha256;
  return { available, rendered, pageCountsMatch, targetChanged, untouchedStable };
}

function stableVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount === 2;
  const pageStable = (index) => pageCountsMatch
    && source.pages?.[index]?.width === output.pages?.[index]?.width
    && source.pages?.[index]?.height === output.pages?.[index]?.height
    && source.pages?.[index]?.pixelSha256 === output.pages?.[index]?.pixelSha256;
  return {
    available,
    rendered,
    pageCountsMatch,
    targetStable: pageStable(0),
    untouchedStable: pageStable(1),
  };
}

function usedTypedPptxRoundTrip(commandText) {
  const directPublicApi = /PresentationFile\.importPptx/i.test(commandText)
    && /PresentationFile\.exportPptx/i.test(commandText);
  return directPublicApi || SHIPPED_TITLE_NOTES_WORKFLOW.test(commandText);
}

function usedTypedSlideNameRoundTrip(commandText) {
  const directPublicApi = /PresentationFile\.importPptx/i.test(commandText)
    && /PresentationFile\.exportPptx/i.test(commandText);
  return directPublicApi || SHIPPED_SLIDE_NAME_WORKFLOW.test(commandText);
}

export function gradePptxTitleNotesEvidence({ evidence, audit, commands }) {
  const fixture = PPTX_TITLE_NOTES_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const visual = visualEvidence(evidence.visual?.source, evidence.visual?.output);
  const changedPaths = packageChanges(source, output);
  const expectedChangedPaths = [source.target?.path, source.targetNotesPath].filter(Boolean).sort();
  const sourceSlideNames = source.slides.map((slide) => slide.name);
  const outputSlideNames = output.slides.map((slide) => slide.name);
  const commandText = commands.join("\n");
  const sourceText = source.target?.texts || [];
  const outputText = output.target?.texts || [];
  return [
    check("pptx-machine:canonical-fixture", "machine", source.slides.length === 2
      && source.target?.title === fixture.originalTitle
      && source.target?.background === fixture.targetBackground
      && source.targetNotes === fixture.originalNotes
      && source.untouched?.background === fixture.untouchedBackground
      && sourceText.includes(fixture.supportingText), {
      sourceTarget: source.target,
      sourceNotes: source.targetNotes,
      sourceSlides: sourceSlideNames,
    }),
    check("pptx-machine:title-and-notes-edited", "machine", output.target?.title === fixture.replacementTitle
      && output.targetNotes === fixture.replacementNotes
      && !outputText.includes(fixture.originalTitle)
      && !String(output.targetNotes || "").includes(fixture.originalNotes), {
      outputTarget: output.target,
      outputNotes: output.targetNotes,
    }),
    check("pptx-machine:target-and-untouched-structure-preserved", "machine", sameArray(sourceSlideNames, outputSlideNames)
      && output.target?.background === fixture.targetBackground
      && outputText.includes(fixture.supportingText)
      && output.untouched?.background === fixture.untouchedBackground
      && source.untouched?.path === output.untouched?.path, {
      sourceSlides: sourceSlideNames,
      outputSlides: outputSlideNames,
      sourceTarget: source.target,
      outputTarget: output.target,
      sourceUntouched: source.untouched,
      outputUntouched: output.untouched,
    }),
    check("pptx-machine:only-bounded-parts-changed", "machine", sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
    }),
    check("pptx-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("pptx-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("pptx-visual:target-change-and-untouched-slide-stable", "visual", visual.targetChanged && visual.untouchedStable, {
      visual: evidence.visual,
    }),
    gate("pptx-security:fixed-topology-and-package-preservation", "security", sameArray(source.paths, output.paths)
      && source.target?.path === output.target?.path
      && source.targetNotesPath === output.targetNotesPath
      && source.partHashes[source.untouched?.path] === output.partHashes[output.untouched?.path]
      && sameArray(changedPaths, expectedChangedPaths), {
      sourcePaths: source.paths,
      outputPaths: output.paths,
      changedPaths,
      targetPath: { source: source.target?.path, output: output.target?.path },
      notesPath: { source: source.targetNotesPath, output: output.targetNotesPath },
      untouchedPath: { source: source.untouched?.path, output: output.untouched?.path },
    }),
    gate("pptx-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("pptx-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("pptx-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("pptx-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("pptx-trace:title-notes-operation", "trace", /title.*notes|notes.*title|speaker/i.test(auditOperation(audit)), {
      operation: auditOperation(audit),
    }),
    check("pptx-trace:typed-roundtrip", "trace", usedTypedPptxRoundTrip(commandText), {
      expected: "public PresentationFile importPptx/exportPptx calls or the integrity-protected published title/notes workflow",
    }),
    check("pptx-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

/**
 * Grade the narrow non-visual mutation separately from title/notes. An Open
 * XML SDK save may canonicalize the target SlidePart, so this oracle checks
 * the semantic p:cSld name and requires every *other* part to remain byte
 * identical. Native page pixels must consequently stay stable on both slides.
 */
export function gradePptxSlideNameEvidence({ evidence, audit, commands }) {
  const fixture = PPTX_SLIDE_NAME_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const visual = stableVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const sourceTargets = source.slides.filter((slide) => slide.name === fixture.expectedName);
  const sourceTarget = sourceTargets.length === 1 ? sourceTargets[0] : null;
  const outputTarget = sourceTarget
    ? output.slides.find((slide) => slide.path === sourceTarget.path) || null
    : null;
  const sourceUntouched = source.slides.find((slide) => slide.name === fixture.untouchedSlideName) || null;
  const outputUntouched = sourceUntouched
    ? output.slides.find((slide) => slide.path === sourceUntouched.path) || null
    : null;
  const sourceSlideNames = source.slides.map((slide) => slide.name);
  const expectedSlideNames = sourceSlideNames.map((name) => name === fixture.expectedName ? fixture.replacementName : name);
  const outputSlideNames = output.slides.map((slide) => slide.name);
  const changedPaths = packageChanges(source, output);
  const expectedChangedPaths = sourceTarget ? [sourceTarget.path] : [];
  const commandText = commands.join("\n");
  const operation = audit?.operation && typeof audit.operation === "object" ? audit.operation : {};
  return [
    check("pptx-name-machine:canonical-fixture", "machine", sourceTargets.length === 1
      && source.slides.length === 2
      && sourceTarget?.title === PPTX_TITLE_NOTES_FIXTURE.originalTitle
      && sourceTarget?.background === PPTX_TITLE_NOTES_FIXTURE.targetBackground
      && source.targetNotes === PPTX_TITLE_NOTES_FIXTURE.originalNotes
      && sourceUntouched?.background === PPTX_TITLE_NOTES_FIXTURE.untouchedBackground
      && sourceTarget?.texts.includes(PPTX_TITLE_NOTES_FIXTURE.supportingText), {
      sourceTargets,
      sourceNotes: source.targetNotes,
      sourceSlideNames,
    }),
    check("pptx-name-machine:native-name-edited", "machine", outputTarget?.name === fixture.replacementName, {
      sourceTarget,
      outputTarget,
    }),
    check("pptx-name-machine:semantic-content-and-order-preserved", "machine", sameArray(outputSlideNames, expectedSlideNames)
      && sourceTarget?.path === outputTarget?.path
      && sourceTarget?.title === outputTarget?.title
      && sameArray(sourceTarget?.texts || [], outputTarget?.texts || [])
      && sourceTarget?.background === outputTarget?.background
      && source.targetNotes === output.targetNotes
      && sourceUntouched?.path === outputUntouched?.path
      && sourceUntouched?.title === outputUntouched?.title
      && sameArray(sourceUntouched?.texts || [], outputUntouched?.texts || [])
      && sourceUntouched?.background === outputUntouched?.background, {
      sourceSlideNames,
      expectedSlideNames,
      outputSlideNames,
      sourceTarget,
      outputTarget,
      sourceUntouched,
      outputUntouched,
    }),
    check("pptx-name-machine:only-target-slide-part-changed", "machine", sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
    }),
    check("pptx-name-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("pptx-name-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("pptx-name-visual:all-pages-pixel-stable", "visual", visual.targetStable && visual.untouchedStable, {
      visual: evidence.visual,
    }),
    gate("pptx-name-security:fixed-topology-and-non-target-byte-preservation", "security", sameArray(source.paths, output.paths)
      && sourceTarget?.path === outputTarget?.path
      && sourceUntouched?.path === outputUntouched?.path
      && source.partHashes[sourceUntouched?.path] === output.partHashes[outputUntouched?.path]
      && sameArray(changedPaths, expectedChangedPaths), {
      sourcePaths: source.paths,
      outputPaths: output.paths,
      changedPaths,
      targetPath: { source: sourceTarget?.path, output: outputTarget?.path },
      untouchedPath: { source: sourceUntouched?.path, output: outputUntouched?.path },
    }),
    gate("pptx-name-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("pptx-name-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("pptx-name-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("pptx-name-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), {
      strategy: auditStrategy(audit),
    }),
    check("pptx-name-trace:source-bound-name-operation", "trace", /slide.*name|name.*slide/i.test(auditOperation(audit))
      && operation.sourcePart === sourceTarget?.path
      && operation.expectedName === fixture.expectedName
      && operation.replacementName === fixture.replacementName
      && operation.nativeAttribute === "p:cSld/@name", {
      operation: audit?.operation || null,
      expected: {
        sourcePart: sourceTarget?.path,
        expectedName: fixture.expectedName,
        replacementName: fixture.replacementName,
        nativeAttribute: "p:cSld/@name",
      },
    }),
    check("pptx-name-trace:typed-roundtrip", "trace", usedTypedSlideNameRoundTrip(commandText), {
      expected: "public PresentationFile importPptx/exportPptx calls or the integrity-protected published slide-name workflow",
    }),
    check("pptx-name-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
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

export async function gradePptxCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  if (!pptxGradedCaseIds.has(item.id)) return { supported: false };
  const isSlideNameCase = item.id === "pptx-source-bound-slide-name-edit";
  const fixture = isSlideNameCase ? PPTX_SLIDE_NAME_FIXTURE : PPTX_TITLE_NOTES_FIXTURE;
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const sourcePath = path.join(workspace, "inputs", fixture.presentationName);
  const outputPath = path.join(workspace, "outputs", isSlideNameCase ? "launch-review-renamed.pptx" : "launch-review-updated.pptx");
  let source;
  let output;
  try {
    [source, output] = await Promise.all([
      inspectTitleNotesPptx(sourcePath),
      inspectTitleNotesPptx(outputPath),
    ]);
  } catch (error) {
    const checks = [
      gate("pptx-machine:readable-output", "machine", false, { error: error.message }),
      gate("pptx-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }

  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(sourcePath, "pptx-source"),
    renderOfficeFile(outputPath, "pptx-output"),
  ]);
  const visualUnavailable = [sourceRender, outputRender].find((result) => !result.available);
  if (visualUnavailable) {
    return {
      supported: true,
      graded: false,
      checks: [],
      evidence: { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage },
      pending: ["native LibreOffice/Poppler presentation rendering"],
      infrastructureErrors: [visualUnavailable.reason],
    };
  }
  const evidence = { source, output, visual: { source: sourceRender, output: outputRender }, finalMessage };
  const checks = isSlideNameCase
    ? gradePptxSlideNameEvidence({ evidence, audit, commands, item })
    : gradePptxTitleNotesEvidence({ evidence, audit, commands, item });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}
