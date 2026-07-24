import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import JSZip from "jszip";

import {
  XLSX_CONNECTION_REFRESH_FIXTURE,
  XLSX_GROWTH_UPDATE_FIXTURE,
  XLSX_PIVOT_REFRESH_FIXTURE,
  XLSX_THREADED_REVIEW_FIXTURE,
} from "./agent-eval-office-fixtures.mjs";
import { renderOfficeFile } from "./agent-eval-office-native-render.mjs";
import { extractCompletedCommands, summarizeCaseScore } from "./agent-eval-pdf-graders.mjs";

export const spreadsheetGradedCaseIds = new Set([
  "xlsx-threaded-reply-resolve",
  "xlsx-growth-assumption-update",
  "xlsx-connection-refresh-on-open",
  "xlsx-pivot-refresh-on-open",
]);

const defaultWeights = { machine: 45, visual: 25, security: 20, trace: 10 };
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/i;
const SHIPPED_THREADED_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/spreadsheets|node_modules\/open-office-artifact-tool\/skills\/spreadsheets\/skills\/spreadsheets)\/examples\/openchestnut-threaded-comment-reply-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_GROWTH_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/spreadsheets|node_modules\/open-office-artifact-tool\/skills\/spreadsheets\/skills\/spreadsheets)\/examples\/openchestnut-growth-assumption-edit-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_CONNECTION_REFRESH_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/spreadsheets|node_modules\/open-office-artifact-tool\/skills\/spreadsheets\/skills\/spreadsheets)\/examples\/openchestnut-connection-refresh-hardening-workflow\.mjs(?:$|[\s"'`])/i;
const SHIPPED_PIVOT_REFRESH_WORKFLOW = /(?:^|[\s"'`])(?:\.?\/)?(?:\.agents\/skills\/spreadsheets|node_modules\/open-office-artifact-tool\/skills\/spreadsheets\/skills\/spreadsheets)\/examples\/openchestnut-pivot-refresh-hardening-workflow\.mjs(?:$|[\s"'`])/i;

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

function normalizeFormula(value) {
  return String(value || "").replace(/^=/, "");
}

function numericValue(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function closeEnough(actual, expected, tolerance = 1e-9) {
  return typeof actual === "number" && Number.isFinite(actual) && Math.abs(actual - expected) <= tolerance;
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

function parseCells(xml = "") {
  const cells = new Map();
  for (const match of String(xml).matchAll(/<[^>]*c\b([^>]*)>([\s\S]*?)<\/[^>]*c>/g)) {
    const attributes = xmlAttributes(match[1]);
    const address = String(attributes.r || "").toUpperCase();
    if (!address) continue;
    const content = match[2];
    const formula = /<[^>]*f\b[^>]*>([\s\S]*?)<\/[^>]*f>/.exec(content)?.[1] || null;
    const rawValue = /<[^>]*v\b[^>]*>([\s\S]*?)<\/[^>]*v>/.exec(content)?.[1] || null;
    cells.set(address, {
      address,
      formula: formula === null ? null : decodeXml(formula),
      rawValue: rawValue === null ? null : decodeXml(rawValue),
      value: rawValue === null ? null : numericValue(decodeXml(rawValue)),
      type: attributes.t || null,
    });
  }
  return cells;
}

function relationshipTargets(xml = "") {
  const targets = new Map();
  for (const match of String(xml).matchAll(/<[^>]*Relationship\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    if (attributes.Id && attributes.Target && !attributes.TargetMode) targets.set(attributes.Id, attributes.Target);
  }
  return targets;
}

function workbookSheets(workbookXml, relsXml) {
  const targets = relationshipTargets(relsXml);
  const sheets = [];
  for (const match of String(workbookXml).matchAll(/<[^>]*sheet\b[^>]*\/?\s*>/g)) {
    const attributes = xmlAttributes(match[0]);
    const target = targets.get(attributes.id);
    if (!attributes.name || !target) continue;
    const normalizedTarget = String(target).replace(/^\/+/, "");
    const part = path.posix.normalize(normalizedTarget.startsWith("xl/") ? normalizedTarget : path.posix.join("xl", normalizedTarget));
    sheets.push({ name: attributes.name, id: attributes.sheetId || null, path: part });
  }
  return sheets;
}

async function packagePartHashes(zip, paths) {
  const hashes = {};
  for (const partPath of paths) hashes[partPath] = sha256(await zip.file(partPath).async("uint8array"));
  return hashes;
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

export async function inspectGrowthWorkbook(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const [workbookXml, workbookRelsXml] = await Promise.all([
    zip.file("xl/workbook.xml")?.async("text") || "",
    zip.file("xl/_rels/workbook.xml.rels")?.async("text") || "",
  ]);
  const sheets = workbookSheets(workbookXml, workbookRelsXml);
  const target = sheets.find((sheet) => sheet.name === XLSX_GROWTH_UPDATE_FIXTURE.targetSheetName) || null;
  const canary = sheets.find((sheet) => sheet.name === XLSX_GROWTH_UPDATE_FIXTURE.canarySheetName) || null;
  const [targetXml, canaryXml] = await Promise.all([
    target ? zip.file(target.path)?.async("text") || "" : "",
    canary ? zip.file(canary.path)?.async("text") || "" : "",
  ]);
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await packagePartHashes(zip, paths),
    sheets,
    target: target ? { ...target, cells: parseCells(targetXml) } : null,
    canary: canary ? { ...canary, cells: parseCells(canaryXml) } : null,
  };
}

function parseWorkbookConnections(xml = "") {
  const connections = [];
  for (const match of String(xml).matchAll(/<(?:[\w.-]+:)?connection\b([^>]*)>([\s\S]*?)<\/(?:[\w.-]+:)?connection>/g)) {
    const attributes = xmlAttributes(match[1]);
    const dbPr = /<(?:[\w.-]+:)?dbPr\b([^>]*)\/?\s*>/i.exec(match[2])?.[1] || "";
    connections.push({
      id: Number(attributes.id),
      name: attributes.name || "",
      type: Number(attributes.type),
      refreshedVersion: Number(attributes.refreshedVersion),
      description: attributes.description || "",
      keepAlive: new Set(["1", "true", "on"]).has(String(attributes.keepAlive || "0").toLowerCase()),
      background: new Set(["1", "true", "on"]).has(String(attributes.background || "0").toLowerCase()),
      refreshOnLoad: new Set(["1", "true", "on"]).has(String(attributes.refreshOnLoad || "0").toLowerCase()),
      saveData: new Set(["1", "true", "on"]).has(String(attributes.saveData || "0").toLowerCase()),
      intervalMinutes: Number(attributes.interval),
      command: xmlAttributes(dbPr).command || "",
      opaqueValue: /<(?:[\w.-]+:)?connectionOpaque\b[^>]*\bvalue="([^"]*)"/i.exec(match[2])?.[1] || "",
    });
  }
  return connections;
}

function normalizeConnectionRefreshConnectionPart(xml, connectionId) {
  let targetCount = 0;
  let targetRefreshAttributeCount = 0;
  const normalized = String(xml)
    .replace(/^<\?xml\b[^>]*\?>/i, "")
    .replace(/\s+xmlns(?::[\w.-]+)?="[^"]*"/g, "")
    .replace(/>\s+</g, "><")
    .replace(/\s+\/>/g, "/>")
    .replace(/<((?:[\w.-]+:)?connection)\b([^>]*)>/gi, (match, tag, attributes) => {
      if (Number(xmlAttributes(attributes).id) !== connectionId) return match;
      targetCount += 1;
      const refreshAttribute = /\s+refreshOnLoad="[^"]*"/i;
      if (!refreshAttribute.test(attributes)) return match;
      targetRefreshAttributeCount += 1;
      return `<${tag}${attributes.replace(refreshAttribute, "")} refreshOnLoad="__connection_refresh__">`;
    });
  return { normalized, targetCount, targetRefreshAttributeCount };
}

export async function inspectConnectionRefreshWorkbook(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const connectionPath = "xl/connections.xml";
  const queryTablePaths = paths.filter((name) => /^xl\/queryTables\/queryTable\d+\.xml$/i.test(name));
  const tablePaths = paths.filter((name) => /^xl\/tables\/table\d+\.xml$/i.test(name));
  const connectionXml = await zip.file(connectionPath)?.async("text") || "";
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await packagePartHashes(zip, paths),
    connectionPath,
    connectionXml,
    connections: parseWorkbookConnections(connectionXml),
    queryTablePaths,
    tablePaths,
  };
}

function normalizePivotRefreshCacheDefinition(xml = "") {
  const roots = [...String(xml).matchAll(/<(?:[\w.-]+:)?pivotCacheDefinition\b[^>]*>/gi)];
  const opening = roots.length === 1 ? roots[0][0] : "";
  const refreshAttributes = [...opening.matchAll(/\s+refreshOnLoad="([^"]*)"/gi)];
  const lexicalValue = refreshAttributes.length === 1 ? String(refreshAttributes[0][1]).toLowerCase() : "";
  const validLexicalValue = new Set(["1", "true", "0", "false"]).has(lexicalValue);
  const normalized = roots.length === 1 && refreshAttributes.length === 1
    ? `${String(xml).slice(0, roots[0].index)}${opening.replace(refreshAttributes[0][0], ' refreshOnLoad="__pivot_refresh__"')}${String(xml).slice(roots[0].index + opening.length)}`
    : String(xml);
  return {
    rootCount: roots.length,
    refreshAttributeCount: refreshAttributes.length,
    enabled: validLexicalValue ? lexicalValue === "1" || lexicalValue === "true" : null,
    normalized,
  };
}

function parsePivotTableDefinition(xml = "") {
  const roots = [...String(xml).matchAll(/<(?:[\w.-]+:)?pivotTableDefinition\b[^>]*>/gi)];
  const attributes = roots.length === 1 ? xmlAttributes(roots[0][0]) : {};
  return { rootCount: roots.length, name: String(attributes.name || "") };
}

export async function inspectPivotRefreshWorkbook(filePath) {
  const bytes = await fs.readFile(filePath);
  const zip = await JSZip.loadAsync(bytes);
  const paths = Object.keys(zip.files).filter((name) => !zip.files[name].dir).sort();
  const cacheDefinitionPaths = paths.filter((name) => /^pivotCache\/pivotCacheDefinition\d*\.xml$/i.test(name));
  const cacheRecordPaths = paths.filter((name) => /^pivotCache\/pivotCacheRecords\d*\.xml$/i.test(name));
  const pivotTablePaths = paths.filter((name) => /^xl\/pivotTables\/pivotTable\d*\.xml$/i.test(name));
  const cacheDefinitionPath = cacheDefinitionPaths.length === 1 ? cacheDefinitionPaths[0] : null;
  const pivotTablePath = pivotTablePaths.length === 1 ? pivotTablePaths[0] : null;
  const cacheXml = cacheDefinitionPath ? await zip.file(cacheDefinitionPath).async("text") : "";
  const pivotXml = pivotTablePath ? await zip.file(pivotTablePath).async("text") : "";
  return {
    bytes: bytes.length,
    sha256: sha256(bytes),
    paths,
    partHashes: await packagePartHashes(zip, paths),
    cacheDefinitionPaths,
    cacheRecordPaths,
    pivotTablePaths,
    cacheDefinitionPath,
    pivotTablePath,
    cache: normalizePivotRefreshCacheDefinition(cacheXml),
    pivot: parsePivotTableDefinition(pivotXml),
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

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function packageChanges(source, output) {
  const paths = [...new Set([...source.paths, ...output.paths])].sort();
  return paths.filter((partPath) => source.partHashes[partPath] !== output.partHashes[partPath]);
}

function usedTypedXlsxRoundTrip(commandText, workflow) {
  const directPublicApi = /SpreadsheetFile\.importXlsx/i.test(commandText) && /SpreadsheetFile\.exportXlsx/i.test(commandText);
  return directPublicApi || workflow.test(commandText);
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

export function gradeXlsxThreadedReplyEvidence({ evidence, audit, commands }) {
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
    check("xlsx-trace:typed-roundtrip", "trace", usedTypedXlsxRoundTrip(commandText, SHIPPED_THREADED_WORKFLOW), {
      expected: "public SpreadsheetFile importXlsx/exportXlsx calls or the integrity-protected published threaded-comment workflow",
    }),
    check("xlsx-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

function growthCells(snapshot) {
  const cells = snapshot?.target?.cells || new Map();
  return {
    growth: cells.get(XLSX_GROWTH_UPDATE_FIXTURE.growthAddress),
    margin: cells.get(XLSX_GROWTH_UPDATE_FIXTURE.marginAddress),
    revenue: XLSX_GROWTH_UPDATE_FIXTURE.revenueFormulas.map((_, index) => cells.get(`B${5 + index}`)),
    grossProfit: [4, 5, 6, 7].map((row) => cells.get(`C${row}`)),
  };
}

function nativeGrowthVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount >= 2;
  const changedTargetPage = pageCountsMatch && source.pages.slice(0, -1).some((page, index) => page.pixelSha256 !== output.pages[index]?.pixelSha256);
  const sourceCanary = source?.pages?.at(-1);
  const outputCanary = output?.pages?.at(-1);
  const canaryStable = pageCountsMatch && sourceCanary?.width === outputCanary?.width
    && sourceCanary?.height === outputCanary?.height
    && sourceCanary?.pixelSha256 === outputCanary?.pixelSha256;
  return { available, rendered, pageCountsMatch, changedTargetPage, canaryStable };
}

export function gradeXlsxGrowthUpdateEvidence({ evidence, audit, commands }) {
  const fixture = XLSX_GROWTH_UPDATE_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceCells = growthCells(source);
  const outputCells = growthCells(output);
  const changedPaths = packageChanges(source, output);
  const expectedChangedPaths = [source.target?.path].filter(Boolean);
  const visual = nativeGrowthVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const formulasPreserved = sourceCells.revenue.every((cell, index) => normalizeFormula(cell?.formula) === normalizeFormula(fixture.revenueFormulas[index]))
    && outputCells.revenue.every((cell, index) => normalizeFormula(cell?.formula) === normalizeFormula(fixture.revenueFormulas[index]));
  const revisedRevenue = outputCells.revenue.map((cell) => cell?.value);
  return [
    check("xlsx-growth-machine:canonical-source", "machine", sameArray(source.sheets.map((sheet) => sheet.name), [fixture.targetSheetName, fixture.canarySheetName])
      && closeEnough(sourceCells.growth?.value, fixture.originalGrowth)
      && closeEnough(sourceCells.margin?.value, fixture.grossMargin)
      && formulasPreserved, {
      sheets: source.sheets,
      growth: sourceCells.growth,
      margin: sourceCells.margin,
      revenue: sourceCells.revenue,
    }),
    check("xlsx-growth-machine:only-growth-updated-and-recalculated", "machine", closeEnough(outputCells.growth?.value, fixture.replacementGrowth)
      && closeEnough(outputCells.margin?.value, fixture.grossMargin)
      && formulasPreserved
      && revisedRevenue.every((value, index) => closeEnough(value, fixture.revisedRevenue[index], 1e-7)), {
      outputGrowth: outputCells.growth,
      outputMargin: outputCells.margin,
      revisedRevenue,
      expectedRevenue: fixture.revisedRevenue,
    }),
    check("xlsx-growth-machine:sheet-topology-and-canary-preserved", "machine", sameArray(source.sheets.map((sheet) => sheet.name), output.sheets.map((sheet) => sheet.name))
      && source.canary?.path === output.canary?.path
      && source.partHashes[source.canary?.path] === output.partHashes[output.canary?.path], {
      sourceSheets: source.sheets,
      outputSheets: output.sheets,
      canary: { source: source.canary?.path, output: output.canary?.path },
    }),
    check("xlsx-growth-machine:only-bounded-parts-changed", "machine", sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
    }),
    check("xlsx-growth-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("xlsx-growth-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("xlsx-growth-visual:target-changed-and-canary-stable", "visual", visual.changedTargetPage && visual.canaryStable, {
      visual: evidence.visual,
    }),
    gate("xlsx-growth-security:package-and-canary-preservation", "security", sameArray(source.paths, output.paths)
      && source.target?.path === output.target?.path
      && source.canary?.path === output.canary?.path
      && source.partHashes[source.canary?.path] === output.partHashes[output.canary?.path]
      && sameArray(changedPaths, expectedChangedPaths), {
      changedPaths,
      expectedChangedPaths,
      sourceCanary: source.canary?.path,
      outputCanary: output.canary?.path,
    }),
    gate("xlsx-growth-security:byte-bound-audit-provenance", "security", auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("xlsx-growth-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("xlsx-growth-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("xlsx-growth-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), { strategy: auditStrategy(audit) }),
    check("xlsx-growth-trace:growth-operation", "trace", /growth|assumption|operating-plan/i.test(auditOperation(audit)), { operation: auditOperation(audit) }),
    check("xlsx-growth-trace:typed-roundtrip", "trace", usedTypedXlsxRoundTrip(commandText, SHIPPED_GROWTH_WORKFLOW), {
      expected: "public SpreadsheetFile importXlsx/exportXlsx calls or the integrity-protected published growth-assumption workflow",
    }),
    check("xlsx-growth-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

function nativeConnectionRefreshVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount >= 1;
  const allPagesStable = pageCountsMatch && source.pages.every((page, index) => page.width === output.pages[index]?.width
    && page.height === output.pages[index]?.height
    && page.pixelSha256 === output.pages[index]?.pixelSha256);
  return { available, rendered, pageCountsMatch, allPagesStable };
}

function usedConnectionRefreshWorkflow(commandText) {
  return (/\SpreadsheetFile\.importXlsx/i.test(commandText)
    && /\SpreadsheetFile\.exportXlsx/i.test(commandText)
    && /\disableConnectionRefreshOnLoad/i.test(commandText))
    || SHIPPED_CONNECTION_REFRESH_WORKFLOW.test(commandText);
}

export function gradeXlsxConnectionRefreshEvidence({ evidence, audit, commands }) {
  const fixture = XLSX_CONNECTION_REFRESH_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const sourceConnection = source.connections.find((connection) => connection.id === fixture.connectionId);
  const outputConnection = output.connections.find((connection) => connection.id === fixture.connectionId);
  const changedPaths = packageChanges(source, output);
  const changedScopeSafe = sameArray(changedPaths, [source.connectionPath]);
  const sourceConnectionResidual = normalizeConnectionRefreshConnectionPart(source.connectionXml, fixture.connectionId);
  const outputConnectionResidual = normalizeConnectionRefreshConnectionPart(output.connectionXml, fixture.connectionId);
  const connectionResidualStable = sourceConnectionResidual.targetCount === 1
    && outputConnectionResidual.targetCount === 1
    && sourceConnectionResidual.targetRefreshAttributeCount === 1
    && outputConnectionResidual.targetRefreshAttributeCount === 1
    && sourceConnectionResidual.normalized === outputConnectionResidual.normalized;
  const visual = nativeConnectionRefreshVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const expectedConnection = sourceConnection && { ...sourceConnection, refreshOnLoad: false };
  const sourceConnectionRecognized = source.connections.length === 1
    && sourceConnection?.name === fixture.connectionName
    && sourceConnection?.refreshOnLoad === true
    && sourceConnection?.command === fixture.connectionCommand
    && sourceConnection?.opaqueValue === fixture.connectionOpaqueValue
    && source.queryTablePaths.length === 1
    && source.tablePaths.length === 1;
  const onlyRefreshChanged = Boolean(outputConnection && expectedConnection && sameJson(outputConnection, expectedConnection));
  return [
    check("xlsx-connection-machine:canonical-source", "machine", sourceConnectionRecognized, {
      connections: source.connections,
      queryTablePaths: source.queryTablePaths,
      tablePaths: source.tablePaths,
    }),
    check("xlsx-connection-machine:refresh-on-open-disabled", "machine", output.connections.length === 1
      && onlyRefreshChanged
      && outputConnection?.refreshOnLoad === false, {
      sourceConnection,
      outputConnection,
    }),
    check("xlsx-connection-machine:querytable-and-table-byte-stable", "machine", source.queryTablePaths.length === 1
      && sameArray(source.queryTablePaths, output.queryTablePaths)
      && sameArray(source.tablePaths, output.tablePaths)
      && source.queryTablePaths.every((partPath) => source.partHashes[partPath] === output.partHashes[partPath])
      && source.tablePaths.every((partPath) => source.partHashes[partPath] === output.partHashes[partPath]), {
      queryTablePaths: { source: source.queryTablePaths, output: output.queryTablePaths },
      tablePaths: { source: source.tablePaths, output: output.tablePaths },
    }),
    check("xlsx-connection-machine:connection-residual-stable", "machine", connectionResidualStable, {
      source: sourceConnectionResidual,
      output: outputConnectionResidual,
    }),
    check("xlsx-connection-machine:only-connections-part-changed", "machine", changedScopeSafe, {
      changedPaths,
      expectedChangedPaths: [source.connectionPath],
    }),
    check("xlsx-connection-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("xlsx-connection-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("xlsx-connection-visual:all-pages-stable", "visual", visual.allPagesStable, {
      visual: evidence.visual,
    }),
    gate("xlsx-connection-security:package-scope-and-source-provenance", "security", sameArray(source.paths, output.paths)
      && changedScopeSafe
      && connectionResidualStable
      && auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      changedPaths,
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("xlsx-connection-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("xlsx-connection-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("xlsx-connection-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), { strategy: auditStrategy(audit) }),
    check("xlsx-connection-trace:typed-operation", "trace", /connection.*refresh|refresh.*connection/i.test(auditOperation(audit)), { operation: auditOperation(audit) }),
    check("xlsx-connection-trace:typed-roundtrip", "trace", usedConnectionRefreshWorkflow(commandText), {
      expected: "public SpreadsheetFile importXlsx/exportXlsx plus workbook.disableConnectionRefreshOnLoad, or the integrity-protected published connection-refresh workflow",
    }),
    check("xlsx-connection-trace:second-import", "trace", audit?.validation?.reimport?.ok === true || audit?.validation?.secondImport?.ok === true, {
      validation: audit?.validation || null,
    }),
  ];
}

function nativePivotRefreshVisualEvidence(source, output) {
  const available = Boolean(source?.available && output?.available);
  const rendered = source?.ok === true && output?.ok === true
    && source.pages?.every((page) => page.nonWhitePixels > 0)
    && output.pages?.every((page) => page.nonWhitePixels > 0);
  const pageCountsMatch = source?.pageCount === output?.pageCount && source?.pageCount >= 1;
  const allPagesStable = pageCountsMatch && source.pages.every((page, index) => page.width === output.pages[index]?.width
    && page.height === output.pages[index]?.height
    && page.pixelSha256 === output.pages[index]?.pixelSha256);
  return { available, rendered, pageCountsMatch, allPagesStable };
}

function usedPivotRefreshWorkflow(commandText) {
  return (/\SpreadsheetFile\.importXlsx/i.test(commandText)
    && /\SpreadsheetFile\.exportXlsx/i.test(commandText)
    && /\.disableRefreshOnLoad\b/i.test(commandText))
    || SHIPPED_PIVOT_REFRESH_WORKFLOW.test(commandText);
}

export function gradeXlsxPivotRefreshEvidence({ evidence, audit, commands }) {
  const fixture = XLSX_PIVOT_REFRESH_FIXTURE;
  const source = evidence.source;
  const output = evidence.output;
  const changedPaths = packageChanges(source, output);
  const changedScopeSafe = source.cacheDefinitionPath !== null
    && sameArray(changedPaths, [source.cacheDefinitionPath]);
  const cacheResidualStable = source.cache.rootCount === 1
    && output.cache.rootCount === 1
    && source.cache.refreshAttributeCount === 1
    && output.cache.refreshAttributeCount === 1
    && source.cache.enabled === true
    && output.cache.enabled === false
    && source.cache.normalized === output.cache.normalized;
  const sourceRecognized = source.cacheDefinitionPaths.length === 1
    && source.cacheRecordPaths.length === 1
    && source.pivotTablePaths.length === 1
    && source.pivot.rootCount === 1
    && source.pivot.name === fixture.pivotName
    && source.cache.enabled === true;
  const outputRecognized = output.cacheDefinitionPaths.length === 1
    && output.cacheRecordPaths.length === 1
    && output.pivotTablePaths.length === 1
    && output.pivot.rootCount === 1
    && output.pivot.name === fixture.pivotName
    && output.cache.enabled === false;
  const immutablePivotGraph = sameArray(source.pivotTablePaths, output.pivotTablePaths)
    && sameArray(source.cacheRecordPaths, output.cacheRecordPaths)
    && source.pivotTablePaths.every((partPath) => source.partHashes[partPath] === output.partHashes[partPath])
    && source.cacheRecordPaths.every((partPath) => source.partHashes[partPath] === output.partHashes[partPath]);
  const visual = nativePivotRefreshVisualEvidence(evidence.visual?.source, evidence.visual?.output);
  const commandText = commands.join("\n");
  const auditPivot = audit?.operation?.pivot || {};
  return [
    check("xlsx-pivot-machine:canonical-source", "machine", sourceRecognized, {
      cacheDefinitionPaths: source.cacheDefinitionPaths,
      cacheRecordPaths: source.cacheRecordPaths,
      pivotTablePaths: source.pivotTablePaths,
      pivot: source.pivot,
      cache: source.cache,
    }),
    check("xlsx-pivot-machine:refresh-on-open-disabled", "machine", outputRecognized
      && auditPivot.worksheetName === fixture.sheetName
      && auditPivot.name === fixture.pivotName
      && auditPivot.previousRefreshOnLoad === true
      && auditPivot.refreshOnLoad === false, {
      output: { pivot: output.pivot, cache: output.cache },
      auditPivot,
    }),
    check("xlsx-pivot-machine:pivot-and-cache-records-byte-stable", "machine", immutablePivotGraph, {
      pivotTablePaths: { source: source.pivotTablePaths, output: output.pivotTablePaths },
      cacheRecordPaths: { source: source.cacheRecordPaths, output: output.cacheRecordPaths },
    }),
    check("xlsx-pivot-machine:cache-residual-stable", "machine", cacheResidualStable, {
      source: source.cache,
      output: output.cache,
    }),
    check("xlsx-pivot-machine:only-cache-definition-changed", "machine", changedScopeSafe, {
      changedPaths,
      expectedChangedPaths: source.cacheDefinitionPath ? [source.cacheDefinitionPath] : [],
    }),
    check("xlsx-pivot-machine:audit-succeeded", "machine", /^(?:success|succeeded|completed)$/i.test(String(audit?.status || "")), {
      status: audit?.status || "unreported",
    }),
    check("xlsx-pivot-visual:native-render", "visual", visual.available && visual.rendered && visual.pageCountsMatch, {
      visual: evidence.visual,
    }),
    check("xlsx-pivot-visual:all-pages-stable", "visual", visual.allPagesStable, {
      visual: evidence.visual,
    }),
    gate("xlsx-pivot-security:package-scope-and-source-provenance", "security", sameArray(source.paths, output.paths)
      && changedScopeSafe
      && immutablePivotGraph
      && cacheResidualStable
      && auditHash(audit, "source") === source.sha256
      && auditHash(audit, "output") === output.sha256
      && source.sha256 !== output.sha256, {
      changedPaths,
      source: { expected: source.sha256, actual: auditHash(audit, "source") },
      output: { expected: output.sha256, actual: auditHash(audit, "output") },
    }),
    check("xlsx-pivot-trace:open-chestnut-provider", "trace", /open[- ]?chestnut/i.test(auditProvider(audit)) && Boolean(auditVersion(audit)), {
      provider: auditProvider(audit),
      version: auditVersion(audit),
    }),
    gate("xlsx-pivot-trace:no-silent-fallback", "trace", auditFallbackIsFalse(audit), { provider: audit?.provider || null }),
    check("xlsx-pivot-trace:rewrite-policy", "trace", /^rewrite$/i.test(auditStrategy(audit)), { strategy: auditStrategy(audit) }),
    check("xlsx-pivot-trace:typed-operation", "trace", /pivot.*refresh|refresh.*pivot/i.test(auditOperation(audit)), { operation: auditOperation(audit) }),
    check("xlsx-pivot-trace:typed-roundtrip", "trace", usedPivotRefreshWorkflow(commandText), {
      expected: "public SpreadsheetFile importXlsx/exportXlsx plus pivot.disableRefreshOnLoad, or the integrity-protected published Pivot refresh workflow",
    }),
    check("xlsx-pivot-trace:second-import", "trace", audit?.validation?.reimport?.ok === true
      && audit?.validation?.reimport?.pivotProjectionPreserved === true
      && audit?.validation?.reimport?.refreshOnLoadDisabled === true
      && audit?.validation?.reimport?.refreshOnLoadHardenableWithdrawn === true, {
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

async function gradeThreadedReplyCase({ item, workspace, finalMessage, trace, weights }) {
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const fixture = XLSX_THREADED_REVIEW_FIXTURE;
  let source;
  let output;
  try {
    source = await inspectThreadedWorkbook(path.join(workspace, "inputs", fixture.workbookName));
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
    renderOfficeFile(path.join(workspace, "inputs", fixture.workbookName), "xlsx-threaded-source"),
    renderOfficeFile(path.join(workspace, "outputs", "reviewed-budget-resolved.xlsx"), "xlsx-threaded-output"),
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
  const checks = gradeXlsxThreadedReplyEvidence({ evidence, audit, commands });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeGrowthUpdateCase({ item, workspace, finalMessage, trace, weights }) {
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const fixture = XLSX_GROWTH_UPDATE_FIXTURE;
  let source;
  let output;
  try {
    source = await inspectGrowthWorkbook(path.join(workspace, "inputs", fixture.workbookName));
    output = await inspectGrowthWorkbook(path.join(workspace, "outputs", "operating-plan-updated.xlsx"));
  } catch (error) {
    const checks = [
      gate("xlsx-growth-machine:readable-output", "machine", false, { error: error.message }),
      gate("xlsx-growth-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(path.join(workspace, "inputs", fixture.workbookName), "xlsx-growth-source"),
    renderOfficeFile(path.join(workspace, "outputs", "operating-plan-updated.xlsx"), "xlsx-growth-output"),
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
  const checks = gradeXlsxGrowthUpdateEvidence({ evidence, audit, commands });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradeConnectionRefreshCase({ item, workspace, finalMessage, trace, weights }) {
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const fixture = XLSX_CONNECTION_REFRESH_FIXTURE;
  let source;
  let output;
  try {
    source = await inspectConnectionRefreshWorkbook(path.join(workspace, "inputs", fixture.workbookName));
    output = await inspectConnectionRefreshWorkbook(path.join(workspace, "outputs", "external-sales-refresh-disabled.xlsx"));
  } catch (error) {
    const checks = [
      gate("xlsx-connection-machine:readable-output", "machine", false, { error: error.message }),
      gate("xlsx-connection-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(path.join(workspace, "inputs", fixture.workbookName), "xlsx-connection-source"),
    renderOfficeFile(path.join(workspace, "outputs", "external-sales-refresh-disabled.xlsx"), "xlsx-connection-output"),
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
  const checks = gradeXlsxConnectionRefreshEvidence({ evidence, audit, commands });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

async function gradePivotRefreshCase({ item, workspace, finalMessage, trace, weights }) {
  const audit = await readAudit(workspace);
  const commands = extractCompletedCommands(trace);
  const fixture = XLSX_PIVOT_REFRESH_FIXTURE;
  let source;
  let output;
  try {
    source = await inspectPivotRefreshWorkbook(path.join(workspace, "inputs", fixture.workbookName));
    output = await inspectPivotRefreshWorkbook(path.join(workspace, "outputs", "regional-revenue-refresh-disabled.xlsx"));
  } catch (error) {
    const checks = [
      gate("xlsx-pivot-machine:readable-output", "machine", false, { error: error.message }),
      gate("xlsx-pivot-security:no-partial-success", "security", false, { error: error.message }),
    ];
    const score = summarizeCaseScore(checks, item.grade, weights, false);
    return { supported: true, graded: true, checks, evidence: { error: error.message }, pending: [], ...score };
  }
  const [sourceRender, outputRender] = await Promise.all([
    renderOfficeFile(path.join(workspace, "inputs", fixture.workbookName), "xlsx-pivot-source"),
    renderOfficeFile(path.join(workspace, "outputs", "regional-revenue-refresh-disabled.xlsx"), "xlsx-pivot-output"),
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
  const checks = gradeXlsxPivotRefreshEvidence({ evidence, audit, commands });
  const score = summarizeCaseScore(checks, item.grade, weights, checks.filter((entry) => entry.gate).every((entry) => entry.passed));
  return { supported: true, graded: true, checks, evidence, pending: [], ...score };
}

export async function gradeSpreadsheetCase({ item, workspace, finalMessage, trace, weights = defaultWeights }) {
  if (!spreadsheetGradedCaseIds.has(item.id)) return { supported: false };
  if (item.id === "xlsx-threaded-reply-resolve") return gradeThreadedReplyCase({ item, workspace, finalMessage, trace, weights });
  if (item.id === "xlsx-connection-refresh-on-open") return gradeConnectionRefreshCase({ item, workspace, finalMessage, trace, weights });
  if (item.id === "xlsx-pivot-refresh-on-open") return gradePivotRefreshCase({ item, workspace, finalMessage, trace, weights });
  return gradeGrowthUpdateCase({ item, workspace, finalMessage, trace, weights });
}
