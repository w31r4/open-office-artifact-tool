import { attributes, decodeXml, rootTag } from "./source-reference-xml.mjs";

const BOOKMARK_NAME_MAX = 40;
const HYPERLINK_ANCHOR_MAX = 255;
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

export function normalizeDocxBookmarkName(value, label = "DOCX bookmark name") {
  const name = String(value || "").trim();
  if (!name) throw new TypeError(`${label} must be non-empty.`);
  if (name.length > BOOKMARK_NAME_MAX) throw new RangeError(`${label} must contain at most ${BOOKMARK_NAME_MAX} characters.`);
  return name;
}

function validBookmarkNativeId(value) {
  const number = Number(value);
  return Number.isInteger(number) && (number >= 0 || number <= -2);
}

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "DOCX", type, severity: "error", ...detail, message };
}

function validateDocumentLinks(xml = "", partPath) {
  const issues = [];
  const startById = new Map();
  const endCounts = new Map();
  const names = new Set();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?bookmarkStart\b[^>]*\/?\s*>/g)) {
    const nativeId = Number(localAttribute(match[0], "id"));
    const name = decodeXml(localAttribute(match[0], "name") || "");
    if (!validBookmarkNativeId(nativeId)) issues.push(issue("docxBookmarkIdInvalid", `DOCX bookmark in ${partPath} has invalid ID ${localAttribute(match[0], "id")}.`, { path: partPath, nativeId }));
    else if (startById.has(nativeId)) issues.push(issue("docxBookmarkIdDuplicate", `DOCX document part ${partPath} contains duplicate bookmark ID ${nativeId}.`, { path: partPath, nativeId }));
    else startById.set(nativeId, { name, offset: match.index });
    if (!name || name.length > BOOKMARK_NAME_MAX) issues.push(issue("docxBookmarkNameInvalid", `DOCX bookmark ${nativeId} name must contain 1 to ${BOOKMARK_NAME_MAX} characters.`, { path: partPath, nativeId, name }));
    else if (names.has(name)) issues.push(issue("docxBookmarkNameDuplicate", `DOCX document part ${partPath} contains duplicate bookmark name ${name}.`, { path: partPath, nativeId, name }));
    else names.add(name);
  }
  const endOffsets = new Map();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?bookmarkEnd\b[^>]*\/?\s*>/g)) {
    const nativeId = Number(localAttribute(match[0], "id"));
    endCounts.set(nativeId, (endCounts.get(nativeId) || 0) + 1);
    if (!endOffsets.has(nativeId)) endOffsets.set(nativeId, match.index);
  }
  for (const [nativeId, start] of startById) {
    if (endCounts.get(nativeId) !== 1) issues.push(issue("docxBookmarkEndMissing", `DOCX bookmark ${start.name || nativeId} requires exactly one matching bookmarkEnd.`, { path: partPath, nativeId, ends: endCounts.get(nativeId) || 0 }));
    else if (endOffsets.get(nativeId) < start.offset) issues.push(issue("docxBookmarkRangeReversed", `DOCX bookmark ${start.name || nativeId} ends before it starts.`, { path: partPath, nativeId }));
  }
  for (const [nativeId, count] of endCounts) if (!startById.has(nativeId)) issues.push(issue("docxBookmarkStartMissing", `DOCX bookmarkEnd ${nativeId} has no matching bookmarkStart.`, { path: partPath, nativeId, ends: count }));
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?hyperlink\b[^>]*>/g)) {
    const anchor = decodeXml(localAttribute(match[0], "anchor") || "");
    const relationshipId = localAttribute(match[0], "id");
    if (anchor.length > HYPERLINK_ANCHOR_MAX) issues.push(issue("docxHyperlinkAnchorInvalid", `DOCX hyperlink anchor must contain at most ${HYPERLINK_ANCHOR_MAX} characters.`, { path: partPath, anchor }));
    if (anchor && !relationshipId && !names.has(anchor)) issues.push(issue("docxHyperlinkAnchorNotFound", `DOCX internal hyperlink anchor ${anchor} does not resolve to a bookmark.`, { path: partPath, anchor }));
  }
  return issues;
}

export function validateDocxLinkPackageSemantics({ bytesByPath }) {
  const issues = [];
  for (const [partPath, bytes] of bytesByPath) {
    const xml = decoder.decode(bytes);
    if (rootTag(xml, "document") || rootTag(xml, "glossaryDocument")) issues.push(...validateDocumentLinks(xml, partPath));
  }
  return issues;
}
