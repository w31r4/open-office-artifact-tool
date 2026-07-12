import { attrEscape, attributes, decodeXml, rootTag } from "./source-reference-xml.mjs";

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

export function planDocxBookmarks(bookmarks = [], blocks = []) {
  const blockById = new Map(blocks.map((block) => [block.id, block]));
  const names = new Set();
  const usedNativeIds = new Set();
  for (const bookmark of bookmarks) {
    const name = normalizeDocxBookmarkName(bookmark.name, `DOCX bookmark ${bookmark.id} name`);
    if (names.has(name)) throw new Error(`Duplicate DOCX bookmark name ${name}.`);
    names.add(name);
    if (!bookmark.nativeId && bookmark.nativeId !== 0) continue;
    const nativeId = Number(bookmark.nativeId);
    if (!validBookmarkNativeId(nativeId)) throw new RangeError(`DOCX bookmark ${bookmark.id} nativeId must be an integer greater than or equal to 0, or less than or equal to -2.`);
    if (usedNativeIds.has(nativeId)) throw new Error(`Duplicate DOCX bookmark nativeId ${nativeId}.`);
    usedNativeIds.add(nativeId);
  }
  let nextNativeId = 0;
  const entries = bookmarks.map((bookmark) => {
    const target = blockById.get(bookmark.targetId);
    const endTarget = blockById.get(bookmark.endTargetId || bookmark.targetId);
    if (!target) throw new Error(`DOCX bookmark ${bookmark.id} references missing start target ${bookmark.targetId}.`);
    if (!endTarget) throw new Error(`DOCX bookmark ${bookmark.id} references missing end target ${bookmark.endTargetId}.`);
    if (target.kind === "table" || endTarget.kind === "table") throw new Error(`DOCX bookmark ${bookmark.id} requires paragraph-backed start and end targets.`);
    let nativeId = bookmark.nativeId === 0 || bookmark.nativeId ? Number(bookmark.nativeId) : undefined;
    if (nativeId === undefined) {
      while (usedNativeIds.has(nextNativeId)) nextNativeId += 1;
      nativeId = nextNativeId;
      usedNativeIds.add(nativeId);
    }
    return { bookmark, id: bookmark.id, name: normalizeDocxBookmarkName(bookmark.name), nativeId, targetId: target.id, endTargetId: endTarget.id };
  });
  const blockIndex = new Map(blocks.map((block, index) => [block.id, index]));
  for (const entry of entries) if (blockIndex.get(entry.targetId) > blockIndex.get(entry.endTargetId)) throw new Error(`DOCX bookmark ${entry.id} end target precedes its start target.`);
  return { entries, byName: new Map(entries.map((entry) => [entry.name, entry])) };
}

export function wrapDocxParagraphBookmarks(xml = "", entries = []) {
  if (!entries.length) return String(xml || "");
  const source = String(xml || "");
  if (!/^<(?:[A-Za-z_][\w.-]*:)?p\b/.test(source) || !/<\/(?:[A-Za-z_][\w.-]*:)?p>\s*$/.test(source)) throw new Error("DOCX bookmark targets must serialize as paragraphs.");
  const starts = entries.filter((entry) => entry.start).map((entry) => `<w:bookmarkStart w:id="${entry.nativeId}" w:name="${attrEscape(entry.name)}"/>`).join("");
  const ends = entries.filter((entry) => entry.end).reverse().map((entry) => `<w:bookmarkEnd w:id="${entry.nativeId}"/>`).join("");
  const paragraphPropertiesEnd = source.indexOf("</w:pPr>");
  const startOffset = paragraphPropertiesEnd >= 0 ? paragraphPropertiesEnd + "</w:pPr>".length : source.indexOf(">") + 1;
  const endOffset = source.search(/<\/w:p>\s*$/);
  return `${source.slice(0, startOffset)}${starts}${source.slice(startOffset, endOffset)}${ends}${source.slice(endOffset)}`;
}

export function docxBookmarkEntriesForBlock(plan, targetId) {
  return plan.entries.flatMap((entry) => {
    const start = entry.targetId === targetId;
    const end = entry.endTargetId === targetId;
    return start || end ? [{ ...entry, start, end }] : [];
  });
}

export function parseDocxBookmarkMarkers(xml = "") {
  const starts = [...String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?bookmarkStart\b[^>]*\/?\s*>/g)].map((match) => ({
    nativeId: Number(localAttribute(match[0], "id")),
    name: decodeXml(localAttribute(match[0], "name") || ""),
  }));
  const ends = [...String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?bookmarkEnd\b[^>]*\/?\s*>/g)].map((match) => ({ nativeId: Number(localAttribute(match[0], "id")) }));
  return { starts, ends };
}

export function docxHyperlinkAttributes(link = {}, relationshipId) {
  const anchor = String(link.anchor || "").trim();
  if (anchor) {
    if (anchor.length > HYPERLINK_ANCHOR_MAX) throw new RangeError(`DOCX hyperlink ${link.id || "(unknown)"} anchor must contain at most ${HYPERLINK_ANCHOR_MAX} characters.`);
    return ` w:anchor="${attrEscape(anchor)}"${link.history === false ? ` w:history="0"` : ` w:history="1"`}${link.tooltip ? ` w:tooltip="${attrEscape(link.tooltip)}"` : ""}`;
  }
  if (!relationshipId) throw new Error(`DOCX external hyperlink ${link.id || "(unknown)"} is missing a relationship ID.`);
  return ` r:id="${attrEscape(relationshipId)}"${link.history === false ? ` w:history="0"` : ` w:history="1"`}${link.tooltip ? ` w:tooltip="${attrEscape(link.tooltip)}"` : ""}`;
}

export function parseDocxHyperlink(xml = "", relationshipsById = new Map()) {
  const source = String(xml || "");
  const opening = /^<(?:[A-Za-z_][\w.-]*:)?hyperlink\b[^>]*>/.exec(source)?.[0] || "";
  const relationshipId = localAttribute(opening, "id");
  const relationship = relationshipsById.get(relationshipId);
  const history = String(localAttribute(opening, "history") || "1").toLowerCase();
  return {
    text: decodeXml([...source.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((match) => match[1]).join("")),
    url: relationship?.target,
    anchor: decodeXml(localAttribute(opening, "anchor") || "") || undefined,
    relationshipId: relationship?.id,
    history: !new Set(["0", "false", "off"]).has(history),
    tooltip: decodeXml(localAttribute(opening, "tooltip") || "") || undefined,
  };
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
    else startById.set(nativeId, name);
    if (!name || name.length > BOOKMARK_NAME_MAX) issues.push(issue("docxBookmarkNameInvalid", `DOCX bookmark ${nativeId} name must contain 1 to ${BOOKMARK_NAME_MAX} characters.`, { path: partPath, nativeId, name }));
    else if (names.has(name)) issues.push(issue("docxBookmarkNameDuplicate", `DOCX document part ${partPath} contains duplicate bookmark name ${name}.`, { path: partPath, nativeId, name }));
    else names.add(name);
  }
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?bookmarkEnd\b[^>]*\/?\s*>/g)) {
    const nativeId = Number(localAttribute(match[0], "id"));
    endCounts.set(nativeId, (endCounts.get(nativeId) || 0) + 1);
  }
  for (const [nativeId, name] of startById) if (endCounts.get(nativeId) !== 1) issues.push(issue("docxBookmarkEndMissing", `DOCX bookmark ${name || nativeId} requires exactly one matching bookmarkEnd.`, { path: partPath, nativeId, ends: endCounts.get(nativeId) || 0 }));
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
