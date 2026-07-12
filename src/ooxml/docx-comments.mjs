import path from "node:path";

import { attrEscape, attributes, decodeXml, rootTag } from "./source-reference-xml.mjs";

export const DOCX_COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
export const DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
export const DOCX_COMMENTS_EXTENDED_PATH = "word/commentsExtended.xml";

const WORDPROCESSINGML_NAMESPACE = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const WORD_2010_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";
const WORD_2012_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const MARKUP_COMPATIBILITY_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const PARA_ID = /^[0-9A-F]{8}$/;
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

function normalizeParaId(value, label) {
  const normalized = String(value || "").trim().toUpperCase();
  if (!PARA_ID.test(normalized)) throw new TypeError(`${label} must be exactly eight hexadecimal digits.`);
  return normalized;
}

function parentIdOf(comment) {
  const parent = comment.parentId ?? comment.replyToId ?? comment.replyTo;
  return typeof parent === "object" ? parent?.id : parent;
}

function generatedParaId(index, used) {
  let value = index + 1;
  while (value <= 0xFFFFFFFF) {
    const candidate = value.toString(16).toUpperCase().padStart(8, "0");
    if (!used.has(candidate)) return candidate;
    value += 1;
  }
  throw new RangeError("DOCX comment paraId space is exhausted.");
}

function validateParentGraph(entries, byId) {
  for (const entry of entries) {
    const visited = new Set([entry.comment.id]);
    let parentId = entry.parentId;
    while (parentId) {
      const parent = byId.get(parentId);
      if (!parent) throw new Error(`DOCX comment ${entry.comment.id} references missing parent comment ${parentId}.`);
      if (visited.has(parentId)) throw new Error(`DOCX comment ${entry.comment.id} has a cyclic parent chain.`);
      if (parent.targetId !== entry.comment.targetId) throw new Error(`DOCX comment ${entry.comment.id} and parent ${parentId} must target the same block.`);
      visited.add(parentId);
      parentId = parentIdOf(parent);
    }
  }
}

export function planDocxComments(comments = []) {
  if (!Array.isArray(comments)) throw new TypeError("DOCX comments must be an array.");
  const byId = new Map();
  for (const comment of comments) {
    const id = String(comment?.id || "").trim();
    if (!id) throw new TypeError("DOCX comment IDs must be non-empty strings.");
    if (byId.has(id)) throw new Error(`Duplicate DOCX comment ID ${id}.`);
    byId.set(id, comment);
  }
  const usedParaIds = new Set();
  for (const comment of comments) {
    if (!comment.paraId) continue;
    const paraId = normalizeParaId(comment.paraId, `DOCX comment ${comment.id} paraId`);
    if (usedParaIds.has(paraId)) throw new Error(`Duplicate DOCX comment paraId ${paraId}.`);
    usedParaIds.add(paraId);
  }
  const entries = comments.map((comment, index) => {
    if (comment.date && Number.isNaN(Date.parse(comment.date))) throw new TypeError(`DOCX comment ${comment.id} date must be a valid date string.`);
    const paraId = comment.paraId ? normalizeParaId(comment.paraId, `DOCX comment ${comment.id} paraId`) : generatedParaId(index, usedParaIds);
    usedParaIds.add(paraId);
    const parentId = parentIdOf(comment);
    return { comment, commentId: String(index), paraId, parentId: parentId ? String(parentId) : undefined };
  });
  validateParentGraph(entries, byId);
  const entryByCommentId = new Map(entries.map((entry) => [entry.comment.id, entry]));
  for (const entry of entries) entry.parentParaId = entry.parentId ? entryByCommentId.get(entry.parentId).paraId : undefined;
  return { entries, entryByCommentId };
}

export function docxCommentsXml(plan) {
  const comments = plan.entries.map(({ comment, commentId, paraId }) => `<w:comment w:id="${attrEscape(commentId)}" w:author="${attrEscape(comment.author || "User")}" w:initials="${attrEscape(comment.initials || "")}"${comment.date ? ` w:date="${attrEscape(comment.date)}"` : ""}><w:p w14:paraId="${paraId}"><w:r><w:t>${attrEscape(comment.text || "")}</w:t></w:r></w:p></w:comment>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:comments xmlns:w="${WORDPROCESSINGML_NAMESPACE}" xmlns:w14="${WORD_2010_NAMESPACE}" xmlns:mc="${MARKUP_COMPATIBILITY_NAMESPACE}" mc:Ignorable="w14">${comments}</w:comments>`;
}

export function docxCommentsExtendedXml(plan) {
  const comments = plan.entries.map(({ comment, paraId, parentParaId }) => `<w15:commentEx w15:paraId="${paraId}"${parentParaId ? ` w15:paraIdParent="${parentParaId}"` : ""} w15:done="${comment.resolved ? 1 : 0}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w15:commentsEx xmlns:w15="${WORD_2012_NAMESPACE}">${comments}</w15:commentsEx>`;
}

function parseCommentsExtended(xml = "") {
  const byParaId = new Map();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentEx\b[^>]*\/?\s*>/g)) {
    const paraId = localAttribute(match[0], "paraId")?.toUpperCase();
    if (!paraId || byParaId.has(paraId)) continue;
    const done = String(localAttribute(match[0], "done") || "0").trim().toLowerCase();
    byParaId.set(paraId, { parentParaId: localAttribute(match[0], "paraIdParent")?.toUpperCase(), resolved: new Set(["1", "true", "on"]).has(done) });
  }
  return byParaId;
}

export function parseDocxComments(classicXml = "", extendedXml = "") {
  const extensions = parseCommentsExtended(extendedXml);
  const comments = new Map();
  for (const match of String(classicXml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?comment\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?comment>/g)) {
    const opening = /^<(?:[A-Za-z_][\w.-]*:)?comment\b[^>]*>/.exec(match[0])?.[0] || "";
    const id = localAttribute(opening, "id");
    if (id === undefined || comments.has(id)) continue;
    const paragraphs = [...match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?p\b[^>]*>/g)];
    const paraId = paragraphs.length ? localAttribute(paragraphs.at(-1)[0], "paraId")?.toUpperCase() : undefined;
    const extension = paraId ? extensions.get(paraId) : undefined;
    comments.set(id, {
      text: decodeXml([...match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((textMatch) => textMatch[1]).join("")),
      author: localAttribute(opening, "author") || "User",
      initials: localAttribute(opening, "initials") || undefined,
      date: localAttribute(opening, "date") || undefined,
      paraId,
      parentParaId: extension?.parentParaId,
      resolved: extension?.resolved || false,
    });
  }
  return comments;
}

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "DOCX", type, severity: "error", ...detail, message };
}

function relationshipSource(partPath) {
  if (partPath === "_rels/.rels") return "";
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/.exec(partPath);
  if (!match) return undefined;
  return match[1] ? `${match[1]}/${match[2]}` : match[2];
}

function resolveTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const sourceDirectory = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDirectory === "." ? "" : sourceDirectory, target));
}

function packageXml(bytesByPath, partPath) {
  const bytes = bytesByPath.get(partPath);
  return bytes ? decoder.decode(bytes) : "";
}

function declaredContentType(contentTypes, partPath) {
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  return contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension);
}

function extendedRelationships(bytesByPath) {
  const entries = [];
  for (const [partPath] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = relationshipSource(partPath);
    if (source == null) continue;
    for (const match of packageXml(bytesByPath, partPath).matchAll(/<Relationship\b[^>]*\/?\s*>/g)) {
      const attrs = attributes(match[0]);
      if (attrs.Type !== DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE) continue;
      entries.push({ source, path: partPath, id: attrs.Id, target: attrs.Target ? resolveTarget(source, attrs.Target) : undefined, external: String(attrs.TargetMode || "").toLowerCase() === "external" });
    }
  }
  return entries;
}

function classicCommentsTarget(bytesByPath, source) {
  const sourceDirectory = source ? path.posix.dirname(source) : "";
  const relationshipPath = source ? `${sourceDirectory === "." ? "" : `${sourceDirectory}/`}_rels/${path.posix.basename(source)}.rels` : "_rels/.rels";
  for (const match of packageXml(bytesByPath, relationshipPath).matchAll(/<Relationship\b[^>]*\/?\s*>/g)) {
    const attrs = attributes(match[0]);
    if (String(attrs.Type || "").endsWith("/comments") && String(attrs.TargetMode || "").toLowerCase() !== "external") return resolveTarget(source, attrs.Target);
  }
  return undefined;
}

function classicCommentParaIds(xml, partPath) {
  const issues = [];
  const paraIds = new Set();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?comment\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?comment>/g)) {
    const paragraphs = [...match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?p\b[^>]*>/g)];
    const paraId = paragraphs.length ? localAttribute(paragraphs.at(-1)[0], "paraId")?.toUpperCase() : undefined;
    if (!paraId || !PARA_ID.test(paraId)) issues.push(issue("docxCommentParaIdMissing", `DOCX comment in ${partPath} requires an eight-digit paraId when commentsExtended is present.`, { path: partPath, paraId }));
    else {
      if (paraIds.has(paraId)) issues.push(issue("docxCommentParaIdDuplicate", `DOCX comments part ${partPath} contains duplicate paraId ${paraId}.`, { path: partPath, paraId }));
      paraIds.add(paraId);
    }
  }
  return { paraIds, issues };
}

function validateExtendedPart(bytesByPath, contentTypes, entry) {
  const issues = [];
  if (entry.external) issues.push(issue("docxCommentsExtendedRelationshipExternal", "DOCX commentsExtended relationship must be internal.", { path: entry.path, relationshipId: entry.id }));
  const sourceXml = packageXml(bytesByPath, entry.source);
  if (!rootTag(sourceXml, "document") && !rootTag(sourceXml, "glossaryDocument")) issues.push(issue("docxCommentsExtendedSourceInvalid", "DOCX commentsExtended relationship must originate from a main or glossary document part.", { path: entry.path, source: entry.source, relationshipId: entry.id }));
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const actualContentType = declaredContentType(contentTypes, entry.target);
  if (actualContentType !== DOCX_COMMENTS_EXTENDED_CONTENT_TYPE) issues.push(issue("docxCommentsExtendedContentTypeInvalid", `DOCX commentsExtended part ${entry.target} must use ${DOCX_COMMENTS_EXTENDED_CONTENT_TYPE}.`, { path: entry.target, contentType: actualContentType, expectedContentType: DOCX_COMMENTS_EXTENDED_CONTENT_TYPE }));
  const extendedXml = packageXml(bytesByPath, entry.target);
  const root = rootTag(extendedXml, "commentsEx");
  const rootAttrs = attributes(root?.[0] || "");
  const rootNamespace = root ? (root[1] ? rootAttrs[`xmlns:${root[1]}`] : rootAttrs.xmlns) : undefined;
  if (rootNamespace !== WORD_2012_NAMESPACE) issues.push(issue("docxCommentsExtendedRootInvalid", `DOCX commentsExtended part ${entry.target} must have a commentsEx root in the Word 2012 namespace.`, { path: entry.target, namespace: rootNamespace }));
  const commentsTarget = classicCommentsTarget(bytesByPath, entry.source);
  if (!commentsTarget || !bytesByPath.has(commentsTarget)) {
    issues.push(issue("docxCommentsExtendedClassicCommentsMissing", "DOCX commentsExtended requires a classic Comments part related from the same document part.", { path: entry.target, source: entry.source }));
    return issues;
  }
  const classic = classicCommentParaIds(packageXml(bytesByPath, commentsTarget), commentsTarget);
  issues.push(...classic.issues);
  const extensionParents = new Map();
  for (const match of extendedXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentEx\b[^>]*\/?\s*>/g)) {
    const paraId = localAttribute(match[0], "paraId")?.toUpperCase();
    const parentParaId = localAttribute(match[0], "paraIdParent")?.toUpperCase();
    if (!paraId || !PARA_ID.test(paraId)) {
      issues.push(issue("docxCommentExParaIdInvalid", `DOCX commentEx in ${entry.target} requires an eight-digit paraId.`, { path: entry.target, paraId }));
      continue;
    }
    if (extensionParents.has(paraId)) issues.push(issue("docxCommentExParaIdDuplicate", `DOCX commentsExtended part ${entry.target} contains duplicate paraId ${paraId}.`, { path: entry.target, paraId }));
    extensionParents.set(paraId, parentParaId);
    if (!classic.paraIds.has(paraId)) issues.push(issue("docxCommentExReferenceNotFound", `DOCX commentEx paraId ${paraId} does not resolve to a classic comment paragraph.`, { path: entry.target, paraId }));
    if (parentParaId && !classic.paraIds.has(parentParaId)) issues.push(issue("docxCommentExParentNotFound", `DOCX commentEx paraId ${paraId} references missing parent paraId ${parentParaId}.`, { path: entry.target, paraId, parentParaId }));
  }
  for (const paraId of extensionParents.keys()) {
    const visited = new Set([paraId]);
    let parent = extensionParents.get(paraId);
    while (parent && extensionParents.has(parent)) {
      if (visited.has(parent)) {
        issues.push(issue("docxCommentExParentCycle", `DOCX commentsExtended part ${entry.target} contains a parent cycle at paraId ${parent}.`, { path: entry.target, paraId, parentParaId: parent }));
        break;
      }
      visited.add(parent);
      parent = extensionParents.get(parent);
    }
  }
  return issues;
}

export function validateDocxCommentPackageSemantics({ bytesByPath, contentTypes }) {
  const entries = extendedRelationships(bytesByPath);
  const issues = [];
  const counts = new Map();
  for (const entry of entries) {
    counts.set(entry.source, (counts.get(entry.source) || 0) + 1);
    issues.push(...validateExtendedPart(bytesByPath, contentTypes, entry));
  }
  for (const [source, count] of counts) if (count > 1) issues.push(issue("docxCommentsExtendedRelationshipMultiplicity", `DOCX document part ${source} has ${count} commentsExtended relationships; at most one is allowed.`, { source, count, maximum: 1 }));
  const targets = new Set(entries.filter((entry) => !entry.external).map((entry) => entry.target));
  for (const partPath of bytesByPath.keys()) if (declaredContentType(contentTypes, partPath) === DOCX_COMMENTS_EXTENDED_CONTENT_TYPE && !targets.has(partPath)) issues.push(issue("docxCommentsExtendedPartOrphaned", `DOCX commentsExtended part ${partPath} is not targeted by its required relationship.`, { path: partPath }));
  return issues;
}
