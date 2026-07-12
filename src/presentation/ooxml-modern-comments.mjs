import path from "node:path";

import { attrEscape, attributes, decodeXml, rootTag } from "../ooxml/source-reference-xml.mjs";

export const PPTX_MODERN_COMMENT_NAMESPACE = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
export const PPTX_MODERN_COMMENT_CONTENT_TYPE = "application/vnd.ms-powerpoint.comments+xml";
export const PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
export const PPTX_MODERN_AUTHOR_CONTENT_TYPE = "application/vnd.ms-powerpoint.authors+xml";
export const PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/authors";

const PRESENTATION_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/presentationml/2006/main",
  "http://purl.oclc.org/ooxml/presentationml/main",
]);
const GUID = /^\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\}$/;
const STATUS = new Set(["active", "resolved", "closed"]);
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

function normalizeGuid(value, label) {
  const guid = String(value || "").trim().toUpperCase();
  if (!GUID.test(guid)) throw new TypeError(`${label} must be a brace-delimited GUID.`);
  return guid;
}

function hashWords(value) {
  let first = 0x811C9DC5;
  let second = 0x9E3779B9;
  for (const character of String(value)) {
    const code = character.codePointAt(0);
    first = Math.imul(first ^ code, 0x01000193) >>> 0;
    second = Math.imul(second ^ (code + first), 0x85EBCA6B) >>> 0;
  }
  return [first, second, (first ^ Math.imul(second, 33)) >>> 0, (second ^ Math.imul(first, 97)) >>> 0];
}

export function deterministicPresentationGuid(seed) {
  const words = hashWords(seed).map((word) => word.toString(16).toUpperCase().padStart(8, "0")).join("");
  return `{${words.slice(0, 8)}-${words.slice(8, 12)}-4${words.slice(13, 16)}-8${words.slice(17, 20)}-${words.slice(20, 32)}}`;
}

function validDate(value, label) {
  const text = String(value || "1970-01-01T00:00:00.000Z");
  if (Number.isNaN(Date.parse(text))) throw new TypeError(`${label} ${text} must be a valid date-time.`);
  return new Date(text).toISOString();
}

function commentPosition(value, label) {
  const x = Number(value?.x ?? 0);
  const y = Number(value?.y ?? 0);
  const minimum = -27_273_042_329_600;
  const maximum = 27_273_042_316_900;
  if (![x, y].every((coordinate) => Number.isFinite(coordinate) && coordinate >= minimum && coordinate <= maximum)) throw new RangeError(`${label} coordinates must be finite DrawingML positions.`);
  return { x: Math.round(x), y: Math.round(y) };
}

function commentAuthor(comment, thread) {
  const person = comment.person || {};
  const name = String(person.name || person.displayName || comment.author || thread.author || "User");
  return {
    id: normalizeGuid(comment.authorId || person.id || deterministicPresentationGuid(`author:${name}`), `PowerPoint modern comment author ${name} id`),
    name,
    initials: String(person.initials || comment.initials || initials(name)),
    userId: String(person.userId ?? comment.userId ?? name),
    providerId: String(person.providerId ?? comment.providerId ?? "None"),
  };
}

function initials(name) {
  const words = String(name || "User").trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words.slice(0, 2).map((word) => [...word][0]) : [...(words[0] || "U")].slice(0, 2)).join("").toUpperCase();
}

export function planPresentationModernComments(slides = []) {
  const authors = new Map();
  const commentIds = new Set();
  const parts = [];
  for (const [slideIndex, slide] of slides.entries()) {
    const roots = [];
    for (const [threadIndex, thread] of slide.comments.items.entries()) {
      if (!thread.comments.length) continue;
      const comments = thread.comments.map((comment, commentIndex) => {
        const author = commentAuthor(comment, thread);
        const prior = authors.get(author.id);
        if (prior && JSON.stringify(prior) !== JSON.stringify(author)) throw new Error(`PowerPoint modern comment author ${author.id} has conflicting metadata.`);
        authors.set(author.id, author);
        const id = normalizeGuid(comment.nativeId || comment.id || deterministicPresentationGuid(`comment:${thread.id || `${slideIndex}:${threadIndex}`}:${commentIndex}`), `PowerPoint modern comment ${commentIndex} id`);
        if (commentIds.has(id)) throw new Error(`Duplicate PowerPoint modern comment id ${id}.`);
        commentIds.add(id);
        const status = String(comment.status || (thread.resolved ? "resolved" : "active")).toLowerCase();
        if (!STATUS.has(status)) throw new TypeError(`PowerPoint modern comment ${id} status ${status} is invalid.`);
        return { id, authorId: author.id, author, status, created: validDate(comment.created || thread.created, `PowerPoint modern comment ${id} date`), text: String(comment.text ?? "") };
      });
      roots.push({
        id: comments[0].id,
        targetId: thread.targetId,
        targetName: slide.resolve(thread.targetId)?.name || "",
        position: commentPosition(thread.position || { x: 100 + threadIndex * 32, y: 100 + threadIndex * 32 }, `PowerPoint modern comment ${comments[0].id} position`),
        root: comments[0],
        replies: comments.slice(1),
      });
    }
    if (roots.length) parts.push({ slideIndex, roots });
  }
  return { authors: [...authors.values()], parts };
}

function textBody(text) {
  return `<p188:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${attrEscape(text)}</a:t></a:r></a:p></p188:txBody>`;
}

export function presentationModernAuthorsXml(plan) {
  const authors = plan.authors.map((author) => `<p188:author id="${attrEscape(author.id)}" name="${attrEscape(author.name)}" initials="${attrEscape(author.initials)}" userId="${attrEscape(author.userId)}" providerId="${attrEscape(author.providerId)}"/>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p188:authorLst xmlns:p188="${PPTX_MODERN_COMMENT_NAMESPACE}">${authors}</p188:authorLst>`;
}

export function presentationModernCommentsXml(part) {
  const comments = part.roots.map((entry) => {
    const replies = entry.replies.length ? `<p188:replyLst>${entry.replies.map((reply) => `<p188:reply id="${attrEscape(reply.id)}" authorId="${attrEscape(reply.authorId)}" status="${reply.status}" created="${attrEscape(reply.created)}">${textBody(reply.text)}</p188:reply>`).join("")}</p188:replyLst>` : "";
    const root = entry.root;
    return `<p188:cm id="${attrEscape(root.id)}" authorId="${attrEscape(root.authorId)}" status="${root.status}" created="${attrEscape(root.created)}"><p188:unknownAnchor/><p188:pos x="${Math.round(entry.position.x)}" y="${Math.round(entry.position.y)}"/>${replies}${textBody(root.text)}</p188:cm>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p188:cmLst xmlns:p188="${PPTX_MODERN_COMMENT_NAMESPACE}" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${comments}</p188:cmLst>`;
}

export function parsePresentationModernAuthors(xml = "") {
  const result = new Map();
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?author\b[^>]*\/?\s*>/g)) {
    const id = localAttribute(match[0], "id")?.toUpperCase();
    if (!id || result.has(id)) continue;
    result.set(id, { id, name: localAttribute(match[0], "name") || "User", initials: localAttribute(match[0], "initials") || undefined, userId: localAttribute(match[0], "userId") || undefined, providerId: localAttribute(match[0], "providerId") || undefined });
  }
  return result;
}

function textFromBody(xml = "") {
  return [...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((match) => decodeXml(match[1])).join("");
}

function parseCommentElement(xml, authors) {
  const opening = /^<(?:[A-Za-z_][\w.-]*:)?(?:cm|reply)\b[^>]*>/.exec(xml)?.[0] || "";
  const authorId = localAttribute(opening, "authorId")?.toUpperCase();
  const person = authors.get(authorId);
  return {
    nativeId: localAttribute(opening, "id")?.toUpperCase(),
    authorId,
    author: person?.name || "User",
    person,
    status: localAttribute(opening, "status") || "active",
    created: localAttribute(opening, "created") || new Date(0).toISOString(),
    text: textFromBody(xml),
  };
}

export function parsePresentationModernComments(xml = "", authors = new Map()) {
  const threads = [];
  for (const match of String(xml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?cm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?cm>/g)) {
    const source = match[0];
    const replyList = /<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/.exec(source)?.[1] || "";
    const rootWithoutReplies = source.replace(/<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/, "");
    const root = parseCommentElement(rootWithoutReplies, authors);
    const replies = [...replyList.matchAll(/<(?:[A-Za-z_][\w.-]*:)?reply\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?reply>/g)].map((reply) => parseCommentElement(reply[0], authors));
    const anchor = /<(?:[A-Za-z_][\w.-]*:)?anchor\b[^>]*\/?\s*>/.exec(source)?.[0] || "";
    const pos = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\/?\s*>/.exec(source)?.[0] || "";
    threads.push({
      id: root.nativeId,
      nativeFormat: "modern",
      targetId: localAttribute(anchor, "targetId") || undefined,
      targetName: localAttribute(anchor, "targetName") || undefined,
      author: root.author,
      resolved: root.status === "resolved" || root.status === "closed",
      created: root.created,
      position: pos ? { x: Number(localAttribute(pos, "x") || 0), y: Number(localAttribute(pos, "y") || 0) } : undefined,
      comments: [root, ...replies],
    });
  }
  return threads;
}

function issue(type, message, detail = {}) {
  return { kind: "ooxmlIssue", family: "PPTX", type, severity: "error", ...detail, message };
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

function rootNamespace(xml, localName) {
  const root = rootTag(xml, localName);
  if (!root) return undefined;
  const attrs = attributes(root[0]);
  return root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns;
}

function relationships(bytesByPath) {
  const result = [];
  for (const [partPath] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = relationshipSource(partPath);
    if (source == null) continue;
    for (const match of packageXml(bytesByPath, partPath).matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/?\s*>/g)) {
      const attrs = attributes(match[0]);
      if (![PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE, PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE].includes(attrs.Type)) continue;
      result.push({ source, path: partPath, id: attrs.Id, type: attrs.Type, target: attrs.Target ? resolveTarget(source, attrs.Target) : undefined, external: String(attrs.TargetMode || "").toLowerCase() === "external" });
    }
  }
  return result;
}

function validateCommentNode(node, authors, ids, partPath, root = false) {
  const issues = [];
  if (!node.nativeId || !GUID.test(node.nativeId)) issues.push(issue("pptxModernCommentIdInvalid", `PPTX modern comment id ${node.nativeId || "(missing)"} must be a brace-delimited GUID.`, { path: partPath, commentId: node.nativeId }));
  else if (ids.has(node.nativeId)) issues.push(issue("pptxModernCommentIdDuplicate", `PPTX modern comment id ${node.nativeId} is duplicated.`, { path: partPath, commentId: node.nativeId }));
  else ids.add(node.nativeId);
  if (!node.authorId || !GUID.test(node.authorId)) issues.push(issue("pptxModernCommentAuthorIdInvalid", `PPTX modern comment authorId ${node.authorId || "(missing)"} must be a brace-delimited GUID.`, { path: partPath, commentId: node.nativeId, authorId: node.authorId }));
  else if (!authors.has(node.authorId)) issues.push(issue("pptxModernCommentAuthorNotFound", `PPTX modern comment references missing author ${node.authorId}.`, { path: partPath, commentId: node.nativeId, authorId: node.authorId }));
  if (!STATUS.has(node.status)) issues.push(issue("pptxModernCommentStatusInvalid", `PPTX modern comment ${node.nativeId || "(missing)"} has invalid status ${node.status}.`, { path: partPath, commentId: node.nativeId, status: node.status }));
  if (Number.isNaN(Date.parse(node.created))) issues.push(issue("pptxModernCommentDateInvalid", `PPTX modern comment ${node.nativeId || "(missing)"} has invalid created date ${node.created}.`, { path: partPath, commentId: node.nativeId, created: node.created }));
  if (root && node.anchorCount !== 1) issues.push(issue("pptxModernCommentAnchorMissing", `PPTX modern root comment ${node.nativeId || "(missing)"} requires exactly one anchor.`, { path: partPath, commentId: node.nativeId, anchorCount: node.anchorCount }));
  if (root && node.position) {
    try { commentPosition(node.position, `PowerPoint modern comment ${node.nativeId || "(missing)"} position`); }
    catch { issues.push(issue("pptxModernCommentPositionInvalid", `PPTX modern comment ${node.nativeId || "(missing)"} has an invalid position.`, { path: partPath, commentId: node.nativeId, position: node.position })); }
  }
  return issues;
}

function relationshipsPartPath(source) {
  const directory = path.posix.dirname(source);
  const base = path.posix.basename(source);
  return path.posix.join(directory === "." ? "" : directory, "_rels", `${base}.rels`);
}

function hasOutgoingRelationships(bytesByPath, source) {
  return /<(?:[A-Za-z_][\w.-]*:)?Relationship\b/.test(packageXml(bytesByPath, relationshipsPartPath(source)));
}

export function validatePresentationModernCommentPackageSemantics({ bytesByPath, contentTypes }) {
  const issues = [];
  const rels = relationships(bytesByPath);
  const authorRels = rels.filter((rel) => rel.type === PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE);
  const commentRels = rels.filter((rel) => rel.type === PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE);
  const authors = new Map();
  if (authorRels.length > 1) issues.push(issue("pptxModernAuthorRelationshipMultiplicity", `PPTX package contains ${authorRels.length} modern author relationships; at most one is allowed.`, { count: authorRels.length }));
  for (const rel of authorRels) {
    if (rel.external) issues.push(issue("pptxModernAuthorRelationshipExternal", "PPTX modern author relationship must be internal.", { path: rel.path, relationshipId: rel.id }));
    if (!PRESENTATION_NAMESPACES.has(rootNamespace(packageXml(bytesByPath, rel.source), "presentation"))) issues.push(issue("pptxModernAuthorRelationshipSourceInvalid", "PPTX modern author relationship must originate from the Presentation part.", { path: rel.path, source: rel.source }));
    if (!rel.target || !bytesByPath.has(rel.target)) continue;
    if (hasOutgoingRelationships(bytesByPath, rel.target)) issues.push(issue("pptxModernAuthorPartRelationshipsForbidden", `PPTX modern author part ${rel.target} must not own package relationships.`, { path: relationshipsPartPath(rel.target), source: rel.target }));
    if (declaredContentType(contentTypes, rel.target) !== PPTX_MODERN_AUTHOR_CONTENT_TYPE) issues.push(issue("pptxModernAuthorContentTypeInvalid", `PPTX modern author part ${rel.target} has the wrong content type.`, { path: rel.target }));
    const xml = packageXml(bytesByPath, rel.target);
    if (rootNamespace(xml, "authorLst") !== PPTX_MODERN_COMMENT_NAMESPACE) issues.push(issue("pptxModernAuthorRootInvalid", `PPTX modern author part ${rel.target} must have an authorLst root in the Office 2021 namespace.`, { path: rel.target }));
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?author\b[^>]*\/?\s*>/g)) {
      const id = localAttribute(match[0], "id")?.toUpperCase();
      if (!id || !GUID.test(id)) issues.push(issue("pptxModernAuthorIdInvalid", `PPTX modern author id ${id || "(missing)"} must be a brace-delimited GUID.`, { path: rel.target, authorId: id }));
      else if (authors.has(id)) issues.push(issue("pptxModernAuthorIdDuplicate", `PPTX modern author id ${id} is duplicated.`, { path: rel.target, authorId: id }));
      if (!localAttribute(match[0], "name") || !localAttribute(match[0], "userId") || !localAttribute(match[0], "providerId")) issues.push(issue("pptxModernAuthorMetadataMissing", `PPTX modern author ${id || "(missing)"} requires name, userId, and providerId.`, { path: rel.target, authorId: id }));
      if (id) authors.set(id, { id });
    }
  }
  const commentCountBySource = new Map();
  const ids = new Set();
  for (const rel of commentRels) {
    commentCountBySource.set(rel.source, (commentCountBySource.get(rel.source) || 0) + 1);
    if (rel.external) issues.push(issue("pptxModernCommentRelationshipExternal", "PPTX modern comment relationship must be internal.", { path: rel.path, relationshipId: rel.id }));
    if (!PRESENTATION_NAMESPACES.has(rootNamespace(packageXml(bytesByPath, rel.source), "sld"))) issues.push(issue("pptxModernCommentRelationshipSourceInvalid", "PPTX modern comment relationship must originate from a Slide part.", { path: rel.path, source: rel.source }));
    if (!rel.target || !bytesByPath.has(rel.target)) continue;
    if (hasOutgoingRelationships(bytesByPath, rel.target)) issues.push(issue("pptxModernCommentPartRelationshipsForbidden", `PPTX modern comment part ${rel.target} must not own package relationships.`, { path: relationshipsPartPath(rel.target), source: rel.target }));
    if (declaredContentType(contentTypes, rel.target) !== PPTX_MODERN_COMMENT_CONTENT_TYPE) issues.push(issue("pptxModernCommentContentTypeInvalid", `PPTX modern comment part ${rel.target} has the wrong content type.`, { path: rel.target }));
    const xml = packageXml(bytesByPath, rel.target);
    if (rootNamespace(xml, "cmLst") !== PPTX_MODERN_COMMENT_NAMESPACE) issues.push(issue("pptxModernCommentRootInvalid", `PPTX modern comment part ${rel.target} must have a cmLst root in the Office 2021 namespace.`, { path: rel.target }));
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?cm>/g)) {
      const source = match[0];
      const repliesXml = /<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/.exec(source)?.[1] || "";
      const rootXml = source.replace(/<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/, "");
      const root = parseCommentElement(rootXml, authors);
      root.anchorCount = [...rootXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:sldMkLst|sldLayoutMkLst|sldMasterMkLst|deMkLst|txBodyMkLst|txMkLst|tcMkLst|trMkLst|gridColMkLst|unknownAnchor)\b/g)].length;
      const position = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\/?\s*>/.exec(rootXml)?.[0];
      if (position) root.position = { x: Number(localAttribute(position, "x")), y: Number(localAttribute(position, "y")) };
      issues.push(...validateCommentNode(root, authors, ids, rel.target, true));
      for (const reply of repliesXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?reply\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?reply>/g)) issues.push(...validateCommentNode(parseCommentElement(reply[0], authors), authors, ids, rel.target));
    }
  }
  for (const [source, count] of commentCountBySource) if (count > 1) issues.push(issue("pptxModernCommentRelationshipMultiplicity", `PPTX slide ${source} has ${count} modern comment relationships; at most one is allowed.`, { source, count }));
  const authorTargets = new Set(authorRels.filter((rel) => !rel.external).map((rel) => rel.target));
  const commentTargets = new Set(commentRels.filter((rel) => !rel.external).map((rel) => rel.target));
  for (const partPath of bytesByPath.keys()) {
    const type = declaredContentType(contentTypes, partPath);
    if (type === PPTX_MODERN_AUTHOR_CONTENT_TYPE && !authorTargets.has(partPath)) issues.push(issue("pptxModernAuthorPartOrphaned", `PPTX modern author part ${partPath} is not relationship-owned.`, { path: partPath }));
    if (type === PPTX_MODERN_COMMENT_CONTENT_TYPE && !commentTargets.has(partPath)) issues.push(issue("pptxModernCommentPartOrphaned", `PPTX modern comment part ${partPath} is not relationship-owned.`, { path: partPath }));
  }
  const authorParts = [...bytesByPath.keys()].filter((partPath) => declaredContentType(contentTypes, partPath) === PPTX_MODERN_AUTHOR_CONTENT_TYPE);
  if (authorParts.length > 1) issues.push(issue("pptxModernAuthorPartMultiplicity", `PPTX package contains ${authorParts.length} modern Author parts; at most one is allowed.`, { paths: authorParts, count: authorParts.length }));
  if (commentRels.length && !authorRels.length) issues.push(issue("pptxModernAuthorPartMissing", "PPTX modern comments require a presentation-owned Author part."));
  return issues;
}
