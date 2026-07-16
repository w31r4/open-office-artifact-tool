import path from "node:path";

import { attributes, decodeXml, rootTag } from "../ooxml/source-reference-xml.mjs";
import { directPresentationChildren } from "./group-shapes.mjs";

export const PPTX_MODERN_COMMENT_NAMESPACE = "http://schemas.microsoft.com/office/powerpoint/2018/8/main";
export const PPTX_MODERN_COMMENT_CONTENT_TYPE = "application/vnd.ms-powerpoint.comments+xml";
export const PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/comments";
export const PPTX_MODERN_AUTHOR_CONTENT_TYPE = "application/vnd.ms-powerpoint.authors+xml";
export const PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/10/relationships/authors";
export const PPTX_DRAWING_COMMAND_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2013/main/command";
export const PPTX_PRESENTATION_COMMAND_NAMESPACE = "http://schemas.microsoft.com/office/powerpoint/2013/main/command";
export const PPTX_CREATION_ID_NAMESPACE = "http://schemas.microsoft.com/office/drawing/2014/main";
export const PPTX_CREATION_ID_EXTENSION_URI = "{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}";

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

function textRangeMonikerXml(xml = "") {
  return /<(?:[A-Za-z_][\w.-]*:)?txMk\b[^>]*(?:\/\s*>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?txMk>)/.exec(String(xml))?.[0] || "";
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

export function planPresentationSlideElementIdentities(slide, entries = [], options = {}) {
  const used = new Set(options.reservedNativeIds || []);
  let nextId = 2;
  for (const { element } of entries) {
    const candidate = Number(element.nativeId);
    if (Number.isInteger(candidate) && candidate >= 2 && candidate <= 4_294_967_295 && !used.has(candidate)) used.add(candidate);
    else element.nativeId = undefined;
  }
  for (const { element, moniker } of entries) {
    if (!element.nativeId) {
      while (used.has(nextId)) nextId += 1;
      element.nativeId = nextId;
      used.add(nextId);
    }
    element.creationId = normalizeGuid(element.creationId || deterministicPresentationGuid(`drawing:${slide.id}:${element.id}`), `PowerPoint drawing element ${element.id} creationId`);
    element.moniker = moniker;
  }
  return entries.map(({ element, moniker, ancestors = [] }) => {
    element.monikerPath = [...ancestors, { element, moniker }].map((entry) => ({ nativeId: entry.element.nativeId, creationId: entry.element.creationId, moniker: entry.moniker }));
    return { element, targetId: element.id, nativeId: element.nativeId, creationId: element.creationId, moniker, monikerPath: element.monikerPath };
  });
}


function presentationElementIdentity(xml = "", moniker) {
  const cNvPr = /<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*>/.exec(String(xml))?.[0] || /<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\/\s*>/.exec(String(xml))?.[0] || "";
  const rawCreationId = localAttribute(/<(?:[A-Za-z_][\w.-]*:)?creationId\b[^>]*\/\s*>/.exec(String(xml))?.[0] || "", "id")?.toUpperCase();
  return {
    nativeId: /^\d+$/.test(localAttribute(cNvPr, "id") || "") ? Number(localAttribute(cNvPr, "id")) : undefined,
    creationId: rawCreationId && GUID.test(rawCreationId) ? rawCreationId : undefined,
    moniker,
  };
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
        anchor: (() => {
          const target = slide.resolve(thread.targetId);
          if (target?.kind === "textRange") {
            const parent = slide.resolve(target.parentId);
            if (parent?.nativeId && parent?.moniker === "spMk") {
              const preserved = thread.nativeAnchor?.type === "textRange" ? thread.nativeAnchor : {};
              return {
                type: "textRange",
                nativeId: parent.nativeId,
                creationId: parent.creationId,
                moniker: parent.moniker,
                slideId: slide.nativeSlideId,
                cp: preserved.cp ?? 0,
                length: preserved.length ?? String(target.text ?? "").length,
                contextLength: preserved.contextLength,
                contextHash: preserved.contextHash,
                monikers: parent.monikerPath,
              };
            }
          }
          if (target?.nativeId && target?.moniker) return { nativeId: target.nativeId, creationId: target.creationId, moniker: target.moniker, monikers: target.monikerPath, slideId: slide.nativeSlideId };
          return thread.nativeAnchor;
        })(),
        position: commentPosition(thread.position || { x: 100 + threadIndex * 32, y: 100 + threadIndex * 32 }, `PowerPoint modern comment ${comments[0].id} position`),
        root: comments[0],
        replies: comments.slice(1),
      });
    }
    if (roots.length) parts.push({ slideIndex, roots });
  }
  return { authors: [...authors.values()], parts };
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

function presentationSlideIds(bytesByPath) {
  const result = new Map();
  for (const [partPath] of bytesByPath) {
    if (!PRESENTATION_NAMESPACES.has(rootNamespace(packageXml(bytesByPath, partPath), "presentation"))) continue;
    const relsXml = packageXml(bytesByPath, relationshipsPartPath(partPath));
    const byId = new Map();
    for (const match of relsXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/\s*>/g)) {
      const attrs = attributes(match[0]);
      if (String(attrs.Type || "").endsWith("/slide") && String(attrs.TargetMode || "").toLowerCase() !== "external") byId.set(attrs.Id, resolveTarget(partPath, attrs.Target));
    }
    for (const match of packageXml(bytesByPath, partPath).matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*\/\s*>/g)) {
      const attrs = attributes(match[0]);
      const relationshipId = Object.entries(attrs).find(([name]) => name.endsWith(":id"))?.[1];
      const target = byId.get(relationshipId);
      const nativeSlideId = Number(attrs.id);
      if (target && Number.isInteger(nativeSlideId)) result.set(target, nativeSlideId);
    }
  }
  return result;
}

function slideElementIdentities(xml = "") {
  const kinds = { sp: "spMk", grpSp: "grpSpMk", graphicFrame: "graphicFrameMk", cxnSp: "cxnSpMk", pic: "picMk" };
  const walk = (source, parentLocalName, ancestors = []) => directPresentationChildren(source, parentLocalName).flatMap((child) => {
    const moniker = kinds[child.localName];
    if (!moniker) return [];
    const identity = { ...presentationElementIdentity(child.xml, moniker), moniker, plainTextLength: moniker === "spMk" ? officeArtPlainText(child.xml).length : undefined };
    const monikerPath = [...ancestors, identity].map((item) => ({ nativeId: item.nativeId, creationId: item.creationId, moniker: item.moniker }));
    const current = { ...identity, monikerPath };
    return child.localName === "grpSp" ? [current, ...walk(child.xml, "grpSp", [...ancestors, identity])] : [current];
  });
  return walk(String(xml), "spTree").filter((identity) => identity.nativeId);
}

function anchorMonikers(xml = "") {
  return [...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?(spMk|graphicFrameMk|cxnSpMk|picMk|grpSpMk|inkMk)\b[^>]*\/\s*>/g)].map((match) => ({
    nativeId: Number(localAttribute(match[0], "id")),
    creationId: localAttribute(match[0], "creationId")?.toUpperCase(),
    moniker: match[1],
  }));
}

function monikerPathMatches(candidate, requested) {
  return candidate.monikerPath?.length === requested.length && candidate.monikerPath.every((part, index) => {
    const target = requested[index];
    return part.moniker === target.moniker && part.nativeId === target.nativeId && (!target.creationId || part.creationId === target.creationId);
  });
}

function officeArtPlainText(xml = "") {
  const body = /<(?:[A-Za-z_][\w.-]*:)?txBody\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?txBody>/.exec(String(xml))?.[1] || "";
  const paragraphs = [...body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?p\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?p>/g)];
  return paragraphs.map((paragraph) => [...paragraph[1].matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>|<(?:[A-Za-z_][\w.-]*:)?(tab|br)\b[^>]*\/\s*>/g)]
    .map((token) => token[1] != null ? decodeXml(token[1]) : token[2] === "tab" ? "\t" : "\v")
    .join("")).join("\r");
}

function validateSlideMonikerChain(anchorXml, expectedSlideId, partPath, commentId) {
  const issues = [];
  const slideMoniker = /<(?:[A-Za-z_][\w.-]*:)?sldMk\b[^>]*\/\s*>/.exec(anchorXml)?.[0] || "";
  const slideId = Number(localAttribute(slideMoniker, "sldId"));
  if (!/<(?:[A-Za-z_][\w.-]*:)?docMk\b[^>]*\/\s*>/.test(anchorXml) || !slideMoniker || !Number.isInteger(slideId) || slideId < 256 || slideId >= 2_147_483_648) issues.push(issue("pptxModernCommentSlideMonikerInvalid", `PPTX modern comment ${commentId || "(missing)"} has an invalid slide moniker chain.`, { path: partPath, commentId, slideId }));
  else if (expectedSlideId && slideId !== expectedSlideId) issues.push(issue("pptxModernCommentSlideMonikerMismatch", `PPTX modern comment ${commentId || "(missing)"} slide moniker ${slideId} does not match relationship source slide ${expectedSlideId}.`, { path: partPath, commentId, slideId, expectedSlideId }));
  return issues;
}

function validateDrawingAnchor(rootXml, slideXml, expectedSlideId, partPath, commentId) {
  const issues = [];
  const anchor = /<(?:[A-Za-z_][\w.-]*:)?deMkLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?deMkLst>/.exec(rootXml);
  if (!anchor) return issues;
  issues.push(...validateSlideMonikerChain(anchor[1], expectedSlideId, partPath, commentId));
  const monikers = anchorMonikers(anchor[1]);
  const invalid = monikers.find((moniker) => !Number.isInteger(moniker.nativeId) || moniker.nativeId < 1 || moniker.nativeId > 4_294_967_295 || (moniker.creationId && !GUID.test(moniker.creationId)));
  const targetMoniker = monikers.at(-1);
  if (!targetMoniker || invalid) {
    issues.push(issue("pptxModernCommentDrawingMonikerInvalid", `PPTX modern comment ${commentId || "(missing)"} has an invalid drawing-element moniker path.`, { path: partPath, commentId, monikers }));
    return issues;
  }
  const identities = slideElementIdentities(slideXml);
  const target = identities.find((identity) => monikerPathMatches(identity, monikers));
  if (!target) {
    const creationResolved = targetMoniker.creationId ? identities.find((identity) => identity.moniker === targetMoniker.moniker && identity.creationId === targetMoniker.creationId) : undefined;
    if (creationResolved) issues.push(issue("pptxModernCommentDrawingIdentityMismatch", `PPTX modern comment ${commentId || "(missing)"} drawing ID, creationId, or ancestor moniker path resolve to different elements.`, { path: partPath, commentId, monikers, resolvedNativeId: creationResolved.nativeId }));
    else issues.push(issue("pptxModernCommentDrawingTargetNotFound", `PPTX modern comment ${commentId || "(missing)"} drawing anchor does not resolve in its source slide.`, { path: partPath, commentId, monikers }));
  }
  return issues;
}

function validateTextRangeAnchor(rootXml, slideXml, expectedSlideId, partPath, commentId) {
  const issues = [];
  const anchor = /<(?:[A-Za-z_][\w.-]*:)?txMkLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?txMkLst>/.exec(rootXml);
  if (!anchor) return issues;
  issues.push(...validateSlideMonikerChain(anchor[1], expectedSlideId, partPath, commentId));
  const monikers = anchorMonikers(anchor[1]);
  const shapeMoniker = monikers.at(-1);
  const invalid = monikers.find((moniker) => !Number.isInteger(moniker.nativeId) || moniker.nativeId < 1 || moniker.nativeId > 4_294_967_295 || (moniker.creationId && !GUID.test(moniker.creationId)));
  if (!shapeMoniker || shapeMoniker.moniker !== "spMk" || invalid) {
    issues.push(issue("pptxModernCommentTextTargetMonikerInvalid", `PPTX modern comment ${commentId || "(missing)"} has an invalid text-range shape moniker path.`, { path: partPath, commentId, monikers }));
    return issues;
  }
  const identities = slideElementIdentities(slideXml);
  const target = identities.find((identity) => monikerPathMatches(identity, monikers));
  if (!target) {
    const creationResolved = shapeMoniker.creationId ? identities.find((identity) => identity.moniker === "spMk" && identity.creationId === shapeMoniker.creationId) : undefined;
    if (creationResolved) issues.push(issue("pptxModernCommentTextIdentityMismatch", `PPTX modern comment ${commentId || "(missing)"} text-range shape identity or ancestor moniker path does not match.`, { path: partPath, commentId, monikers, resolvedNativeId: creationResolved.nativeId }));
    else issues.push(issue("pptxModernCommentTextTargetNotFound", `PPTX modern comment ${commentId || "(missing)"} text-range anchor does not resolve to a shape in its source slide.`, { path: partPath, commentId, monikers }));
    return issues;
  }
  const rangeMoniker = textRangeMonikerXml(anchor[1]);
  const rangeOpening = rootTag(rangeMoniker, "txMk")?.[0] || "";
  const cp = Number(localAttribute(rangeOpening, "cp"));
  const length = Number(localAttribute(rangeOpening, "len") || 0);
  if (!rangeMoniker || !Number.isInteger(cp) || cp < 0 || cp > 2_147_483_647 || !Number.isInteger(length) || length < 0 || length > 2_147_483_647 - cp) {
    issues.push(issue("pptxModernCommentTextRangeMonikerInvalid", `PPTX modern comment ${commentId || "(missing)"} has an invalid text character-range moniker.`, { path: partPath, commentId, cp, length }));
  } else if (cp + length > target.plainTextLength) {
    issues.push(issue("pptxModernCommentTextRangeOutOfBounds", `PPTX modern comment ${commentId || "(missing)"} text range ${cp}:${cp + length} exceeds its shape text length ${target.plainTextLength}.`, { path: partPath, commentId, cp, length, textLength: target.plainTextLength }));
  }
  const context = rootTag(rangeMoniker, "context")?.[0] || "";
  if (context) {
    const contextHash = Number(localAttribute(context, "hash"));
    const contextLengthValue = localAttribute(context, "len");
    const contextLength = contextLengthValue == null ? undefined : Number(contextLengthValue);
    if (!Number.isInteger(contextHash) || contextHash < 0 || contextHash > 4_294_967_295 || (contextLength != null && (!Number.isInteger(contextLength) || contextLength < 0 || contextLength > 4_294_967_295))) issues.push(issue("pptxModernCommentTextRangeContextInvalid", `PPTX modern comment ${commentId || "(missing)"} has an invalid text-range context hash or length.`, { path: partPath, commentId, contextHash, contextLength }));
  }
  return issues;
}

function validateSlideElementIdentityUniqueness(slideXml, source) {
  const issues = [];
  const nativeIds = new Set();
  const creationIds = new Set();
  for (const identity of slideElementIdentities(slideXml)) {
    if (nativeIds.has(identity.nativeId)) issues.push(issue("pptxModernDrawingNativeIdDuplicate", `PPTX slide ${source} contains duplicate drawing native ID ${identity.nativeId}.`, { path: source, nativeId: identity.nativeId }));
    nativeIds.add(identity.nativeId);
    if (!identity.creationId) continue;
    if (creationIds.has(identity.creationId)) issues.push(issue("pptxModernDrawingCreationIdDuplicate", `PPTX slide ${source} contains duplicate drawing creationId ${identity.creationId}.`, { path: source, creationId: identity.creationId }));
    creationIds.add(identity.creationId);
  }
  return issues;
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
  const slideIds = presentationSlideIds(bytesByPath);
  const validatedSlideIdentities = new Set();
  for (const rel of commentRels) {
    commentCountBySource.set(rel.source, (commentCountBySource.get(rel.source) || 0) + 1);
    if (rel.external) issues.push(issue("pptxModernCommentRelationshipExternal", "PPTX modern comment relationship must be internal.", { path: rel.path, relationshipId: rel.id }));
    if (!PRESENTATION_NAMESPACES.has(rootNamespace(packageXml(bytesByPath, rel.source), "sld"))) issues.push(issue("pptxModernCommentRelationshipSourceInvalid", "PPTX modern comment relationship must originate from a Slide part.", { path: rel.path, source: rel.source }));
    if (!rel.target || !bytesByPath.has(rel.target)) continue;
    if (!validatedSlideIdentities.has(rel.source)) {
      issues.push(...validateSlideElementIdentityUniqueness(packageXml(bytesByPath, rel.source), rel.source));
      validatedSlideIdentities.add(rel.source);
    }
    if (hasOutgoingRelationships(bytesByPath, rel.target)) issues.push(issue("pptxModernCommentPartRelationshipsForbidden", `PPTX modern comment part ${rel.target} must not own package relationships.`, { path: relationshipsPartPath(rel.target), source: rel.target }));
    if (declaredContentType(contentTypes, rel.target) !== PPTX_MODERN_COMMENT_CONTENT_TYPE) issues.push(issue("pptxModernCommentContentTypeInvalid", `PPTX modern comment part ${rel.target} has the wrong content type.`, { path: rel.target }));
    const xml = packageXml(bytesByPath, rel.target);
    if (rootNamespace(xml, "cmLst") !== PPTX_MODERN_COMMENT_NAMESPACE) issues.push(issue("pptxModernCommentRootInvalid", `PPTX modern comment part ${rel.target} must have a cmLst root in the Office 2021 namespace.`, { path: rel.target }));
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cm\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?cm>/g)) {
      const source = match[0];
      const repliesXml = /<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/.exec(source)?.[1] || "";
      const rootXml = source.replace(/<(?:[A-Za-z_][\w.-]*:)?replyLst\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?replyLst>/, "");
      const root = parseCommentElement(rootXml, authors);
      const anchorNames = [...rootXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(sldMkLst|sldLayoutMkLst|sldMasterMkLst|deMkLst|txBodyMkLst|txMkLst|tcMkLst|trMkLst|gridColMkLst|unknownAnchor)\b/g)].map((match) => match[1]);
      const contentMonikerLists = new Set(["deMkLst", "txBodyMkLst", "txMkLst", "tcMkLst", "trMkLst", "gridColMkLst"]);
      const hasContentMonikerList = anchorNames.some((name) => contentMonikerLists.has(name));
      root.anchorCount = anchorNames.filter((name) => name !== "sldMkLst" || !hasContentMonikerList).length;
      const position = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\/?\s*>/.exec(rootXml)?.[0];
      if (position) root.position = { x: Number(localAttribute(position, "x")), y: Number(localAttribute(position, "y")) };
      issues.push(...validateCommentNode(root, authors, ids, rel.target, true));
      issues.push(...validateDrawingAnchor(rootXml, packageXml(bytesByPath, rel.source), slideIds.get(rel.source), rel.target, root.nativeId));
      issues.push(...validateTextRangeAnchor(rootXml, packageXml(bytesByPath, rel.source), slideIds.get(rel.source), rel.target, root.nativeId));
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
