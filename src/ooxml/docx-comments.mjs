import path from "node:path";

import { attributes, rootTag } from "./source-reference-xml.mjs";

export const DOCX_COMMENTS_EXTENDED_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml";
export const DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/commentsExtended";
export const DOCX_COMMENTS_IDS_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml";
export const DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2016/09/relationships/commentsIds";
export const DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtensible+xml";
export const DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2018/08/relationships/commentsExtensible";
export const DOCX_PEOPLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.people+xml";
export const DOCX_PEOPLE_RELATIONSHIP_TYPE = "http://schemas.microsoft.com/office/2011/relationships/people";

const WORD_2012_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const WORD_2016_COMMENT_ID_NAMESPACE = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const WORD_2018_COMMENT_EXTENSIBLE_NAMESPACE = "http://schemas.microsoft.com/office/word/2018/wordml/cex";
const PARA_ID = /^[0-9A-F]{8}$/;
const decoder = new TextDecoder();

function localAttribute(tag, localName) {
  return Object.entries(attributes(tag)).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
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

function relationshipsOfType(bytesByPath, relationshipType) {
  const entries = [];
  for (const [partPath] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = relationshipSource(partPath);
    if (source == null) continue;
    for (const match of packageXml(bytesByPath, partPath).matchAll(/<Relationship\b[^>]*\/?\s*>/g)) {
      const attrs = attributes(match[0]);
      if (attrs.Type !== relationshipType) continue;
      entries.push({ source, path: partPath, id: attrs.Id, target: attrs.Target ? resolveTarget(source, attrs.Target) : undefined, external: String(attrs.TargetMode || "").toLowerCase() === "external" });
    }
  }
  return entries;
}

function relatedTarget(bytesByPath, source, relationshipType) {
  return relationshipsOfType(bytesByPath, relationshipType).find((entry) => entry.source === source && !entry.external)?.target;
}

function partRootNamespace(xml, localName) {
  const root = rootTag(xml, localName);
  const attrs = attributes(root?.[0] || "");
  return root ? (root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns) : undefined;
}

function relationshipEnvelopeIssues(bytesByPath, contentTypes, entry, config) {
  const issues = [];
  if (entry.external) issues.push(issue(`${config.prefix}RelationshipExternal`, `DOCX ${config.label} relationship must be internal.`, { path: entry.path, relationshipId: entry.id }));
  const sourceXml = packageXml(bytesByPath, entry.source);
  if (!rootTag(sourceXml, "document") && !rootTag(sourceXml, "glossaryDocument")) issues.push(issue(`${config.prefix}SourceInvalid`, `DOCX ${config.label} relationship must originate from a main or glossary document part.`, { path: entry.path, source: entry.source, relationshipId: entry.id }));
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const actualContentType = declaredContentType(contentTypes, entry.target);
  if (actualContentType !== config.contentType) issues.push(issue(`${config.prefix}ContentTypeInvalid`, `DOCX ${config.label} part ${entry.target} must use ${config.contentType}.`, { path: entry.target, contentType: actualContentType, expectedContentType: config.contentType }));
  const namespace = partRootNamespace(packageXml(bytesByPath, entry.target), config.root);
  if (namespace !== config.namespace) issues.push(issue(`${config.prefix}RootInvalid`, `DOCX ${config.label} part ${entry.target} must have a ${config.root} root in ${config.namespace}.`, { path: entry.target, namespace }));
  return issues;
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
    if (!validParagraphId(paraId)) issues.push(issue("docxCommentParaIdMissing", `DOCX comment in ${partPath} requires a paraId from 00000001 through 7FFFFFFF when commentsExtended is present.`, { path: partPath, paraId }));
    else {
      if (paraIds.has(paraId)) issues.push(issue("docxCommentParaIdDuplicate", `DOCX comments part ${partPath} contains duplicate paraId ${paraId}.`, { path: partPath, paraId }));
      paraIds.add(paraId);
    }
  }
  return { paraIds, issues };
}

function validateExtendedPart(bytesByPath, contentTypes, entry) {
  const issues = relationshipEnvelopeIssues(bytesByPath, contentTypes, entry, { prefix: "docxCommentsExtended", label: "commentsExtended", contentType: DOCX_COMMENTS_EXTENDED_CONTENT_TYPE, root: "commentsEx", namespace: WORD_2012_NAMESPACE });
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const extendedXml = packageXml(bytesByPath, entry.target);
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
    if (!validParagraphId(paraId)) {
      issues.push(issue("docxCommentExParaIdInvalid", `DOCX commentEx in ${entry.target} requires a paraId from 00000001 through 7FFFFFFF.`, { path: entry.target, paraId }));
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

function validDurableId(value) {
  return PARA_ID.test(String(value || "")) && Number.parseInt(value, 16) > 0 && Number.parseInt(value, 16) < 0x7FFFFFFF;
}

function validParagraphId(value) {
  return PARA_ID.test(String(value || "")) && Number.parseInt(value, 16) > 0 && Number.parseInt(value, 16) < 0x80000000;
}

function validateCommentsIdsPart(bytesByPath, contentTypes, entry) {
  const issues = relationshipEnvelopeIssues(bytesByPath, contentTypes, entry, { prefix: "docxCommentsIds", label: "commentsIds", contentType: DOCX_COMMENTS_IDS_CONTENT_TYPE, root: "commentsIds", namespace: WORD_2016_COMMENT_ID_NAMESPACE });
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const commentsTarget = classicCommentsTarget(bytesByPath, entry.source);
  if (!commentsTarget || !bytesByPath.has(commentsTarget)) {
    issues.push(issue("docxCommentsIdsClassicCommentsMissing", "DOCX commentsIds requires a classic Comments part related from the same document part.", { path: entry.target, source: entry.source }));
    return issues;
  }
  const classic = classicCommentParaIds(packageXml(bytesByPath, commentsTarget), commentsTarget);
  issues.push(...classic.issues);
  const mappedParaIds = new Set();
  const durableIds = new Set();
  for (const match of packageXml(bytesByPath, entry.target).matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentId\b[^>]*\/?\s*>/g)) {
    const paraId = localAttribute(match[0], "paraId")?.toUpperCase();
    const durableId = localAttribute(match[0], "durableId")?.toUpperCase();
    if (!validParagraphId(paraId)) issues.push(issue("docxCommentIdParaIdInvalid", `DOCX commentId in ${entry.target} requires a paraId from 00000001 through 7FFFFFFF.`, { path: entry.target, paraId }));
    else {
      if (mappedParaIds.has(paraId)) issues.push(issue("docxCommentIdParaIdDuplicate", `DOCX commentsIds part ${entry.target} contains duplicate paraId ${paraId}.`, { path: entry.target, paraId }));
      mappedParaIds.add(paraId);
      if (!classic.paraIds.has(paraId)) issues.push(issue("docxCommentIdReferenceNotFound", `DOCX commentId paraId ${paraId} does not resolve to a classic comment paragraph.`, { path: entry.target, paraId }));
    }
    if (!validDurableId(durableId)) issues.push(issue("docxCommentDurableIdInvalid", `DOCX commentId in ${entry.target} requires a durableId greater than 0 and less than 7FFFFFFF.`, { path: entry.target, durableId }));
    else {
      if (durableIds.has(durableId)) issues.push(issue("docxCommentDurableIdDuplicate", `DOCX commentsIds part ${entry.target} contains duplicate durableId ${durableId}.`, { path: entry.target, durableId }));
      durableIds.add(durableId);
    }
  }
  for (const paraId of classic.paraIds) if (!mappedParaIds.has(paraId)) issues.push(issue("docxCommentIdMappingMissing", `DOCX commentsIds part ${entry.target} does not identify classic comment paraId ${paraId}.`, { path: entry.target, paraId }));
  return issues;
}

function commentsIdsMaps(bytesByPath, source) {
  const target = relatedTarget(bytesByPath, source, DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE);
  const durableToPara = new Map();
  const paraToDurable = new Map();
  for (const match of packageXml(bytesByPath, target).matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentId\b[^>]*\/?\s*>/g)) {
    const paraId = localAttribute(match[0], "paraId")?.toUpperCase();
    const durableId = localAttribute(match[0], "durableId")?.toUpperCase();
    if (paraId && durableId) { durableToPara.set(durableId, paraId); paraToDurable.set(paraId, durableId); }
  }
  return { target, durableToPara, paraToDurable };
}

function replyParaIds(bytesByPath, source) {
  const target = relatedTarget(bytesByPath, source, DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE);
  const result = new Set();
  for (const match of packageXml(bytesByPath, target).matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentEx\b[^>]*\/?\s*>/g)) if (localAttribute(match[0], "paraIdParent")) result.add(localAttribute(match[0], "paraId")?.toUpperCase());
  return result;
}

function validateCommentsExtensiblePart(bytesByPath, contentTypes, entry) {
  const issues = relationshipEnvelopeIssues(bytesByPath, contentTypes, entry, { prefix: "docxCommentsExtensible", label: "commentsExtensible", contentType: DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE, root: "commentsExtensible", namespace: WORD_2018_COMMENT_EXTENSIBLE_NAMESPACE });
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const ids = commentsIdsMaps(bytesByPath, entry.source);
  if (!ids.target || !bytesByPath.has(ids.target)) {
    issues.push(issue("docxCommentsExtensibleCommentsIdsMissing", "DOCX commentsExtensible requires a commentsIds part related from the same document part.", { path: entry.target, source: entry.source }));
    return issues;
  }
  const replies = replyParaIds(bytesByPath, entry.source);
  const seen = new Set();
  for (const match of packageXml(bytesByPath, entry.target).matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentExtensible\b[^>]*\/?\s*>/g)) {
    const durableId = localAttribute(match[0], "durableId")?.toUpperCase();
    const dateUtc = localAttribute(match[0], "dateUtc");
    const intelligent = String(localAttribute(match[0], "intelligentPlaceholder") || "0").trim().toLowerCase();
    const intelligentPlaceholder = new Set(["1", "true", "on"]).has(intelligent);
    if (!validDurableId(durableId)) issues.push(issue("docxCommentExtensibleDurableIdInvalid", `DOCX commentExtensible in ${entry.target} requires a valid durableId.`, { path: entry.target, durableId }));
    else {
      if (seen.has(durableId)) issues.push(issue("docxCommentExtensibleDurableIdDuplicate", `DOCX commentsExtensible part ${entry.target} contains duplicate durableId ${durableId}.`, { path: entry.target, durableId }));
      seen.add(durableId);
      if (!ids.durableToPara.has(durableId)) issues.push(issue("docxCommentExtensibleReferenceNotFound", `DOCX commentExtensible durableId ${durableId} does not resolve through commentsIds.`, { path: entry.target, durableId }));
      if (intelligentPlaceholder && replies.has(ids.durableToPara.get(durableId))) issues.push(issue("docxCommentExtensibleReplyPlaceholderInvalid", `DOCX reply ${durableId} must not be an intelligent placeholder.`, { path: entry.target, durableId }));
    }
    if (dateUtc && (Number.isNaN(Date.parse(dateUtc)) || !/(?:Z|[+-]00:00)$/i.test(dateUtc))) issues.push(issue("docxCommentExtensibleDateUtcInvalid", `DOCX commentExtensible ${durableId || "(missing)"} dateUtc must be a valid UTC date-time.`, { path: entry.target, durableId, dateUtc }));
  }
  return issues;
}

function usedDocumentAuthors(bytesByPath, source) {
  const authors = new Set();
  const commentsTarget = classicCommentsTarget(bytesByPath, source);
  for (const xml of [packageXml(bytesByPath, source), packageXml(bytesByPath, commentsTarget)]) {
    for (const match of xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:comment|ins|del|moveFrom|moveTo)\b[^>]*>/g)) {
      const author = localAttribute(match[0], "author");
      if (author) authors.add(author);
    }
  }
  return authors;
}

function validatePeoplePart(bytesByPath, contentTypes, entry) {
  const issues = relationshipEnvelopeIssues(bytesByPath, contentTypes, entry, { prefix: "docxPeople", label: "people", contentType: DOCX_PEOPLE_CONTENT_TYPE, root: "people", namespace: WORD_2012_NAMESPACE });
  if (!entry.target || !bytesByPath.has(entry.target)) return issues;
  const authors = usedDocumentAuthors(bytesByPath, entry.source);
  const seen = new Set();
  const xml = packageXml(bytesByPath, entry.target);
  const people = [...xml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?person\b[^>]*\/>|<(?:[A-Za-z_][\w.-]*:)?person\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?person>/g)];
  for (const match of people) {
    const opening = /^<(?:[A-Za-z_][\w.-]*:)?person\b[^>]*[>]/.exec(match[0])?.[0] || "";
    const author = localAttribute(opening, "author");
    if (!author) issues.push(issue("docxPersonAuthorMissing", `DOCX person in ${entry.target} requires an author.`, { path: entry.target }));
    else {
      if (seen.has(author)) issues.push(issue("docxPersonAuthorDuplicate", `DOCX people part ${entry.target} contains duplicate author ${author}.`, { path: entry.target, author }));
      seen.add(author);
      if (!authors.has(author)) issues.push(issue("docxPersonAuthorNotFound", `DOCX person author ${author} does not match a comment or revision author.`, { path: entry.target, author }));
    }
    const presence = /<(?:[A-Za-z_][\w.-]*:)?presenceInfo\b[^>]*\/?\s*>/.exec(match[0])?.[0];
    if (!presence) continue;
    const providerId = localAttribute(presence, "providerId");
    const userId = localAttribute(presence, "userId");
    if (!providerId || providerId.length > 100) issues.push(issue("docxPersonProviderIdInvalid", `DOCX person ${author || "(missing)"} providerId must contain 1 to 100 characters.`, { path: entry.target, author, providerId }));
    if (!userId || userId.length > 300) issues.push(issue("docxPersonUserIdInvalid", `DOCX person ${author || "(missing)"} userId must contain 1 to 300 characters.`, { path: entry.target, author, userId }));
  }
  return issues;
}

const COMMENT_PART_CONFIGS = [
  { relationshipType: DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE, contentType: DOCX_COMMENTS_EXTENDED_CONTENT_TYPE, prefix: "docxCommentsExtended", label: "commentsExtended", validate: validateExtendedPart },
  { relationshipType: DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE, contentType: DOCX_COMMENTS_IDS_CONTENT_TYPE, prefix: "docxCommentsIds", label: "commentsIds", validate: validateCommentsIdsPart },
  { relationshipType: DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE, contentType: DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE, prefix: "docxCommentsExtensible", label: "commentsExtensible", validate: validateCommentsExtensiblePart },
  { relationshipType: DOCX_PEOPLE_RELATIONSHIP_TYPE, contentType: DOCX_PEOPLE_CONTENT_TYPE, prefix: "docxPeople", label: "people", validate: validatePeoplePart },
];

export function validateDocxCommentPackageSemantics({ bytesByPath, contentTypes }) {
  const issues = [];
  for (const config of COMMENT_PART_CONFIGS) {
    const entries = relationshipsOfType(bytesByPath, config.relationshipType);
    const counts = new Map();
    for (const entry of entries) {
      counts.set(entry.source, (counts.get(entry.source) || 0) + 1);
      issues.push(...config.validate(bytesByPath, contentTypes, entry));
    }
    for (const [source, count] of counts) if (count > 1) issues.push(issue(`${config.prefix}RelationshipMultiplicity`, `DOCX document part ${source} has ${count} ${config.label} relationships; at most one is allowed.`, { source, count, maximum: 1 }));
    const targets = new Set(entries.filter((entry) => !entry.external).map((entry) => entry.target));
    for (const partPath of bytesByPath.keys()) if (declaredContentType(contentTypes, partPath) === config.contentType && !targets.has(partPath)) issues.push(issue(`${config.prefix}PartOrphaned`, `DOCX ${config.label} part ${partPath} is not targeted by its required relationship.`, { path: partPath }));
  }
  return issues;
}
