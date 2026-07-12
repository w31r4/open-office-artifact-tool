import path from "node:path";

import { attributes, decodeXml, rootTag } from "./source-reference-xml.mjs";

const PRESENTATION_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/presentationml/2006/main",
  "http://purl.oclc.org/ooxml/presentationml/main",
]);

const CONTENT_TYPES = {
  notesSlide: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml",
  comments: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml",
  commentAuthors: "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml",
};

const RELATIONSHIP_KINDS = {
  notesSlide: { sourceRoot: "sld", targetRoot: "notes", maximumPerSource: 1 },
  comments: { sourceRoot: "sld", targetRoot: "cmLst", maximumPerSource: 1 },
  commentAuthors: { sourceRoot: "presentation", targetRoot: "cmAuthorLst", maximumPerSource: 1 },
};

const decoder = new TextDecoder();

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
  const sourceDir = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDir === "." ? "" : sourceDir, target));
}

function xml(bytesByPath, partPath) {
  const bytes = bytesByPath.get(partPath);
  return bytes ? decoder.decode(bytes) : "";
}

function rootNamespace(sourceXml, localName) {
  const root = rootTag(sourceXml, localName);
  if (!root) return undefined;
  const attrs = attributes(root[0]);
  return root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns;
}

function hasPresentationRoot(sourceXml, localName) {
  return PRESENTATION_NAMESPACES.has(rootNamespace(sourceXml, localName));
}

function relationshipKind(type) {
  return Object.keys(RELATIONSHIP_KINDS).find((kind) => String(type || "").endsWith(`/${kind}`));
}

function relationships(bytesByPath) {
  const result = [];
  for (const [partPath] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = relationshipSource(partPath);
    if (source == null) continue;
    for (const match of xml(bytesByPath, partPath).matchAll(/<Relationship\b[^>]*\/?\s*>/g)) {
      const attrs = attributes(match[0]);
      const kind = relationshipKind(attrs.Type);
      if (!kind) continue;
      result.push({ kind, source, path: partPath, id: attrs.Id, target: attrs.Target ? resolveTarget(source, attrs.Target) : undefined, external: String(attrs.TargetMode || "").toLowerCase() === "external" });
    }
  }
  return result;
}

function declaredContentType(contentTypes, partPath) {
  const extension = path.posix.extname(partPath).slice(1).toLowerCase();
  return contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension);
}

function unsignedInteger(value, minimum = 0) {
  return /^\d+$/.test(String(value ?? "")) && Number(value) >= minimum && Number(value) <= 4_294_967_295;
}

function validateRelationshipShape(entries, bytesByPath, contentTypes) {
  const issues = [];
  const bySourceKind = new Map();
  for (const entry of entries) {
    const spec = RELATIONSHIP_KINDS[entry.kind];
    const key = `${entry.source}\u0000${entry.kind}`;
    bySourceKind.set(key, (bySourceKind.get(key) || 0) + 1);
    if (entry.external) {
      issues.push(issue("pptxSemanticRelationshipExternal", `PPTX ${entry.kind} relationship ${entry.id || "(unknown)"} must be internal.`, { path: entry.path, relationshipId: entry.id, relationshipKind: entry.kind }));
      continue;
    }
    if (!hasPresentationRoot(xml(bytesByPath, entry.source), spec.sourceRoot)) {
      issues.push(issue("pptxSemanticRelationshipSourceInvalid", `PPTX ${entry.kind} relationship must originate from a ${spec.sourceRoot} part.`, { path: entry.path, source: entry.source, relationshipId: entry.id, relationshipKind: entry.kind }));
    }
    if (!entry.target || !bytesByPath.has(entry.target)) continue;
    if (!hasPresentationRoot(xml(bytesByPath, entry.target), spec.targetRoot)) {
      issues.push(issue("pptxSemanticPartRootInvalid", `PPTX ${entry.kind} part ${entry.target} must have a ${spec.targetRoot} root in a PresentationML namespace.`, { path: entry.target, relationshipId: entry.id, relationshipKind: entry.kind, expectedRoot: spec.targetRoot }));
    }
    const actualContentType = declaredContentType(contentTypes, entry.target);
    if (actualContentType !== CONTENT_TYPES[entry.kind]) {
      issues.push(issue("pptxSemanticContentTypeInvalid", `PPTX ${entry.kind} part ${entry.target} must use content type ${CONTENT_TYPES[entry.kind]}.`, { path: entry.target, relationshipId: entry.id, relationshipKind: entry.kind, contentType: actualContentType, expectedContentType: CONTENT_TYPES[entry.kind] }));
    }
  }
  for (const [key, count] of bySourceKind) {
    const [source, kind] = key.split("\u0000");
    const maximum = RELATIONSHIP_KINDS[kind].maximumPerSource;
    if (count > maximum) issues.push(issue("pptxSemanticRelationshipMultiplicity", `PPTX ${source} has ${count} ${kind} relationships; at most ${maximum} is allowed.`, { source, relationshipKind: kind, count, maximum }));
  }
  return issues;
}

function typedParts(bytesByPath, contentTypes, kind) {
  return [...bytesByPath.keys()].filter((partPath) => declaredContentType(contentTypes, partPath) === CONTENT_TYPES[kind]);
}

function validatePartReachability(entries, bytesByPath, contentTypes) {
  const issues = [];
  for (const kind of Object.keys(CONTENT_TYPES)) {
    const parts = typedParts(bytesByPath, contentTypes, kind);
    if (kind === "commentAuthors" && parts.length > 1) issues.push(issue("pptxCommentAuthorsMultiplicity", `PPTX package contains ${parts.length} Comment Authors parts; at most one is allowed.`, { paths: parts, count: parts.length, maximum: 1 }));
    const targets = new Set(entries.filter((entry) => entry.kind === kind && !entry.external).map((entry) => entry.target));
    for (const partPath of parts) {
      if (!targets.has(partPath)) issues.push(issue("pptxSemanticPartOrphaned", `PPTX ${kind} part ${partPath} is not targeted by its required relationship.`, { path: partPath, relationshipKind: kind }));
    }
  }
  return issues;
}

function parseAuthors(authorXml, partPath) {
  const issues = [];
  const authors = new Map();
  for (const match of String(authorXml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?cmAuthor\b[^>]*\/?>/g)) {
    const attrs = attributes(match[0]);
    const id = String(attrs.id ?? "");
    if (!unsignedInteger(id)) issues.push(issue("pptxCommentAuthorAttributeInvalid", `PPTX comment author in ${partPath} has an invalid unsigned id.`, { path: partPath, attribute: "id", value: attrs.id }));
    else if (authors.has(id)) issues.push(issue("pptxCommentAuthorDuplicateId", `PPTX Comment Authors part ${partPath} contains duplicate author id ${id}.`, { path: partPath, authorId: id }));
    for (const name of ["name", "initials"]) if (!String(attrs[name] ?? "").length) issues.push(issue("pptxCommentAuthorAttributeInvalid", `PPTX comment author ${id || "(unknown)"} in ${partPath} requires ${name}.`, { path: partPath, authorId: id || undefined, attribute: name }));
    for (const name of ["lastIdx", "clrIdx"]) if (!unsignedInteger(attrs[name])) issues.push(issue("pptxCommentAuthorAttributeInvalid", `PPTX comment author ${id || "(unknown)"} in ${partPath} has an invalid unsigned ${name}.`, { path: partPath, authorId: id || undefined, attribute: name, value: attrs[name] }));
    if (unsignedInteger(id) && !authors.has(id)) authors.set(id, { id, lastIndex: unsignedInteger(attrs.lastIdx) ? Number(attrs.lastIdx) : undefined, partPath });
  }
  return { authors, issues };
}

function validateComments(entries, bytesByPath, contentTypes) {
  const issues = [];
  const authorParts = typedParts(bytesByPath, contentTypes, "commentAuthors");
  const authorPart = authorParts[0];
  const parsed = authorPart ? parseAuthors(xml(bytesByPath, authorPart), authorPart) : { authors: new Map(), issues: [] };
  issues.push(...parsed.issues);
  const commentParts = typedParts(bytesByPath, contentTypes, "comments");
  if (commentParts.length && !authorPart) issues.push(issue("pptxCommentAuthorsMissing", "PPTX comments require a Comment Authors part related from the Presentation part.", { commentParts }));
  const seenIndexes = new Set();
  const maximumIndexByAuthor = new Map();
  for (const partPath of commentParts) {
    const sourceXml = xml(bytesByPath, partPath);
    for (const match of sourceXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cm\b([^>]*)>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?cm>/g)) {
      const attrs = attributes(`<cm ${match[1]}>`);
      const authorId = String(attrs.authorId ?? "");
      const index = String(attrs.idx ?? "");
      if (!unsignedInteger(authorId)) issues.push(issue("pptxCommentAuthorIdInvalid", `PPTX comment in ${partPath} has an invalid unsigned authorId.`, { path: partPath, authorId: attrs.authorId }));
      else if (!parsed.authors.has(authorId)) issues.push(issue("pptxCommentAuthorReferenceNotFound", `PPTX comment in ${partPath} references missing authorId ${authorId}.`, { path: partPath, authorId }));
      if (!unsignedInteger(index, 1)) issues.push(issue("pptxCommentIndexInvalid", `PPTX comment in ${partPath} has invalid idx ${index || "(missing)"}; indexes start at 1.`, { path: partPath, authorId: authorId || undefined, commentIndex: attrs.idx }));
      if (unsignedInteger(authorId) && unsignedInteger(index, 1)) {
        const key = `${authorId}\u0000${index}`;
        if (seenIndexes.has(key)) issues.push(issue("pptxCommentIndexDuplicate", `PPTX comment index ${index} is duplicated for authorId ${authorId}.`, { path: partPath, authorId, commentIndex: Number(index) }));
        seenIndexes.add(key);
        maximumIndexByAuthor.set(authorId, Math.max(maximumIndexByAuthor.get(authorId) || 0, Number(index)));
      }
      if (!String(attrs.dt ?? "").length || Number.isNaN(Date.parse(attrs.dt))) issues.push(issue("pptxCommentDateInvalid", `PPTX comment in ${partPath} requires a valid XML date-time in dt.`, { path: partPath, authorId: authorId || undefined, value: attrs.dt }));
      const body = match[2];
      const position = /<(?:[A-Za-z_][\w.-]*:)?pos\b[^>]*\/?>/.exec(body)?.[0];
      const positionAttrs = attributes(position || "");
      if (!position || !/^-?\d+$/.test(String(positionAttrs.x ?? "")) || !/^-?\d+$/.test(String(positionAttrs.y ?? ""))) issues.push(issue("pptxCommentPositionInvalid", `PPTX comment in ${partPath} requires a pos child with integer x and y.`, { path: partPath, authorId: authorId || undefined, commentIndex: unsignedInteger(index, 1) ? Number(index) : undefined }));
      if (!/<(?:[A-Za-z_][\w.-]*:)?text\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?text>/.test(body)) issues.push(issue("pptxCommentTextMissing", `PPTX comment in ${partPath} requires a text child.`, { path: partPath, authorId: authorId || undefined, commentIndex: unsignedInteger(index, 1) ? Number(index) : undefined }));
    }
  }
  for (const [authorId, maximumIndex] of maximumIndexByAuthor) {
    const author = parsed.authors.get(authorId);
    if (author?.lastIndex != null && author.lastIndex < maximumIndex) issues.push(issue("pptxCommentLastIndexTooSmall", `PPTX comment author ${authorId} lastIdx ${author.lastIndex} is smaller than used comment index ${maximumIndex}.`, { path: author.partPath, authorId, lastIndex: author.lastIndex, maximumCommentIndex: maximumIndex }));
  }
  return issues;
}

function validateNotes(bytesByPath, contentTypes) {
  const issues = [];
  for (const partPath of typedParts(bytesByPath, contentTypes, "notesSlide")) {
    if (!/<(?:[A-Za-z_][\w.-]*:)?cSld\b/.test(xml(bytesByPath, partPath))) issues.push(issue("pptxNotesCommonSlideDataMissing", `PPTX Notes Slide part ${partPath} requires a cSld child.`, { path: partPath }));
  }
  return issues;
}

export function validatePptxPackageSemantics({ bytesByPath, contentTypes }) {
  const entries = relationships(bytesByPath);
  return [
    ...validateRelationshipShape(entries, bytesByPath, contentTypes),
    ...validatePartReachability(entries, bytesByPath, contentTypes),
    ...validateComments(entries, bytesByPath, contentTypes),
    ...validateNotes(bytesByPath, contentTypes),
  ];
}
