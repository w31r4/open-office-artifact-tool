import path from "node:path";

import JSZip from "jszip";

import {
  DOCX_COMMENTS_EXTENDED_CONTENT_TYPE,
  DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE,
  DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE,
  DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE,
  DOCX_COMMENTS_IDS_CONTENT_TYPE,
  DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE,
  DOCX_PEOPLE_CONTENT_TYPE,
  DOCX_PEOPLE_RELATIONSHIP_TYPE,
} from "./docx-comments.mjs";
import {
  mutateOoxmlSourceReference,
  mutateOoxmlSourceReferenceTarget,
  supportedOoxmlSourceReferenceSummary,
  supportsOoxmlSourceReference,
  validateOoxmlSourceReferenceTarget,
} from "./source-references.mjs";
import {
  PPTX_MODERN_AUTHOR_CONTENT_TYPE,
  PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE,
  PPTX_MODERN_COMMENT_CONTENT_TYPE,
  PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE,
} from "../presentation/ooxml-modern-comments.mjs";
import { decoder, encoder, toUint8Array } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { imageContentTypeFromExtension } from "../shared/images.mjs";
import { ndjson } from "../shared/inspection.mjs";
import { attrEscape } from "../shared/xml.mjs";

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

export function ooxmlSafePartPath(partPath, family = "OOXML") {
  const raw = String(partPath || "").replaceAll("\\", "/").trim();
  if (!raw || raw.startsWith("/") || raw.includes("\0")) throw new Error(`Unsafe ${family} part path: ${partPath}`);
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized.startsWith("../") || normalized.includes("/../") || normalized === "..") throw new Error(`Unsafe ${family} part path: ${partPath}`);
  if (normalized.length > 1024) throw new Error(`Unsafe ${family} part path: path exceeds 1024 characters`);
  return normalized;
}

function ooxmlXmlAttributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function ooxmlContentTypeMaps(xml = "") {
  const defaults = new Map();
  const overrides = new Map();
  for (const match of String(xml).matchAll(/<Default\b[^>]*\/?\s*>/g)) {
    const attrs = ooxmlXmlAttributes(match[0]);
    if (attrs.Extension && attrs.ContentType) defaults.set(String(attrs.Extension).toLowerCase(), attrs.ContentType);
  }
  for (const match of String(xml).matchAll(/<Override\b[^>]*\/?\s*>/g)) {
    const attrs = ooxmlXmlAttributes(match[0]);
    if (attrs.PartName && attrs.ContentType) overrides.set(String(attrs.PartName).replace(/^\//, ""), attrs.ContentType);
  }
  return { defaults, overrides };
}

function ooxmlFallbackContentType(partPath) {
  if (partPath.endsWith(".rels")) return "application/vnd.openxmlformats-package.relationships+xml";
  if (partPath.endsWith(".xml")) return "application/xml";
  if (partPath.endsWith(".json")) return "application/json";
  return imageContentTypeFromExtension(path.posix.extname(partPath).slice(1));
}

function ooxmlPartExtension(partPath) {
  if (partPath.endsWith(".rels")) return "rels";
  return path.posix.extname(partPath).slice(1).toLowerCase();
}

const OOXML_RELATIONSHIP_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const OOXML_MICROSOFT_RELATIONSHIP_BASE = "http://schemas.microsoft.com/office/2017/10/relationships";
const OOXML_COMMON_PART_RECIPES = {
  image: { relationshipType: `${OOXML_RELATIONSHIP_BASE}/image` },
  chart: { contentType: "application/vnd.openxmlformats-officedocument.drawingml.chart+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/chart` },
  theme: { contentType: "application/vnd.openxmlformats-officedocument.theme+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/theme` },
  customxml: { contentType: "application/xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/customXml` },
};
const OOXML_FAMILY_PART_RECIPES = {
  DOCX: {
    header: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/header` },
    footer: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/footer` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    commentsextended: { contentType: DOCX_COMMENTS_EXTENDED_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_EXTENDED_RELATIONSHIP_TYPE },
    commentsids: { contentType: DOCX_COMMENTS_IDS_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_IDS_RELATIONSHIP_TYPE },
    commentsextensible: { contentType: DOCX_COMMENTS_EXTENSIBLE_CONTENT_TYPE, relationshipType: DOCX_COMMENTS_EXTENSIBLE_RELATIONSHIP_TYPE },
    people: { contentType: DOCX_PEOPLE_CONTENT_TYPE, relationshipType: DOCX_PEOPLE_RELATIONSHIP_TYPE },
    numbering: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/numbering` },
    styles: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/styles` },
    settings: { contentType: "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/settings` },
  },
  XLSX: {
    drawing: { contentType: "application/vnd.openxmlformats-officedocument.drawing+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/drawing` },
    worksheet: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/worksheet` },
    styles: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/styles` },
    sharedstrings: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/sharedStrings` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    table: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/table` },
    pivottable: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotTable` },
    pivotcachedefinition: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotCacheDefinition` },
    pivotcacherecords: { contentType: "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheRecords+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/pivotCacheRecords` },
    threadedcomments: { contentType: "application/vnd.ms-excel.threadedcomments+xml", relationshipType: `${OOXML_MICROSOFT_RELATIONSHIP_BASE}/threadedComment` },
    person: { contentType: "application/vnd.ms-excel.person+xml", relationshipType: `${OOXML_MICROSOFT_RELATIONSHIP_BASE}/person` },
  },
  PPTX: {
    slide: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slide+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slide` },
    slidelayout: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slideLayout` },
    slidemaster: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/slideMaster` },
    notesslide: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/notesSlide` },
    comments: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.comments+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/comments` },
    commentauthors: { contentType: "application/vnd.openxmlformats-officedocument.presentationml.commentAuthors+xml", relationshipType: `${OOXML_RELATIONSHIP_BASE}/commentAuthors` },
    moderncomments: { contentType: PPTX_MODERN_COMMENT_CONTENT_TYPE, relationshipType: PPTX_MODERN_COMMENT_RELATIONSHIP_TYPE },
    modernauthors: { contentType: PPTX_MODERN_AUTHOR_CONTENT_TYPE, relationshipType: PPTX_MODERN_AUTHOR_RELATIONSHIP_TYPE },
  },
};

function ooxmlRecipeKind(value) {
  return String(value || "").trim().toLowerCase().replace(/[^a-z0-9]+/g, "");
}

function applyOoxmlPartRecipe(patch, family) {
  const rawRecipe = patch.recipe ?? patch.partRecipe ?? patch.partKind;
  if (rawRecipe == null) return patch;
  const recipe = typeof rawRecipe === "object" ? rawRecipe : { kind: rawRecipe };
  const kind = ooxmlRecipeKind(recipe.kind || recipe.name || recipe.type);
  const spec = OOXML_FAMILY_PART_RECIPES[family]?.[kind] || OOXML_COMMON_PART_RECIPES[kind];
  if (!spec) {
    const supported = [...Object.keys(OOXML_COMMON_PART_RECIPES), ...Object.keys(OOXML_FAMILY_PART_RECIPES[family] || {})].sort().join(", ");
    throw new Error(`${family} OOXML part recipe ${recipe.kind || recipe.name || recipe.type || "(missing)"} is unsupported. Supported recipes: ${supported}.`);
  }
  const derivedRelationship = recipe.relationship === false ? undefined : {
    source: recipe.source ?? recipe.sourcePart,
    id: recipe.id ?? recipe.relationshipId,
    type: spec.relationshipType,
    target: recipe.target,
    targetMode: recipe.targetMode,
  };
  let relationship = patch.relationship;
  let relationships = patch.relationships;
  if (relationship) relationship = { ...derivedRelationship, ...relationship, type: relationship.type || spec.relationshipType };
  else if (relationships) relationships = relationships.map((item) => ({ ...derivedRelationship, ...item, type: item.type || spec.relationshipType }));
  else if (derivedRelationship?.source !== undefined) relationship = derivedRelationship;
  return { ...patch, contentType: patch.contentType || patch.mimeType || patch.type || spec.contentType, relationship, relationships, sourceReference: patch.sourceReference ?? recipe.sourceReference, recipeKind: kind };
}

function ooxmlRelationshipSource(partPath) {
  if (partPath === "_rels/.rels") return "";
  const match = /^(?:(.*)\/)?_rels\/([^/]+)\.rels$/.exec(partPath);
  if (!match) return undefined;
  return match[1] ? `${match[1]}/${match[2]}` : match[2];
}

export function ooxmlResolveRelationshipTarget(source, rawTarget) {
  const target = String(rawTarget || "").split("#")[0];
  if (target.startsWith("/")) return target.slice(1);
  const sourceDir = source ? path.posix.dirname(source) : "";
  return path.posix.normalize(path.posix.join(sourceDir === "." ? "" : sourceDir, target));
}

const OOXML_RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

function ooxmlRelationshipReferences(xml = "") {
  const source = String(xml || "");
  const tags = [...source.matchAll(/<[^!?][^>]*>/g)].map((match) => match[0]);
  const references = [];
  const namespaceStack = [new Map()];
  for (const tag of tags) {
    if (/^<\//.test(tag)) { if (namespaceStack.length > 1) namespaceStack.pop(); continue; }
    const namespaces = new Map(namespaceStack.at(-1));
    for (const match of tag.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2/g)) namespaces.set(match[1], decodeXml(match[3]));
    for (const [prefix, namespace] of namespaces) {
      if (!OOXML_RELATIONSHIP_NAMESPACES.has(namespace)) continue;
      const escapedPrefix = prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
      const pattern = new RegExp(`\\b${escapedPrefix}:(id|embed|link|dm|lo|qs|cs)\\s*=\\s*(["'])(.*?)\\2`, "g");
      for (const match of tag.matchAll(pattern)) references.push({ attribute: `${prefix}:${match[1]}`, id: decodeXml(match[3]) });
    }
    if (!/\/\s*>$/.test(tag)) namespaceStack.push(namespaces);
  }
  return references.filter((reference) => reference.id);
}

function ooxmlPackageIssues(files, bytesByPath, contentTypes, family) {
  const paths = new Set(files.map((file) => file.name));
  const issues = [];
  if (!paths.has("[Content_Types].xml")) issues.push({ kind: "ooxmlIssue", family, type: "missingContentTypes", severity: "error", message: `${family} package is missing [Content_Types].xml.` });
  if (!paths.has("_rels/.rels")) issues.push({ kind: "ooxmlIssue", family, type: "missingRootRelationships", severity: "error", message: `${family} package is missing _rels/.rels.` });
  for (const partPath of contentTypes.overrides.keys()) {
    if (!paths.has(partPath)) issues.push({ kind: "ooxmlIssue", family, type: "contentTypeTargetNotFound", severity: "error", path: partPath, message: `${family} [Content_Types].xml declares missing part ${partPath}.` });
  }
  for (const partPath of paths) {
    if (partPath === "[Content_Types].xml") continue;
    const extension = ooxmlPartExtension(partPath);
    if (!contentTypes.overrides.has(partPath) && !contentTypes.defaults.has(extension)) {
      issues.push({ kind: "ooxmlIssue", family, type: "missingContentType", severity: "error", path: partPath, message: `${family} part ${partPath} has no [Content_Types].xml declaration.` });
    }
  }
  const relationshipsBySource = new Map();
  for (const [partPath, bytes] of bytesByPath) {
    if (!partPath.endsWith(".rels")) continue;
    const source = ooxmlRelationshipSource(partPath);
    if (source == null) continue;
    const xml = decoder.decode(bytes);
    const relationshipEntries = ooxmlRelationshipEntries(xml);
    relationshipsBySource.set(source, { path: partPath, ids: new Set(relationshipEntries.map((entry) => entry.attrs.Id).filter(Boolean)) });
    if (source && !paths.has(source)) issues.push({ kind: "ooxmlIssue", family, type: "relationshipSourceNotFound", severity: "error", path: partPath, source, message: `${family} relationship part ${partPath} belongs to missing source part ${source}.` });
    const relationshipIds = new Set();
    for (const { attrs } of relationshipEntries) {
      if (attrs.Id && relationshipIds.has(attrs.Id)) issues.push({ kind: "ooxmlIssue", family, type: "duplicateRelationshipId", severity: "error", path: partPath, relationshipId: attrs.Id, message: `${family} relationship part ${partPath} contains duplicate Id ${attrs.Id}.` });
      if (attrs.Id) relationshipIds.add(attrs.Id);
      if (String(attrs.TargetMode || "").toLowerCase() === "external") continue;
      if (!attrs.Target) {
        issues.push({ kind: "ooxmlIssue", family, type: "relationshipTargetMissing", severity: "error", path: partPath, relationshipId: attrs.Id, message: `${family} relationship ${attrs.Id || "(unknown)"} in ${partPath} has no target.` });
        continue;
      }
      const target = ooxmlResolveRelationshipTarget(source, attrs.Target);
      if (!paths.has(target)) issues.push({ kind: "ooxmlIssue", family, type: "relationshipTargetNotFound", severity: "error", path: partPath, relationshipId: attrs.Id, target, message: `${family} relationship ${attrs.Id || "(unknown)"} in ${partPath} targets missing part ${target}.` });
    }
  }
  for (const [partPath, bytes] of bytesByPath) {
    if (!partPath.endsWith(".xml") || partPath === "[Content_Types].xml") continue;
    const references = ooxmlRelationshipReferences(decoder.decode(bytes));
    if (!references.length) continue;
    const relationshipPart = relationshipsBySource.get(partPath);
    if (!relationshipPart) {
      issues.push({ kind: "ooxmlIssue", family, type: "relationshipReferencePartNotFound", severity: "error", path: partPath, relationshipIds: [...new Set(references.map((reference) => reference.id))], message: `${family} source part ${partPath} contains relationship references but has no corresponding .rels part.` });
      continue;
    }
    const seen = new Set();
    for (const reference of references) {
      const key = `${reference.attribute}\u0000${reference.id}`;
      if (seen.has(key) || relationshipPart.ids.has(reference.id)) continue;
      seen.add(key);
      issues.push({ kind: "ooxmlIssue", family, type: "relationshipReferenceIdNotFound", severity: "error", path: partPath, relationshipsPath: relationshipPart.path, relationshipId: reference.id, referenceAttribute: reference.attribute, message: `${family} source part ${partPath} references missing relationship Id ${reference.id} through ${reference.attribute}.` });
    }
  }
  return issues;
}

async function ooxmlPackageRecords(zip, options = {}, config = {}) {
  const includeText = Boolean(options.includeText || options.preview || options.includeXml);
  const maxPreviewChars = Math.max(0, Number(options.maxPreviewChars ?? 400) || 0);
  const files = Object.values(zip.files).filter((file) => !file.dir).sort((a, b) => a.name.localeCompare(b.name));
  const family = config.family || "OOXML";
  const maxParts = Math.max(1, Number(options.maxParts ?? 5000) || 5000);
  const maxPartBytes = Math.max(1, Number(options.maxPartBytes ?? 64 * 1024 * 1024) || 64 * 1024 * 1024);
  const maxTotalBytes = Math.max(1, Number(options.maxTotalBytes ?? 256 * 1024 * 1024) || 256 * 1024 * 1024);
  if (files.length > maxParts) throw new Error(`${family} package has ${files.length} parts; maxParts is ${maxParts}.`);
  const safePaths = new Map();
  let declaredTotalBytes = 0;
  for (const file of files) {
    const partPath = ooxmlSafePartPath(file.name, family);
    safePaths.set(file, partPath);
    const declaredSize = Number(file._data?.uncompressedSize);
    if (!Number.isFinite(declaredSize)) continue;
    if (declaredSize > maxPartBytes) throw new Error(`${family} part ${partPath} exceeds maxPartBytes (${maxPartBytes}).`);
    declaredTotalBytes += declaredSize;
    if (declaredTotalBytes > maxTotalBytes) throw new Error(`${family} package exceeds maxTotalBytes (${maxTotalBytes}).`);
  }
  const contentTypesEntry = zip.file("[Content_Types].xml");
  const contentTypesText = contentTypesEntry ? await contentTypesEntry.async("text") : "";
  const contentTypes = ooxmlContentTypeMaps(contentTypesText);
  const counts = Object.fromEntries(Object.entries(config.counts || {}).map(([name, pattern]) => [name, files.filter((file) => pattern.test(file.name)).length]));
  const records = [{ kind: config.packageKind || "ooxmlPackage", family, parts: files.length, ...counts }];
  const bytesByPath = new Map();
  let totalBytes = 0;
  for (const file of files) {
    const partPath = safePaths.get(file);
    const bytes = await file.async("uint8array");
    if (bytes.byteLength > maxPartBytes) throw new Error(`${family} part ${partPath} exceeds maxPartBytes (${maxPartBytes}).`);
    totalBytes += bytes.byteLength;
    if (totalBytes > maxTotalBytes) throw new Error(`${family} package exceeds maxTotalBytes (${maxTotalBytes}).`);
    const extension = ooxmlPartExtension(partPath);
    const contentType = contentTypes.overrides.get(partPath) || contentTypes.defaults.get(extension) || ooxmlFallbackContentType(partPath);
    const record = { kind: config.partKind || "ooxmlPart", path: partPath, size: bytes.byteLength, contentType };
    if (includeText && /\.(xml|json|rels)$/i.test(partPath)) record.textPreview = decoder.decode(bytes).slice(0, maxPreviewChars);
    records.push(record);
    bytesByPath.set(partPath, bytes);
  }
  const issues = ooxmlPackageIssues(files, bytesByPath, contentTypes, family);
  const semanticIssues = config.semanticIssues ? await config.semanticIssues({ bytesByPath, contentTypes, family }) : [];
  issues.push(...semanticIssues);
  records[0].uncompressedBytes = totalBytes;
  records[0].relationshipReferences = [...bytesByPath].reduce((count, [partPath, bytes]) => count + (partPath.endsWith(".xml") && partPath !== "[Content_Types].xml" ? ooxmlRelationshipReferences(decoder.decode(bytes)).length : 0), 0);
  records[0].relationshipReferenceIssues = issues.filter((issue) => issue.type === "relationshipReferencePartNotFound" || issue.type === "relationshipReferenceIdNotFound").length;
  if (config.semanticIssues) {
    records[0].semanticValidation = true;
    records[0].semanticIssues = semanticIssues.length;
  }
  records[0].ok = issues.length === 0;
  records[0].issues = issues.length;
  records.push(...issues);
  return records;
}

function ooxmlPatchData(patch, family = "OOXML") {
  if (patch.json !== undefined) return encoder.encode(JSON.stringify(patch.json, null, 2));
  if (patch.text !== undefined || patch.xml !== undefined) return encoder.encode(String(patch.text ?? patch.xml));
  if (patch.bytes !== undefined || patch.data !== undefined || patch.buffer !== undefined) return toUint8Array(patch.bytes ?? patch.data ?? patch.buffer);
  if (patch.content !== undefined) return typeof patch.content === "string" ? encoder.encode(patch.content) : toUint8Array(patch.content);
  if (patch.remove || patch.delete) return undefined;
  throw new Error(`${family} patch for ${patch.path || patch.part || "unknown part"} has no content or remove flag.`);
}

function ooxmlRemoveContentTypeOverride(xml, partPath) {
  return String(xml).replace(/<Override\b[^>]*\/?\s*>/g, (tag) => {
    const partName = ooxmlXmlAttributes(tag).PartName;
    return String(partName || "").replace(/^\//, "") === partPath ? "" : tag;
  });
}

function ooxmlInferredPatchContentType(patch, partPath) {
  const explicit = patch.contentType || patch.mimeType || patch.type;
  if (explicit) return String(explicit);
  const extension = ooxmlPartExtension(partPath);
  if (patch.json !== undefined || extension === "json") return "application/json";
  if (patch.xml !== undefined || extension === "xml" || extension === "rels") return extension === "rels" ? "application/vnd.openxmlformats-package.relationships+xml" : "application/xml";
  if (patch.text !== undefined || extension === "txt") return "text/plain";
  const imageType = imageContentTypeFromExtension(extension);
  return imageType === "application/octet-stream" ? undefined : imageType;
}

async function syncOoxmlPatchContentTypes(zip, normalizedPatches, options, family) {
  if (options.syncContentTypes === false) return 0;
  const entry = zip.file("[Content_Types].xml");
  if (!entry) throw new Error(`${family} package is missing [Content_Types].xml; cannot synchronize patch content types.`);
  let xml = await entry.async("text");
  let updates = 0;
  for (const { patch, partPath } of normalizedPatches) {
    if (partPath === "[Content_Types].xml") continue;
    const withoutOverride = ooxmlRemoveContentTypeOverride(xml, partPath);
    const removedOverride = withoutOverride !== xml;
    if (patch.remove || patch.delete) {
      if (removedOverride) { xml = withoutOverride; updates += 1; }
      continue;
    }
    const requestedType = ooxmlInferredPatchContentType(patch, partPath);
    const declarations = ooxmlContentTypeMaps(xml);
    const extension = ooxmlPartExtension(partPath);
    const existingType = declarations.overrides.get(partPath) || declarations.defaults.get(extension);
    if (!requestedType || (existingType && !patch.contentType && !patch.mimeType && !patch.type)) continue;
    if (declarations.defaults.get(extension) === requestedType && !patch.contentType && !patch.mimeType && !patch.type) {
      if (removedOverride) { xml = withoutOverride; updates += 1; }
      continue;
    }
    xml = withoutOverride.replace(/<\/Types>\s*$/, `<Override PartName="/${attrEscape(partPath)}" ContentType="${attrEscape(requestedType)}"/></Types>`);
    updates += 1;
  }
  if (updates) zip.file("[Content_Types].xml", xml);
  return updates;
}

function ooxmlRelationshipPartPath(source, family) {
  if (!source) return "_rels/.rels";
  const safeSource = ooxmlSafePartPath(source, family);
  const dir = path.posix.dirname(safeSource);
  return `${dir === "." ? "" : `${dir}/`}_rels/${path.posix.basename(safeSource)}.rels`;
}

function ooxmlRelationshipEntries(xml = "") {
  return [...String(xml).matchAll(/<Relationship\b[^>]*\/?\s*>/g)].map((match) => ({
    tag: match[0],
    attrs: ooxmlXmlAttributes(match[0]),
  }));
}

function ooxmlEmptyRelationshipsXml() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
}

async function syncOoxmlPatchRelationships(zip, normalizedPatches, options, family) {
  if (options.syncRelationships === false) return 0;
  let updates = 0;
  const removedParts = new Set(normalizedPatches.filter(({ patch }) => patch.remove || patch.delete).map(({ partPath }) => partPath));
  const removedRelationshipIds = new Map();
  for (const removedPart of removedParts) {
    const outgoingRelationships = ooxmlRelationshipPartPath(removedPart, family);
    if (zip.file(outgoingRelationships)) { zip.remove(outgoingRelationships); updates += 1; }
  }
  for (const file of Object.values(zip.files).filter((item) => !item.dir && item.name.endsWith(".rels"))) {
    const relsPath = ooxmlSafePartPath(file.name, family);
    const source = ooxmlRelationshipSource(relsPath);
    if (source == null) continue;
    const xml = await file.async("text");
    const next = xml.replace(/<Relationship\b[^>]*\/?\s*>/g, (tag) => {
      const entry = ooxmlRelationshipEntries(tag)[0];
      if (!entry || String(entry.attrs.TargetMode || "").toLowerCase() === "external") return tag;
      const resolvedTarget = ooxmlResolveRelationshipTarget(source, entry.attrs.Target);
      if (!removedParts.has(resolvedTarget)) return tag;
      const key = `${source}\u0000${resolvedTarget}`;
      if (entry.attrs.Id) removedRelationshipIds.set(key, [...(removedRelationshipIds.get(key) || []), entry.attrs.Id]);
      return "";
    });
    if (next !== xml) { zip.file(relsPath, next); updates += 1; }
  }
  for (const { patch, partPath } of normalizedPatches) {
    const relationshipConfigs = patch.relationships || (patch.relationship ? [patch.relationship] : []);
    for (const relationship of relationshipConfigs) {
      const source = relationship.source || relationship.sourcePart || "";
      if (source && !zip.file(ooxmlSafePartPath(source, family))) throw new Error(`${family} relationship source part not found: ${source}`);
      const relsPath = ooxmlRelationshipPartPath(source, family);
      const relsEntry = zip.file(relsPath);
      let xml = relsEntry ? await relsEntry.async("text") : ooxmlEmptyRelationshipsXml();
      const entries = ooxmlRelationshipEntries(xml);
      const target = relationship.target || (source ? path.posix.relative(path.posix.dirname(source), partPath) : partPath);
      const remove = relationship.remove === true || relationship.delete === true || patch.remove || patch.delete;
      const matches = (entry) => (relationship.id && entry.attrs.Id === relationship.id) || (!relationship.id && ooxmlResolveRelationshipTarget(source, entry.attrs.Target) === partPath);
      const existingById = relationship.id ? entries.find((entry) => entry.attrs.Id === relationship.id) : undefined;
      if (!remove && existingById && ooxmlResolveRelationshipTarget(source, existingById.attrs.Target) !== partPath && relationship.replaceExisting !== true && relationship.replace !== true) throw new Error(`${family} relationship Id ${relationship.id} in ${relsPath} already targets ${existingById.attrs.Target}; pass replaceExisting:true only for an intentional rebind.`);
      relationship.resolvedIds = [...new Set([...entries.filter(matches).map((entry) => entry.attrs.Id).filter(Boolean), ...(removedRelationshipIds.get(`${source}\u0000${partPath}`) || [])])];
      const withoutMatch = xml.replace(/<Relationship\b[^>]*\/?\s*>/g, (tag) => {
        const entry = ooxmlRelationshipEntries(tag)[0];
        return entry && matches(entry) ? "" : tag;
      });
      if (remove) {
        if (withoutMatch !== xml) { zip.file(relsPath, withoutMatch); updates += 1; }
        continue;
      }
      if (!relationship.type) throw new Error(`${family} relationship for ${partPath} requires type.`);
      const usedIds = new Set(entries.map((entry) => entry.attrs.Id));
      let id = relationship.id;
      if (!id) { let index = 1; while (usedIds.has(`rId${index}`)) index += 1; id = `rId${index}`; }
      relationship.resolvedId = id;
      const targetMode = relationship.targetMode ? ` TargetMode="${attrEscape(relationship.targetMode)}"` : "";
      const tag = `<Relationship Id="${attrEscape(id)}" Type="${attrEscape(relationship.type)}" Target="${attrEscape(target)}"${targetMode}/>`;
      const next = withoutMatch.replace(/<\/Relationships>\s*$/, `${tag}</Relationships>`);
      zip.file(relsPath, next);
      updates += 1;
    }
  }
  return updates;
}

function ooxmlReferenceConfig(value) {
  if (value === true) return {};
  if (value && typeof value === "object") return value;
  return undefined;
}

async function syncOoxmlSourceReferences(zip, normalizedPatches, options, family) {
  if (options.syncSourceReferences === false) return 0;
  let updates = 0;
  for (const { patch, partPath } of normalizedPatches) {
    const config = ooxmlReferenceConfig(patch.sourceReference);
    if (!config) continue;
    if (!supportsOoxmlSourceReference(family, patch.recipeKind)) throw new Error(`${family} sourceReference is not supported for recipe ${patch.recipeKind || "(missing)"}. Supported recipes: ${supportedOoxmlSourceReferenceSummary()}.`);
    const relationship = patch.relationship || patch.relationships?.[0];
    const source = relationship?.source || relationship?.sourcePart;
    if (!source) throw new Error(`${family} ${patch.recipeKind} sourceReference requires recipe.source.`);
    const safeSource = ooxmlSafePartPath(source, family);
    const sourceEntry = zip.file(safeSource);
    if (!sourceEntry) throw new Error(`${family} sourceReference source part not found: ${safeSource}`);
    const remove = patch.remove === true || patch.delete === true;
    const resolvedIds = new Set([...(relationship.resolvedIds || []), relationship.id].filter(Boolean));
    const addId = remove ? undefined : relationship.resolvedId || relationship.id;
    if (!remove && !addId) throw new Error(`${family} ${patch.recipeKind} sourceReference could not resolve a relationship Id.`);
    const validatesTarget = (family === "DOCX" && ["comments", "numbering", "settings"].includes(patch.recipeKind)) || (family === "PPTX" && patch.recipeKind === "chart");
    if (!remove && validatesTarget) {
      const targetEntry = zip.file(partPath);
      if (!targetEntry) throw new Error(`${family} sourceReference target part not found: ${partPath}`);
      const targetXml = await targetEntry.async("text");
      if (family === "DOCX" && patch.recipeKind === "settings") {
        const nextTarget = mutateOoxmlSourceReferenceTarget({ family, recipeKind: patch.recipeKind, targetXml, config });
        if (nextTarget !== targetXml) { zip.file(partPath, nextTarget); updates += 1; }
        continue;
      }
      validateOoxmlSourceReferenceTarget({ family, recipeKind: patch.recipeKind, targetXml, config });
    }
    const xml = await sourceEntry.async("text");
    const next = mutateOoxmlSourceReference({ family, recipeKind: patch.recipeKind, xml, relationshipIds: resolvedIds, addId, config });
    if (next !== xml) { zip.file(safeSource, next); updates += 1; }
  }
  return updates;
}

export async function inspectOoxmlPackage(blobOrBuffer, options = {}, config = {}) {
  const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
  const zip = await JSZip.loadAsync(bytes);
  const records = await ooxmlPackageRecords(zip, options, config);
  const partKind = config.partKind || "ooxmlPart";
  return { ok: records[0].ok, issues: records.filter((record) => record.kind === "ooxmlIssue"), parts: records.filter((record) => record.kind === partKind), records, ...ndjson(records, options.maxChars ?? Infinity) };
}

export async function patchOoxmlPackage(blobOrBuffer, patches = [], options = {}, config = {}) {
  const family = config.family || "OOXML";
  const bytes = blobOrBuffer instanceof FileBlob ? new Uint8Array(await blobOrBuffer.arrayBuffer()) : toUint8Array(blobOrBuffer);
  const zip = await JSZip.loadAsync(bytes);
  const list = Array.isArray(patches) ? patches : Object.entries(patches || {}).map(([partPath, content]) => (
    content && typeof content === "object" && !(content instanceof Uint8Array) && !(content instanceof ArrayBuffer) && !ArrayBuffer.isView(content)
      ? { path: partPath, ...content }
      : { path: partPath, content }
  ));
  const preparedList = list.map((patch) => applyOoxmlPartRecipe(patch, family));
  const maxPatchBytes = Math.max(1, Number(options.maxPatchBytes ?? 5 * 1024 * 1024) || 5 * 1024 * 1024);
  const maxParts = Math.max(1, Number(options.maxParts ?? 5000) || 5000);
  const existingParts = new Set(Object.values(zip.files).filter((file) => !file.dir).map((file) => ooxmlSafePartPath(file.name, family)));
  const normalizedPatches = preparedList.map((patch) => ({ patch, partPath: ooxmlSafePartPath(patch.path || patch.part || patch.name, family) }));
  const resultingParts = new Set(existingParts);
  for (const { patch, partPath } of normalizedPatches) {
    if (patch.remove || patch.delete) resultingParts.delete(partPath);
    else resultingParts.add(partPath);
  }
  if (resultingParts.size > maxParts) throw new Error(`${family} patch would create ${resultingParts.size} parts; maxParts is ${maxParts}.`);
  for (const { patch, partPath } of normalizedPatches) {
    if (patch.remove || patch.delete) { zip.remove(partPath); continue; }
    const data = ooxmlPatchData(patch, family);
    if (data?.byteLength > maxPatchBytes) throw new Error(`${family} patch for ${partPath} exceeds maxPatchBytes (${maxPatchBytes}).`);
    zip.file(partPath, data);
  }
  const contentTypesUpdated = await syncOoxmlPatchContentTypes(zip, normalizedPatches, options, family);
  const relationshipsUpdated = await syncOoxmlPatchRelationships(zip, normalizedPatches, options, family);
  const sourceReferencesUpdated = await syncOoxmlSourceReferences(zip, normalizedPatches, options, family);
  const validated = options.validateResult !== false;
  let validationIssues = [];
  if (validated) {
    const records = await ooxmlPackageRecords(zip, { maxParts, maxPartBytes: options.maxPartBytes, maxTotalBytes: options.maxTotalBytes }, { ...config, family });
    validationIssues = records.filter((record) => record.kind === "ooxmlIssue");
    if (validationIssues.length) {
      const summary = validationIssues.slice(0, 5).map((issue) => `${issue.type}${issue.path ? `:${issue.path}` : ""}`).join(", ");
      throw new Error(`${family} patch produced an invalid OOXML package (${validationIssues.length} issue${validationIssues.length === 1 ? "" : "s"}): ${summary}. Pass { validateResult: false } only when intentionally constructing an invalid fixture.`);
    }
  }
  return { bytes: await zip.generateAsync({ type: "uint8array", compression: "DEFLATE" }), patchedParts: preparedList.length, recipesApplied: preparedList.filter((patch) => patch.recipeKind).length, contentTypesUpdated, relationshipsUpdated, sourceReferencesUpdated, validated, validationIssues: validationIssues.length };
}
