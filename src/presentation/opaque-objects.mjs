import path from "node:path";

const RELATIONSHIP_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function attrEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function safePartPath(partPath) {
  const raw = String(partPath || "").replaceAll("\\", "/").replace(/^\//, "");
  const normalized = path.posix.normalize(raw).replace(/^\.\//, "");
  if (!normalized || normalized === "." || normalized === ".." || normalized.startsWith("../") || normalized.includes("/../") || normalized.includes("\0")) throw new Error(`Unsafe PPTX native-object part path: ${partPath}`);
  return normalized;
}

function relationshipPartPath(source) {
  const safeSource = safePartPath(source);
  const dir = path.posix.dirname(safeSource);
  return `${dir === "." ? "" : `${dir}/`}_rels/${path.posix.basename(safeSource)}.rels`;
}

function resolveTarget(source, target) {
  const raw = String(target || "").split("#")[0];
  if (raw.startsWith("/")) return safePartPath(raw);
  return safePartPath(path.posix.join(path.posix.dirname(source), raw));
}

function relativeTarget(source, target) {
  const relative = path.posix.relative(path.posix.dirname(source), target);
  return relative || path.posix.basename(target);
}

function parseRelationships(xml = "") {
  return [...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?Relationship\b[^>]*\/?\s*>/g)].map((match) => {
    const attrs = attributes(match[0]);
    return { id: attrs.Id || "", type: attrs.Type || "", target: attrs.Target || "", targetMode: attrs.TargetMode || "" };
  }).filter((relationship) => relationship.id);
}

function relationshipsXml(relationships = []) {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${relationships.map((relationship) => `<Relationship Id="${attrEscape(relationship.id)}" Type="${attrEscape(relationship.type)}" Target="${attrEscape(relationship.target)}"${relationship.targetMode ? ` TargetMode="${attrEscape(relationship.targetMode)}"` : ""}/>`).join("")}</Relationships>`;
}

export function presentationNamespaceMap(xml = "") {
  const root = /<(?![!?/])[A-Za-z_][\w:.-]*\b[^>]*>/.exec(String(xml))?.[0] || "";
  return new Map([...root.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function presentationNamespaceDeclarations(xml = "") {
  return new Map([...String(xml).matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

export function selfContainedPresentationFragment(fragment = "", sourceXml = "") {
  const xml = String(fragment);
  const opening = /<(?![!?/])[A-Za-z_][\w:.-]*\b[^>]*>/.exec(xml)?.[0];
  if (!opening) return xml;
  const namespaces = presentationNamespaceMap(sourceXml);
  for (const [prefix, namespace] of presentationNamespaceMap(xml)) namespaces.set(prefix, namespace);
  const declared = new Set([...opening.matchAll(/\bxmlns:([A-Za-z_][\w.-]*)\s*=/g)].map((match) => match[1]));
  const used = new Set([...xml.matchAll(/(?:<\/?|\s)([A-Za-z_][\w.-]*):[A-Za-z_][\w.-]*/g)].map((match) => match[1]));
  const additions = [...used].filter((prefix) => !declared.has(prefix) && namespaces.has(prefix)).map((prefix) => ` xmlns:${prefix}="${attrEscape(namespaces.get(prefix))}"`).join("");
  if (!additions) return xml;
  const closing = /\/\s*>$/.test(opening) ? "/>" : ">";
  return xml.replace(opening, opening.replace(/\/?\s*>$/, `${additions}${closing}`));
}

export function presentationRelationshipReferences(fragment = "", sourceXml = "") {
  const namespaces = presentationNamespaceMap(sourceXml);
  for (const [prefix, namespace] of presentationNamespaceDeclarations(fragment)) namespaces.set(prefix, namespace);
  const references = [];
  for (const [prefix, namespace] of namespaces) {
    if (!RELATIONSHIP_NAMESPACES.has(namespace)) continue;
    const escaped = prefix.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const pattern = new RegExp(`\\b${escaped}:(id|embed|link|dm|lo|qs|cs)\\s*=\\s*(["'])(.*?)\\2`, "g");
    for (const match of String(fragment).matchAll(pattern)) references.push({ attribute: `${prefix}:${match[1]}`, id: decodeXml(match[3]) });
  }
  return references.filter((reference, index, list) => list.findIndex((candidate) => candidate.id === reference.id) === index);
}

function contentTypeMaps(xml = "") {
  const defaults = new Map();
  const overrides = new Map();
  for (const match of String(xml).matchAll(/<Default\b[^>]*\/?\s*>/g)) {
    const attrs = attributes(match[0]);
    if (attrs.Extension && attrs.ContentType) defaults.set(String(attrs.Extension).toLowerCase(), attrs.ContentType);
  }
  for (const match of String(xml).matchAll(/<Override\b[^>]*\/?\s*>/g)) {
    const attrs = attributes(match[0]);
    if (attrs.PartName && attrs.ContentType) overrides.set(String(attrs.PartName).replace(/^\//, ""), attrs.ContentType);
  }
  return { defaults, overrides };
}

function contentTypeFor(partPath, maps) {
  return maps.overrides.get(partPath) || maps.defaults.get(path.posix.extname(partPath).slice(1).toLowerCase()) || "application/octet-stream";
}

export async function capturePresentationOpaqueObject({ zip, slidePath, slideXml, fragment, relationships, contentTypesXml }) {
  const references = presentationRelationshipReferences(fragment, slideXml);
  const relationshipById = new Map(relationships.map((relationship) => [relationship.id, relationship]));
  const missingReference = references.find((reference) => !relationshipById.has(reference.id));
  if (missingReference) throw new Error(`PPTX native object references missing slide relationship ${missingReference.id}.`);
  const rootRelationships = references.map((reference) => ({ ...relationshipById.get(reference.id) }));
  const maps = contentTypeMaps(contentTypesXml);
  const parts = new Map();

  async function capture(partPath) {
    const safePath = safePartPath(partPath);
    if (parts.has(safePath)) return;
    const entry = zip.file(safePath);
    if (!entry) throw new Error(`PPTX native object references missing part ${safePath}.`);
    const bytes = await entry.async("uint8array");
    const relsPath = relationshipPartPath(safePath);
    const partRelationships = parseRelationships(await zip.file(relsPath)?.async("text"));
    const record = { path: safePath, bytes, contentType: contentTypeFor(safePath, maps), relationships: partRelationships };
    parts.set(safePath, record);
    for (const relationship of partRelationships) {
      if (relationship.targetMode.toLowerCase() === "external") continue;
      await capture(resolveTarget(safePath, relationship.target));
    }
  }

  for (const relationship of rootRelationships) {
    if (relationship.targetMode.toLowerCase() === "external") continue;
    await capture(resolveTarget(slidePath, relationship.target));
  }
  return {
    sourcePart: safePartPath(slidePath),
    rawXml: selfContainedPresentationFragment(fragment, slideXml),
    relationshipReferences: references,
    rootRelationships,
    parts: [...parts.values()],
  };
}

function uniquePreservedPath(original, reserved, assigned, sequence) {
  if (!reserved.has(original) && !assigned.has(original)) return original;
  const extension = path.posix.extname(original);
  const stem = path.posix.basename(original, extension).replace(/[^A-Za-z0-9_.-]+/g, "-") || "part";
  let candidate;
  do candidate = `ppt/preserved/native${sequence.value++}/${stem}${extension}`;
  while (reserved.has(candidate) || assigned.has(candidate));
  return candidate;
}

function rewriteFragmentRelationships(xml, replacements, relationshipReferences = []) {
  if (!replacements.size) return xml;
  const attributes = new Set(relationshipReferences.map((reference) => reference.attribute));
  return String(xml).replace(/\b([A-Za-z_][\w.-]*):(id|embed|link|dm|lo|qs|cs)\s*=\s*(["'])(.*?)\3/g, (match, prefix, localName, quote, value) => {
    if (attributes.size && !attributes.has(`${prefix}:${localName}`)) return match;
    const replacement = replacements.get(decodeXml(value));
    return replacement ? `${prefix}:${localName}=${quote}${attrEscape(replacement)}${quote}` : match;
  });
}

export function planPresentationOpaqueParts(slides, options = {}) {
  const reserved = new Set(options.reservedPaths || []);
  const allParts = new Map();
  for (const slide of slides) for (const object of options.objectsForSlide(slide)) for (const part of object.parts || []) {
    const existing = allParts.get(part.path);
    if (existing && (existing.bytes.length !== part.bytes.length || !existing.bytes.every((byte, index) => byte === part.bytes[index]))) throw new Error(`PPTX native objects contain conflicting bytes for ${part.path}.`);
    allParts.set(part.path, part);
  }
  const pathMap = new Map();
  const assigned = new Set();
  const sequence = { value: 1 };
  for (const original of [...allParts.keys()].sort()) {
    const output = uniquePreservedPath(original, reserved, assigned, sequence);
    pathMap.set(original, output);
    assigned.add(output);
  }
  const parts = [...allParts.values()].map((part) => {
    const outputPath = pathMap.get(part.path);
    const relationships = part.relationships.map((relationship) => {
      if (relationship.targetMode.toLowerCase() === "external") return { ...relationship };
      const originalTarget = resolveTarget(part.path, relationship.target);
      return { ...relationship, target: relativeTarget(outputPath, pathMap.get(originalTarget)) };
    });
    return { ...part, outputPath, relationships, relationshipsXml: relationships.length ? relationshipsXml(relationships) : undefined };
  });
  const bySlide = new Map();
  for (const [slideIndex, slide] of slides.entries()) {
    let nextRelationshipIndex = Number(options.startRelationshipIndex(slide, slideIndex));
    const entries = new Map();
    for (const object of options.objectsForSlide(slide)) {
      const replacements = new Map();
      const rootRelationships = object.rootRelationships.map((relationship) => {
        const id = `rId${nextRelationshipIndex++}`;
        replacements.set(relationship.id, id);
        if (relationship.targetMode.toLowerCase() === "external") return { ...relationship, id };
        const originalTarget = resolveTarget(object.sourcePart || options.slidePath(slide, slideIndex), relationship.target);
        return { ...relationship, id, target: relativeTarget(options.slidePath(slide, slideIndex), pathMap.get(originalTarget)) };
      });
      entries.set(object.id, { xml: rewriteFragmentRelationships(object.rawXml, replacements, object.relationshipReferences), relationships: rootRelationships });
    }
    bySlide.set(slide, { entries, relationships: [...entries.values()].flatMap((entry) => entry.relationships), nextRelationshipIndex });
  }
  return { bySlide, parts, contentTypeOverrides: parts.map((part) => ({ path: part.outputPath, contentType: part.contentType })) };
}

export function presentationOpaqueContentTypeXml(overrides = []) {
  return overrides.map((entry) => `<Override PartName="/${attrEscape(entry.path)}" ContentType="${attrEscape(entry.contentType)}"/>`).join("");
}
