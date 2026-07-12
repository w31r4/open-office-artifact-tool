import {
  appendToRoot,
  attrEscape,
  attributes,
  ensureNamespacePrefix,
  ensureRelationshipPrefix,
  insertBeforeOrAppend,
  qname,
  removeReferenceTags,
  rootPrefix,
  rootTag,
} from "./source-reference-xml.mjs";

const PRESENTATION_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/presentationml/2006/main",
  "http://purl.oclc.org/ooxml/presentationml/main",
]);
const DRAWING_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/main",
  "http://purl.oclc.org/ooxml/drawingml/main",
]);
const CHART_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/drawingml/2006/chart",
  "http://purl.oclc.org/ooxml/drawingml/chart",
]);
const SUPPORTED = new Set(["slide", "slidemaster", "slidelayout", "image", "chart"]);

function namespacePrefix(xml, rootLocalName, namespaces, preferred) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`PPTX sourceReference could not find root element ${rootLocalName}.`);
  for (const [name, value] of Object.entries(attributes(root[0]))) {
    if (!namespaces.has(value)) continue;
    if (name === "xmlns") return { xml: String(xml), prefix: "", namespace: value };
    if (name.startsWith("xmlns:")) return { xml: String(xml), prefix: name.slice(6), namespace: value };
  }
  const namespace = [...namespaces][0];
  const ensured = ensureNamespacePrefix(xml, rootLocalName, namespace, preferred, false);
  return { ...ensured, namespace };
}

function assertPresentationRoot(xml, localName) {
  const root = rootTag(xml, localName);
  if (!root) throw new Error(`PPTX sourceReference target must have a ${localName} root element.`);
  const attrs = attributes(root[0]);
  const namespace = root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns;
  if (!PRESENTATION_NAMESPACES.has(namespace)) throw new Error(`PPTX sourceReference ${localName} root must use a PresentationML namespace.`);
  return root;
}

function mutateSlideReference(xml, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "presentation");
  let next = removeReferenceTags(xml, "sldId", ids);
  if (!addId) return next;
  const existingIds = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldId\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).id)).filter(Number.isFinite);
  const slideId = Number(config.slideId ?? Math.max(255, ...existingIds) + 1);
  if (!Number.isInteger(slideId) || slideId < 256 || slideId > 2_147_483_647) throw new Error("PPTX slide sourceReference slideId must be an integer from 256 through 2147483647.");
  if (existingIds.includes(slideId)) throw new Error(`PPTX slide sourceReference slideId ${slideId} already exists.`);
  const ensured = ensureRelationshipPrefix(next, "presentation");
  next = ensured.xml;
  const slideTag = `<${qname(prefix, "sldId")} id="${slideId}" ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const listName = qname(prefix, "sldIdLst");
  const block = new RegExp(`<${listName}\\b[^>]*>[\\s\\S]*?</${listName}>`).exec(next)?.[0];
  if (block) return next.replace(block, block.replace(new RegExp(`</${listName}>$`), `${slideTag}</${listName}>`));
  if (new RegExp(`<${listName}\\b[^>]*\\/>`).test(next)) return next.replace(new RegExp(`<${listName}\\b[^>]*\\/>`), (tag) => `${tag.replace(/\/>$/, ">")}${slideTag}</${listName}>`);
  return appendToRoot(next, "presentation", `<${listName}>${slideTag}</${listName}>`);
}

function mutateMasterReference(xml, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "presentation");
  let next = removeReferenceTags(xml, "sldMasterId", ids);
  if (!addId) return next.replace(new RegExp(`<${qname(prefix, "sldMasterIdLst")}\\b[^>]*>\\s*</${qname(prefix, "sldMasterIdLst")}>`), "");
  const existingIds = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldMasterId\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).id)).filter(Number.isFinite);
  const masterId = Number(config.masterId ?? Math.max(2_147_483_647, ...existingIds) + 1);
  if (!Number.isInteger(masterId) || masterId < 2_147_483_648 || masterId > 4_294_967_295) throw new Error("PPTX slideMaster sourceReference masterId must be an integer from 2147483648 through 4294967295.");
  if (existingIds.includes(masterId)) throw new Error(`PPTX slideMaster sourceReference masterId ${masterId} already exists.`);
  const ensured = ensureRelationshipPrefix(next, "presentation");
  next = ensured.xml;
  const tag = `<${qname(prefix, "sldMasterId")} id="${masterId}" ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const listName = qname(prefix, "sldMasterIdLst");
  const block = new RegExp(`<${listName}\\b[^>]*>[\\s\\S]*?</${listName}>`).exec(next)?.[0];
  if (block) return next.replace(block, block.replace(new RegExp(`</${listName}>$`), `${tag}</${listName}>`));
  if (new RegExp(`<${listName}\\b[^>]*\\/>`).test(next)) return next.replace(new RegExp(`<${listName}\\b[^>]*\\/>`), (item) => `${item.replace(/\/\s*>$/, ">")}${tag}</${listName}>`);
  return insertBeforeOrAppend(next, "presentation", `<${listName}>${tag}</${listName}>`, ["sldIdLst", "sldSz", "notesSz", "embeddedFontLst", "custShowLst", "photoAlbum", "custDataLst", "kinsoku", "defaultTextStyle", "modifyVerifier", "extLst"]);
}

function mutateLayoutReference(xml, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "sldMaster");
  let next = removeReferenceTags(xml, "sldLayoutId", ids);
  if (!addId) return next.replace(new RegExp(`<${qname(prefix, "sldLayoutIdLst")}\\b[^>]*>\\s*</${qname(prefix, "sldLayoutIdLst")}>`), "");
  const existingIds = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sldLayoutId\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).id)).filter(Number.isFinite);
  const layoutId = Number(config.layoutId ?? Math.max(2_147_483_647, ...existingIds) + 1);
  if (!Number.isInteger(layoutId) || layoutId < 2_147_483_648 || layoutId > 4_294_967_295) throw new Error("PPTX slideLayout sourceReference layoutId must be an integer from 2147483648 through 4294967295.");
  if (existingIds.includes(layoutId)) throw new Error(`PPTX slideLayout sourceReference layoutId ${layoutId} already exists.`);
  const ensured = ensureRelationshipPrefix(next, "sldMaster");
  next = ensured.xml;
  const tag = `<${qname(prefix, "sldLayoutId")} id="${layoutId}" ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const listName = qname(prefix, "sldLayoutIdLst");
  const block = new RegExp(`<${listName}\\b[^>]*>[\\s\\S]*?</${listName}>`).exec(next)?.[0];
  if (block) return next.replace(block, block.replace(new RegExp(`</${listName}>$`), `${tag}</${listName}>`));
  if (new RegExp(`<${listName}\\b[^>]*\\/>`).test(next)) return next.replace(new RegExp(`<${listName}\\b[^>]*\\/>`), (item) => `${item.replace(/\/\s*>$/, ">")}${tag}</${listName}>`);
  return insertBeforeOrAppend(next, "sldMaster", `<${listName}>${tag}</${listName}>`, ["transition", "timing", "hf", "txStyles", "extLst"]);
}

function removeSlideObjects(xml, ids) {
  if (!ids.size) return String(xml);
  return String(xml).replace(/<(?:[A-Za-z_][\w.-]*:)?(pic|graphicFrame)\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?\1>/g, (object) => {
    const references = [...object.matchAll(/\b[A-Za-z_][\w.-]*:(?:id|embed|link)\s*=\s*(["'])(.*?)\1/g)].map((match) => match[2]);
    return references.some((id) => ids.has(id)) ? "" : object;
  });
}

function objectPosition(config = {}, kind) {
  const position = config.position;
  if (!position || typeof position !== "object") throw new Error(`PPTX ${kind} sourceReference requires an explicit position with left, top, width, and height in pixels.`);
  const values = Object.fromEntries(["left", "top", "width", "height"].map((name) => [name, Number(position[name])]));
  if (!Number.isFinite(values.left) || !Number.isFinite(values.top)) throw new Error(`PPTX ${kind} sourceReference position left/top must be finite numbers.`);
  if (!Number.isFinite(values.width) || !Number.isFinite(values.height) || values.width <= 0 || values.height <= 0) throw new Error(`PPTX ${kind} sourceReference position width/height must be positive finite numbers.`);
  return Object.fromEntries(Object.entries(values).map(([name, value]) => [name, Math.round(value * 9525)]));
}

function appendToShapeTree(xml, objectXml) {
  const prefix = rootPrefix(xml, "sld");
  const treeName = qname(prefix, "spTree");
  const block = new RegExp(`<${treeName}\\b[^>]*>[\\s\\S]*?</${treeName}>`).exec(xml)?.[0];
  if (!block) throw new Error("PPTX image/chart sourceReference requires a slide with a non-empty p:spTree.");
  return String(xml).replace(block, block.replace(new RegExp(`</${treeName}>$`), `${objectXml}</${treeName}>`));
}

function mutateSlideObject(xml, kind, ids, addId, config = {}) {
  assertPresentationRoot(xml, "sld");
  let next = removeSlideObjects(xml, ids);
  if (!addId) return next;
  const existingIds = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).id)).filter(Number.isFinite);
  const objectId = Number(config.objectId ?? Math.max(0, ...existingIds) + 1);
  if (!Number.isInteger(objectId) || objectId < 1 || objectId > 4_294_967_295) throw new Error(`PPTX ${kind} sourceReference objectId must be an integer from 1 through 4294967295.`);
  if (existingIds.includes(objectId)) throw new Error(`PPTX ${kind} sourceReference objectId ${objectId} already exists.`);
  const position = objectPosition(config, kind);
  const p = rootPrefix(next, "sld");
  let ensured = namespacePrefix(next, "sld", DRAWING_NAMESPACES, "a");
  next = ensured.xml;
  const a = ensured.prefix;
  ensured = ensureRelationshipPrefix(next, "sld");
  next = ensured.xml;
  const r = ensured.prefix;
  const name = attrEscape(config.name || `${kind === "image" ? "Image" : "Chart"} ${objectId}`);
  const alt = attrEscape(config.alt || config.description || "");
  const transform = `<${qname(a, "off")} x="${position.left}" y="${position.top}"/><${qname(a, "ext")} cx="${position.width}" cy="${position.height}"/>`;
  if (kind === "image") {
    const object = `<${qname(p, "pic")}><${qname(p, "nvPicPr")}><${qname(p, "cNvPr")} id="${objectId}" name="${name}" descr="${alt}"/><${qname(p, "cNvPicPr")}><${qname(a, "picLocks")} noChangeAspect="1"/></${qname(p, "cNvPicPr")}><${qname(p, "nvPr")}/></${qname(p, "nvPicPr")}><${qname(p, "blipFill")}><${qname(a, "blip")} ${r}:embed="${attrEscape(addId)}"/><${qname(a, "stretch")}><${qname(a, "fillRect")}/></${qname(a, "stretch")}></${qname(p, "blipFill")}><${qname(p, "spPr")}><${qname(a, "xfrm")}>${transform}</${qname(a, "xfrm")}><${qname(a, "prstGeom")} prst="rect"><${qname(a, "avLst")}/></${qname(a, "prstGeom")}></${qname(p, "spPr")}></${qname(p, "pic")}>`;
    return appendToShapeTree(next, object);
  }
  ensured = namespacePrefix(next, "sld", CHART_NAMESPACES, "c");
  next = ensured.xml;
  const c = ensured.prefix;
  const chartNamespace = ensured.namespace;
  const object = `<${qname(p, "graphicFrame")}><${qname(p, "nvGraphicFramePr")}><${qname(p, "cNvPr")} id="${objectId}" name="${name}" descr="${alt}"/><${qname(p, "cNvGraphicFramePr")}><${qname(a, "graphicFrameLocks")} noGrp="1"/></${qname(p, "cNvGraphicFramePr")}><${qname(p, "nvPr")}/></${qname(p, "nvGraphicFramePr")}><${qname(p, "xfrm")}>${transform}</${qname(p, "xfrm")}><${qname(a, "graphic")}><${qname(a, "graphicData")} uri="${chartNamespace}"><${qname(c, "chart")} ${r}:id="${attrEscape(addId)}"/></${qname(a, "graphicData")}></${qname(a, "graphic")}></${qname(p, "graphicFrame")}>`;
  return appendToShapeTree(next, object);
}

export function supportsPptxSourceReference(recipeKind) {
  return SUPPORTED.has(recipeKind);
}

export function validatePptxSourceReferenceTarget(recipeKind, targetXml) {
  if (recipeKind !== "chart") return;
  const root = rootTag(targetXml, "chartSpace");
  if (!root) throw new Error("PPTX chart sourceReference target part must have a chartSpace root element.");
  const attrs = attributes(root[0]);
  const namespace = root[1] ? attrs[`xmlns:${root[1]}`] : attrs.xmlns;
  if (!CHART_NAMESPACES.has(namespace)) throw new Error("PPTX chart sourceReference target root must use a DrawingML chart namespace.");
}

export function mutatePptxSourceReference(recipeKind, xml, relationshipIds, addId, config = {}) {
  if (recipeKind === "slide") return mutateSlideReference(xml, relationshipIds, addId, config);
  if (recipeKind === "slidemaster") return mutateMasterReference(xml, relationshipIds, addId, config);
  if (recipeKind === "slidelayout") return mutateLayoutReference(xml, relationshipIds, addId, config);
  return mutateSlideObject(xml, recipeKind, relationshipIds, addId, config);
}
