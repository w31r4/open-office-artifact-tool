const RELATIONSHIP_NAMESPACE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELATIONSHIP_NAMESPACES = new Set([
  RELATIONSHIP_NAMESPACE,
  "http://purl.oclc.org/ooxml/officeDocument/relationships",
]);
const DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/main";
const CHART_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const SPREADSHEET_DRAWING_NAMESPACE = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";

const SUPPORTED = new Set([
  "DOCX:header",
  "DOCX:footer",
  "DOCX:comments",
  "XLSX:worksheet",
  "XLSX:table",
  "XLSX:drawing",
  "XLSX:image",
  "XLSX:chart",
  "XLSX:pivotcachedefinition",
  "XLSX:pivotcacherecords",
  "PPTX:slide",
  "PPTX:slidemaster",
  "PPTX:slidelayout",
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
    .replaceAll('"', "&quot;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function qname(prefix, localName) {
  return prefix ? `${prefix}:${localName}` : localName;
}

function rootTag(xml, localName) {
  return new RegExp(`<(?:([A-Za-z_][\\w.-]*):)?${localName}\\b[^>]*\\/?>`).exec(String(xml));
}

function rootPrefix(xml, localName) {
  const root = rootTag(xml, localName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${localName}.`);
  return root[1] || "";
}

function ensureNamespacePrefix(xml, rootLocalName, namespace, preferred, allowDefault = true) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const attrs = attributes(root[0]);
  for (const [name, value] of Object.entries(attrs)) {
    if (value !== namespace) continue;
    if (name === "xmlns" && allowDefault) return { xml: String(xml), prefix: "" };
    if (name.startsWith("xmlns:")) return { xml: String(xml), prefix: name.slice(6) };
  }
  let prefix = preferred;
  let index = 1;
  while (attrs[`xmlns:${prefix}`]) prefix = `${preferred}${index++}`;
  const updatedRoot = root[0].replace(/\s*\/?>$/, (ending) => ` xmlns:${prefix}="${namespace}"${ending}`);
  return { xml: `${String(xml).slice(0, root.index)}${updatedRoot}${String(xml).slice(root.index + root[0].length)}`, prefix };
}

function ensureRelationshipPrefix(xml, rootLocalName) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const attrs = attributes(root[0]);
  for (const [name, namespace] of Object.entries(attrs)) {
    if (name.startsWith("xmlns:") && RELATIONSHIP_NAMESPACES.has(namespace)) return { xml: String(xml), prefix: name.slice(6) };
  }
  return ensureNamespacePrefix(xml, rootLocalName, RELATIONSHIP_NAMESPACE, "r", false);
}

function relationshipId(tag = "") {
  return Object.entries(attributes(tag)).find(([name]) => /:(?:id|embed|link)$/.test(name))?.[1];
}

function removeReferenceTags(xml, localName, ids) {
  if (!ids.size) return String(xml);
  return String(xml).replace(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`, "g"), (tag) => ids.has(relationshipId(tag)) ? "" : tag);
}

function setAttribute(tag, name, value) {
  const escapedName = String(name).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const pattern = new RegExp(`(\\b${escapedName}\\s*=\\s*)(["'])(.*?)\\2`);
  if (pattern.test(tag)) return tag.replace(pattern, `$1"${attrEscape(value)}"`);
  return tag.replace(/\s*\/?>$/, (ending) => ` ${name}="${attrEscape(value)}"${ending}`);
}

function appendToRoot(xml, rootLocalName, content) {
  const root = rootTag(xml, rootLocalName);
  if (!root) throw new Error(`OOXML source reference could not find root element ${rootLocalName}.`);
  const prefix = root[1] || "";
  if (/\/\s*>$/.test(root[0])) {
    const opening = root[0].replace(/\/\s*>$/, ">");
    return `${String(xml).slice(0, root.index)}${opening}${content}</${qname(prefix, rootLocalName)}>${String(xml).slice(root.index + root[0].length)}`;
  }
  const closing = new RegExp(`</${prefix ? `${prefix}:` : ""}${rootLocalName}>`);
  if (!closing.test(String(xml))) throw new Error(`OOXML source reference could not find closing root element ${rootLocalName}.`);
  return String(xml).replace(closing, `${content}</${qname(prefix, rootLocalName)}>`);
}

function insertBeforeOrAppend(xml, rootLocalName, content, followingLocalNames = []) {
  if (followingLocalNames.length) {
    const alternatives = followingLocalNames.join("|");
    const following = new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?(?:${alternatives})\\b`).exec(String(xml));
    if (following) return `${String(xml).slice(0, following.index)}${content}${String(xml).slice(following.index)}`;
  }
  return appendToRoot(xml, rootLocalName, content);
}

function regexEscape(value) {
  return String(value).replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function wordprocessingId(tag = "") {
  return Object.entries(attributes(tag)).find(([name]) => name === "id" || name.endsWith(":id"))?.[1];
}

function removeDocxCommentAnchors(xml, commentIds) {
  const removeAll = commentIds == null;
  const ids = commentIds || new Set();
  const matchesId = (tag) => removeAll || ids.has(String(wordprocessingId(tag)));
  let next = String(xml).replace(/<(?:[A-Za-z_][\w.-]*:)?r\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?r>/g, (run) => {
    const references = [...run.matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentReference\b[^>]*\/?\s*>/g)];
    if (!references.some((match) => matchesId(match[0]))) return run;
    const withoutReferences = run.replace(/<(?:[A-Za-z_][\w.-]*:)?commentReference\b[^>]*\/?\s*>/g, (tag) => matchesId(tag) ? "" : tag);
    const substantive = withoutReferences
      .replace(/^<(?:[A-Za-z_][\w.-]*:)?r\b[^>]*>|<\/(?:[A-Za-z_][\w.-]*:)?r>$/g, "")
      .replace(/<(?:[A-Za-z_][\w.-]*:)?rPr\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?rPr>)/g, "")
      .trim();
    return substantive ? withoutReferences : "";
  });
  next = next.replace(/<(?:[A-Za-z_][\w.-]*:)?(?:commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/?\s*>/g, (tag) => matchesId(tag) ? "" : tag);
  return next;
}

function nonNegativeIndex(value, label) {
  const index = Number(value);
  if (!Number.isInteger(index) || index < 0) throw new RangeError(`DOCX comments sourceReference ${label} must be a non-negative integer.`);
  return index;
}

function indexedXmlMatch(xml, pattern, index, label) {
  const matches = [...String(xml).matchAll(pattern)];
  if (index >= matches.length) throw new RangeError(`DOCX comments sourceReference ${label} ${index} is out of range; found ${matches.length}.`);
  return matches[index];
}

function replaceXmlMatch(xml, match, replacement) {
  return `${String(xml).slice(0, match.index)}${replacement}${String(xml).slice(match.index + match[0].length)}`;
}

function anchorDocxParagraph(paragraphXml, prefix, commentId) {
  const paragraphName = qname(prefix, "p");
  let paragraph = String(paragraphXml);
  if (/\/\s*>$/.test(paragraph)) paragraph = `${paragraph.replace(/\/\s*>$/, ">")}</${paragraphName}>`;
  const opening = new RegExp(`^<${regexEscape(paragraphName)}\\b[^>]*>`).exec(paragraph)?.[0];
  const closing = `</${paragraphName}>`;
  if (!opening || !paragraph.endsWith(closing)) throw new Error("DOCX comments sourceReference target paragraph is malformed.");
  const start = `<${qname(prefix, "commentRangeStart")} ${prefix ? `${prefix}:` : ""}id="${attrEscape(commentId)}"/>`;
  const end = `<${qname(prefix, "commentRangeEnd")} ${prefix ? `${prefix}:` : ""}id="${attrEscape(commentId)}"/>`;
  const reference = `<${qname(prefix, "r")}><${qname(prefix, "rPr")}><${qname(prefix, "rStyle")} ${prefix ? `${prefix}:` : ""}val="CommentReference"/></${qname(prefix, "rPr")}><${qname(prefix, "commentReference")} ${prefix ? `${prefix}:` : ""}id="${attrEscape(commentId)}"/></${qname(prefix, "r")}>`;
  const properties = new RegExp(`^<${regexEscape(paragraphName)}\\b[^>]*><${regexEscape(qname(prefix, "pPr"))}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${regexEscape(qname(prefix, "pPr"))}>)`).exec(paragraph)?.[0];
  const insertAt = properties?.length || opening.length;
  paragraph = `${paragraph.slice(0, insertAt)}${start}${paragraph.slice(insertAt)}`;
  return paragraph.replace(new RegExp(`${regexEscape(closing)}$`), `${end}${reference}${closing}`);
}

function addDocxCommentAnchor(xml, config = {}) {
  const commentId = config.commentId ?? config.id;
  if (commentId === undefined || commentId === null || String(commentId) === "") throw new Error("DOCX comments sourceReference commentId is required.");
  const normalizedId = String(commentId);
  if (!/^-?\d+$/.test(normalizedId) || Number(normalizedId) < 0) throw new Error("DOCX comments sourceReference commentId must be a non-negative integer.");
  const existingIds = new Set([...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/?\s*>/g)].map((match) => String(wordprocessingId(match[0]))));
  if (existingIds.has(normalizedId)) throw new Error(`DOCX comments sourceReference commentId ${normalizedId} already exists.`);
  const prefix = rootPrefix(xml, "document");
  const paragraphName = regexEscape(qname(prefix, "p"));
  const tableName = regexEscape(qname(prefix, "tbl"));
  const rowName = regexEscape(qname(prefix, "tr"));
  const cellName = regexEscape(qname(prefix, "tc"));
  const target = config.target && typeof config.target === "object" ? config.target : config;
  const type = String(target.type || target.kind || (target.tableIndex !== undefined ? "tableCell" : target.paragraphIndex !== undefined ? "paragraph" : "block")).toLowerCase().replace(/[^a-z]/g, "");
  if (type === "paragraph") {
    const index = nonNegativeIndex(target.index ?? target.paragraphIndex, "paragraphIndex");
    const paragraph = indexedXmlMatch(xml, new RegExp(`<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g"), index, "paragraphIndex");
    return replaceXmlMatch(xml, paragraph, anchorDocxParagraph(paragraph[0], prefix, normalizedId));
  }
  if (type === "block") {
    const index = nonNegativeIndex(target.index ?? target.blockIndex, "blockIndex");
    const block = indexedXmlMatch(xml, new RegExp(`<${tableName}\\b[^>]*>[\\s\\S]*?</${tableName}>|<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g"), index, "blockIndex");
    if (new RegExp(`^<${paragraphName}\\b`).test(block[0])) return replaceXmlMatch(xml, block, anchorDocxParagraph(block[0], prefix, normalizedId));
    const paragraph = indexedXmlMatch(block[0], new RegExp(`<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g"), 0, `blockIndex ${index} paragraphIndex`);
    return replaceXmlMatch(xml, block, replaceXmlMatch(block[0], paragraph, anchorDocxParagraph(paragraph[0], prefix, normalizedId)));
  }
  if (type === "tablecell") {
    const tableIndex = nonNegativeIndex(target.tableIndex, "tableIndex");
    const rowIndex = nonNegativeIndex(target.rowIndex ?? target.row, "rowIndex");
    const columnIndex = nonNegativeIndex(target.columnIndex ?? target.column ?? target.col, "columnIndex");
    const table = indexedXmlMatch(xml, new RegExp(`<${tableName}\\b[^>]*>[\\s\\S]*?</${tableName}>`, "g"), tableIndex, "tableIndex");
    const row = indexedXmlMatch(table[0], new RegExp(`<${rowName}\\b[^>]*>[\\s\\S]*?</${rowName}>`, "g"), rowIndex, `tableIndex ${tableIndex} rowIndex`);
    const cell = indexedXmlMatch(row[0], new RegExp(`<${cellName}\\b[^>]*>[\\s\\S]*?</${cellName}>`, "g"), columnIndex, `tableIndex ${tableIndex} rowIndex ${rowIndex} columnIndex`);
    const paragraph = indexedXmlMatch(cell[0], new RegExp(`<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g"), 0, `tableIndex ${tableIndex} rowIndex ${rowIndex} columnIndex ${columnIndex} paragraphIndex`);
    const nextCell = replaceXmlMatch(cell[0], paragraph, anchorDocxParagraph(paragraph[0], prefix, normalizedId));
    const nextRow = replaceXmlMatch(row[0], cell, nextCell);
    return replaceXmlMatch(xml, table, replaceXmlMatch(table[0], row, nextRow));
  }
  throw new Error("DOCX comments sourceReference target type must be block, paragraph, or tableCell.");
}

function mutateDocxCommentReferences(xml, addId, config = {}) {
  const anchors = Array.isArray(config.anchors) ? config.anchors : [config];
  if (!addId) return removeDocxCommentAnchors(xml);
  if (!anchors.length) throw new Error("DOCX comments sourceReference anchors must contain at least one anchor.");
  return anchors.reduce((next, anchor) => addDocxCommentAnchor(next, anchor), String(xml));
}

function mutateDocxSectionReference(xml, kind, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "document");
  const tagName = qname(prefix, `${kind}Reference`);
  let next = removeReferenceTags(xml, `${kind}Reference`, ids);
  if (!addId) return next;
  const referenceType = String(config.type || config.referenceType || "default");
  if (!new Set(["default", "first", "even"]).has(referenceType)) throw new Error(`DOCX ${kind} sourceReference type must be default, first, or even.`);
  const ensured = ensureRelationshipPrefix(next, "document");
  next = ensured.xml;
  const referenceTag = `<${tagName} ${prefix ? `${prefix}:` : ""}type="${referenceType}" ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const sectionName = qname(prefix, "sectPr");
  const sections = [...next.matchAll(new RegExp(`<${sectionName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${sectionName}>)`, "g"))];
  const requestedIndex = config.sectionIndex === undefined ? sections.length - 1 : Number(config.sectionIndex);
  if (sections.length) {
    if (!Number.isInteger(requestedIndex) || requestedIndex < 0 || requestedIndex >= sections.length) throw new RangeError(`DOCX ${kind} sourceReference sectionIndex must be an integer from 0 through ${sections.length - 1}.`);
    const section = sections[requestedIndex][0];
    const expanded = section.endsWith("/>") ? `${section.replace(/\/>$/, ">")}</${sectionName}>` : section;
    const withoutSameType = expanded.replace(new RegExp(`<${tagName}\\b[^>]*\\/?>`, "g"), (tag) => attributes(tag)[`${prefix ? `${prefix}:` : ""}type`] === referenceType ? "" : tag);
    const titlePageName = qname(prefix, "titlePg");
    const titlePage = referenceType === "first" && !new RegExp(`<${titlePageName}\\b`).test(withoutSameType) ? `<${titlePageName}/>` : "";
    const updated = withoutSameType.replace(new RegExp(`^<${sectionName}\\b[^>]*>`), (opening) => `${opening}${referenceTag}${titlePage}`);
    return `${next.slice(0, sections[requestedIndex].index)}${updated}${next.slice(sections[requestedIndex].index + section.length)}`;
  }
  if (config.sectionIndex !== undefined && Number(config.sectionIndex) !== 0) throw new RangeError(`DOCX ${kind} sourceReference sectionIndex must be 0 when the document has no existing w:sectPr.`);
  const bodyClosing = `</${qname(prefix, "body")}>`;
  if (!next.includes(bodyClosing)) throw new Error("DOCX header/footer sourceReference requires w:body or w:sectPr.");
  return next.replace(bodyClosing, `<${sectionName}>${referenceTag}${referenceType === "first" ? `<${qname(prefix, "titlePg")}/>` : ""}</${sectionName}>${bodyClosing}`);
}

function mutateXlsxCountedReference(xml, rootLocalName, containerLocalName, itemLocalName, ids, addId) {
  const prefix = rootPrefix(xml, rootLocalName);
  let next = removeReferenceTags(xml, itemLocalName, ids);
  const containerName = qname(prefix, containerLocalName);
  if (addId) {
    const ensured = ensureRelationshipPrefix(next, rootLocalName);
    next = ensured.xml;
    const itemTag = `<${qname(prefix, itemLocalName)} ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
    const blockPattern = new RegExp(`<${containerName}\\b[^>]*>[\\s\\S]*?</${containerName}>`);
    const block = blockPattern.exec(next)?.[0];
    if (block) next = next.replace(block, block.replace(new RegExp(`</${containerName}>$`), `${itemTag}</${containerName}>`));
    else if (new RegExp(`<${containerName}\\b[^>]*\\/>`).test(next)) next = next.replace(new RegExp(`<${containerName}\\b[^>]*\\/>`), (tag) => `${tag.replace(/\/>$/, ">")}${itemTag}</${containerName}>`);
    else next = insertBeforeOrAppend(next, rootLocalName, `<${containerName} count="1">${itemTag}</${containerName}>`, ["extLst"]);
  }
  const blockPattern = new RegExp(`<${containerName}\\b[^>]*>[\\s\\S]*?</${containerName}>`);
  const block = blockPattern.exec(next)?.[0];
  if (!block) return next;
  const count = [...block.matchAll(new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${itemLocalName}\\b[^>]*\\/?>`, "g"))].length;
  if (!count) return next.replace(block, "");
  const opening = new RegExp(`^<${containerName}\\b[^>]*>`).exec(block)?.[0];
  return opening ? next.replace(block, block.replace(opening, setAttribute(opening, "count", count))) : next;
}

function mutateXlsxTableReference(xml, ids, addId) {
  return mutateXlsxCountedReference(xml, "worksheet", "tableParts", "tablePart", ids, addId);
}

function mutateXlsxPivotCacheReference(xml, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "workbook");
  let next = removeReferenceTags(xml, "pivotCache", ids);
  const containerName = qname(prefix, "pivotCaches");
  if (addId) {
    const cacheId = Number(config.cacheId);
    if (!Number.isInteger(cacheId) || cacheId < 0) throw new Error("XLSX pivotCacheDefinition sourceReference cacheId must be a non-negative integer.");
    const existing = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?pivotCache\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).cacheId)).filter(Number.isFinite);
    if (existing.includes(cacheId)) throw new Error(`XLSX pivotCacheDefinition sourceReference cacheId ${cacheId} already exists.`);
    const ensured = ensureRelationshipPrefix(next, "workbook");
    next = ensured.xml;
    const cacheTag = `<${qname(prefix, "pivotCache")} cacheId="${cacheId}" ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
    const block = new RegExp(`<${containerName}\\b[^>]*>[\\s\\S]*?</${containerName}>`).exec(next)?.[0];
    if (block) next = next.replace(block, block.replace(new RegExp(`</${containerName}>$`), `${cacheTag}</${containerName}>`));
    else if (new RegExp(`<${containerName}\\b[^>]*\\/>`).test(next)) next = next.replace(new RegExp(`<${containerName}\\b[^>]*\\/>`), (tag) => `${tag.replace(/\/>$/, ">")}${cacheTag}</${containerName}>`);
    else next = insertBeforeOrAppend(next, "workbook", `<${containerName}>${cacheTag}</${containerName}>`, ["smartTagPr", "smartTagTypes", "webPublishing", "fileRecoveryPr", "webPublishObjects", "extLst"]);
  }
  const block = new RegExp(`<${containerName}\\b[^>]*>[\\s\\S]*?</${containerName}>`).exec(next)?.[0];
  if (!block) return next;
  return /<(?:[A-Za-z_][\w.-]*:)?pivotCache\b[^>]*\/?>/.test(block) ? next : next.replace(block, "");
}

function mutateXlsxPivotCacheRecordsReference(xml, ids, addId) {
  let next = String(xml);
  let root = rootTag(next, "pivotCacheDefinition");
  if (!root) throw new Error("OOXML source reference could not find root element pivotCacheDefinition.");
  const attrs = attributes(root[0]);
  let rootText = root[0];
  for (const [name, value] of Object.entries(attrs)) {
    if (/:id$/.test(name) && ids.has(value)) rootText = rootText.replace(new RegExp(`\\s+${name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\s*=\\s*(["']).*?\\1`), "");
  }
  if (rootText !== root[0]) next = `${next.slice(0, root.index)}${rootText}${next.slice(root.index + root[0].length)}`;
  if (!addId) return next;
  const ensured = ensureRelationshipPrefix(next, "pivotCacheDefinition");
  next = ensured.xml;
  root = rootTag(next, "pivotCacheDefinition");
  const updated = setAttribute(root[0], `${ensured.prefix}:id`, addId);
  return `${next.slice(0, root.index)}${updated}${next.slice(root.index + root[0].length)}`;
}

function mutateXlsxWorksheetReference(xml, ids, addId, config = {}) {
  const prefix = rootPrefix(xml, "workbook");
  let next = removeReferenceTags(xml, "sheet", ids);
  if (!addId) return next;
  const name = String(config.name || config.sheetName || "").trim();
  if (!name) throw new Error("XLSX worksheet sourceReference requires name or sheetName.");
  const state = config.state == null ? undefined : String(config.state);
  if (state && !new Set(["visible", "hidden", "veryHidden"]).has(state)) throw new Error("XLSX worksheet sourceReference state must be visible, hidden, or veryHidden.");
  const existingSheets = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?sheet\b[^>]*\/?>/g)].map((match) => attributes(match[0]));
  const existingIds = existingSheets.map((item) => Number(item.sheetId)).filter(Number.isFinite);
  const sheetId = Number(config.sheetId ?? Math.max(0, ...existingIds) + 1);
  if (!Number.isInteger(sheetId) || sheetId < 1) throw new Error("XLSX worksheet sourceReference sheetId must be a positive integer.");
  if (existingIds.includes(sheetId)) throw new Error(`XLSX worksheet sourceReference sheetId ${sheetId} already exists.`);
  if (existingSheets.some((item) => String(item.name || "").toLowerCase() === name.toLowerCase())) throw new Error(`XLSX worksheet sourceReference name ${name} already exists.`);
  const ensured = ensureRelationshipPrefix(next, "workbook");
  next = ensured.xml;
  const sheetTag = `<${qname(prefix, "sheet")} name="${attrEscape(name)}" sheetId="${sheetId}"${state ? ` state="${state}"` : ""} ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const sheetsName = qname(prefix, "sheets");
  const block = new RegExp(`<${sheetsName}\\b[^>]*>[\\s\\S]*?</${sheetsName}>`).exec(next)?.[0];
  if (block) return next.replace(block, block.replace(new RegExp(`</${sheetsName}>$`), `${sheetTag}</${sheetsName}>`));
  if (new RegExp(`<${sheetsName}\\b[^>]*\\/>`).test(next)) return next.replace(new RegExp(`<${sheetsName}\\b[^>]*\\/>`), (tag) => `${tag.replace(/\/>$/, ">")}${sheetTag}</${sheetsName}>`);
  return appendToRoot(next, "workbook", `<${sheetsName}>${sheetTag}</${sheetsName}>`);
}

function mutatePptxSlideReference(xml, ids, addId, config = {}) {
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

function mutatePptxMasterReference(xml, ids, addId, config = {}) {
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

function mutatePptxLayoutReference(xml, ids, addId, config = {}) {
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

function mutateXlsxDrawingReference(xml, ids, addId) {
  const prefix = rootPrefix(xml, "worksheet");
  let next = removeReferenceTags(xml, "drawing", ids);
  if (!addId) return next;
  if (/<(?:[A-Za-z_][\w.-]*:)?drawing\b[^>]*\/?>/.test(next)) throw new Error("XLSX drawing sourceReference requires a worksheet without another drawing reference.");
  const ensured = ensureRelationshipPrefix(next, "worksheet");
  next = ensured.xml;
  const tag = `<${qname(prefix, "drawing")} ${ensured.prefix}:id="${attrEscape(addId)}"/>`;
  const following = /<(?:[A-Za-z_][\w.-]*:)?(?:legacyDrawing|legacyDrawingHF|picture|oleObjects|controls|webPublishItems|tableParts|extLst)\b/.exec(next);
  if (following) return `${next.slice(0, following.index)}${tag}${next.slice(following.index)}`;
  return appendToRoot(next, "worksheet", tag);
}

function finiteNumber(value, label, { positive = false, integer = false } = {}) {
  const number = Number(value);
  if (!Number.isFinite(number) || (positive && number <= 0) || (integer && !Number.isInteger(number))) throw new Error(`XLSX drawing sourceReference ${label} must be ${positive ? "a positive" : "a finite"}${integer ? " integer" : " number"}.`);
  return number;
}

function markerXml(prefix, name, marker = {}) {
  const col = finiteNumber(marker.col, `${name}.col`, { integer: true });
  const row = finiteNumber(marker.row, `${name}.row`, { integer: true });
  if (col < 0 || row < 0) throw new Error(`XLSX drawing sourceReference ${name} row/col must be non-negative integers.`);
  const colOff = Math.round(finiteNumber(marker.colOffsetPx ?? 0, `${name}.colOffsetPx`) * 9525);
  const rowOff = Math.round(finiteNumber(marker.rowOffsetPx ?? 0, `${name}.rowOffsetPx`) * 9525);
  return `<${qname(prefix, name)}><${qname(prefix, "col")}>${col}</${qname(prefix, "col")}><${qname(prefix, "colOff")}>${colOff}</${qname(prefix, "colOff")}><${qname(prefix, "row")}>${row}</${qname(prefix, "row")}><${qname(prefix, "rowOff")}>${rowOff}</${qname(prefix, "rowOff")}></${qname(prefix, name)}>`;
}

function drawingAnchorXml(prefixes, kind, relationship, config, objectId) {
  const anchor = config.anchor;
  if (!anchor || typeof anchor !== "object") throw new Error(`XLSX ${kind} sourceReference requires an explicit anchor.`);
  const type = String(anchor.type || anchor.anchorType || "oneCell").toLowerCase().replace(/[^a-z]/g, "");
  const widthPx = finiteNumber(anchor.extent?.widthPx ?? anchor.widthPx, "anchor.extent.widthPx", { positive: true });
  const heightPx = finiteNumber(anchor.extent?.heightPx ?? anchor.heightPx, "anchor.extent.heightPx", { positive: true });
  const cx = Math.round(widthPx * 9525);
  const cy = Math.round(heightPx * 9525);
  let opening;
  let position;
  if (type === "absolute" || type === "absoluteanchor") {
    const leftPx = finiteNumber(anchor.position?.leftPx ?? anchor.leftPx, "anchor.position.leftPx");
    const topPx = finiteNumber(anchor.position?.topPx ?? anchor.topPx, "anchor.position.topPx");
    opening = qname(prefixes.xdr, "absoluteAnchor");
    position = `<${qname(prefixes.xdr, "pos")} x="${Math.round(leftPx * 9525)}" y="${Math.round(topPx * 9525)}"/><${qname(prefixes.xdr, "ext")} cx="${cx}" cy="${cy}"/>`;
  } else if (type === "twocell" || type === "twocellanchor") {
    opening = qname(prefixes.xdr, "twoCellAnchor");
    position = `${markerXml(prefixes.xdr, "from", anchor.from)}${markerXml(prefixes.xdr, "to", anchor.to)}`;
  } else if (type === "onecell" || type === "onecellanchor") {
    opening = qname(prefixes.xdr, "oneCellAnchor");
    position = `${markerXml(prefixes.xdr, "from", anchor.from)}<${qname(prefixes.xdr, "ext")} cx="${cx}" cy="${cy}"/>`;
  } else throw new Error("XLSX drawing sourceReference anchor type must be oneCell, twoCell, or absolute.");
  const name = attrEscape(config.name || `${kind === "image" ? "Image" : "Chart"} ${objectId}`);
  const object = kind === "image"
    ? `<${qname(prefixes.xdr, "pic")}><${qname(prefixes.xdr, "nvPicPr")}><${qname(prefixes.xdr, "cNvPr")} id="${objectId}" name="${name}" descr="${attrEscape(config.alt || config.description || "")}"/><${qname(prefixes.xdr, "cNvPicPr")}/></${qname(prefixes.xdr, "nvPicPr")}><${qname(prefixes.xdr, "blipFill")}><${qname(prefixes.a, "blip")} ${prefixes.r}:embed="${attrEscape(relationship)}"/><${qname(prefixes.a, "stretch")}><${qname(prefixes.a, "fillRect")}/></${qname(prefixes.a, "stretch")}></${qname(prefixes.xdr, "blipFill")}><${qname(prefixes.xdr, "spPr")}><${qname(prefixes.a, "prstGeom")} prst="rect"><${qname(prefixes.a, "avLst")}/></${qname(prefixes.a, "prstGeom")}></${qname(prefixes.xdr, "spPr")}></${qname(prefixes.xdr, "pic")}>`
    : `<${qname(prefixes.xdr, "graphicFrame")}><${qname(prefixes.xdr, "nvGraphicFramePr")}><${qname(prefixes.xdr, "cNvPr")} id="${objectId}" name="${name}"/><${qname(prefixes.xdr, "cNvGraphicFramePr")}/></${qname(prefixes.xdr, "nvGraphicFramePr")}><${qname(prefixes.xdr, "xfrm")}><${qname(prefixes.a, "off")} x="0" y="0"/><${qname(prefixes.a, "ext")} cx="${cx}" cy="${cy}"/></${qname(prefixes.xdr, "xfrm")}><${qname(prefixes.a, "graphic")}><${qname(prefixes.a, "graphicData")} uri="${CHART_NAMESPACE}"><${qname(prefixes.c, "chart")} ${prefixes.r}:id="${attrEscape(relationship)}"/></${qname(prefixes.a, "graphicData")}></${qname(prefixes.a, "graphic")}></${qname(prefixes.xdr, "graphicFrame")}>`;
  return `<${opening}>${position}${object}<${qname(prefixes.xdr, "clientData")}/></${opening}>`;
}

function removeDrawingAnchors(xml, ids) {
  if (!ids.size) return String(xml);
  return String(xml).replace(/<(?:[A-Za-z_][\w.-]*:)?(oneCellAnchor|twoCellAnchor|absoluteAnchor)\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?\1>/g, (anchor) => {
    const references = [...anchor.matchAll(/\b[A-Za-z_][\w.-]*:(?:id|embed|link)\s*=\s*(["'])(.*?)\1/g)].map((match) => decodeXml(match[2]));
    return references.some((id) => ids.has(id)) ? "" : anchor;
  });
}

function mutateXlsxDrawingObject(xml, kind, ids, addId, config = {}) {
  let next = removeDrawingAnchors(xml, ids);
  if (!addId) return next;
  const existingObjectIds = [...next.matchAll(/<(?:[A-Za-z_][\w.-]*:)?cNvPr\b[^>]*\/?>/g)].map((match) => Number(attributes(match[0]).id)).filter(Number.isFinite);
  const objectId = Number(config.objectId ?? Math.max(1, ...existingObjectIds) + 1);
  if (!Number.isInteger(objectId) || objectId < 1) throw new Error(`XLSX ${kind} sourceReference objectId must be a positive integer.`);
  if (existingObjectIds.includes(objectId)) throw new Error(`XLSX ${kind} sourceReference objectId ${objectId} already exists.`);
  let ensured = ensureNamespacePrefix(next, "wsDr", SPREADSHEET_DRAWING_NAMESPACE, "xdr");
  next = ensured.xml;
  const prefixes = { xdr: ensured.prefix || rootPrefix(next, "wsDr") };
  ensured = ensureNamespacePrefix(next, "wsDr", DRAWING_NAMESPACE, "a", false);
  next = ensured.xml;
  prefixes.a = ensured.prefix;
  ensured = ensureRelationshipPrefix(next, "wsDr");
  next = ensured.xml;
  prefixes.r = ensured.prefix;
  if (kind === "chart") {
    ensured = ensureNamespacePrefix(next, "wsDr", CHART_NAMESPACE, "c", false);
    next = ensured.xml;
    prefixes.c = ensured.prefix;
  }
  return appendToRoot(next, "wsDr", drawingAnchorXml(prefixes, kind, addId, config, objectId));
}

export function supportsOoxmlSourceReference(family, recipeKind) {
  return SUPPORTED.has(`${family}:${recipeKind}`);
}

export function supportedOoxmlSourceReferenceSummary() {
  return "DOCX header/footer/comments, XLSX worksheet/table/drawing/image/chart/pivotCacheDefinition/pivotCacheRecords, PPTX slide/slideMaster/slideLayout";
}

export function validateOoxmlSourceReferenceTarget({ family, recipeKind, targetXml, config = {} }) {
  if (family !== "DOCX" || recipeKind !== "comments") return;
  const ids = [...String(targetXml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?comment\b[^>]*>/g)].map((match) => wordprocessingId(match[0])).filter((id) => id !== undefined).map(String);
  const duplicates = ids.filter((id, index) => ids.indexOf(id) !== index);
  if (duplicates.length) throw new Error(`DOCX comments sourceReference target part has duplicate commentId ${duplicates[0]}.`);
  const declared = new Set(ids);
  const anchors = Array.isArray(config.anchors) ? config.anchors : [config];
  for (const anchor of anchors) {
    const commentId = anchor?.commentId ?? anchor?.id;
    if (commentId === undefined || commentId === null || String(commentId) === "") throw new Error("DOCX comments sourceReference commentId is required.");
    if (!declared.has(String(commentId))) throw new Error(`DOCX comments sourceReference commentId ${commentId} is not declared in the Comments part.`);
  }
}

export function mutateOoxmlSourceReference({ family, recipeKind, xml, relationshipIds = new Set(), addId, config = {} }) {
  const key = `${family}:${recipeKind}`;
  if (!SUPPORTED.has(key)) throw new Error(`${family} sourceReference is not supported for recipe ${recipeKind || "(missing)"}. Supported recipes: ${supportedOoxmlSourceReferenceSummary()}.`);
  if (family === "DOCX" && recipeKind === "comments") return mutateDocxCommentReferences(xml, addId, config);
  if (family === "DOCX") return mutateDocxSectionReference(xml, recipeKind, relationshipIds, addId, config);
  if (family === "PPTX" && recipeKind === "slide") return mutatePptxSlideReference(xml, relationshipIds, addId, config);
  if (family === "PPTX" && recipeKind === "slidemaster") return mutatePptxMasterReference(xml, relationshipIds, addId, config);
  if (family === "PPTX") return mutatePptxLayoutReference(xml, relationshipIds, addId, config);
  if (recipeKind === "worksheet") return mutateXlsxWorksheetReference(xml, relationshipIds, addId, config);
  if (recipeKind === "table") return mutateXlsxTableReference(xml, relationshipIds, addId);
  if (recipeKind === "drawing") return mutateXlsxDrawingReference(xml, relationshipIds, addId);
  if (recipeKind === "pivotcachedefinition") return mutateXlsxPivotCacheReference(xml, relationshipIds, addId, config);
  if (recipeKind === "pivotcacherecords") return mutateXlsxPivotCacheRecordsReference(xml, relationshipIds, addId);
  return mutateXlsxDrawingObject(xml, recipeKind, relationshipIds, addId, config);
}
