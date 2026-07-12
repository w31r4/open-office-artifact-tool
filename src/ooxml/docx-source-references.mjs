import { attrEscape, attributes, ensureRelationshipPrefix, qname, regexEscape, removeReferenceTags, rootPrefix } from "./source-reference-xml.mjs";

const SUPPORTED = new Set(["header", "footer", "comments", "numbering"]);

function wordprocessingId(tag = "") {
  return Object.entries(attributes(tag)).find(([name]) => name === "id" || name.endsWith(":id"))?.[1];
}

function removeCommentAnchors(xml) {
  let next = String(xml).replace(/<(?:[A-Za-z_][\w.-]*:)?r\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?r>/g, (run) => {
    const references = [...run.matchAll(/<(?:[A-Za-z_][\w.-]*:)?commentReference\b[^>]*\/?\s*>/g)];
    if (!references.length) return run;
    const withoutReferences = run.replace(/<(?:[A-Za-z_][\w.-]*:)?commentReference\b[^>]*\/?\s*>/g, "");
    const substantive = withoutReferences
      .replace(/^<(?:[A-Za-z_][\w.-]*:)?r\b[^>]*>|<\/(?:[A-Za-z_][\w.-]*:)?r>$/g, "")
      .replace(/<(?:[A-Za-z_][\w.-]*:)?rPr\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?rPr>)/g, "")
      .trim();
    return substantive ? withoutReferences : "";
  });
  next = next.replace(/<(?:[A-Za-z_][\w.-]*:)?(?:commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/?\s*>/g, "");
  return next;
}

function nonNegativeIndex(value, label) {
  const index = Number(value);
  if (!Number.isInteger(index) || index < 0) throw new RangeError(`DOCX sourceReference ${label} must be a non-negative integer.`);
  return index;
}

function indexedXmlMatch(xml, pattern, index, label) {
  const matches = [...String(xml).matchAll(pattern)];
  if (index >= matches.length) throw new RangeError(`DOCX sourceReference ${label} ${index} is out of range; found ${matches.length}.`);
  return matches[index];
}

function replaceXmlMatch(xml, match, replacement) {
  return `${String(xml).slice(0, match.index)}${replacement}${String(xml).slice(match.index + match[0].length)}`;
}

function anchorParagraph(paragraphXml, prefix, commentId) {
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

function targetDescriptor(rawTarget = {}) {
  const target = rawTarget && typeof rawTarget === "object" ? rawTarget : {};
  const type = String(target.type || target.kind || (target.tableIndex !== undefined ? "tableCell" : target.paragraphIndex !== undefined ? "paragraph" : "block")).toLowerCase().replace(/[^a-z]/g, "");
  if (type === "paragraph") return { type, paragraphIndex: nonNegativeIndex(target.index ?? target.paragraphIndex, "paragraphIndex") };
  if (type === "block") return { type, blockIndex: nonNegativeIndex(target.index ?? target.blockIndex, "blockIndex") };
  if (type === "tablecell") return { type, tableIndex: nonNegativeIndex(target.tableIndex, "tableIndex"), rowIndex: nonNegativeIndex(target.rowIndex ?? target.row, "rowIndex"), columnIndex: nonNegativeIndex(target.columnIndex ?? target.column ?? target.col, "columnIndex") };
  throw new Error("DOCX sourceReference target type must be block, paragraph, or tableCell.");
}

function targetKey(target) {
  return JSON.stringify(targetDescriptor(target));
}

function mutateTargetParagraph(xml, rawTarget, mutateParagraph) {
  const target = targetDescriptor(rawTarget);
  const prefix = rootPrefix(xml, "document");
  const paragraphName = regexEscape(qname(prefix, "p"));
  const tableName = regexEscape(qname(prefix, "tbl"));
  const rowName = regexEscape(qname(prefix, "tr"));
  const cellName = regexEscape(qname(prefix, "tc"));
  const paragraphPattern = () => new RegExp(`<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g");
  if (target.type === "paragraph") {
    const paragraph = indexedXmlMatch(xml, paragraphPattern(), target.paragraphIndex, "paragraphIndex");
    return replaceXmlMatch(xml, paragraph, mutateParagraph(paragraph[0], prefix));
  }
  if (target.type === "block") {
    const block = indexedXmlMatch(xml, new RegExp(`<${tableName}\\b[^>]*>[\\s\\S]*?</${tableName}>|<${paragraphName}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${paragraphName}>)`, "g"), target.blockIndex, "blockIndex");
    if (new RegExp(`^<${paragraphName}\\b`).test(block[0])) return replaceXmlMatch(xml, block, mutateParagraph(block[0], prefix));
    const paragraph = indexedXmlMatch(block[0], paragraphPattern(), 0, `blockIndex ${target.blockIndex} paragraphIndex`);
    return replaceXmlMatch(xml, block, replaceXmlMatch(block[0], paragraph, mutateParagraph(paragraph[0], prefix)));
  }
  const table = indexedXmlMatch(xml, new RegExp(`<${tableName}\\b[^>]*>[\\s\\S]*?</${tableName}>`, "g"), target.tableIndex, "tableIndex");
  const row = indexedXmlMatch(table[0], new RegExp(`<${rowName}\\b[^>]*>[\\s\\S]*?</${rowName}>`, "g"), target.rowIndex, `tableIndex ${target.tableIndex} rowIndex`);
  const cell = indexedXmlMatch(row[0], new RegExp(`<${cellName}\\b[^>]*>[\\s\\S]*?</${cellName}>`, "g"), target.columnIndex, `tableIndex ${target.tableIndex} rowIndex ${target.rowIndex} columnIndex`);
  const paragraph = indexedXmlMatch(cell[0], paragraphPattern(), 0, `tableIndex ${target.tableIndex} rowIndex ${target.rowIndex} columnIndex ${target.columnIndex} paragraphIndex`);
  const nextCell = replaceXmlMatch(cell[0], paragraph, mutateParagraph(paragraph[0], prefix));
  const nextRow = replaceXmlMatch(row[0], cell, nextCell);
  return replaceXmlMatch(xml, table, replaceXmlMatch(table[0], row, nextRow));
}

function addCommentAnchor(xml, config = {}) {
  const commentId = config.commentId ?? config.id;
  if (commentId === undefined || commentId === null || String(commentId) === "") throw new Error("DOCX comments sourceReference commentId is required.");
  const normalizedId = String(commentId);
  if (!/^-?\d+$/.test(normalizedId) || Number(normalizedId) < 0) throw new Error("DOCX comments sourceReference commentId must be a non-negative integer.");
  const existingIds = new Set([...String(xml).matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:commentRangeStart|commentRangeEnd|commentReference)\b[^>]*\/?\s*>/g)].map((match) => String(wordprocessingId(match[0]))));
  if (existingIds.has(normalizedId)) throw new Error(`DOCX comments sourceReference commentId ${normalizedId} already exists.`);
  const target = config.target && typeof config.target === "object" ? config.target : config;
  return mutateTargetParagraph(xml, target, (paragraph, prefix) => anchorParagraph(paragraph, prefix, normalizedId));
}

function mutateCommentReferences(xml, addId, config = {}) {
  const anchors = Array.isArray(config.anchors) ? config.anchors : [config];
  if (!addId) return removeCommentAnchors(xml);
  if (!anchors.length) throw new Error("DOCX comments sourceReference anchors must contain at least one anchor.");
  return anchors.reduce((next, anchor) => addCommentAnchor(next, anchor), String(xml));
}

function numberingAssignments(config = {}) {
  return Array.isArray(config.assignments) ? config.assignments : [config];
}

function normalizeNumberingAssignment(assignment = {}) {
  const numId = String(assignment.numId ?? assignment.numberingId ?? "");
  if (!/^\d+$/.test(numId)) throw new Error("DOCX numbering sourceReference numId must be a non-negative integer.");
  const level = Number(assignment.level ?? assignment.ilvl ?? 0);
  if (!Number.isInteger(level) || level < 0 || level > 8) throw new RangeError("DOCX numbering sourceReference level must be an integer from 0 through 8.");
  const target = assignment.target && typeof assignment.target === "object" ? assignment.target : assignment;
  return { numId, level, target };
}

function removeNumberingProperties(xml) {
  let next = String(xml).replace(/<(?:[A-Za-z_][\w.-]*:)?numPr\b[^>]*(?:\/>|>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?numPr>)/g, "");
  next = next.replace(/<(?:[A-Za-z_][\w.-]*:)?pPr\b[^>]*>\s*<\/(?:[A-Za-z_][\w.-]*:)?pPr>/g, "");
  return next;
}

function numberParagraph(paragraphXml, prefix, numId, level) {
  const paragraphName = qname(prefix, "p");
  const propertiesName = qname(prefix, "pPr");
  let paragraph = String(paragraphXml);
  if (/\/\s*>$/.test(paragraph)) paragraph = `${paragraph.replace(/\/\s*>$/, ">")}</${paragraphName}>`;
  const opening = new RegExp(`^<${regexEscape(paragraphName)}\\b[^>]*>`).exec(paragraph)?.[0];
  if (!opening) throw new Error("DOCX numbering sourceReference target paragraph is malformed.");
  const valueAttribute = prefix ? `${prefix}:val` : "val";
  const numPr = `<${qname(prefix, "numPr")}><${qname(prefix, "ilvl")} ${valueAttribute}="${level}"/><${qname(prefix, "numId")} ${valueAttribute}="${attrEscape(numId)}"/></${qname(prefix, "numPr")}>`;
  const propertiesPattern = new RegExp(`^<${regexEscape(paragraphName)}\\b[^>]*><${regexEscape(propertiesName)}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${regexEscape(propertiesName)}>)`);
  const propertiesMatch = propertiesPattern.exec(paragraph);
  if (!propertiesMatch) return `${paragraph.slice(0, opening.length)}<${propertiesName}>${numPr}</${propertiesName}>${paragraph.slice(opening.length)}`;
  const wholePrefix = propertiesMatch[0];
  const paragraphOpening = new RegExp(`^<${regexEscape(paragraphName)}\\b[^>]*>`).exec(wholePrefix)[0];
  let properties = wholePrefix.slice(paragraphOpening.length);
  if (/\/\s*>$/.test(properties)) properties = `${properties.replace(/\/\s*>$/, ">")}</${propertiesName}>`;
  properties = properties.replace(new RegExp(`<${regexEscape(qname(prefix, "numPr"))}\\b[^>]*(?:\\/>|>[\\s\\S]*?</${regexEscape(qname(prefix, "numPr"))}>)`), "");
  const followingNames = ["suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens", "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE", "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind", "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection", "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr", "sectPr", "pPrChange"];
  const following = new RegExp(`<${regexEscape(prefix ? `${prefix}:` : "")}(?:${followingNames.join("|")})\\b`).exec(properties);
  properties = following ? `${properties.slice(0, following.index)}${numPr}${properties.slice(following.index)}` : properties.replace(new RegExp(`</${regexEscape(propertiesName)}>$`), `${numPr}</${propertiesName}>`);
  return `${paragraphOpening}${properties}${paragraph.slice(wholePrefix.length)}`;
}

function mutateNumberingReferences(xml, addId, config = {}) {
  if (!addId) return removeNumberingProperties(xml);
  const assignments = numberingAssignments(config);
  if (!assignments.length) throw new Error("DOCX numbering sourceReference assignments must contain at least one assignment.");
  const normalized = assignments.map(normalizeNumberingAssignment);
  const targetKeys = normalized.map((assignment) => targetKey(assignment.target));
  const duplicateTarget = targetKeys.find((key, index) => targetKeys.indexOf(key) !== index);
  if (duplicateTarget) throw new Error(`DOCX numbering sourceReference target is assigned more than once: ${duplicateTarget}.`);
  return normalized.reduce((next, assignment) => mutateTargetParagraph(next, assignment.target, (paragraph, prefix) => numberParagraph(paragraph, prefix, assignment.numId, assignment.level)), String(xml));
}

function mutateSectionReference(xml, kind, ids, addId, config = {}) {
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

export function supportsDocxSourceReference(recipeKind) {
  return SUPPORTED.has(recipeKind);
}

export function validateDocxSourceReferenceTarget(recipeKind, targetXml, config = {}) {
  if (recipeKind === "comments") {
    const ids = [...String(targetXml || "").matchAll(/<(?:[A-Za-z_][\w.-]*:)?comment\b[^>]*>/g)].map((match) => wordprocessingId(match[0])).filter((id) => id !== undefined).map(String);
    const duplicate = ids.find((id, index) => ids.indexOf(id) !== index);
    if (duplicate) throw new Error(`DOCX comments sourceReference target part has duplicate commentId ${duplicate}.`);
    const declared = new Set(ids);
    for (const anchor of Array.isArray(config.anchors) ? config.anchors : [config]) {
      const commentId = anchor?.commentId ?? anchor?.id;
      if (commentId === undefined || commentId === null || String(commentId) === "") throw new Error("DOCX comments sourceReference commentId is required.");
      if (!declared.has(String(commentId))) throw new Error(`DOCX comments sourceReference commentId ${commentId} is not declared in the Comments part.`);
    }
    return;
  }
  if (recipeKind !== "numbering") return;
  const source = String(targetXml || "");
  const abstracts = new Map();
  for (const match of source.matchAll(/<(?:[A-Za-z_][\w.-]*:)?abstractNum\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?abstractNum>/g)) {
    const opening = /^<(?:[A-Za-z_][\w.-]*:)?abstractNum\b[^>]*>/.exec(match[0])?.[0] || "";
    const abstractId = Object.entries(attributes(opening)).find(([name]) => name === "abstractNumId" || name.endsWith(":abstractNumId"))?.[1];
    if (abstractId === undefined) continue;
    if (abstracts.has(String(abstractId))) throw new Error(`DOCX numbering sourceReference target part has duplicate abstractNumId ${abstractId}.`);
    const levels = new Set([...match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?lvl\b[^>]*>/g)].map((levelMatch) => Object.entries(attributes(levelMatch[0])).find(([name]) => name === "ilvl" || name.endsWith(":ilvl"))?.[1]).filter((level) => level !== undefined).map(String));
    abstracts.set(String(abstractId), levels);
  }
  const instances = new Map();
  for (const match of source.matchAll(/<(?:[A-Za-z_][\w.-]*:)?num\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?num>/g)) {
    const opening = /^<(?:[A-Za-z_][\w.-]*:)?num\b[^>]*>/.exec(match[0])?.[0] || "";
    const numId = Object.entries(attributes(opening)).find(([name]) => name === "numId" || name.endsWith(":numId"))?.[1];
    if (numId === undefined) continue;
    if (instances.has(String(numId))) throw new Error(`DOCX numbering sourceReference target part has duplicate numId ${numId}.`);
    const abstractTag = /<(?:[A-Za-z_][\w.-]*:)?abstractNumId\b[^>]*\/?\s*>/.exec(match[0])?.[0] || "";
    const abstractId = Object.entries(attributes(abstractTag)).find(([name]) => name === "val" || name.endsWith(":val"))?.[1];
    const levels = new Set(abstracts.get(String(abstractId)) || []);
    for (const override of match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?lvlOverride\b[^>]*>/g)) {
      const level = Object.entries(attributes(override[0])).find(([name]) => name === "ilvl" || name.endsWith(":ilvl"))?.[1];
      if (level !== undefined) levels.add(String(level));
    }
    instances.set(String(numId), { levels });
  }
  for (const assignment of numberingAssignments(config).map(normalizeNumberingAssignment)) {
    const instance = instances.get(assignment.numId);
    if (!instance) throw new Error(`DOCX numbering sourceReference numId ${assignment.numId} is not declared in the Numbering part.`);
    if (!instance.levels.has(String(assignment.level))) throw new Error(`DOCX numbering sourceReference level ${assignment.level} is not declared for numId ${assignment.numId}.`);
  }
}

export function mutateDocxSourceReference(recipeKind, xml, relationshipIds, addId, config = {}) {
  if (recipeKind === "comments") return mutateCommentReferences(xml, addId, config);
  if (recipeKind === "numbering") return mutateNumberingReferences(xml, addId, config);
  return mutateSectionReference(xml, recipeKind, relationshipIds, addId, config);
}
