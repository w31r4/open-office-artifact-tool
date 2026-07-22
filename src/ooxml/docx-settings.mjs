import { attrEscape, attributes, qname, rootTag } from "./source-reference-xml.mjs";

const SETTING_ORDER = `
writeProtection view zoom removePersonalInformation removeDateAndTime
doNotDisplayPageBoundaries displayBackgroundShape printPostScriptOverText
printFractionalCharacterWidth printFormsData embedTrueTypeFonts embedSystemFonts
saveSubsetFonts saveFormsData mirrorMargins alignBordersAndEdges
bordersDoNotSurroundHeader bordersDoNotSurroundFooter gutterAtTop
hideSpellingErrors hideGrammaticalErrors activeWritingStyle proofState formsDesign
attachedTemplate linkStyles stylePaneFormatFilter stylePaneSortMethod documentType
mailMerge revisionView trackRevisions doNotTrackMoves doNotTrackFormatting
documentProtection autoFormatOverride styleLockTheme styleLockQFSet defaultTabStop
autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps
showEnvelope summaryLength clickAndTypeStyle defaultTableStyle evenAndOddHeaders
bookFoldRevPrinting bookFoldPrinting bookFoldPrintingSheets drawingGridHorizontalSpacing
drawingGridVerticalSpacing displayHorizontalDrawingGrid displayVerticalDrawingGrid
doNotUseMarginsForDrawingGridOrigin drawingGridHorizontalOrigin drawingGridVerticalOrigin
doNotShadeFormData noPunctuationKerning characterSpacingControl printTwoOnOne
strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore savePreviewPicture
doNotValidateAgainstSchema saveInvalidXml ignoreMixedContent alwaysShowPlaceholderText
doNotDemarcateInvalidXml saveXmlDataOnly useXSLTWhenSaving saveThroughXslt
showXMLTags alwaysMergeEmptyNamespace updateFields hdrShapeDefaults footnotePr
endnotePr compat docVars rsids mathPr uiCompat97To2003 attachedSchema themeFontLang
clrSchemeMapping doNotIncludeSubdocsInStats doNotAutoCompressPictures forceUpgrade
captions readModeInkLockDown schemaLibrary shapeDefaults decimalSymbol listSeparator
documentId discardImageEditingData defaultImageDpi conflictMode chartTrackingRefBased
persistentDocumentId
`.trim().split(/\s+/);

const ORDER_INDEX = new Map(SETTING_ORDER.map((name, index) => [name, index]));
const BOOLEAN_SETTINGS = ["trackRevisions", "updateFields", "evenAndOddHeaders", "mirrorMargins", "gutterAtTop"];
const PROTECTION_MODES = new Set(["none", "readOnly", "comments", "trackedChanges", "forms"]);
const CONFIG_KEYS = new Set([...BOOLEAN_SETTINGS, "updateFieldsOnOpen", "documentProtection"]);
const WORDPROCESSING_NAMESPACES = new Set([
  "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  "http://purl.oclc.org/ooxml/wordprocessingml/main",
]);

function localName(tag = "") {
  return /^<\/?([^\s/>]+)/.exec(tag)?.[1]?.split(":").at(-1);
}

function settingsChildren(xml) {
  const source = String(xml || "");
  const root = rootTag(source, "settings");
  if (!root) throw new Error("DOCX settings sourceReference target part must have a w:settings root element.");
  const rootAttributes = attributes(root[0]);
  const namespace = root[1] ? rootAttributes[`xmlns:${root[1]}`] : rootAttributes.xmlns;
  if (!WORDPROCESSING_NAMESPACES.has(namespace)) throw new Error("DOCX settings sourceReference target root must use a WordprocessingML namespace.");
  if (/\/\s*>$/.test(root[0])) return { source, root, closingIndex: root.index + root[0].length, children: [] };
  const rootName = /^<([^\s/>]+)/.exec(root[0])?.[1];
  const closing = `</${rootName}>`;
  const closingIndex = source.lastIndexOf(closing);
  if (closingIndex < root.index + root[0].length) throw new Error("DOCX settings sourceReference target part has no closing w:settings element.");
  const innerStart = root.index + root[0].length;
  const inner = source.slice(innerStart, closingIndex);
  const tokens = [...inner.matchAll(/<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<\/?[A-Za-z_][^>]*>/g)];
  const children = [];
  let depth = 0;
  let current;
  for (const token of tokens) {
    const text = token[0];
    if (text.startsWith("<!--") || text.startsWith("<?") || text.startsWith("<![CDATA[")) continue;
    const closingTag = text.startsWith("</");
    const selfClosing = /\/\s*>$/.test(text);
    if (closingTag) {
      if (depth > 0) depth -= 1;
      if (depth === 0 && current) {
        current.end = innerStart + token.index + text.length;
        children.push(current);
        current = undefined;
      }
      continue;
    }
    if (depth === 0) {
      current = { name: localName(text), start: innerStart + token.index, end: innerStart + token.index + text.length, tag: text };
      if (selfClosing) {
        children.push(current);
        current = undefined;
      }
    }
    if (!selfClosing) depth += 1;
  }
  if (depth || current) throw new Error("DOCX settings sourceReference target part contains malformed child XML.");
  return { source, root, closingIndex, children };
}

function removeDirectSetting(xml, name) {
  const state = settingsChildren(xml);
  const matches = state.children.filter((child) => child.name === name).sort((a, b) => b.start - a.start);
  return matches.reduce((next, child) => `${next.slice(0, child.start)}${next.slice(child.end)}`, state.source);
}

function insertOrderedSetting(xml, name, tag) {
  let source = String(xml);
  let state = settingsChildren(source);
  if (/\/\s*>$/.test(state.root[0])) {
    const prefix = state.root[1] || "";
    const opening = state.root[0].replace(/\/\s*>$/, ">");
    source = `${source.slice(0, state.root.index)}${opening}</${qname(prefix, "settings")}>${source.slice(state.root.index + state.root[0].length)}`;
    state = settingsChildren(source);
  }
  const order = ORDER_INDEX.get(name);
  const following = state.children.find((child) => (ORDER_INDEX.get(child.name) ?? Number.POSITIVE_INFINITY) > order);
  const position = following?.start ?? state.closingIndex;
  return `${source.slice(0, position)}${tag}${source.slice(position)}`;
}

function booleanValue(value, label) {
  if (typeof value !== "boolean") throw new TypeError(`DOCX settings sourceReference ${label} must be a boolean.`);
  return value;
}

function normalizeProtection(value, { partial = false } = {}) {
  if (value === undefined) return partial ? undefined : null;
  if (value === false || value === null || value === "off") return null;
  const raw = typeof value === "string" ? { edit: value } : value;
  if (!raw || typeof raw !== "object" || Array.isArray(raw)) throw new TypeError("DOCX documentProtection must be false, a protection mode, or an object.");
  const unknown = Object.keys(raw).filter((key) => !new Set(["edit", "mode", "enforcement", "formatting"]).has(key));
  if (unknown.length) throw new Error(`DOCX documentProtection has unsupported fields: ${unknown.join(", ")}. Password hashing is intentionally unsupported.`);
  const edit = String(raw.edit ?? raw.mode ?? "readOnly");
  if (!PROTECTION_MODES.has(edit)) throw new Error("DOCX documentProtection edit must be none, readOnly, comments, trackedChanges, or forms.");
  return {
    edit,
    enforcement: raw.enforcement === undefined ? true : booleanValue(raw.enforcement, "documentProtection.enforcement"),
    formatting: raw.formatting === undefined ? false : booleanValue(raw.formatting, "documentProtection.formatting"),
  };
}

function normalizeSettingsConfig(config = {}, { partial = false } = {}) {
  if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("DOCX settings sourceReference must be an object.");
  const unknown = Object.keys(config).filter((key) => !CONFIG_KEYS.has(key));
  if (unknown.length) throw new Error(`DOCX settings sourceReference has unsupported fields: ${unknown.join(", ")}.`);
  const result = {};
  for (const name of BOOLEAN_SETTINGS) {
    const raw = name === "updateFields" ? (config.updateFields ?? config.updateFieldsOnOpen) : config[name];
    if (raw !== undefined) result[name] = booleanValue(raw, name);
    else if (!partial) result[name] = false;
  }
  const protection = normalizeProtection(config.documentProtection, { partial });
  if (protection !== undefined) result.documentProtection = protection;
  else if (!partial) result.documentProtection = null;
  return result;
}

function booleanSettingXml(prefix, name) {
  return `<${qname(prefix, name)}/>`;
}

function protectionXml(prefix, protection) {
  const attribute = (name, value) => ` ${qname(prefix, name)}="${attrEscape(value)}"`;
  return `<${qname(prefix, "documentProtection")}${attribute("edit", protection.edit)}${attribute("enforcement", protection.enforcement ? "1" : "0")}${attribute("formatting", protection.formatting ? "1" : "0")}/>`;
}

export function normalizeDocxSettings(settings = {}) {
  return normalizeSettingsConfig(settings);
}

export function mutateDocxSettings(xml, config = {}) {
  const normalized = normalizeSettingsConfig(config, { partial: true });
  if (!Object.keys(normalized).length) throw new Error("DOCX settings sourceReference requires at least one supported setting.");
  let next = String(xml);
  const prefix = settingsChildren(next).root[1] || "";
  for (const name of BOOLEAN_SETTINGS) {
    if (!(name in normalized)) continue;
    next = removeDirectSetting(next, name);
    if (normalized[name]) next = insertOrderedSetting(next, name, booleanSettingXml(prefix, name));
  }
  if ("documentProtection" in normalized) {
    next = removeDirectSetting(next, "documentProtection");
    if (normalized.documentProtection) next = insertOrderedSetting(next, "documentProtection", protectionXml(prefix, normalized.documentProtection));
  }
  return next;
}
