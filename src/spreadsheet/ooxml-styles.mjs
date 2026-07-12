// Clean-room SpreadsheetML style codec and deterministic display formatter.

import { resolveColorToken } from "../shared/colors.mjs";

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function parseAttrs(attrs = "") {
  return Object.fromEntries([...String(attrs).matchAll(/\b([A-Za-z_:][\w:.-]*)="([^"]*)"/g)].map((match) => [match[1], decodeXml(match[2])]));
}

function attrEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");
}

export function normalizeXlsxColor(value, fallback = "000000") {
  const raw = String(value || fallback).trim();
  const hex = String(resolveColorToken(raw, raw) || fallback).replace(/^#/, "");
  const rgb = /^[0-9a-fA-F]{8}$/.test(hex) ? hex.slice(2) : hex;
  return rgb.slice(0, 6).padEnd(6, "0").toUpperCase();
}

function normalizeBorder(style = {}) {
  const raw = style.border || style.borders;
  if (!raw) return undefined;
  const base = raw.outside || raw.all || raw;
  return { style: base.style || base.lineStyle || base.weight || "thin", color: base.color || base.fill || base.borderColor || "#CBD5E1" };
}

function normalizeAlignment(style = {}) {
  const raw = style.alignment || style.align || {};
  const result = {};
  const horizontal = raw.horizontal || style.horizontalAlignment || style.textAlign;
  const vertical = raw.vertical || style.verticalAlignment;
  if (horizontal) result.horizontal = horizontal;
  if (vertical) result.vertical = vertical;
  const fields = {
    wrapText: raw.wrapText ?? style.wrapText,
    textRotation: raw.textRotation ?? style.textRotation,
    indent: raw.indent ?? style.indent,
    shrinkToFit: raw.shrinkToFit ?? style.shrinkToFit,
    readingOrder: raw.readingOrder ?? style.readingOrder,
  };
  if (fields.wrapText != null) result.wrapText = Boolean(fields.wrapText);
  if (fields.shrinkToFit != null) result.shrinkToFit = Boolean(fields.shrinkToFit);
  for (const name of ["textRotation", "indent", "readingOrder"]) if (fields[name] != null && Number.isFinite(Number(fields[name]))) result[name] = Number(fields[name]);
  return Object.keys(result).length ? result : undefined;
}

export function normalizeXlsxStyle(style = {}) {
  const font = style.font || {};
  return {
    font: {
      bold: Boolean(style.bold ?? font.bold),
      italic: Boolean(style.italic ?? font.italic),
      underline: style.underline ?? font.underline ?? undefined,
      strike: Boolean(style.strike ?? font.strike),
      color: style.fontColor || font.color || style.color || undefined,
      size: Number(style.fontSize || font.size || 11),
      name: style.fontFamily || font.name || "Aptos",
    },
    fill: style.fill || style.backgroundColor || style.fillColor || undefined,
    numberFormat: style.numberFormat || style.numFmt || undefined,
    alignment: normalizeAlignment(style),
    border: normalizeBorder(style),
    protection: style.protection ? { locked: style.protection.locked == null ? undefined : Boolean(style.protection.locked), hidden: style.protection.hidden == null ? undefined : Boolean(style.protection.hidden) } : undefined,
  };
}

export function xlsxStyleKey(style = {}) {
  const normalized = normalizeXlsxStyle(style);
  if (!normalized.font.bold && !normalized.font.italic && !normalized.font.underline && !normalized.font.strike && !normalized.font.color && normalized.font.size === 11 && normalized.font.name === "Aptos" && !normalized.fill && !normalized.numberFormat && !normalized.alignment && !normalized.border && !normalized.protection) return "";
  return JSON.stringify(normalized);
}

function fontXml(style = {}) {
  const font = normalizeXlsxStyle(style).font;
  const underline = font.underline ? `<u${typeof font.underline === "string" && font.underline !== "single" ? ` val="${attrEscape(font.underline)}"` : ""}/>` : "";
  return `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}${underline}${font.strike ? "<strike/>" : ""}<sz val="${Number(font.size) || 11}"/><color rgb="FF${normalizeXlsxColor(font.color, "000000")}"/><name val="${attrEscape(font.name || "Aptos")}"/></font>`;
}

function fillXml(style = {}) {
  const fill = normalizeXlsxStyle(style).fill;
  if (!fill) return `<fill><patternFill patternType="none"/></fill>`;
  return `<fill><patternFill patternType="solid"><fgColor rgb="FF${normalizeXlsxColor(fill, "FFFFFF")}"/><bgColor indexed="64"/></patternFill></fill>`;
}

function borderXml(style = {}) {
  const border = normalizeXlsxStyle(style).border;
  if (!border) return `<border/>`;
  const color = normalizeXlsxColor(border.color, "CBD5E1"), lineStyle = attrEscape(border.style || "thin");
  const edge = (name) => `<${name} style="${lineStyle}"><color rgb="FF${color}"/></${name}>`;
  return `<border>${edge("left")}${edge("right")}${edge("top")}${edge("bottom")}<diagonal/></border>`;
}

function dxfXml(style = {}) {
  const normalized = normalizeXlsxStyle(style), font = normalized.font || {};
  const underline = font.underline ? `<u${typeof font.underline === "string" && font.underline !== "single" ? ` val="${attrEscape(font.underline)}"` : ""}/>` : "";
  const fontOutput = (font.bold || font.italic || font.underline || font.strike || font.color || font.size || font.name)
    ? `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}${underline}${font.strike ? "<strike/>" : ""}${font.size ? `<sz val="${Number(font.size) || 11}"/>` : ""}${font.color ? `<color rgb="FF${normalizeXlsxColor(font.color, "000000")}"/>` : ""}${font.name ? `<name val="${attrEscape(font.name)}"/>` : ""}</font>`
    : "";
  return `<dxf>${fontOutput}${normalized.fill ? fillXml(normalized) : ""}${normalized.numberFormat ? `<numFmt numFmtId="0" formatCode="${attrEscape(normalized.numberFormat)}"/>` : ""}</dxf>`;
}

export function xlsxStylesXml(styleTable = {}) {
  const styles = styleTable.styles || [{}], dxfs = styleTable.dxfs || [];
  const customFormats = new Map();
  styles.forEach((style) => { if (style.numberFormat && !customFormats.has(style.numberFormat)) customFormats.set(style.numberFormat, 164 + customFormats.size); });
  const numFmts = customFormats.size ? `<numFmts count="${customFormats.size}">${[...customFormats.entries()].map(([code, id]) => `<numFmt numFmtId="${id}" formatCode="${attrEscape(code)}"/>`).join("")}</numFmts>` : "";
  const fonts = styles.map((style, index) => index === 0 ? `<font><sz val="11"/><name val="Aptos"/></font>` : fontXml(style)).join("");
  const fills = [`<fill><patternFill patternType="none"/></fill>`, `<fill><patternFill patternType="gray125"/></fill>`, ...styles.slice(1).map(fillXml)].join("");
  const borders = [`<border/>`, ...styles.slice(1).map(borderXml)].join("");
  const xfs = styles.map((style, index) => {
    if (index === 0) return `<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`;
    const normalized = normalizeXlsxStyle(style), numFmtId = normalized.numberFormat ? customFormats.get(normalized.numberFormat) : 0;
    const fillId = normalized.fill ? index + 1 : 0, borderId = normalized.border ? index : 0;
    const alignment = normalized.alignment;
    const alignmentXml = alignment ? `<alignment${alignment.horizontal ? ` horizontal="${attrEscape(alignment.horizontal)}"` : ""}${alignment.vertical ? ` vertical="${attrEscape(alignment.vertical)}"` : ""}${alignment.wrapText != null ? ` wrapText="${alignment.wrapText ? 1 : 0}"` : ""}${alignment.textRotation != null ? ` textRotation="${alignment.textRotation}"` : ""}${alignment.indent != null ? ` indent="${alignment.indent}"` : ""}${alignment.shrinkToFit != null ? ` shrinkToFit="${alignment.shrinkToFit ? 1 : 0}"` : ""}${alignment.readingOrder != null ? ` readingOrder="${alignment.readingOrder}"` : ""}/>` : "";
    const protectionXml = normalized.protection ? `<protection${normalized.protection.locked != null ? ` locked="${normalized.protection.locked ? 1 : 0}"` : ""}${normalized.protection.hidden != null ? ` hidden="${normalized.protection.hidden ? 1 : 0}"` : ""}/>` : "";
    const attrs = `numFmtId="${numFmtId}" fontId="${index}" fillId="${fillId}" borderId="${borderId}" xfId="0"${numFmtId ? ` applyNumberFormat="1"` : ""} applyFont="1"${fillId ? ` applyFill="1"` : ""}${borderId ? ` applyBorder="1"` : ""}${alignment ? ` applyAlignment="1"` : ""}${normalized.protection ? ` applyProtection="1"` : ""}`;
    return alignmentXml || protectionXml ? `<xf ${attrs}>${alignmentXml}${protectionXml}</xf>` : `<xf ${attrs}/>`;
  }).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts}<fonts count="${styles.length}">${fonts}</fonts><fills count="${styles.length + 1}">${fills}</fills><borders count="${styles.length}">${borders}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfs}</cellXfs><dxfs count="${dxfs.length}">${dxfs.map(dxfXml).join("")}</dxfs></styleSheet>`;
}

function booleanAttribute(attrs, name) {
  if (attrs[name] == null) return undefined;
  return !["0", "false", "off"].includes(String(attrs[name]).trim().toLowerCase());
}

function booleanElement(body, name) {
  const match = new RegExp(`<${name}\\b([^>]*)/?>`).exec(body);
  if (!match) return false;
  const value = parseAttrs(match[1]).val;
  return value == null || !["0", "false", "off"].includes(String(value).toLowerCase());
}

function rgbColor(body, elementName) {
  const value = new RegExp(`<${elementName}[^>]*rgb="([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8})"`).exec(body)?.[1];
  return value ? `#${value.slice(-6)}` : undefined;
}

function parseFont(body = "") {
  const underlineMatch = /<u\b([^>]*)\/?>/.exec(body);
  return {
    bold: booleanElement(body, "b"),
    italic: booleanElement(body, "i"),
    underline: underlineMatch ? parseAttrs(underlineMatch[1]).val || "single" : undefined,
    strike: booleanElement(body, "strike"),
    color: rgbColor(body, "color"),
    size: Number(/<sz[^>]*val="([^"]+)"/.exec(body)?.[1] || 11),
    name: /<name[^>]*val="([^"]+)"/.exec(body)?.[1] || "Aptos",
  };
}

function parseFill(body = "") {
  return rgbColor(body, "fgColor");
}

function parseBorder(body = "") {
  const edge = /<(left|right|top|bottom)\b([^>]*)>([\s\S]*?)<\/\1>/.exec(body);
  if (!edge) return undefined;
  const attrs = parseAttrs(edge[2]);
  const color = rgbColor(edge[3], "color");
  return { style: attrs.style || "thin", ...(color ? { color } : {}) };
}

function parseAlignment(body = "") {
  const attrs = parseAttrs(/<alignment\b([^>]*)\/?\s*>/.exec(body)?.[1] || "");
  if (!Object.keys(attrs).length) return undefined;
  const style = {};
  if (attrs.horizontal != null) style.horizontal = attrs.horizontal;
  if (attrs.vertical != null) style.vertical = attrs.vertical;
  for (const name of ["wrapText", "shrinkToFit"]) if (attrs[name] != null) style[name] = booleanAttribute(attrs, name);
  for (const name of ["textRotation", "indent", "readingOrder"]) if (attrs[name] != null && Number.isFinite(Number(attrs[name]))) style[name] = Number(attrs[name]);
  return Object.keys(style).length ? style : undefined;
}

function parseProtection(body = "") {
  const attrs = parseAttrs(/<protection\b([^>]*)\/?\s*>/.exec(body)?.[1] || "");
  if (!Object.keys(attrs).length) return undefined;
  return {
    ...(attrs.locked != null ? { locked: booleanAttribute(attrs, "locked") } : {}),
    ...(attrs.hidden != null ? { hidden: booleanAttribute(attrs, "hidden") } : {}),
  };
}

function parseDxf(body = "") {
  const style = {};
  const fontBody = /<font>([\s\S]*?)<\/font>/.exec(body)?.[1];
  const fillBody = /<fill>([\s\S]*?)<\/fill>/.exec(body)?.[1];
  const numFmt = /<numFmt\b([^>]*)\/?>(?:<\/numFmt>)?/.exec(body)?.[1];
  if (fontBody != null) style.font = parseFont(fontBody);
  const fill = fillBody != null ? parseFill(fillBody) : undefined;
  if (fill) style.fill = fill;
  if (numFmt) style.numberFormat = parseAttrs(numFmt).formatCode;
  return style;
}

export const XLSX_BUILTIN_NUMBER_FORMATS = new Map([
  [0, "General"], [1, "0"], [2, "0.00"], [3, "#,##0"], [4, "#,##0.00"],
  [5, "$#,##0_);($#,##0)"], [6, "$#,##0_);[Red]($#,##0)"], [7, "$#,##0.00_);($#,##0.00)"], [8, "$#,##0.00_);[Red]($#,##0.00)"],
  [9, "0%"], [10, "0.00%"], [11, "0.00E+00"], [12, "# ?/?"], [13, "# ??/??"],
  [14, "mm-dd-yy"], [15, "d-mmm-yy"], [16, "d-mmm"], [17, "mmm-yy"],
  [18, "h:mm AM/PM"], [19, "h:mm:ss AM/PM"], [20, "h:mm"], [21, "h:mm:ss"], [22, "m/d/yy h:mm"],
  [37, "#,##0;(#,##0)"], [38, "#,##0;[Red](#,##0)"], [39, "#,##0.00;(#,##0.00)"], [40, "#,##0.00;[Red](#,##0.00)"],
  [41, "_(* #,##0_);_(* \\(#,##0\\);_(* \"-\"_);_(@_)"], [42, "_(\"$\"* #,##0_);_(\"$\"* \\(#,##0\\);_(\"$\"* \"-\"_);_(@_)"],
  [43, "_(* #,##0.00_);_(* \\(#,##0.00\\);_(* \"-\"??_);_(@_)"], [44, "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)"],
  [45, "mm:ss"], [46, "[h]:mm:ss"], [47, "mmss.0"], [48, "##0.0E+0"], [49, "@"],
]);

function xfRecords(body = "") {
  return [...String(body).matchAll(/<xf\b([^>]*)\/>|<xf\b([^>]*)>([\s\S]*?)<\/xf>/g)].map((match) => ({ attrs: parseAttrs(match[1] || match[2]), body: match[3] || "" }));
}

function xfComponents(record, resources) {
  const numberFormatId = Number(record.attrs.numFmtId || 0);
  return {
    font: resources.fonts[Number(record.attrs.fontId || 0)] || {},
    fill: resources.fills[Number(record.attrs.fillId || 0)],
    border: resources.borders[Number(record.attrs.borderId || 0)],
    numberFormat: numberFormatId === 0 ? undefined : resources.numberFormats.get(numberFormatId),
    alignment: parseAlignment(record.body),
    protection: parseProtection(record.body),
    ids: { font: Number(record.attrs.fontId || 0), fill: Number(record.attrs.fillId || 0), border: Number(record.attrs.borderId || 0), numberFormat: numberFormatId },
  };
}

function effectiveStyle(record, baseRecord, resources) {
  const direct = xfComponents(record, resources);
  const base = baseRecord ? xfComponents(baseRecord, resources) : undefined;
  const style = {};
  const select = (name, applyName) => {
    const apply = booleanAttribute(record.attrs, applyName);
    if (base && apply === false) return base[name];
    if (direct[name] != null && (apply === true || !base || direct.ids[name] !== base.ids[name] || name === "alignment" || name === "protection")) return direct[name];
    return base?.[name] ?? direct[name];
  };
  const font = select("font", "applyFont");
  if (font && Object.keys(font).length) style.font = { ...font };
  const fill = select("fill", "applyFill");
  if (fill) style.fill = fill;
  const border = select("border", "applyBorder");
  if (border) style.border = border;
  const numberFormat = select("numberFormat", "applyNumberFormat");
  if (numberFormat) style.numberFormat = numberFormat;
  const alignment = select("alignment", "applyAlignment");
  if (alignment) style.alignment = { ...alignment };
  const protection = select("protection", "applyProtection");
  if (protection) style.protection = { ...protection };
  return style;
}

export function parseXlsxStylesXml(xml = "") {
  const text = String(xml);
  const numberFormats = new Map(XLSX_BUILTIN_NUMBER_FORMATS);
  for (const match of text.matchAll(/<numFmt\b([^>]*)\/>/g)) { const attrs = parseAttrs(match[1]); numberFormats.set(Number(attrs.numFmtId), attrs.formatCode); }
  const fontsBody = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(text)?.[1] || "";
  const fonts = [...fontsBody.matchAll(/<font>([\s\S]*?)<\/font>/g)].map((match) => parseFont(match[1]));
  const fillsBody = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(text)?.[1] || "";
  const fills = [...fillsBody.matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((match) => parseFill(match[1]));
  const bordersBody = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(text)?.[1] || "";
  const borders = [...bordersBody.matchAll(/<border\b[^>]*>([\s\S]*?)<\/border>/g)].map((match) => parseBorder(match[1]));
  const resources = { fonts, fills, borders, numberFormats };
  const styleXfs = xfRecords(/<cellStyleXfs\b[^>]*>([\s\S]*?)<\/cellStyleXfs>/.exec(text)?.[1] || "");
  const styles = xfRecords(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(text)?.[1] || "").map((record) => effectiveStyle(record, styleXfs[Number(record.attrs.xfId || 0)], resources));
  const dxfsBody = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(text)?.[1] || "";
  styles.dxfs = [...dxfsBody.matchAll(/<dxf>([\s\S]*?)<\/dxf>/g)].map((match) => parseDxf(match[1]));
  return styles;
}

function formatSections(format = "") {
  const sections = [];
  let current = "", quoted = false, bracketDepth = 0;
  for (const character of String(format)) {
    if (character === '"') quoted = !quoted;
    if (!quoted && character === "[") bracketDepth += 1;
    if (!quoted && character === "]") bracketDepth = Math.max(0, bracketDepth - 1);
    if (!quoted && bracketDepth === 0 && character === ";") { sections.push(current); current = ""; continue; }
    current += character;
  }
  sections.push(current);
  return sections;
}

function literalCleanup(format = "") {
  return String(format).replace(/\[(?![hms]\])[^\]]*\]/gi, "").replace(/_./g, " ").replace(/\*./g, "").replace(/\\(.)/g, "$1").replace(/"([^"]*)"/g, "$1");
}

function dateParts(serialValue, dateSystem = "1900") {
  const serial = Math.floor(serialValue);
  if (!Number.isFinite(serial) || serial < 0) return undefined;
  if (dateSystem === "1904") {
    const date = new Date(Date.UTC(1904, 0, 1) + serial * 86_400_000);
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }
  if (serial === 0) return { year: 1900, month: 1, day: 0 };
  if (serial === 60) return { year: 1900, month: 2, day: 29 };
  const date = new Date(Date.UTC(1899, 11, 31) + (serial > 60 ? serial - 1 : serial) * 86_400_000);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

function dateDisplay(value, format, dateSystem) {
  const parts = dateParts(value, dateSystem);
  if (!parts) return String(value);
  const totalSeconds = Math.round((((Number(value) % 1) + 1) % 1) * 86_400) % 86_400;
  const hour24 = Math.floor(totalSeconds / 3600), minute = Math.floor(totalSeconds % 3600 / 60), second = totalSeconds % 60;
  const short = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  const long = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
  const meridiem = /AM\/PM/i.test(format);
  const cleaned = literalCleanup(format).replace(/\[h\]/gi, "__ELAPSED_HOURS__");
  return cleaned.replace(/__ELAPSED_HOURS__|AM\/PM|yyyy|mmmm|mmm|yy|mm|dd|hh|ss|m|d|h|s/gi, (token, offset, source) => {
    const lower = token.toLowerCase(), before = source[offset - 1], after = source[offset + token.length];
    const minutes = lower.startsWith("m") && (before === ":" || after === ":") && /h/i.test(source);
    if (token === "__ELAPSED_HOURS__") return String(Math.floor(Math.abs(Number(value)) * 24));
    if (lower === "am/pm") return hour24 < 12 ? "AM" : "PM";
    if (lower === "yyyy") return String(parts.year).padStart(4, "0");
    if (lower === "yy") return String(parts.year % 100).padStart(2, "0");
    if (lower === "mmmm") return long[parts.month - 1];
    if (lower === "mmm") return short[parts.month - 1];
    if (lower === "mm") return String(minutes ? minute : parts.month).padStart(2, "0");
    if (lower === "m") return String(minutes ? minute : parts.month);
    if (lower === "dd") return String(parts.day).padStart(2, "0");
    if (lower === "d") return String(parts.day);
    if (lower === "hh") return String(meridiem ? (hour24 % 12 || 12) : hour24).padStart(2, "0");
    if (lower === "h") return String(meridiem ? (hour24 % 12 || 12) : hour24);
    if (lower === "ss") return String(second).padStart(2, "0");
    if (lower === "s") return String(second);
    return token;
  });
}

export function formatSpreadsheetDisplayValue(value, style = {}, options = {}) {
  if (value == null) return "";
  const format = String(style?.numberFormat || "General");
  if (!format || /^general$/i.test(format)) return typeof value === "boolean" ? (value ? "TRUE" : "FALSE") : String(value);
  const sections = formatSections(format);
  const section = typeof value === "string" ? sections[3] || sections[0] || "@" : Number(value) < 0 && sections[1] != null ? sections[1] : Number(value) === 0 && sections[2] != null ? sections[2] : sections[0] || "General";
  if (typeof value === "string") return section.includes("@") ? literalCleanup(section).replace("@", value) : value;
  const number = Number(value);
  if (!Number.isFinite(number)) return String(value);
  if (/[yd]/i.test(section) || /(?:h+|s+|\[h\])(?=:|\b)/i.test(section)) return dateDisplay(number, section, options.dateSystem || "1900");
  const cleaned = literalCleanup(section);
  if (/@/.test(cleaned)) return cleaned.replace("@", String(value));
  const percent = cleaned.includes("%"), scaled = (percent ? 100 : 1) * Math.abs(number);
  const exponent = /E[+-]0+/i.test(cleaned), decimals = /[0#?]+\.([0#?]+)/.exec(cleaned)?.[1]?.length || 0;
  let rendered = exponent
    ? scaled.toExponential(decimals).replace(/e([+-]?)(\d+)/i, (_, sign, digits) => `E${sign || "+"}${String(digits).padStart(2, "0")}`)
    : scaled.toLocaleString("en-US", { useGrouping: cleaned.includes(","), minimumFractionDigits: decimals, maximumFractionDigits: decimals });
  const prefix = cleaned.slice(0, Math.max(0, cleaned.search(/[0#?]/))).replace(/[()\s]/g, "");
  rendered = `${prefix}${rendered}${percent ? "%" : ""}`;
  if (number < 0) rendered = cleaned.includes("(") ? `(${rendered})` : `-${rendered}`;
  return rendered;
}
