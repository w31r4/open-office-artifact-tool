// Clean-room SpreadsheetML style codec and deterministic display formatter.

import { resolveColorToken } from "../shared/colors.mjs";

export const XLSX_THEME_COLOR_NAMES = ["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"];

const XLSX_DEFAULT_THEME_COLORS = {
  dk1: "#000000", lt1: "#FFFFFF", dk2: "#1F497D", lt2: "#EEECE1",
  accent1: "#4F81BD", accent2: "#C0504D", accent3: "#9BBB59", accent4: "#8064A2",
  accent5: "#4BACC6", accent6: "#F79646", hlink: "#0000FF", folHlink: "#800080",
};

const XLSX_PATTERN_TYPES = new Set([
  "none", "solid", "mediumGray", "darkGray", "lightGray", "darkHorizontal", "darkVertical", "darkDown", "darkUp", "darkGrid", "darkTrellis",
  "lightHorizontal", "lightVertical", "lightDown", "lightUp", "lightGrid", "lightTrellis", "gray125", "gray0625",
]);

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
  if (!/^[0-9a-fA-F]{6}$/.test(rgb)) throw new TypeError("Spreadsheet color must be a supported color token or six/eight-digit RGB value.");
  return rgb.toUpperCase();
}

function normalizedThemeColors(theme = {}) {
  const input = Array.isArray(theme) ? Object.fromEntries(XLSX_THEME_COLOR_NAMES.map((name, index) => [name, theme[index]])) : theme.colors || theme;
  return Object.fromEntries(XLSX_THEME_COLOR_NAMES.map((name) => [name, `#${normalizeXlsxColor(input?.[name], XLSX_DEFAULT_THEME_COLORS[name])}`]));
}

export function normalizeXlsxThemeConfig(theme = {}) {
  return { name: String(theme.name || "Office Clean Room"), colors: normalizedThemeColors(theme) };
}

export function xlsxThemeXml(theme = {}) {
  const normalized = normalizeXlsxThemeConfig(theme);
  const colors = XLSX_THEME_COLOR_NAMES.map((name) => `<a:${name}><a:srgbClr val="${normalizeXlsxColor(normalized.colors[name])}"/></a:${name}>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${attrEscape(normalized.name)}"><a:themeElements><a:clrScheme name="${attrEscape(normalized.name)}">${colors}</a:clrScheme><a:fontScheme name="Office Clean Room"><a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface=""/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Aptos"/><a:ea typeface=""/><a:cs typeface=""/></a:minorFont></a:fontScheme><a:fmtScheme name="Office Clean Room"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements></a:theme>`;
}

export function normalizeXlsxColorReference(value, fallback) {
  if (value == null || value === "") return fallback == null ? undefined : normalizeXlsxColorReference(fallback);
  if (typeof value === "string") { normalizeXlsxColor(value); return String(value); }
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Spreadsheet color must be a color string or reference object.");
  const kinds = [value.rgb != null, value.theme != null, value.indexed != null, value.auto === true].filter(Boolean).length;
  if (kinds !== 1) throw new TypeError("Spreadsheet color reference requires exactly one of rgb, theme, indexed, or auto.");
  const tint = value.tint == null ? undefined : Number(value.tint);
  if (tint != null && (!Number.isFinite(tint) || tint < -1 || tint > 1)) throw new RangeError("Spreadsheet color tint must be between -1 and 1.");
  if (value.rgb != null) return `#${normalizeXlsxColor(value.rgb)}`;
  if (value.theme != null && (!Number.isInteger(Number(value.theme)) || Number(value.theme) < 0 || Number(value.theme) >= XLSX_THEME_COLOR_NAMES.length)) throw new RangeError("Spreadsheet theme color index must be an integer from 0 through 11.");
  if (value.indexed != null && (!Number.isInteger(Number(value.indexed)) || Number(value.indexed) < 0 || Number(value.indexed) > 65)) throw new RangeError("Spreadsheet indexed color must be an integer from 0 through 65.");
  return {
    ...(value.theme != null ? { theme: Number(value.theme) } : {}),
    ...(value.indexed != null ? { indexed: Number(value.indexed) } : {}),
    ...(value.auto === true ? { auto: true } : {}),
    ...(tint == null || tint === 0 ? {} : { tint }),
    ...(value.resolved ? { resolved: `#${normalizeXlsxColor(value.resolved)}` } : {}),
  };
}

export function normalizeXlsxFill(value) {
  if (value == null || value === "") return undefined;
  if (typeof value === "string") { normalizeXlsxColor(value, "FFFFFF"); return String(value); }
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Spreadsheet fill must be a color string or pattern object.");
  const patternType = String(value.patternType || value.pattern || "solid");
  if (!XLSX_PATTERN_TYPES.has(patternType)) throw new RangeError(`Unsupported SpreadsheetML patternType ${patternType}.`);
  const foreground = normalizeXlsxColorReference(value.foreground ?? value.fgColor ?? value.color ?? value.fill);
  const background = normalizeXlsxColorReference(value.background ?? value.bgColor);
  if (patternType === "solid" && !foreground) throw new TypeError("A solid spreadsheet fill requires foreground/color.");
  return { patternType, ...(foreground ? { foreground } : {}), ...(background ? { background } : {}) };
}

function normalizeBorder(style = {}) {
  const raw = style.border || style.borders;
  if (!raw) return undefined;
  const base = raw.outside || raw.all || raw;
  const edgeNames = ["left", "right", "top", "bottom", "diagonal", "start", "end", "horizontal", "vertical"];
  if (edgeNames.some((name) => base[name] != null)) {
    const result = {};
    for (const name of edgeNames) {
      const edge = base[name];
      if (edge == null || edge === false) continue;
      if (typeof edge === "string") result[name] = { style: edge, color: "#CBD5E1" };
      else if (edge.style || edge.lineStyle || edge.weight) result[name] = { style: edge.style || edge.lineStyle || edge.weight, color: normalizeXlsxColorReference(edge.color || edge.fill || edge.borderColor || "#CBD5E1") };
    }
    for (const name of ["diagonalUp", "diagonalDown", "outline"]) if (base[name] != null) result[name] = Boolean(base[name]);
    return Object.keys(result).length ? result : undefined;
  }
  return { style: base.style || base.lineStyle || base.weight || "thin", color: normalizeXlsxColorReference(base.color || base.fill || base.borderColor || "#CBD5E1") };
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
      color: normalizeXlsxColorReference(style.fontColor || font.color || style.color),
      size: Number(style.fontSize || font.size || 11),
      name: style.fontFamily || font.name || "Aptos",
    },
    fill: normalizeXlsxFill(style.fill || style.backgroundColor || style.fillColor),
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
  return `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}${underline}${font.strike ? "<strike/>" : ""}<sz val="${Number(font.size) || 11}"/>${xlsxColorElementXml("color", font.color || "#000000")}<name val="${attrEscape(font.name || "Aptos")}"/></font>`;
}

function fillXml(style = {}) {
  const fill = normalizeXlsxStyle(style).fill;
  if (!fill) return `<fill><patternFill patternType="none"/></fill>`;
  const normalized = typeof fill === "string" ? { patternType: "solid", foreground: fill } : fill;
  return `<fill><patternFill patternType="${attrEscape(normalized.patternType)}">${normalized.foreground ? xlsxColorElementXml("fgColor", normalized.foreground) : ""}${normalized.background ? xlsxColorElementXml("bgColor", normalized.background) : normalized.patternType === "solid" ? '<bgColor indexed="64"/>' : ""}</patternFill></fill>`;
}

function borderXml(style = {}) {
  const border = normalizeXlsxStyle(style).border;
  if (!border) return `<border/>`;
  const edgeXml = (name, edge) => edge?.style ? `<${name} style="${attrEscape(edge.style)}">${xlsxColorElementXml("color", edge.color || "#CBD5E1")}</${name}>` : `<${name}/>`;
  if (border.style) {
    const edge = { style: border.style, color: border.color };
    return `<border>${edgeXml("left", edge)}${edgeXml("right", edge)}${edgeXml("top", edge)}${edgeXml("bottom", edge)}<diagonal/></border>`;
  }
  const attrs = ["diagonalUp", "diagonalDown", "outline"].filter((name) => border[name] != null).map((name) => `${name}="${border[name] ? 1 : 0}"`).join(" ");
  const perimeter = ["left", "right", "top", "bottom"].map((name) => edgeXml(name, border[name])).join("");
  const diagonal = edgeXml("diagonal", border.diagonal);
  const extras = ["start", "end", "horizontal", "vertical"].filter((name) => border[name]).map((name) => edgeXml(name, border[name])).join("");
  return `<border${attrs ? ` ${attrs}` : ""}>${perimeter}${diagonal}${extras}</border>`;
}

function dxfXml(style = {}) {
  const normalized = normalizeXlsxStyle(style), font = normalized.font || {};
  const underline = font.underline ? `<u${typeof font.underline === "string" && font.underline !== "single" ? ` val="${attrEscape(font.underline)}"` : ""}/>` : "";
  const fontOutput = (font.bold || font.italic || font.underline || font.strike || font.color || font.size || font.name)
    ? `<font>${font.bold ? "<b/>" : ""}${font.italic ? "<i/>" : ""}${underline}${font.strike ? "<strike/>" : ""}${font.size ? `<sz val="${Number(font.size) || 11}"/>` : ""}${font.color ? xlsxColorElementXml("color", font.color) : ""}${font.name ? `<name val="${attrEscape(font.name)}"/>` : ""}</font>`
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
  const indexedColors = Array.isArray(styleTable.indexedColors) && styleTable.indexedColors.length
    ? `<colors><indexedColors>${styleTable.indexedColors.map((color) => `<rgbColor rgb="FF${normalizeXlsxColor(color)}"/>`).join("")}</indexedColors></colors>`
    : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">${numFmts}<fonts count="${styles.length}">${fonts}</fonts><fills count="${styles.length + 1}">${fills}</fills><borders count="${styles.length}">${borders}</borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${styles.length}">${xfs}</cellXfs><dxfs count="${dxfs.length}">${dxfs.map(dxfXml).join("")}</dxfs>${indexedColors}</styleSheet>`;
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

const XLSX_INDEXED_COLORS = [
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "000000", "FFFFFF", "FF0000", "00FF00", "0000FF", "FFFF00", "FF00FF", "00FFFF",
  "800000", "008000", "000080", "808000", "800080", "008080", "C0C0C0", "808080",
  "9999FF", "993366", "FFFFCC", "CCFFFF", "660066", "FF8080", "0066CC", "CCCCFF",
  "000080", "FF00FF", "FFFF00", "00FFFF", "800080", "800000", "008080", "0000FF",
  "00CCFF", "CCFFFF", "CCFFCC", "FFFF99", "99CCFF", "FF99CC", "CC99FF", "FFCC99",
  "3366FF", "33CCCC", "99CC00", "FFCC00", "FF9900", "FF6600", "666699", "969696",
  "003366", "339966", "003300", "333300", "993300", "993366", "333399", "333333",
].map((value) => `#${value}`);

function rgbToHsl(color) {
  const value = normalizeXlsxColor(color, "000000");
  const [red, green, blue] = [0, 2, 4].map((offset) => Number.parseInt(value.slice(offset, offset + 2), 16) / 255);
  const maximum = Math.max(red, green, blue), minimum = Math.min(red, green, blue);
  const lightness = (maximum + minimum) / 2;
  if (maximum === minimum) return [0, 0, lightness];
  const delta = maximum - minimum;
  const saturation = lightness > 0.5 ? delta / (2 - maximum - minimum) : delta / (maximum + minimum);
  const hue = maximum === red ? ((green - blue) / delta + (green < blue ? 6 : 0)) / 6 : maximum === green ? ((blue - red) / delta + 2) / 6 : ((red - green) / delta + 4) / 6;
  return [hue, saturation, lightness];
}

function hslToRgb(hue, saturation, lightness) {
  if (saturation === 0) {
    const value = Math.round(lightness * 255).toString(16).padStart(2, "0");
    return `#${value}${value}${value}`.toUpperCase();
  }
  const q = lightness < 0.5 ? lightness * (1 + saturation) : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  const channel = (offset) => {
    let value = hue + offset;
    if (value < 0) value += 1;
    if (value > 1) value -= 1;
    const result = value < 1 / 6 ? p + (q - p) * 6 * value : value < 1 / 2 ? q : value < 2 / 3 ? p + (q - p) * (2 / 3 - value) * 6 : p;
    return Math.round(result * 255 + 1e-9).toString(16).padStart(2, "0");
  };
  return `#${channel(1 / 3)}${channel(0)}${channel(-1 / 3)}`.toUpperCase();
}

function tintedColor(color, tintValue) {
  const tint = Number(tintValue);
  if (!Number.isFinite(tint) || tint === 0) return color;
  const [hue, saturation, lightness] = rgbToHsl(color);
  const nextLightness = Math.max(0, Math.min(1, tint < 0 ? lightness * (1 + tint) : lightness * (1 - tint) + tint));
  return hslToRgb(hue, saturation, nextLightness);
}

function xlsxColorElementXml(name, value) {
  const normalized = normalizeXlsxColorReference(value, "#000000");
  if (typeof normalized === "string") return `<${name} rgb="FF${normalizeXlsxColor(normalized)}"/>`;
  const tint = normalized.tint == null ? "" : ` tint="${normalized.tint}"`;
  if (normalized.theme != null) return `<${name} theme="${normalized.theme}"${tint}/>`;
  if (normalized.indexed != null) return `<${name} indexed="${normalized.indexed}"${tint}/>`;
  return `<${name} auto="1"${tint}/>`;
}

export function xlsxColorCss(value, resources = {}) {
  const normalized = normalizeXlsxColorReference(value, resources.fallback || "#000000");
  if (typeof normalized === "string") return normalized;
  if (normalized.resolved) return normalized.resolved;
  const themeColors = Array.isArray(resources.themeColors)
    ? resources.themeColors
    : XLSX_THEME_COLOR_NAMES.map((name) => normalizeXlsxThemeConfig(resources.theme || {}).colors[name]);
  let color;
  if (normalized.theme != null) color = themeColors[normalized.theme];
  else if (normalized.indexed != null) color = normalized.indexed === 64 ? resources.autoColor || "#000000" : normalized.indexed === 65 ? resources.background || "#FFFFFF" : (resources.indexedColors || XLSX_INDEXED_COLORS)[normalized.indexed];
  else if (normalized.auto) color = resources.autoColor || "#000000";
  return tintedColor(color || resources.fallback || "#000000", normalized.tint);
}

export function xlsxFillPaint(value, resources = {}) {
  const fill = normalizeXlsxFill(value) || { patternType: "none" };
  if (typeof fill === "string") return { patternType: "solid", foreground: xlsxColorCss(fill, resources), background: xlsxColorCss(fill, resources) };
  const foreground = fill.foreground ? xlsxColorCss(fill.foreground, resources) : resources.foreground || "#000000";
  const background = fill.background ? xlsxColorCss(fill.background, resources) : fill.patternType === "solid" ? foreground : resources.background || "#FFFFFF";
  return { patternType: fill.patternType, foreground, background };
}

export function xlsxFillSvgPaint(value, id, resources = {}) {
  const fill = xlsxFillPaint(value, resources);
  if (fill.patternType === "none" || fill.patternType === "solid") return { paint: fill.background };
  const safeId = String(id || "xlsx-fill").replace(/[^A-Za-z0-9_-]/g, "-");
  const light = fill.patternType.startsWith("light"), strokeWidth = light ? 1 : 2;
  const horizontal = /Horizontal|Grid|Trellis/.test(fill.patternType);
  const vertical = /Vertical|Grid|Trellis/.test(fill.patternType);
  const down = /Down|Trellis/.test(fill.patternType);
  const up = /Up|Trellis/.test(fill.patternType);
  const grayOpacity = { mediumGray: 0.5, darkGray: 0.75, lightGray: 0.25, gray125: 0.125, gray0625: 0.0625 }[fill.patternType];
  const marks = grayOpacity != null
    ? `<rect width="8" height="8" fill="${fill.foreground}" opacity="${grayOpacity}"/>`
    : `${horizontal ? `<path d="M0 4H8"/>` : ""}${vertical ? `<path d="M4 0V8"/>` : ""}${down ? `<path d="M-2 -2L10 10M-2 6L2 10M6 -2L10 2"/>` : ""}${up ? `<path d="M-2 10L10 -2M-2 2L6 10M6 -2L10 2"/>` : ""}`;
  const stroke = grayOpacity == null ? ` stroke="${fill.foreground}" stroke-width="${strokeWidth}" fill="none"` : "";
  return { paint: `url(#${safeId})`, definition: `<pattern id="${safeId}" patternUnits="userSpaceOnUse" width="8" height="8"><rect width="8" height="8" fill="${fill.background}"/><g${stroke}>${marks}</g></pattern>` };
}

function parsedIndexedColors(stylesXml = "") {
  const body = /<indexedColors\b[^>]*>([\s\S]*?)<\/indexedColors>/.exec(String(stylesXml))?.[1] || "";
  const colors = [...body.matchAll(/<rgbColor\b([^>]*)\/?\s*>/g)].map((match) => {
    const value = parseAttrs(match[1]).rgb;
    return value ? `#${value.slice(-6)}` : undefined;
  });
  return colors.length ? colors : XLSX_INDEXED_COLORS;
}

function parsedColor(body, elementName, resources = {}) {
  const match = new RegExp(`<${elementName}\\b([^>]*)\\/?\\s*>`).exec(body);
  if (!match) return undefined;
  const attrs = parseAttrs(match[1]);
  if (attrs.rgb) return `#${attrs.rgb.slice(-6).toUpperCase()}`;
  const tint = attrs.tint == null || Number(attrs.tint) === 0 ? undefined : Number(attrs.tint);
  if (attrs.theme != null) {
    const theme = Number(attrs.theme), resolved = resources.themeColors?.[theme];
    return normalizeXlsxColorReference({ theme, ...(tint == null ? {} : { tint }), ...(resolved ? { resolved: tintedColor(resolved, tint) } : {}) });
  }
  if (attrs.indexed != null) {
    const indexed = Number(attrs.indexed), resolved = indexed === 64 ? resources.autoColor || "#000000" : indexed === 65 ? resources.background || "#FFFFFF" : resources.indexedColors?.[indexed];
    return normalizeXlsxColorReference({ indexed, ...(tint == null ? {} : { tint }), ...(resolved ? { resolved: tintedColor(resolved, tint) } : {}) });
  }
  if (booleanAttribute(attrs, "auto")) return normalizeXlsxColorReference({ auto: true, ...(tint == null ? {} : { tint }), resolved: tintedColor(resources.autoColor || "#000000", tint) });
  return undefined;
}

export function parseXlsxThemeColors(xml = "") {
  const text = String(xml || "");
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";
  const scheme = new RegExp(`<${prefix}clrScheme\\b[^>]*>([\\s\\S]*?)<\\/${prefix}clrScheme>`).exec(text)?.[1] || "";
  return XLSX_THEME_COLOR_NAMES.map((name) => {
    const body = new RegExp(`<${prefix}${name}\\b[^>]*>([\\s\\S]*?)<\\/${prefix}${name}>`).exec(scheme)?.[1] || "";
    const srgb = new RegExp(`<${prefix}srgbClr\\b[^>]*\\bval="([0-9A-Fa-f]{6})"`).exec(body)?.[1];
    const system = new RegExp(`<${prefix}sysClr\\b([^>]*)\\/?\\s*>`).exec(body);
    const systemAttrs = parseAttrs(system?.[1] || "");
    const value = srgb || systemAttrs.lastClr || systemAttrs.val;
    return value && /^[0-9A-Fa-f]{6}$/.test(value) ? `#${value.toUpperCase()}` : undefined;
  });
}

export function parseXlsxThemeConfig(xml = "") {
  const text = String(xml || "");
  const name = decodeXml(/<(?:[A-Za-z_][\w.-]*:)?theme\b[^>]*\bname="([^"]*)"/.exec(text)?.[1] || "Imported Office Theme");
  const parsed = parseXlsxThemeColors(text);
  return normalizeXlsxThemeConfig({ name, colors: Object.fromEntries(XLSX_THEME_COLOR_NAMES.map((colorName, index) => [colorName, parsed[index]])) });
}

function parseFont(body = "", resources = {}) {
  const underlineMatch = /<u\b([^>]*)\/?>/.exec(body);
  return {
    bold: booleanElement(body, "b"),
    italic: booleanElement(body, "i"),
    underline: underlineMatch ? parseAttrs(underlineMatch[1]).val || "single" : undefined,
    strike: booleanElement(body, "strike"),
    color: parsedColor(body, "color", resources),
    size: Number(/<sz[^>]*val="([^"]+)"/.exec(body)?.[1] || 11),
    name: /<name[^>]*val="([^"]+)"/.exec(body)?.[1] || "Aptos",
  };
}

function parseFill(body = "", resources = {}) {
  const pattern = /<patternFill\b([^>]*)\/?\s*>/.exec(body);
  if (!pattern) return undefined;
  const patternType = parseAttrs(pattern[1]).patternType || "none";
  if (patternType === "none") return undefined;
  const parsedForeground = parsedColor(body, "fgColor", resources);
  const parsedBackground = parsedColor(body, "bgColor", resources);
  const background = patternType === "solid" && parsedBackground?.indexed === 64 ? undefined : parsedBackground;
  const foreground = parsedForeground || (patternType === "solid" ? background : undefined);
  if (patternType === "solid" && typeof foreground === "string") return foreground;
  return normalizeXlsxFill({ patternType, foreground, ...(parsedForeground ? { background } : {}) });
}

function parseBorder(body = "", resources = {}, borderAttrs = {}) {
  const edges = {};
  for (const name of ["left", "right", "top", "bottom", "diagonal", "start", "end", "horizontal", "vertical"]) {
    const match = new RegExp(`<${name}\\b([^>]*)(?:\\/>|>([\\s\\S]*?)<\\/${name}>)`).exec(body);
    if (!match) continue;
    const attrs = parseAttrs(match[1]);
    if (!attrs.style) continue;
    edges[name] = { style: attrs.style, color: parsedColor(match[2] || "", "color", resources) || resources.autoColor || "#000000" };
  }
  const flags = {};
  for (const name of ["diagonalUp", "diagonalDown", "outline"]) if (borderAttrs[name] != null) flags[name] = booleanAttribute(borderAttrs, name);
  const perimeter = ["left", "right", "top", "bottom"].map((name) => edges[name]);
  const samePerimeter = perimeter.every(Boolean) && perimeter.every((edge) => edge.style === perimeter[0].style && edge.color === perimeter[0].color);
  const hasExtras = ["diagonal", "start", "end", "horizontal", "vertical"].some((name) => edges[name]) || Object.keys(flags).length;
  if (samePerimeter && !hasExtras) return { ...perimeter[0] };
  return Object.keys(edges).length || Object.keys(flags).length ? { ...edges, ...flags } : undefined;
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

function parseDxf(body = "", resources = {}) {
  const style = {};
  const fontBody = /<font>([\s\S]*?)<\/font>/.exec(body)?.[1];
  const fillBody = /<fill>([\s\S]*?)<\/fill>/.exec(body)?.[1];
  const numFmt = /<numFmt\b([^>]*)\/?>(?:<\/numFmt>)?/.exec(body)?.[1];
  if (fontBody != null) style.font = parseFont(fontBody, resources);
  const fill = fillBody != null ? parseFill(fillBody, resources) : undefined;
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

export function parseXlsxStylesXml(xml = "", options = {}) {
  const text = String(xml);
  const numberFormats = new Map(XLSX_BUILTIN_NUMBER_FORMATS);
  for (const match of text.matchAll(/<numFmt\b([^>]*)\/>/g)) { const attrs = parseAttrs(match[1]); numberFormats.set(Number(attrs.numFmtId), attrs.formatCode); }
  const colorResources = {
    themeColors: Array.isArray(options.themeColors) ? options.themeColors : [],
    indexedColors: parsedIndexedColors(text),
    autoColor: options.autoColor || "#000000",
  };
  const fontsBody = /<fonts\b[^>]*>([\s\S]*?)<\/fonts>/.exec(text)?.[1] || "";
  const fonts = [...fontsBody.matchAll(/<font>([\s\S]*?)<\/font>/g)].map((match) => parseFont(match[1], colorResources));
  const fillsBody = /<fills\b[^>]*>([\s\S]*?)<\/fills>/.exec(text)?.[1] || "";
  const fills = [...fillsBody.matchAll(/<fill>([\s\S]*?)<\/fill>/g)].map((match) => parseFill(match[1], colorResources));
  const bordersBody = /<borders\b[^>]*>([\s\S]*?)<\/borders>/.exec(text)?.[1] || "";
  const borders = [...bordersBody.matchAll(/<border\b([^>]*)\/\s*>|<border\b([^>]*)>([\s\S]*?)<\/border>/g)].map((match) => parseBorder(match[3] || "", colorResources, parseAttrs(match[1] || match[2] || "")));
  const resources = { ...colorResources, fonts, fills, borders, numberFormats };
  const styleXfs = xfRecords(/<cellStyleXfs\b[^>]*>([\s\S]*?)<\/cellStyleXfs>/.exec(text)?.[1] || "");
  const styles = xfRecords(/<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(text)?.[1] || "").map((record) => effectiveStyle(record, styleXfs[Number(record.attrs.xfId || 0)], resources));
  const dxfsBody = /<dxfs\b[^>]*>([\s\S]*?)<\/dxfs>/.exec(text)?.[1] || "";
  styles.dxfs = [...dxfsBody.matchAll(/<dxf>([\s\S]*?)<\/dxf>/g)].map((match) => parseDxf(match[1], resources));
  styles.indexedColors = colorResources.indexedColors;
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
