// Spreadsheet style normalization, display formatting, and SVG paint helpers.

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
  const cleaned = literalCleanup(section);
  // Remove color/condition brackets before deciding whether this is a date.
  // `[Red]($#,##0)` otherwise looks like a date only because "Red" has `d`.
  if (/[yd]/i.test(cleaned) || /(?:h+|s+|\[h\])(?=:|\b)/i.test(cleaned)) return dateDisplay(number, section, options.dateSystem || "1900");
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
