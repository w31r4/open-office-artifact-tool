// Clean-room SpreadsheetML style codec and deterministic display formatter.

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

function parseFont(body = "") {
  const underlineMatch = /<u\b([^>]*)\/?>/.exec(body);
  return {
    bold: booleanElement(body, "b"),
    italic: booleanElement(body, "i"),
    underline: underlineMatch ? parseAttrs(underlineMatch[1]).val || "single" : undefined,
    strike: booleanElement(body, "strike"),
    color: /<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1] ? `#${/<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1]}` : undefined,
    size: Number(/<sz[^>]*val="([^"]+)"/.exec(body)?.[1] || 11),
    name: /<name[^>]*val="([^"]+)"/.exec(body)?.[1] || "Aptos",
  };
}

function parseFill(body = "") {
  const color = /<fgColor[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(body)?.[1];
  return color ? `#${color}` : undefined;
}

function parseBorder(body = "") {
  const edge = /<(left|right|top|bottom)\b([^>]*)>([\s\S]*?)<\/\1>/.exec(body);
  if (!edge) return undefined;
  const attrs = parseAttrs(edge[2]);
  const color = /<color[^>]*rgb="(?:FF)?([0-9A-Fa-f]{6})"/.exec(edge[3])?.[1];
  return { style: attrs.style || "thin", ...(color ? { color: `#${color}` } : {}) };
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
  [9, "0%"], [10, "0.00%"], [11, "0.00E+00"], [12, "# ?/?"], [13, "# ??/??"],
  [14, "mm-dd-yy"], [15, "d-mmm-yy"], [16, "d-mmm"], [17, "mmm-yy"],
  [18, "h:mm AM/PM"], [19, "h:mm:ss AM/PM"], [20, "h:mm"], [21, "h:mm:ss"], [22, "m/d/yy h:mm"],
  [37, "#,##0;(#,##0)"], [38, "#,##0;[Red](#,##0)"], [39, "#,##0.00;(#,##0.00)"], [40, "#,##0.00;[Red](#,##0.00)"],
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
    numberFormat: resources.numberFormats.get(numberFormatId),
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
