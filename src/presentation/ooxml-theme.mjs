import { attributes, attrEscape, decodeXml } from "../ooxml/source-reference-xml.mjs";
import { resolveColorToken } from "../shared/colors.mjs";

const COLOR_DEFAULTS = Object.freeze({
  tx1: "#0f172a",
  bg1: "#ffffff",
  tx2: "#1e293b",
  bg2: "#f8fafc",
  accent1: "#156082",
  accent2: "#0ea5e9",
  accent3: "#64748b",
  accent4: "#7c3aed",
  accent5: "#16a34a",
  accent6: "#f97316",
  hlink: "#0563c1",
  folHlink: "#954f72",
});

const FONT_DEFAULTS = Object.freeze({
  major: "Aptos Display",
  minor: "Aptos",
  majorEastAsia: "",
  majorComplexScript: "",
  minorEastAsia: "",
  minorComplexScript: "",
});

const TEXT_STYLE_DEFAULTS = Object.freeze({
  title: Object.freeze({ fontSize: 32, bold: true, italic: false, color: "tx1", fontFamily: "+mj-lt", alignment: "left" }),
  body: Object.freeze({ fontSize: 18, bold: false, italic: false, color: "tx1", fontFamily: "+mn-lt", alignment: "left" }),
  other: Object.freeze({ fontSize: 18, bold: false, italic: false, color: "tx1", fontFamily: "+mn-lt", alignment: "left" }),
});

const COLOR_MAP_DEFAULTS = Object.freeze({
  bg1: "lt1",
  tx1: "dk1",
  bg2: "lt2",
  tx2: "dk2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
});

const COLOR_KEYS = new Set(Object.keys(COLOR_DEFAULTS));
const COLOR_ALIASES = Object.freeze({ dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" });
const FONT_KEYS = new Set(Object.keys(FONT_DEFAULTS));
const TEXT_STYLE_KEYS = new Set(Object.keys(TEXT_STYLE_DEFAULTS));
const TEXT_STYLE_FIELDS = new Set(["fontSize", "bold", "italic", "color", "fontFamily", "alignment"]);
const COLOR_MAP_VALUES = new Set(["dk1", "lt1", "dk2", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);
const ALIGNMENTS = new Set(["left", "center", "right", "justify"]);

const THEME_COLOR_ELEMENTS = Object.freeze({
  tx1: "dk1",
  bg1: "lt1",
  tx2: "dk2",
  bg2: "lt2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hlink: "hlink",
  folHlink: "folHlink",
});

function assertObject(value, label) {
  if (value == null) return {};
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  return value;
}

function normalizeHexColor(value, label) {
  const resolved = String(resolveColorToken(value, value) || "").trim();
  const short = /^#?([0-9a-f]{3})$/i.exec(resolved)?.[1];
  if (short) return `#${[...short].map((char) => char.repeat(2)).join("").toLowerCase()}`;
  const full = /^#?([0-9a-f]{6})$/i.exec(resolved)?.[1];
  if (!full) throw new TypeError(`${label} must be a six-digit RGB color or supported color token.`);
  return `#${full.toLowerCase()}`;
}

function normalizeColors(value, base) {
  const colors = { ...base };
  for (const [rawKey, rawValue] of Object.entries(assertObject(value, "presentation theme colors"))) {
    const key = COLOR_ALIASES[rawKey] || rawKey;
    if (!COLOR_KEYS.has(key)) throw new TypeError(`Unsupported presentation theme color ${rawKey}.`);
    colors[key] = normalizeHexColor(rawValue, `presentation theme color ${rawKey}`);
  }
  return colors;
}

function normalizeFonts(value, base) {
  const fonts = { ...base };
  for (const [key, rawValue] of Object.entries(assertObject(value, "presentation theme fonts"))) {
    if (!FONT_KEYS.has(key)) throw new TypeError(`Unsupported presentation theme font ${key}.`);
    const font = String(rawValue ?? "");
    if ((key === "major" || key === "minor") && !font.trim()) throw new TypeError(`Presentation theme font ${key} must not be empty.`);
    fonts[key] = font;
  }
  return fonts;
}

function normalizeTextStyle(value, base, key) {
  const style = { ...base };
  for (const [field, rawValue] of Object.entries(assertObject(value, `presentation theme ${key} text style`))) {
    if (!TEXT_STYLE_FIELDS.has(field)) throw new TypeError(`Unsupported presentation theme ${key} text style field ${field}.`);
    if (field === "fontSize") {
      const number = Number(rawValue);
      if (!Number.isFinite(number) || number <= 0 || number > 400) throw new RangeError(`Presentation theme ${key} fontSize must be greater than 0 and at most 400.`);
      style.fontSize = number;
    } else if (field === "bold" || field === "italic") {
      if (typeof rawValue !== "boolean") throw new TypeError(`Presentation theme ${key} ${field} must be a boolean.`);
      style[field] = rawValue;
    } else if (field === "color") {
      const color = String(rawValue || "");
      style.color = COLOR_MAP_VALUES.has(color) || color === "tx1" || color === "tx2" || color === "bg1" || color === "bg2" ? color : normalizeHexColor(color, `presentation theme ${key} text color`);
    } else if (field === "fontFamily") {
      if (!String(rawValue || "").trim()) throw new TypeError(`Presentation theme ${key} fontFamily must not be empty.`);
      style.fontFamily = String(rawValue);
    } else {
      if (!ALIGNMENTS.has(rawValue)) throw new TypeError(`Presentation theme ${key} alignment must be left, center, right, or justify.`);
      style.alignment = rawValue;
    }
  }
  return style;
}

function normalizeTextStyles(value, base) {
  const styles = Object.fromEntries(Object.entries(base).map(([key, style]) => [key, { ...style }]));
  for (const [key, rawValue] of Object.entries(assertObject(value, "presentation theme textStyles"))) {
    if (!TEXT_STYLE_KEYS.has(key)) throw new TypeError(`Unsupported presentation theme text style ${key}.`);
    styles[key] = normalizeTextStyle(rawValue, styles[key], key);
  }
  return styles;
}

function normalizeColorMap(value, base) {
  const colorMap = { ...base };
  for (const [key, rawValue] of Object.entries(assertObject(value, "presentation theme colorMap"))) {
    if (!(key in COLOR_MAP_DEFAULTS)) throw new TypeError(`Unsupported presentation theme colorMap key ${key}.`);
    const target = String(rawValue || "");
    if (!COLOR_MAP_VALUES.has(target)) throw new TypeError(`Presentation theme colorMap ${key} has unsupported target ${target}.`);
    colorMap[key] = target;
  }
  return colorMap;
}

export function normalizePresentationThemeConfig(config = {}, base = {}) {
  const input = assertObject(config, "presentation theme");
  const existing = assertObject(base, "presentation theme base");
  const name = input.name ?? existing.name ?? "Open Office Clean Room";
  if (!String(name).trim()) throw new TypeError("Presentation theme name must not be empty.");
  return {
    name: String(name),
    colors: normalizeColors(input.colors, existing.colors || COLOR_DEFAULTS),
    fonts: normalizeFonts(input.fonts, existing.fonts || FONT_DEFAULTS),
    textStyles: normalizeTextStyles(input.textStyles, existing.textStyles || TEXT_STYLE_DEFAULTS),
    colorMap: normalizeColorMap(input.colorMap, existing.colorMap || COLOR_MAP_DEFAULTS),
  };
}

function colorValue(value) {
  return normalizeHexColor(value, "presentation theme color").slice(1).toUpperCase();
}

function schemeColorXml(value) {
  const scheme = { tx1: "tx1", tx2: "tx2", bg1: "bg1", bg2: "bg2" }[value] || (COLOR_MAP_VALUES.has(value) ? value : undefined);
  return scheme ? `<a:schemeClr val="${scheme}"/>` : `<a:srgbClr val="${colorValue(value)}"/>`;
}

function formatSchemeXml(name) {
  const fills = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="50000"/><a:satMod val="200000"/></a:schemeClr></a:solidFill>`;
  const lines = [6350, 12700, 19050].map((width) => `<a:ln w="${width}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join("");
  const effects = "<a:effectStyle><a:effectLst/></a:effectStyle>".repeat(3);
  const backgrounds = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/><a:satMod val="200000"/></a:schemeClr></a:solidFill>`;
  return `<a:fmtScheme name="${attrEscape(name)}"><a:fillStyleLst>${fills}</a:fillStyleLst><a:lnStyleLst>${lines}</a:lnStyleLst><a:effectStyleLst>${effects}</a:effectStyleLst><a:bgFillStyleLst>${backgrounds}</a:bgFillStyleLst></a:fmtScheme>`;
}

export function presentationThemeXml(theme = {}) {
  const normalized = normalizePresentationThemeConfig(theme);
  const colors = Object.entries(THEME_COLOR_ELEMENTS).map(([key, element]) => `<a:${element}><a:srgbClr val="${colorValue(normalized.colors[key])}"/></a:${element}>`).join("");
  const fontGroup = (kind, latin, eastAsia, complexScript) => `<a:${kind}Font><a:latin typeface="${attrEscape(latin)}"/><a:ea typeface="${attrEscape(eastAsia)}"/><a:cs typeface="${attrEscape(complexScript)}"/></a:${kind}Font>`;
  const fonts = `${fontGroup("major", normalized.fonts.major, normalized.fonts.majorEastAsia, normalized.fonts.majorComplexScript)}${fontGroup("minor", normalized.fonts.minor, normalized.fonts.minorEastAsia, normalized.fonts.minorComplexScript)}`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${attrEscape(normalized.name)}"><a:themeElements><a:clrScheme name="${attrEscape(normalized.name)}">${colors}</a:clrScheme><a:fontScheme name="${attrEscape(normalized.name)}">${fonts}</a:fontScheme>${formatSchemeXml(normalized.name)}</a:themeElements><a:objectDefaults/></a:theme>`;
}

function textStyleXml(kind, style) {
  const tag = { title: "titleStyle", body: "bodyStyle", other: "otherStyle" }[kind];
  const alignment = { left: "l", center: "ctr", right: "r", justify: "just" }[style.alignment];
  const fontSize = Math.round(style.fontSize * 100);
  const font = attrEscape(style.fontFamily);
  const eastAsia = font.endsWith("-lt") ? `${font.slice(0, -2)}ea` : font;
  const complexScript = font.endsWith("-lt") ? `${font.slice(0, -2)}cs` : font;
  const run = `<a:defRPr sz="${fontSize}" b="${style.bold ? 1 : 0}" i="${style.italic ? 1 : 0}"><a:solidFill>${schemeColorXml(style.color)}</a:solidFill><a:latin typeface="${font}"/><a:ea typeface="${eastAsia}"/><a:cs typeface="${complexScript}"/></a:defRPr>`;
  const levels = Array.from({ length: 9 }, (_, index) => `<a:lvl${index + 1}pPr marL="${index * 457200}" indent="${index ? -228600 : 0}" algn="${alignment}" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${run}</a:lvl${index + 1}pPr>`).join("");
  return `<p:${tag}>${levels}</p:${tag}>`;
}

export function presentationSlideMasterXml(layoutParts = [], theme = {}, options = {}) {
  const normalized = normalizePresentationThemeConfig(theme);
  const ids = layoutParts.map((part, index) => `<p:sldLayoutId id="${2147483649 + index}" r:id="${attrEscape(part.masterRelId)}"/>`).join("");
  const colorMap = Object.entries(normalized.colorMap).map(([key, value]) => `${key}="${attrEscape(value)}"`).join(" ");
  const textStyles = Object.entries(normalized.textStyles).map(([kind, style]) => textStyleXml(kind, style)).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sldMaster xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" preserve="1"><p:cSld name="${attrEscape(options.name || "Default Master")}">${options.backgroundXml || ""}<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr/>${options.placeholdersXml || ""}</p:spTree></p:cSld><p:clrMap ${colorMap}/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst><p:txStyles>${textStyles}</p:txStyles></p:sldMaster>`;
}

function elementBlock(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml || ""))?.[0] || "";
}

function elementOpening(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml || ""))?.[0] || "";
}

function parsedColor(block) {
  const srgb = /<(?:[A-Za-z_][\w.-]*:)?srgbClr\b[^>]*\bval="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
  const system = /<(?:[A-Za-z_][\w.-]*:)?sysClr\b[^>]*\blastClr="([0-9A-Fa-f]{6})"/.exec(block)?.[1];
  return srgb || system ? `#${String(srgb || system).toLowerCase()}` : undefined;
}

function parseFontGroup(xml, kind) {
  const block = elementBlock(xml, `${kind}Font`);
  return {
    latin: decodeXml(attributes(elementOpening(block, "latin")).typeface || ""),
    eastAsia: decodeXml(attributes(elementOpening(block, "ea")).typeface || ""),
    complexScript: decodeXml(attributes(elementOpening(block, "cs")).typeface || ""),
  };
}

export function parsePresentationThemeXml(xml = "") {
  const themeTag = elementOpening(xml, "theme");
  const colors = {};
  for (const [key, element] of Object.entries(THEME_COLOR_ELEMENTS)) {
    const value = parsedColor(elementBlock(xml, element));
    if (value) colors[key] = value;
  }
  const major = parseFontGroup(xml, "major");
  const minor = parseFontGroup(xml, "minor");
  return {
    name: decodeXml(attributes(themeTag).name || "Imported Theme"),
    colors,
    fonts: {
      major: major.latin || FONT_DEFAULTS.major,
      minor: minor.latin || FONT_DEFAULTS.minor,
      majorEastAsia: major.eastAsia,
      majorComplexScript: major.complexScript,
      minorEastAsia: minor.eastAsia,
      minorComplexScript: minor.complexScript,
    },
  };
}

function parseMasterTextStyle(xml, kind) {
  const block = elementBlock(xml, { title: "titleStyle", body: "bodyStyle", other: "otherStyle" }[kind]);
  if (!block) return undefined;
  const levelTag = elementOpening(block, "lvl1pPr");
  const runTag = elementOpening(block, "defRPr");
  const level = attributes(levelTag);
  const run = attributes(runTag);
  const colorBlock = elementBlock(block, "solidFill");
  const scheme = attributes(elementOpening(colorBlock, "schemeClr")).val;
  const rgb = attributes(elementOpening(colorBlock, "srgbClr")).val;
  const fontFamily = decodeXml(attributes(elementOpening(block, "latin")).typeface || TEXT_STYLE_DEFAULTS[kind].fontFamily);
  const alignment = { l: "left", ctr: "center", r: "right", just: "justify" }[level.algn] || "left";
  return {
    fontSize: run.sz && Number.isFinite(Number(run.sz)) ? Number(run.sz) / 100 : TEXT_STYLE_DEFAULTS[kind].fontSize,
    bold: run.b === "1" || run.b === "true",
    italic: run.i === "1" || run.i === "true",
    color: scheme || (rgb ? `#${rgb.toLowerCase()}` : TEXT_STYLE_DEFAULTS[kind].color),
    fontFamily,
    alignment,
  };
}

export function parsePresentationSlideMasterThemeXml(xml = "") {
  const colorMapTag = elementOpening(xml, "clrMap");
  const colorMapAttrs = attributes(colorMapTag);
  const colorMap = Object.fromEntries(Object.keys(COLOR_MAP_DEFAULTS).filter((key) => COLOR_MAP_VALUES.has(colorMapAttrs[key])).map((key) => [key, colorMapAttrs[key]]));
  const textStyles = Object.fromEntries([...TEXT_STYLE_KEYS].map((key) => [key, parseMasterTextStyle(xml, key)]).filter(([, value]) => value));
  return { colorMap, textStyles };
}
