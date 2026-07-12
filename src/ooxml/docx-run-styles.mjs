import { attrEscape, attributes, decodeXml } from "./source-reference-xml.mjs";
import { resolveColorToken } from "../shared/colors.mjs";

const COLOR_DEFAULTS = Object.freeze({
  tx1: "#000000",
  bg1: "#ffffff",
  tx2: "#44546a",
  bg2: "#e7e6e6",
  accent1: "#4472c4",
  accent2: "#ed7d31",
  accent3: "#a5a5a5",
  accent4: "#ffc000",
  accent5: "#5b9bd5",
  accent6: "#70ad47",
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

const WORD_THEME_COLORS = Object.freeze({
  dark1: "tx1",
  light1: "bg1",
  dark2: "tx2",
  light2: "bg2",
  accent1: "accent1",
  accent2: "accent2",
  accent3: "accent3",
  accent4: "accent4",
  accent5: "accent5",
  accent6: "accent6",
  hyperlink: "hlink",
  followedHyperlink: "folHlink",
  none: undefined,
});

const THEME_FONTS = Object.freeze({
  majorAscii: "major",
  majorHAnsi: "major",
  majorEastAsia: "majorEastAsia",
  majorBidi: "majorComplexScript",
  minorAscii: "minor",
  minorHAnsi: "minor",
  minorEastAsia: "minorEastAsia",
  minorBidi: "minorComplexScript",
  none: undefined,
});

const DERIVED_RUN_STYLE_KEYS = new Set([
  "resolvedColor",
  "resolvedFontFamily",
  "resolvedFontFamilyEastAsia",
  "resolvedFontFamilyComplexScript",
  "effectiveColor",
  "effectiveFontFamily",
  "effectiveBold",
  "effectiveItalic",
  "effectiveFontSize",
]);

const STRUCTURAL_STYLE_KEYS = new Set(["id", "name", "type", "basedOn", "parent", "extends"]);

function objectValue(value, label) {
  if (value == null) return {};
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${label} must be an object.`);
  return value;
}

function normalizeHexColor(value, label) {
  const resolved = String(resolveColorToken(value, value) || "").trim();
  const short = /^#?([0-9a-f]{3})$/i.exec(resolved)?.[1];
  if (short) return `#${[...short].map((char) => char.repeat(2)).join("").toLowerCase()}`;
  const full = /^#?([0-9a-f]{6})$/i.exec(resolved)?.[1];
  if (!full) throw new TypeError(`${label} must be a three- or six-digit RGB color.`);
  return `#${full.toLowerCase()}`;
}

function normalizeByteHex(value, label) {
  if (value === undefined || value === null || value === "") return undefined;
  const raw = String(value).trim();
  if (!/^[0-9a-f]{2}$/i.test(raw)) throw new TypeError(`${label} must be a two-digit hexadecimal byte.`);
  return raw.toUpperCase();
}

function elementBlock(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml || ""))?.[0] || "";
}

function elementOpening(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml || ""))?.[0] || "";
}

function attribute(tag, name) {
  return Object.entries(attributes(tag)).find(([key]) => key === name || key.endsWith(`:${name}`))?.[1];
}

function parsedColor(block) {
  const srgb = attribute(elementOpening(block, "srgbClr"), "val");
  const system = attribute(elementOpening(block, "sysClr"), "lastClr");
  return /^[0-9a-f]{6}$/i.test(srgb || system || "") ? `#${String(srgb || system).toLowerCase()}` : undefined;
}

function fontGroup(xml, kind) {
  const block = elementBlock(xml, `${kind}Font`);
  return {
    latin: decodeXml(attribute(elementOpening(block, "latin"), "typeface") || ""),
    eastAsia: decodeXml(attribute(elementOpening(block, "ea"), "typeface") || ""),
    complexScript: decodeXml(attribute(elementOpening(block, "cs"), "typeface") || ""),
  };
}

export function normalizeDocxThemeConfig(config = {}, base = {}) {
  const input = objectValue(config, "DOCX theme");
  const existing = objectValue(base, "DOCX theme base");
  const colors = { ...COLOR_DEFAULTS, ...objectValue(existing.colors, "DOCX theme base colors") };
  for (const [key, value] of Object.entries(objectValue(input.colors, "DOCX theme colors"))) {
    if (!(key in COLOR_DEFAULTS)) throw new TypeError(`Unsupported DOCX theme color ${key}.`);
    colors[key] = normalizeHexColor(value, `DOCX theme color ${key}`);
  }
  const fonts = { ...FONT_DEFAULTS, ...objectValue(existing.fonts, "DOCX theme base fonts") };
  for (const [key, value] of Object.entries(objectValue(input.fonts, "DOCX theme fonts"))) {
    if (!(key in FONT_DEFAULTS)) throw new TypeError(`Unsupported DOCX theme font ${key}.`);
    fonts[key] = String(value ?? "");
  }
  if (!fonts.major.trim() || !fonts.minor.trim()) throw new TypeError("DOCX theme major and minor fonts must not be empty.");
  const name = String(input.name ?? existing.name ?? "Open Office Clean Room");
  if (!name.trim()) throw new TypeError("DOCX theme name must not be empty.");
  return { name, colors, fonts };
}

export function docxThemeXml(theme = {}) {
  const normalized = normalizeDocxThemeConfig(theme);
  const colors = Object.entries(THEME_COLOR_ELEMENTS).map(([key, element]) => `<a:${element}><a:srgbClr val="${normalized.colors[key].slice(1).toUpperCase()}"/></a:${element}>`).join("");
  const group = (kind, latin, eastAsia, complexScript) => `<a:${kind}Font><a:latin typeface="${attrEscape(latin)}"/><a:ea typeface="${attrEscape(eastAsia)}"/><a:cs typeface="${attrEscape(complexScript)}"/></a:${kind}Font>`;
  const fonts = `${group("major", normalized.fonts.major, normalized.fonts.majorEastAsia, normalized.fonts.majorComplexScript)}${group("minor", normalized.fonts.minor, normalized.fonts.minorEastAsia, normalized.fonts.minorComplexScript)}`;
  const format = '<a:fmtScheme name="Open Office Clean Room"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="25400"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln><a:ln w="38100"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:bgFillStyleLst></a:fmtScheme>';
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${attrEscape(normalized.name)}"><a:themeElements><a:clrScheme name="${attrEscape(normalized.name)}">${colors}</a:clrScheme><a:fontScheme name="${attrEscape(normalized.name)}">${fonts}</a:fontScheme>${format}</a:themeElements><a:objectDefaults/></a:theme>`;
}

export function parseDocxThemeXml(xml = "") {
  const colors = {};
  for (const [key, element] of Object.entries(THEME_COLOR_ELEMENTS)) {
    const value = parsedColor(elementBlock(xml, element));
    if (value) colors[key] = value;
  }
  const major = fontGroup(xml, "major");
  const minor = fontGroup(xml, "minor");
  return normalizeDocxThemeConfig({
    name: decodeXml(attribute(elementOpening(xml, "theme"), "name") || "Imported Theme"),
    colors,
    fonts: {
      major: major.latin || FONT_DEFAULTS.major,
      minor: minor.latin || FONT_DEFAULTS.minor,
      majorEastAsia: major.eastAsia,
      majorComplexScript: major.complexScript,
      minorEastAsia: minor.eastAsia,
      minorComplexScript: minor.complexScript,
    },
  });
}

function transformedColor(color, tint, shade) {
  const channels = color.slice(1).match(/../g).map((value) => Number.parseInt(value, 16));
  const shadeByte = shade === undefined ? undefined : Number.parseInt(shade, 16);
  const tintByte = tint === undefined ? undefined : Number.parseInt(tint, 16);
  return `#${channels.map((channel) => {
    let value = channel;
    if (shadeByte !== undefined) value *= shadeByte / 255;
    if (tintByte !== undefined) value += (255 - value) * (tintByte / 255);
    return Math.max(0, Math.min(255, Math.round(value))).toString(16).padStart(2, "0");
  }).join("")}`;
}

function resolveThemeFont(value, theme) {
  return theme.fonts[THEME_FONTS[value]] || undefined;
}

export function resolveDocxRunStyle(style = {}, theme = {}) {
  const normalizedTheme = normalizeDocxThemeConfig(theme);
  const resolved = { ...style };
  const colorKey = WORD_THEME_COLORS[style.themeColor];
  const themeColor = colorKey ? normalizedTheme.colors[colorKey] : undefined;
  if (themeColor) resolved.resolvedColor = transformedColor(themeColor, style.themeTint, style.themeShade);
  else if (style.color && style.color !== "auto") resolved.resolvedColor = normalizeHexColor(style.color, "DOCX run color");
  const font = resolveThemeFont(style.fontTheme, normalizedTheme) || style.fontFamily;
  const highAnsi = resolveThemeFont(style.fontThemeHighAnsi, normalizedTheme) || style.fontFamilyHighAnsi || font;
  const eastAsia = resolveThemeFont(style.fontThemeEastAsia, normalizedTheme) || style.fontFamilyEastAsia || font;
  const complexScript = resolveThemeFont(style.fontThemeComplexScript, normalizedTheme) || style.fontFamilyComplexScript || font;
  if (font || highAnsi) resolved.resolvedFontFamily = font || highAnsi;
  if (eastAsia) resolved.resolvedFontFamilyEastAsia = eastAsia;
  if (complexScript) resolved.resolvedFontFamilyComplexScript = complexScript;
  return resolved;
}

export function normalizeDocxRunStyle(style = {}, theme = {}) {
  const input = objectValue(style, "DOCX run style");
  const normalized = { ...input };
  for (const key of ["bold", "italic", "boldComplexScript", "italicComplexScript"]) {
    if (key in normalized && typeof normalized[key] !== "boolean") throw new TypeError(`DOCX run style ${key} must be a boolean.`);
  }
  for (const key of ["fontSize", "fontSizeComplexScript"]) {
    if (!(key in normalized)) continue;
    const value = Number(normalized[key]);
    if (!Number.isFinite(value) || value <= 0 || value > 3276) throw new RangeError(`DOCX run style ${key} must be greater than 0 and at most 3276 half-points.`);
    normalized[key] = value;
  }
  if (normalized.color && normalized.color !== "auto") normalized.color = normalizeHexColor(normalized.color, "DOCX run color");
  if (normalized.themeColor && !(normalized.themeColor in WORD_THEME_COLORS)) throw new TypeError(`Unsupported DOCX run themeColor ${normalized.themeColor}.`);
  const themeTint = normalizeByteHex(normalized.themeTint, "DOCX run themeTint");
  const themeShade = normalizeByteHex(normalized.themeShade, "DOCX run themeShade");
  if (themeTint === undefined) delete normalized.themeTint;
  else normalized.themeTint = themeTint;
  if (themeShade === undefined) delete normalized.themeShade;
  else normalized.themeShade = themeShade;
  for (const key of ["fontTheme", "fontThemeHighAnsi", "fontThemeEastAsia", "fontThemeComplexScript"]) {
    if (normalized[key] && !(normalized[key] in THEME_FONTS)) throw new TypeError(`Unsupported DOCX run ${key} ${normalized[key]}.`);
  }
  if (normalized.fontHint && !new Set(["default", "eastAsia", "cs"]).has(normalized.fontHint)) throw new TypeError(`Unsupported DOCX run fontHint ${normalized.fontHint}.`);
  return resolveDocxRunStyle(normalized, theme);
}

function sourceRunStyle(style = {}) {
  return Object.fromEntries(Object.entries(style).filter(([key, value]) => value !== undefined && !DERIVED_RUN_STYLE_KEYS.has(key) && !STRUCTURAL_STYLE_KEYS.has(key)));
}

export function mergeDocxRunStyleCascade(styles = [], theme = {}) {
  const merged = {};
  for (const style of styles) {
    const source = sourceRunStyle(style || {});
    if ("color" in source || "themeColor" in source) {
      for (const key of ["color", "themeColor", "themeTint", "themeShade"]) delete merged[key];
    }
    for (const [direct, themed] of [
      ["fontFamily", "fontTheme"],
      ["fontFamilyHighAnsi", "fontThemeHighAnsi"],
      ["fontFamilyEastAsia", "fontThemeEastAsia"],
      ["fontFamilyComplexScript", "fontThemeComplexScript"],
    ]) {
      if (direct in source || themed in source) {
        delete merged[direct];
        delete merged[themed];
      }
    }
    Object.assign(merged, source);
  }
  return normalizeDocxRunStyle(merged, theme);
}

function onOffXml(tag, style, key) {
  if (!(key in style)) return "";
  return style[key] ? `<w:${tag}/>` : `<w:${tag} w:val="0"/>`;
}

export function docxRunPropertiesXml(style = {}, theme = {}) {
  const value = normalizeDocxRunStyle(style, theme);
  const fontHint = value.fontHint || (value.fontThemeComplexScript || value.fontFamilyComplexScript ? "cs" : value.fontThemeEastAsia || value.fontFamilyEastAsia ? "eastAsia" : undefined);
  const fontAttrs = [
    ["ascii", value.fontFamily], ["hAnsi", value.fontFamilyHighAnsi ?? value.fontFamily], ["eastAsia", value.fontFamilyEastAsia], ["cs", value.fontFamilyComplexScript],
    ["asciiTheme", value.fontTheme], ["hAnsiTheme", value.fontThemeHighAnsi ?? value.fontTheme], ["eastAsiaTheme", value.fontThemeEastAsia], ["cstheme", value.fontThemeComplexScript],
    ["hint", fontHint],
  ].filter(([, item]) => item !== undefined && item !== "").map(([key, item]) => ` w:${key}="${attrEscape(item)}"`).join("");
  const fonts = fontAttrs ? `<w:rFonts${fontAttrs}/>` : "";
  const colorAttrs = [];
  if (value.color) colorAttrs.push(`w:val="${attrEscape(value.color === "auto" ? "auto" : value.color.slice(1))}"`);
  else if (value.themeColor) colorAttrs.push(`w:val="${attrEscape((value.resolvedColor || "#000000").slice(1))}"`);
  if (value.themeColor) colorAttrs.push(`w:themeColor="${attrEscape(value.themeColor)}"`);
  if (value.themeTint) colorAttrs.push(`w:themeTint="${value.themeTint}"`);
  if (value.themeShade) colorAttrs.push(`w:themeShade="${value.themeShade}"`);
  const color = colorAttrs.length ? `<w:color ${colorAttrs.join(" ")}/>` : "";
  const runStyle = value.runStyleId ? `<w:rStyle w:val="${attrEscape(value.runStyleId)}"/>` : "";
  const body = `${runStyle}${fonts}${onOffXml("b", value, "bold")}${onOffXml("bCs", value, "boldComplexScript")}${onOffXml("i", value, "italic")}${onOffXml("iCs", value, "italicComplexScript")}${color}${value.fontSize ? `<w:sz w:val="${Math.round(value.fontSize)}"/>` : ""}${value.fontSizeComplexScript ? `<w:szCs w:val="${Math.round(value.fontSizeComplexScript)}"/>` : ""}`;
  return body ? `<w:rPr>${body}</w:rPr>` : "";
}

function onOffProperty(xml, localName) {
  const tag = elementOpening(xml, localName);
  if (!tag) return undefined;
  return !new Set(["false", "0", "off", "no"]).has(String(attribute(tag, "val") ?? "true").toLowerCase());
}

export function parseDocxRunPropertiesXml(xml = "", theme = {}) {
  const fontTag = elementOpening(xml, "rFonts");
  const colorTag = elementOpening(xml, "color");
  const size = Number(attribute(elementOpening(xml, "sz"), "val"));
  const sizeComplexScript = Number(attribute(elementOpening(xml, "szCs"), "val"));
  const style = {
    ...(attribute(fontTag, "ascii") ? { fontFamily: decodeXml(attribute(fontTag, "ascii")) } : {}),
    ...(attribute(fontTag, "hAnsi") ? { fontFamilyHighAnsi: decodeXml(attribute(fontTag, "hAnsi")) } : {}),
    ...(attribute(fontTag, "eastAsia") ? { fontFamilyEastAsia: decodeXml(attribute(fontTag, "eastAsia")) } : {}),
    ...(attribute(fontTag, "cs") ? { fontFamilyComplexScript: decodeXml(attribute(fontTag, "cs")) } : {}),
    ...(attribute(fontTag, "asciiTheme") ? { fontTheme: attribute(fontTag, "asciiTheme") } : {}),
    ...(attribute(fontTag, "hAnsiTheme") ? { fontThemeHighAnsi: attribute(fontTag, "hAnsiTheme") } : {}),
    ...(attribute(fontTag, "eastAsiaTheme") ? { fontThemeEastAsia: attribute(fontTag, "eastAsiaTheme") } : {}),
    ...(attribute(fontTag, "cstheme") ? { fontThemeComplexScript: attribute(fontTag, "cstheme") } : {}),
    ...(attribute(fontTag, "hint") ? { fontHint: attribute(fontTag, "hint") } : {}),
    ...(attribute(colorTag, "val") ? { color: attribute(colorTag, "val") === "auto" ? "auto" : `#${attribute(colorTag, "val").toLowerCase()}` } : {}),
    ...(attribute(colorTag, "themeColor") ? { themeColor: attribute(colorTag, "themeColor") } : {}),
    ...(attribute(colorTag, "themeTint") ? { themeTint: attribute(colorTag, "themeTint") } : {}),
    ...(attribute(colorTag, "themeShade") ? { themeShade: attribute(colorTag, "themeShade") } : {}),
    ...(Number.isFinite(size) && size > 0 ? { fontSize: size } : {}),
    ...(Number.isFinite(sizeComplexScript) && sizeComplexScript > 0 ? { fontSizeComplexScript: sizeComplexScript } : {}),
  };
  for (const [tag, key] of [["b", "bold"], ["bCs", "boldComplexScript"], ["i", "italic"], ["iCs", "italicComplexScript"]]) {
    const value = onOffProperty(xml, tag);
    if (value !== undefined) style[key] = value;
  }
  return normalizeDocxRunStyle(style, theme);
}

export function parseDocxRunStyleId(xml = "") {
  const value = decodeXml(attribute(elementOpening(xml, "rStyle"), "val") || "").trim();
  return value || undefined;
}

export function parseDocxDefaultRunPropertiesXml(stylesXml = "", theme = {}) {
  const defaults = elementBlock(stylesXml, "docDefaults");
  const runDefaults = elementBlock(defaults, "rPrDefault");
  const properties = elementBlock(runDefaults, "rPr");
  return properties ? parseDocxRunPropertiesXml(properties, theme) : normalizeDocxRunStyle({}, theme);
}

export function effectiveDocxRunStyle(style = {}, text = "", theme = {}) {
  const value = normalizeDocxRunStyle(style, theme);
  const complexScript = /[\u0590-\u08ff]/u.test(text);
  const eastAsia = /[\u3040-\u30ff\u3400-\u9fff\uac00-\ud7af]/u.test(text);
  return {
    ...value,
    effectiveFontFamily: complexScript ? (value.resolvedFontFamilyComplexScript || value.resolvedFontFamily) : eastAsia ? (value.resolvedFontFamilyEastAsia || value.resolvedFontFamily) : value.resolvedFontFamily,
    effectiveBold: complexScript && value.boldComplexScript !== undefined ? value.boldComplexScript : value.bold,
    effectiveItalic: complexScript && value.italicComplexScript !== undefined ? value.italicComplexScript : value.italic,
    effectiveFontSize: complexScript && value.fontSizeComplexScript ? value.fontSizeComplexScript : value.fontSize,
    effectiveColor: value.resolvedColor || (value.color !== "auto" ? value.color : undefined),
  };
}
