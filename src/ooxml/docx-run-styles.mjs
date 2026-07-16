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
  if ("runStyleId" in normalized) {
    if (typeof normalized.runStyleId !== "string" || !normalized.runStyleId.trim()) throw new TypeError("DOCX run style runStyleId must be a non-empty string.");
    normalized.runStyleId = normalized.runStyleId.trim();
  }
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
