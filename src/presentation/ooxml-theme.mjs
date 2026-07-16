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
