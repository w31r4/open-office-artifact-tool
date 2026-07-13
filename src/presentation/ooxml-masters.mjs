import { attributes, attrEscape } from "../ooxml/source-reference-xml.mjs";
import { resolveColorToken } from "../shared/colors.mjs";

const SCHEME_COLORS = new Set(["dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);

function colorValue(value) {
  const resolved = String(resolveColorToken(value, value) || "").trim();
  const short = /^#?([0-9a-f]{3})$/i.exec(resolved)?.[1];
  if (short) return [...short].map((char) => char.repeat(2)).join("").toUpperCase();
  const full = /^#?([0-9a-f]{6})$/i.exec(resolved)?.[1];
  if (!full) throw new TypeError("Presentation background fill must be a scheme color, six-digit RGB color, or supported color token.");
  return full.toUpperCase();
}

function fillXml(fill) {
  return SCHEME_COLORS.has(fill) ? `<a:schemeClr val="${attrEscape(fill)}"/>` : `<a:srgbClr val="${colorValue(fill)}"/>`;
}

export function presentationColorXml(value, fallback = "#0f172a") {
  const fill = String(value || fallback).trim();
  return `<a:solidFill>${fillXml(fill)}</a:solidFill>`;
}

export function normalizePresentationBackground(value, fallback) {
  if (value == null) return fallback == null ? undefined : normalizePresentationBackground(fallback);
  const input = typeof value === "string" ? { fill: value } : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Presentation background must be a color string or object.");
  const fill = String(input.fill || input.color || "").trim();
  if (!fill) throw new TypeError("Presentation background requires fill.");
  if (!SCHEME_COLORS.has(fill)) colorValue(fill);
  const mode = input.mode || input.type || "solid";
  if (!new Set(["solid", "reference"]).has(mode)) throw new TypeError("Presentation background mode must be solid or reference.");
  const index = mode === "reference" ? Number(input.index ?? input.idx ?? 1001) : undefined;
  if (mode === "reference" && (!Number.isInteger(index) || index < 0 || index > 4_294_967_295)) throw new RangeError("Presentation background reference index must be an unsigned 32-bit integer.");
  return { fill, mode, ...(index == null ? {} : { index }) };
}

export function presentationBackgroundXml(background) {
  if (!background) return "";
  const normalized = normalizePresentationBackground(background);
  return normalized.mode === "reference"
    ? `<p:bg><p:bgRef idx="${normalized.index}">${fillXml(normalized.fill)}</p:bgRef></p:bg>`
    : `<p:bg><p:bgPr><a:solidFill>${fillXml(normalized.fill)}</a:solidFill><a:effectLst/></p:bgPr></p:bg>`;
}

function elementBlock(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml || ""))?.[0] || "";
}

function elementOpening(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml || ""))?.[0] || "";
}

function parsedFill(xml) {
  const scheme = attributes(elementOpening(xml, "schemeClr")).val;
  if (scheme) return scheme;
  const rgb = attributes(elementOpening(xml, "srgbClr")).val;
  return rgb ? `#${String(rgb).toLowerCase()}` : undefined;
}

export function parsePresentationBackgroundXml(xml = "") {
  const block = elementBlock(xml, "bg");
  if (!block) return undefined;
  const reference = elementOpening(block, "bgRef");
  const fill = parsedFill(block);
  if (!fill) return undefined;
  if (reference) return normalizePresentationBackground({ fill, mode: "reference", index: Number(attributes(reference).idx || 1001) });
  return normalizePresentationBackground({ fill, mode: "solid" });
}

export function parsePresentationPlaceholderStyleXml(xml = "") {
  const runBlock = elementBlock(xml, "rPr") || elementBlock(xml, "defRPr");
  const run = attributes(elementOpening(runBlock || xml, runBlock ? (/defRPr/.test(runBlock) ? "defRPr" : "rPr") : "rPr"));
  const paragraph = attributes(elementOpening(xml, "pPr"));
  const fill = parsedFill(runBlock);
  const fontFamily = attributes(elementOpening(runBlock, "latin")).typeface;
  const alignment = { l: "left", ctr: "center", r: "right", just: "justify" }[paragraph.algn];
  return {
    ...(run.sz && Number.isFinite(Number(run.sz)) ? { fontSize: Number(run.sz) / 75 } : {}),
    ...(run.b != null ? { bold: ["1", "true", "on"].includes(String(run.b).toLowerCase()) } : {}),
    ...(run.i != null ? { italic: ["1", "true", "on"].includes(String(run.i).toLowerCase()) } : {}),
    ...(fill ? { color: fill } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(alignment ? { alignment } : {}),
  };
}

function placeholderKey(placeholder) {
  return `${placeholder.type || "body"}:${Number(placeholder.idx || 1)}`;
}

export function mergePresentationPlaceholders(masterPlaceholders = [], layoutPlaceholders = []) {
  const cloneParagraphStyles = (styles = {}) => Object.fromEntries(Object.entries(styles).map(([level, style]) => [level, { ...style, style: { ...(style.style || {}) } }]));
  const mergeParagraphStyles = (base = {}, overrides = {}) => {
    const result = cloneParagraphStyles(base);
    for (const [level, style] of Object.entries(overrides)) result[level] = { ...(result[level] || {}), ...style, style: { ...(result[level]?.style || {}), ...(style.style || {}) } };
    return result;
  };
  const merged = new Map(masterPlaceholders.map((placeholder) => [placeholderKey(placeholder), { ...placeholder, position: placeholder.position && { ...placeholder.position }, style: { ...(placeholder.style || {}) }, paragraphStyles: cloneParagraphStyles(placeholder.paragraphStyles) }]));
  for (const placeholder of layoutPlaceholders) {
    const key = placeholderKey(placeholder);
    const inherited = merged.get(key) || {};
    merged.set(key, {
      ...inherited,
      ...placeholder,
      position: { ...(inherited.position || {}), ...(placeholder.position || {}) },
      style: { ...(inherited.style || {}), ...(placeholder.style || {}) },
      paragraphStyles: mergeParagraphStyles(inherited.paragraphStyles, placeholder.paragraphStyles),
    });
  }
  return [...merged.values()];
}

export function resolvePresentationBackgroundColor(background, theme = {}) {
  const fill = normalizePresentationBackground(background, theme.colors?.bg1 || "#ffffff").fill;
  const mapped = theme.colorMap?.[fill] || fill;
  const aliases = { dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" };
  return theme.colors?.[aliases[mapped] || mapped] || resolveColorToken(fill, fill);
}
