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


function placeholderKey(placeholder) {
  return `${placeholder.type || "body"}:${Number(placeholder.idx ?? 0)}`;
}

export function mergePresentationPlaceholders(masterPlaceholders = [], layoutPlaceholders = []) {
  const cloneFrame = (frame) => frame ? { ...frame } : undefined;
  const cloneParagraphStyles = (styles = {}) => Object.fromEntries(Object.entries(styles).map(([level, style]) => [level, { ...style, style: { ...(style.style || {}) } }]));
  const mergeParagraphStyles = (base = {}, overrides = {}) => {
    const result = cloneParagraphStyles(base);
    for (const [level, style] of Object.entries(overrides)) {
      const inherited = { ...(result[level] || {}) };
      if (["bulletCharacter", "bulletImage", "autoNumber", "bulletNone"].some((field) => Object.hasOwn(style, field))) {
        delete inherited.bulletCharacter;
        delete inherited.bulletImage;
        delete inherited.autoNumber;
        delete inherited.bulletNone;
      }
      for (const fields of [["bulletFont", "bulletFontFollowText"], ["bulletColor", "bulletColorFollowText"], ["bulletSize", "bulletSizePercent", "bulletSizeFollowText"]]) {
        if (!fields.some((field) => Object.hasOwn(style, field))) continue;
        for (const field of fields) delete inherited[field];
      }
      result[level] = { ...inherited, ...style, style: { ...(inherited.style || {}), ...(style.style || {}) } };
    }
    return result;
  };
  const merged = new Map(masterPlaceholders.map((placeholder) => {
    const position = cloneFrame(placeholder.position);
    const transform = placeholder.transform ? { ...placeholder.transform } : undefined;
    const effective = {
      ...placeholder,
      ...(position ? { position, geometrySource: "master" } : {}),
      ...(transform ? { transform } : {}),
      style: { ...(placeholder.style || {}) },
      paragraphStyles: cloneParagraphStyles(placeholder.paragraphStyles),
    };
    if (!position) delete effective.position;
    if (!transform) delete effective.transform;
    return [placeholderKey(placeholder), effective];
  }));
  for (const placeholder of layoutPlaceholders) {
    const key = placeholderKey(placeholder);
    const inherited = merged.get(key) || {};
    const hasDirectFrame = Boolean(placeholder.position);
    const position = cloneFrame(hasDirectFrame ? placeholder.position : inherited.position);
    // a:xfrm is one atomic frame slot. When a layout owns a direct frame, an
    // omitted rot/flip attribute does not inherit from the master's a:xfrm.
    const transform = hasDirectFrame
      ? (placeholder.transform ? { ...placeholder.transform } : undefined)
      : (inherited.transform ? { ...inherited.transform } : undefined);
    const effective = {
      ...inherited,
      ...placeholder,
      ...(position ? { position, geometrySource: hasDirectFrame ? "layout" : inherited.geometrySource } : {}),
      ...(transform ? { transform } : {}),
      style: { ...(inherited.style || {}), ...(placeholder.style || {}) },
      paragraphStyles: mergeParagraphStyles(inherited.paragraphStyles, placeholder.paragraphStyles),
    };
    if (!position) {
      delete effective.position;
      delete effective.geometrySource;
    }
    if (!transform) delete effective.transform;
    merged.set(key, effective);
  }
  return [...merged.values()];
}

export function resolvePresentationBackgroundColor(background, theme = {}) {
  const fill = normalizePresentationBackground(background, theme.colors?.bg1 || "#ffffff").fill;
  const mapped = theme.colorMap?.[fill] || fill;
  const aliases = { dk1: "tx1", lt1: "bg1", dk2: "tx2", lt2: "bg2" };
  return theme.colors?.[aliases[mapped] || mapped] || resolveColorToken(fill, fill);
}
