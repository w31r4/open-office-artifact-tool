const DEFAULT_FONT_WEIGHT = 400;
const DEFAULT_FONT_STYLE = "normal";
const DEFAULT_FONT_WIDTH = "normal";

let baseMetrics = [];
const scopedMetrics = [];

function normalizedMetric(entry) {
  if (!entry || typeof entry !== "object") return undefined;
  const family = String(entry.family ?? "").trim();
  const weight = Number(entry.weight);
  const unitsPerEm = Number(entry.unitsPerEm);
  const ascent = Number(entry.ascent);
  const descent = Number(entry.descent);
  const lineGap = entry.lineGap == null ? 0 : Number(entry.lineGap);
  if (!family || family.length > 255) return undefined;
  if (!Number.isFinite(weight) || weight <= 0 || weight > 1_000) return undefined;
  if (!Number.isFinite(unitsPerEm) || unitsPerEm <= 0) return undefined;
  if (!Number.isFinite(ascent) || ascent <= 0) return undefined;
  if (!Number.isFinite(descent) || descent < 0) return undefined;
  if (!Number.isFinite(lineGap)) return undefined;
  const style = String(entry.style || DEFAULT_FONT_STYLE).trim().toLowerCase() || DEFAULT_FONT_STYLE;
  const width = String(entry.width || DEFAULT_FONT_WIDTH).trim().toLowerCase() || DEFAULT_FONT_WIDTH;
  return {
    family,
    weight,
    unitsPerEm,
    ascent,
    descent,
    lineGap,
    style,
    width,
    familyKey: family.toLocaleLowerCase(),
  };
}

function normalizedMetrics(entries) {
  const output = [];
  for (const entry of entries) {
    const metric = normalizedMetric(entry);
    if (metric) output.push(metric);
  }
  return output;
}

export function setOfficeFontDesignMetrics(entries) {
  baseMetrics = normalizedMetrics(entries);
}

export function registerScopedOfficeFontDesignMetrics(entries) {
  const scope = { metrics: normalizedMetrics(entries) };
  scopedMetrics.push(scope);
  let disposed = false;
  return () => {
    if (disposed) return;
    disposed = true;
    const index = scopedMetrics.indexOf(scope);
    if (index >= 0) scopedMetrics.splice(index, 1);
  };
}

export function clearOfficeFontDesignMetrics() {
  baseMetrics = [];
  scopedMetrics.length = 0;
}

function candidatesForFamily(familyKey) {
  const candidates = [];
  for (let index = scopedMetrics.length - 1; index >= 0; index -= 1) {
    candidates.push(...scopedMetrics[index].metrics.filter((metric) => metric.familyKey === familyKey));
  }
  candidates.push(...baseMetrics.filter((metric) => metric.familyKey === familyKey));
  return candidates;
}

export function resolveOfficeFontDesignMetrics(request) {
  const families = request.family;
  if (!Array.isArray(families) || families.length === 0) return undefined;
  const family = String(families[0] ?? "").trim();
  if (!family) return undefined;
  const requestedStyle = String(request.style || DEFAULT_FONT_STYLE).trim().toLowerCase() || DEFAULT_FONT_STYLE;
  const requestedWeight = Number.isFinite(Number(request.weight)) ? Number(request.weight) : DEFAULT_FONT_WEIGHT;
  const familyCandidates = candidatesForFamily(family.toLocaleLowerCase());
  if (familyCandidates.length === 0) return undefined;
  const styleCandidates = familyCandidates.filter((metric) => metric.style === requestedStyle);
  const candidates = styleCandidates.length ? styleCandidates : familyCandidates.filter((metric) => metric.style === DEFAULT_FONT_STYLE);
  const ranked = (candidates.length ? candidates : familyCandidates)
    .map((metric, order) => ({ metric, order }))
    .sort((left, right) => {
      const distance = Math.abs(left.metric.weight - requestedWeight) - Math.abs(right.metric.weight - requestedWeight);
      if (distance) return distance;
      if (left.metric.weight !== right.metric.weight) return left.metric.weight - right.metric.weight;
      return left.order - right.order;
    });
  return ranked.length ? { ...ranked[0].metric } : undefined;
}

export function skiaPaintBaselineCompensationPx(value) {
  const pixels = Number(value);
  return Number.isFinite(pixels) ? pixels - Math.round(pixels) : 0;
}

const FONT_VALUE_KEYS = new Set([
  "fontfamily",
  "fontfamilyhighansi",
  "fontfamilyeastasia",
  "fontfamilycomplexscript",
  "resolvedfontfamily",
  "resolvedfontfamilyeastasia",
  "resolvedfontfamilycomplexscript",
  "effectivefontfamily",
  "bulletfont",
  "bulletfontfamily",
  "typeface",
]);

function rememberFamily(families, value) {
  if (typeof value !== "string") return;
  const family = value.trim();
  if (!family || family.startsWith("+") || family === "none") return;
  const key = family.toLocaleLowerCase();
  if (!families.has(key)) families.set(key, family);
}

function visitFontValues(value, families, seen, parentKey = "") {
  if (!value || typeof value !== "object") return;
  if (seen.has(value)) return;
  seen.add(value);
  if (Array.isArray(value)) {
    for (const item of value) visitFontValues(item, families, seen, parentKey);
    return;
  }
  for (const [key, child] of Object.entries(value)) {
    const normalizedKey = key.toLocaleLowerCase();
    if (FONT_VALUE_KEYS.has(normalizedKey)) rememberFamily(families, child);
    else if (normalizedKey === "name" && parentKey === "font") rememberFamily(families, child);
    visitFontValues(child, families, seen, normalizedKey);
  }
}

export function officeFontFamilies(values = [], defaults = []) {
  const families = new Map();
  for (const family of defaults) rememberFamily(families, family);
  visitFontValues(values, families, new WeakSet());
  return [...families.values()].sort((left, right) => left.toLocaleLowerCase().localeCompare(right.toLocaleLowerCase()));
}
