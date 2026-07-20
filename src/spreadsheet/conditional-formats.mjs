const BASE_ICON_SET_COUNTS = Object.freeze({
  "3Arrows": 3,
  "3ArrowsGray": 3,
  "3Flags": 3,
  "3TrafficLights1": 3,
  "3TrafficLights2": 3,
  "3Signs": 3,
  "3Symbols": 3,
  "3Symbols2": 3,
  "4Arrows": 4,
  "4ArrowsGray": 4,
  "4RedToBlack": 4,
  "4Rating": 4,
  "4TrafficLights": 4,
  "5Arrows": 5,
  "5ArrowsGray": 5,
  "5Rating": 5,
  "5Quarters": 5,
});

const X14_ICON_SETS = new Set(["3Triangles", "3Stars", "5Boxes"]);
const THRESHOLD_TYPES = new Set(["min", "max", "num", "percent", "percentile"]);

const ICON_GLYPHS = Object.freeze({
  "3Arrows": ["▼", "➜", "▲"],
  "3ArrowsGray": ["▼", "➜", "▲"],
  "3Flags": ["⚑", "⚑", "⚑"],
  "3TrafficLights1": ["●", "●", "●"],
  "3TrafficLights2": ["●", "●", "●"],
  "3Signs": ["◆", "▲", "●"],
  "3Symbols": ["✕", "!", "✓"],
  "3Symbols2": ["✕", "!", "✓"],
  "4Arrows": ["▼", "↘", "↗", "▲"],
  "4ArrowsGray": ["▼", "↘", "↗", "▲"],
  "4RedToBlack": ["●", "●", "●", "●"],
  "4Rating": ["◔", "◑", "◕", "●"],
  "4TrafficLights": ["●", "●", "●", "●"],
  "5Arrows": ["▼", "↘", "➜", "↗", "▲"],
  "5ArrowsGray": ["▼", "↘", "➜", "↗", "▲"],
  "5Rating": ["○", "◔", "◑", "◕", "●"],
  "5Quarters": ["○", "◔", "◑", "◕", "●"],
});

const ICON_COLORS = Object.freeze({
  gray: ["#6B7280", "#6B7280", "#6B7280", "#6B7280", "#6B7280"],
  status: ["#DC2626", "#F97316", "#EAB308", "#84CC16", "#16A34A"],
  dark: ["#DC2626", "#6B7280", "#374151", "#111827"],
});

function profileError(location, message) {
  throw new TypeError(`${location} ${message}`);
}

function finiteNumber(value, location) {
  const number = typeof value === "number" ? value : Number(value);
  if (!Number.isFinite(number)) profileError(location, "requires a finite numeric value.");
  return number;
}

function optionalBoolean(source, key, location) {
  if (source[key] == null) return undefined;
  if (typeof source[key] !== "boolean") profileError(location, `${key} must be a boolean.`);
  return source[key];
}

export function normalizeConditionalFormatThreshold(input, location = "Conditional-format threshold") {
  let type;
  let value;
  if (typeof input === "number") {
    type = "num";
    value = finiteNumber(input, location);
  } else if (typeof input === "string") {
    const text = input.trim();
    if (text === "min" || text === "max") type = text;
    else if (/^[+-]?(?:\d+(?:\.\d+)?|\.\d+)%$/.test(text)) {
      type = "percent";
      value = finiteNumber(text.slice(0, -1), location);
    } else profileError(location, `does not support ${JSON.stringify(input)}; use min, max, a number, a percentage, or a typed threshold.`);
  } else if (input && typeof input === "object" && !Array.isArray(input)) {
    type = String(input.type || "");
    value = input.value;
  } else profileError(location, "must be min, max, a number, a percentage, or a typed threshold.");

  if (!THRESHOLD_TYPES.has(type)) profileError(location, `has unsupported type ${type || "(empty)"}.`);
  if (type === "min" || type === "max") {
    if (value != null && value !== "") profileError(location, `${type} must not carry a value.`);
    return { type };
  }
  const number = finiteNumber(value, location);
  if ((type === "percent" || type === "percentile") && (number < 0 || number > 100)) profileError(location, `${type} must be between 0 and 100.`);
  return { type, value: number };
}

function normalizeThresholds(inputs, defaults, count, location) {
  const source = inputs == null ? defaults : inputs;
  if (!Array.isArray(source) || source.length !== count) profileError(location, `requires exactly ${count} thresholds.`);
  const thresholds = source.map((input, index) => normalizeConditionalFormatThreshold(input, `${location} threshold ${index + 1}`));
  const comparable = thresholds.every((item) => item.type === thresholds[0].type && item.value != null);
  if (comparable && thresholds.some((item, index) => index > 0 && item.value < thresholds[index - 1].value)) profileError(location, "threshold values must be nondecreasing.");
  return thresholds;
}

export function spreadsheetIconSetCount(name) {
  return BASE_ICON_SET_COUNTS[String(name || "")];
}

export function normalizeDataBarConfig(config = {}, location = "dataBar conditional format") {
  const source = config?.dataBar || config || {};
  const showValue = optionalBoolean(source, "showValue", location);
  const gradient = optionalBoolean(source, "gradient", location);
  if (gradient === false) profileError(location, "gradient=false requires the x14 solid-data-bar extension and is outside the editable profile.");
  const thresholds = normalizeThresholds(source.thresholds, ["min", "max"], 2, location);
  const color = source.color ?? "#638EC6";
  return {
    color,
    thresholds,
    ...(showValue == null ? {} : { showValue }),
    ...(gradient == null ? {} : { gradient }),
  };
}

export function normalizeIconSetConfig(config = {}, location = "iconSet conditional format") {
  const source = config?.iconSet && typeof config.iconSet === "object" ? config.iconSet : config || {};
  const name = String(typeof source.iconSet === "string" ? source.iconSet : source.name || "");
  const count = spreadsheetIconSetCount(name);
  if (!count) {
    if (X14_ICON_SETS.has(name)) profileError(location, `${name} requires the x14 extension namespace and is outside the editable profile.`);
    profileError(location, `uses unsupported icon set ${name || "(empty)"}.`);
  }
  const defaultValues = count === 3 ? [0, 33, 67] : count === 4 ? [0, 25, 50, 75] : [0, 20, 40, 60, 80];
  const defaults = defaultValues.map((value) => ({ type: "percent", value }));
  const showValue = optionalBoolean(source, "showValue", location);
  const reverse = optionalBoolean(source, "reverse", location);
  return {
    iconSet: name,
    thresholds: normalizeThresholds(source.thresholds, defaults, count, location),
    ...(showValue == null ? {} : { showValue }),
    ...(reverse == null ? {} : { reverse }),
  };
}

function percentile(values, amount) {
  const sorted = [...values].sort((a, b) => a - b);
  if (!sorted.length) return undefined;
  const position = Math.max(0, Math.min(1, amount / 100)) * (sorted.length - 1);
  const lower = Math.floor(position);
  const upper = Math.ceil(position);
  return sorted[lower] + (sorted[upper] - sorted[lower]) * (position - lower);
}

function thresholdValue(threshold, values) {
  const min = Math.min(...values);
  const max = Math.max(...values);
  if (threshold.type === "min") return min;
  if (threshold.type === "max") return max;
  if (threshold.type === "num") return threshold.value;
  if (threshold.type === "percent") return min + (max - min) * threshold.value / 100;
  if (threshold.type === "percentile") return percentile(values, threshold.value);
  return undefined;
}

function numericInputs(value, values) {
  const number = typeof value === "number" ? value : Number.NaN;
  const finite = (values || []).filter((item) => typeof item === "number" && Number.isFinite(item));
  return Number.isFinite(number) && finite.length ? { number, values: finite } : undefined;
}

export function dataBarVisual(config, value, values) {
  const inputs = numericInputs(value, values);
  if (!inputs) return undefined;
  const normalized = normalizeDataBarConfig(config);
  const lower = thresholdValue(normalized.thresholds[0], inputs.values);
  const upper = thresholdValue(normalized.thresholds[1], inputs.values);
  const ratio = upper === lower ? (inputs.number >= upper ? 1 : 0) : Math.max(0, Math.min(1, (inputs.number - lower) / (upper - lower)));
  return { kind: "dataBar", ratio, color: normalized.color, showValue: normalized.showValue !== false, gradient: true, thresholds: normalized.thresholds };
}

function iconPalette(name, count) {
  if (name.endsWith("Gray")) return ICON_COLORS.gray.slice(0, count);
  if (name === "4RedToBlack") return ICON_COLORS.dark;
  const colors = ICON_COLORS.status;
  if (count === 3) return [colors[0], colors[2], colors[4]];
  if (count === 4) return [colors[0], colors[1], colors[2], colors[4]];
  return colors.slice(0, count);
}

export function iconSetVisual(config, value, values) {
  const inputs = numericInputs(value, values);
  if (!inputs) return undefined;
  const normalized = normalizeIconSetConfig(config);
  const count = spreadsheetIconSetCount(normalized.iconSet);
  const resolved = normalized.thresholds.map((threshold) => thresholdValue(threshold, inputs.values));
  let index = 0;
  for (let candidate = 1; candidate < resolved.length; candidate += 1) if (inputs.number >= resolved[candidate]) index = candidate;
  if (normalized.reverse === true) index = count - 1 - index;
  const glyphs = ICON_GLYPHS[normalized.iconSet] || Array.from({ length: count }, () => "●");
  const palette = iconPalette(normalized.iconSet, count);
  return { kind: "iconSet", iconSet: normalized.iconSet, index, glyph: glyphs[index], color: palette[index], showValue: normalized.showValue !== false, reverse: normalized.reverse === true, thresholds: normalized.thresholds };
}
