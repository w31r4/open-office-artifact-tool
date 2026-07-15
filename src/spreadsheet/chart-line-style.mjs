const MAX_WIDTH_POINTS = 1_584;

const STYLE_TO_DRAWINGML = new Map([
  ["solid", "solid"],
  ["dashed", "dash"],
  ["dotted", "dot"],
  ["dash-dot", "dashDot"],
  ["dash-dot-dot", "lgDashDotDot"],
]);

export const SPREADSHEET_CHART_LINE_STYLES = Object.freeze([...STYLE_TO_DRAWINGML.keys()]);
export const SPREADSHEET_CHART_LINE_MAX_WIDTH_POINTS = MAX_WIDTH_POINTS;

function lineError(name, message) {
  throw new TypeError(`Worksheet chart ${name} ${message}`);
}

function normalizedLine(value, name, aliases) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) lineError(name, "must be an object.");
  const allowed = new Set(Object.values(aliases));
  const unsupported = Object.keys(value).filter((key) => !allowed.has(key) && value[key] != null);
  if (unsupported.length) lineError(name, `supports only ${[...allowed].join(", ")}; received ${unsupported.join(", ")}.`);

  const output = {};
  const fill = value[aliases.fill];
  if (fill != null) {
    if (typeof fill !== "string" || !/^#[0-9a-f]{6}$/i.test(fill)) lineError(name, `${aliases.fill} must be a #RRGGBB color.`);
    output.fill = fill.toUpperCase();
  }
  const style = value[aliases.style];
  if (style != null) {
    const normalized = String(style).toLowerCase();
    if (!STYLE_TO_DRAWINGML.has(normalized)) lineError(name, `${aliases.style} must be one of ${SPREADSHEET_CHART_LINE_STYLES.join(", ")}.`);
    output.style = normalized;
  }
  const width = value[aliases.width];
  if (width != null) {
    const normalized = Number(width);
    if (!Number.isFinite(normalized) || normalized < 0 || normalized > MAX_WIDTH_POINTS) lineError(name, `${aliases.width} must be from 0 through ${MAX_WIDTH_POINTS} points.`);
    output.width = normalized;
  }
  return output;
}

export function normalizeSpreadsheetChartLineStyle(value, name = "series.line") {
  return normalizedLine(value, name, { fill: "fill", style: "style", width: "width" });
}

export function normalizeSpreadsheetChartSeriesLine(series) {
  const line = normalizeSpreadsheetChartLineStyle(series?.line, "series.line");
  const stroke = normalizedLine(series?.stroke, "series.stroke", { fill: "color", style: "style", width: "weight" });
  if (line != null && stroke != null && JSON.stringify(line) !== JSON.stringify(stroke)) lineError("series", "line and stroke aliases must describe the same style when both are present.");
  return line ?? stroke;
}

export function spreadsheetChartLineStyleXml(value, name = "series.line") {
  const line = normalizeSpreadsheetChartLineStyle(value, name);
  if (line == null) return "";
  const width = line.width == null ? "" : ` w="${Math.round(line.width * 12_700)}"`;
  const fill = line.fill == null ? "" : `<a:solidFill><a:srgbClr val="${line.fill.slice(1)}"/></a:solidFill>`;
  const dash = line.style == null ? "" : `<a:prstDash val="${STYLE_TO_DRAWINGML.get(line.style)}"/>`;
  return `<a:ln${width}>${fill}${dash}</a:ln>`;
}

export function spreadsheetChartSeriesLineXml(series) {
  const line = normalizeSpreadsheetChartSeriesLine(series);
  return line == null ? "" : spreadsheetChartLineStyleXml(line);
}

export function spreadsheetChartLineDashArray(style) {
  return {
    dashed: "8 5",
    dotted: "2 4",
    "dash-dot": "8 4 2 4",
    "dash-dot-dot": "8 4 2 4 2 4",
  }[style] || "";
}
