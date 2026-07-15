function lineOptionsError(message) {
  throw new TypeError(`Worksheet chart lineOptions ${message}`);
}

export const SPREADSHEET_CHART_LINE_GROUPINGS = Object.freeze(["standard", "stacked", "percentStacked"]);
const GROUPINGS = new Set(SPREADSHEET_CHART_LINE_GROUPINGS);

export function normalizeSpreadsheetChartLineOptions(value) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) lineOptionsError("must be an object.");
  const unsupported = Object.keys(value).filter((key) => !["grouping", "smooth", "varyColors"].includes(key) && value[key] != null);
  if (unsupported.length) lineOptionsError(`supports only grouping, smooth, and varyColors; received ${unsupported.join(", ")}.`);
  const output = {};
  if (value.grouping != null) {
    if (typeof value.grouping !== "string" || !GROUPINGS.has(value.grouping)) lineOptionsError(`grouping must be one of ${SPREADSHEET_CHART_LINE_GROUPINGS.join(", ")}.`);
    output.grouping = value.grouping;
  }
  if (value.smooth != null) {
    if (typeof value.smooth !== "boolean") lineOptionsError("smooth must be a boolean.");
    output.smooth = value.smooth;
  }
  if (value.varyColors != null) {
    if (typeof value.varyColors !== "boolean") lineOptionsError("varyColors must be a boolean.");
    if (value.varyColors) output.varyColors = true;
  }
  return Object.keys(output).length ? output : null;
}

export function spreadsheetChartSmoothLinePath(points = []) {
  if (points.length === 0) return "";
  if (points.length === 1) return `M ${points[0].x} ${points[0].y}`;
  let path = `M ${points[0].x} ${points[0].y}`;
  for (let index = 1; index < points.length; index += 1) {
    const previous = points[index - 1];
    const current = points[index];
    const midpoint = (previous.x + current.x) / 2;
    path += ` C ${midpoint} ${previous.y} ${midpoint} ${current.y} ${current.x} ${current.y}`;
  }
  return path;
}
