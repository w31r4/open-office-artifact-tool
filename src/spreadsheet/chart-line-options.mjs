function lineOptionsError(message) {
  throw new TypeError(`Worksheet chart lineOptions ${message}`);
}

export function normalizeSpreadsheetChartLineOptions(value) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) lineOptionsError("must be an object.");
  const unsupported = Object.keys(value).filter((key) => key !== "smooth" && value[key] != null);
  if (unsupported.length) lineOptionsError(`supports only smooth; received ${unsupported.join(", ")}.`);
  if (value.smooth == null) return null;
  if (typeof value.smooth !== "boolean") lineOptionsError("smooth must be a boolean.");
  return { smooth: value.smooth };
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
