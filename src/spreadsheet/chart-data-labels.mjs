function dataLabelsError(message) {
  throw new TypeError(`Worksheet chart dataLabels ${message}`);
}

export const SPREADSHEET_CHART_DATA_LABEL_POSITIONS = Object.freeze([
  "bestFit", "bottom", "center", "insideBase", "insideEnd", "left", "outsideEnd", "right", "top",
]);

const POSITION_ALIASES = new Map([
  ...SPREADSHEET_CHART_DATA_LABEL_POSITIONS.map((value) => [value, value]),
  ["b", "bottom"], ["ctr", "center"], ["inBase", "insideBase"], ["inEnd", "insideEnd"],
  ["l", "left"], ["outEnd", "outsideEnd"], ["r", "right"], ["t", "top"],
]);

export function normalizeSpreadsheetChartDataLabels(value) {
  if (value == null) return undefined;
  if (typeof value === "boolean") return { showValue: value, showCategoryName: false };
  if (typeof value !== "object" || Array.isArray(value)) dataLabelsError("must be a boolean or object.");
  const supported = new Set(["showValue", "showCategoryName", "showSeriesName", "position"]);
  const unsupported = Object.keys(value).filter((key) => !supported.has(key) && value[key] != null);
  if (unsupported.length) dataLabelsError(`supports only showValue, showCategoryName, showSeriesName, and position; received ${unsupported.join(", ")}.`);
  const present = [...supported].filter((key) => value[key] != null);
  if (present.length === 0) dataLabelsError("must define showValue, showCategoryName, showSeriesName, or position.");
  for (const key of ["showValue", "showCategoryName", "showSeriesName"]) if (value[key] != null && typeof value[key] !== "boolean") dataLabelsError(`${key} must be a boolean.`);
  const position = value.position == null ? undefined : POSITION_ALIASES.get(value.position);
  if (value.position != null && position == null) dataLabelsError(`position must be one of: ${SPREADSHEET_CHART_DATA_LABEL_POSITIONS.join(", ")}.`);
  return {
    showValue: value.showValue === true,
    showCategoryName: value.showCategoryName === true,
    ...(value.showSeriesName == null ? {} : { showSeriesName: value.showSeriesName }),
    ...(position == null ? {} : { position }),
  };
}

export function spreadsheetChartDataLabelText(dataLabels, category, value, context = {}) {
  const normalized = normalizeSpreadsheetChartDataLabels(dataLabels);
  if (!normalized?.showValue && !normalized?.showCategoryName && !normalized?.showSeriesName) return "";
  return [
    normalized.showSeriesName ? context.seriesName : undefined,
    normalized.showCategoryName ? category : undefined,
    normalized.showValue ? value : undefined,
  ].filter((item) => item != null).map(String).join(": ");
}

export function spreadsheetChartDataLabelSvgPlacement(dataLabels, geometry = {}) {
  const position = normalizeSpreadsheetChartDataLabels(dataLabels)?.position || "outsideEnd";
  const x = Number(geometry.x || 0);
  const y = Number(geometry.y || 0);
  const width = Number(geometry.width || 0);
  const height = Number(geometry.height || 0);
  const top = Number(geometry.plotTop || 0);
  const baseY = Number(geometry.baseY ?? y + height);
  const point = geometry.kind === "point";
  if (position === "bottom") return { x: point ? x : x + width / 2, y: point ? y + 14 : baseY - 3, textAnchor: "middle", position };
  if (position === "center") return { x: point ? x : x + width / 2, y: point ? y + 3 : y + height / 2 + 3, textAnchor: "middle", position };
  if (position === "insideBase") return { x: point ? x : x + width / 2, y: point ? y + 11 : baseY - 4, textAnchor: "middle", position };
  if (position === "insideEnd") return { x: point ? x : x + width / 2, y: point ? y - 5 : y + 12, textAnchor: "middle", position };
  if (position === "left") return { x: point ? x - 7 : x - 4, y: point ? y + 3 : y + height / 2 + 3, textAnchor: "end", position };
  if (position === "right") return { x: point ? x + 7 : x + width + 4, y: point ? y + 3 : y + height / 2 + 3, textAnchor: "start", position };
  return { x: point ? x : x + width / 2, y: Math.max(top + 10, y - 4), textAnchor: "middle", position };
}
