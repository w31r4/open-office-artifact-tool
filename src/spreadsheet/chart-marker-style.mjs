import { normalizeSpreadsheetChartLineStyle, spreadsheetChartLineDashArray, spreadsheetChartLineStyleXml } from "./chart-line-style.mjs";

export const SPREADSHEET_CHART_MARKER_SYMBOLS = Object.freeze([
  "none",
  "dot",
  "circle",
  "square",
  "diamond",
  "triangle",
  "x",
  "star",
  "plus",
  "dash",
]);

export const SPREADSHEET_CHART_MARKER_MIN_SIZE = 2;
export const SPREADSHEET_CHART_MARKER_MAX_SIZE = 72;

function markerError(message) {
  throw new TypeError(`Worksheet chart series.marker ${message}`);
}

export function normalizeSpreadsheetChartSeriesMarker(value) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) markerError("must be an object.");
  const unsupported = Object.keys(value).filter((key) => !["symbol", "size", "fill", "line"].includes(key) && value[key] != null);
  if (unsupported.length) markerError(`supports only symbol, size, fill, and line; received ${unsupported.join(", ")}.`);
  const output = {};
  if (value.symbol != null) {
    const symbol = String(value.symbol).toLowerCase();
    if (!SPREADSHEET_CHART_MARKER_SYMBOLS.includes(symbol)) markerError(`symbol must be one of ${SPREADSHEET_CHART_MARKER_SYMBOLS.join(", ")}.`);
    output.symbol = symbol;
  }
  if (value.size != null) {
    const size = Number(value.size);
    if (!Number.isInteger(size) || size < SPREADSHEET_CHART_MARKER_MIN_SIZE || size > SPREADSHEET_CHART_MARKER_MAX_SIZE) markerError(`size must be an integer from ${SPREADSHEET_CHART_MARKER_MIN_SIZE} through ${SPREADSHEET_CHART_MARKER_MAX_SIZE}.`);
    output.size = size;
  }
  if (value.fill != null) {
    if (typeof value.fill !== "string" || !/^#[0-9a-f]{6}$/i.test(value.fill)) markerError("fill must be a #RRGGBB color.");
    output.fill = value.fill.toUpperCase();
  }
  const line = normalizeSpreadsheetChartLineStyle(value.line, "series.marker.line");
  if (line != null) output.line = line;
  return output;
}

export function spreadsheetChartSeriesMarkerXml(value) {
  const marker = normalizeSpreadsheetChartSeriesMarker(value);
  if (marker == null) return "";
  const symbol = marker.symbol == null ? "" : `<c:symbol val="${marker.symbol}"/>`;
  const size = marker.size == null ? "" : `<c:size val="${marker.size}"/>`;
  const fill = marker.fill == null ? "" : `<a:solidFill><a:srgbClr val="${marker.fill.slice(1)}"/></a:solidFill>`;
  const line = spreadsheetChartLineStyleXml(marker.line, "series.marker.line");
  const shapeProperties = fill || line ? `<c:spPr>${fill}${line}</c:spPr>` : "";
  return `<c:marker>${symbol}${size}${shapeProperties}</c:marker>`;
}

export function spreadsheetChartMarkerSvg(value, x, y, color) {
  const marker = normalizeSpreadsheetChartSeriesMarker(value);
  if (marker == null || marker.symbol === "none") return "";
  const symbol = marker.symbol || "circle";
  const size = marker.size || 5;
  const half = size / 2;
  const fill = marker.fill || color;
  const line = marker.line;
  const strokeColor = line?.fill || color;
  const strokeWidth = line?.width ?? 1.5;
  const dash = spreadsheetChartLineDashArray(line?.style);
  const stroke = `stroke="${strokeColor}" stroke-width="${strokeWidth}"${dash ? ` stroke-dasharray="${dash}"` : ""}`;
  if (symbol === "circle" || symbol === "dot") {
    const radius = symbol === "dot" ? Math.max(1, half / 2) : half;
    return `<circle cx="${x}" cy="${y}" r="${radius}" fill="${fill}" ${stroke}/>`;
  }
  if (symbol === "square") return `<rect x="${x - half}" y="${y - half}" width="${size}" height="${size}" fill="${fill}" ${stroke}/>`;
  if (symbol === "diamond") return `<polygon points="${x},${y - half} ${x + half},${y} ${x},${y + half} ${x - half},${y}" fill="${fill}" ${stroke}/>`;
  if (symbol === "triangle") return `<polygon points="${x},${y - half} ${x + half},${y + half} ${x - half},${y + half}" fill="${fill}" ${stroke}/>`;
  if (symbol === "star") {
    const points = Array.from({ length: 10 }, (_, index) => {
      const angle = -Math.PI / 2 + index * Math.PI / 5;
      const radius = index % 2 === 0 ? half : half * 0.45;
      return `${x + Math.cos(angle) * radius},${y + Math.sin(angle) * radius}`;
    }).join(" ");
    return `<polygon points="${points}" fill="${fill}" ${stroke}/>`;
  }
  if (symbol === "dash") return `<line x1="${x - half}" y1="${y}" x2="${x + half}" y2="${y}" ${stroke}/>`;
  const vertical = symbol === "plus" ? `<line x1="${x}" y1="${y - half}" x2="${x}" y2="${y + half}" ${stroke}/>` : `<line x1="${x - half}" y1="${y - half}" x2="${x + half}" y2="${y + half}" ${stroke}/>`;
  const horizontal = symbol === "plus" ? `<line x1="${x - half}" y1="${y}" x2="${x + half}" y2="${y}" ${stroke}/>` : `<line x1="${x + half}" y1="${y - half}" x2="${x - half}" y2="${y + half}" ${stroke}/>`;
  return `${vertical}${horizontal}`;
}
