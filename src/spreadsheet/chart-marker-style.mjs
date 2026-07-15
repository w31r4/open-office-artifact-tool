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
  const unsupported = Object.keys(value).filter((key) => !["symbol", "size"].includes(key) && value[key] != null);
  if (unsupported.length) markerError(`supports only symbol and size; received ${unsupported.join(", ")}.`);
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
  return output;
}

export function spreadsheetChartSeriesMarkerXml(value) {
  const marker = normalizeSpreadsheetChartSeriesMarker(value);
  if (marker == null) return "";
  const symbol = marker.symbol == null ? "" : `<c:symbol val="${marker.symbol}"/>`;
  const size = marker.size == null ? "" : `<c:size val="${marker.size}"/>`;
  return `<c:marker>${symbol}${size}</c:marker>`;
}

export function spreadsheetChartMarkerSvg(value, x, y, color) {
  const marker = normalizeSpreadsheetChartSeriesMarker(value);
  if (marker == null || marker.symbol === "none") return "";
  const symbol = marker.symbol || "circle";
  const size = marker.size || 5;
  const half = size / 2;
  const stroke = `stroke="${color}" stroke-width="1.5"`;
  if (symbol === "circle" || symbol === "dot") {
    const radius = symbol === "dot" ? Math.max(1, half / 2) : half;
    return `<circle cx="${x}" cy="${y}" r="${radius}" fill="${color}" ${stroke}/>`;
  }
  if (symbol === "square") return `<rect x="${x - half}" y="${y - half}" width="${size}" height="${size}" fill="${color}" ${stroke}/>`;
  if (symbol === "diamond") return `<polygon points="${x},${y - half} ${x + half},${y} ${x},${y + half} ${x - half},${y}" fill="${color}" ${stroke}/>`;
  if (symbol === "triangle") return `<polygon points="${x},${y - half} ${x + half},${y + half} ${x - half},${y + half}" fill="${color}" ${stroke}/>`;
  if (symbol === "star") {
    const points = Array.from({ length: 10 }, (_, index) => {
      const angle = -Math.PI / 2 + index * Math.PI / 5;
      const radius = index % 2 === 0 ? half : half * 0.45;
      return `${x + Math.cos(angle) * radius},${y + Math.sin(angle) * radius}`;
    }).join(" ");
    return `<polygon points="${points}" fill="${color}" ${stroke}/>`;
  }
  if (symbol === "dash") return `<line x1="${x - half}" y1="${y}" x2="${x + half}" y2="${y}" ${stroke}/>`;
  const vertical = symbol === "plus" ? `<line x1="${x}" y1="${y - half}" x2="${x}" y2="${y + half}" ${stroke}/>` : `<line x1="${x - half}" y1="${y - half}" x2="${x + half}" y2="${y + half}" ${stroke}/>`;
  const horizontal = symbol === "plus" ? `<line x1="${x - half}" y1="${y}" x2="${x + half}" y2="${y}" ${stroke}/>` : `<line x1="${x + half}" y1="${y - half}" x2="${x - half}" y2="${y + half}" ${stroke}/>`;
  return `${vertical}${horizontal}`;
}
