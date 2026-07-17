import { normalizeSpreadsheetChartDataLabels, spreadsheetChartDataLabelSvgPlacement, spreadsheetChartDataLabelText } from "./chart-data-labels.mjs";
import { normalizeSpreadsheetChartLineOptions, spreadsheetChartSmoothLinePath } from "./chart-line-options.mjs";
import { normalizeSpreadsheetChartSeriesLine, spreadsheetChartLineDashArray } from "./chart-line-style.mjs";
import { normalizeSpreadsheetChartSeriesMarker, spreadsheetChartMarkerSvg } from "./chart-marker-style.mjs";
import { resolvedWorksheetChartCategories, resolvedWorksheetChartSeriesValues } from "./chart-source-data.mjs";
import { xmlEscape } from "../shared/xml.mjs";

const PREVIEW_PALETTE = ["#38BDF8", "#F97316", "#22C55E", "#A855F7", "#E11D48", "#0F766E"];
const CIRCULAR_TYPES = new Set(["pie", "doughnut"]);

function seriesColor(series, index) {
  return /^#[0-9a-f]{6}$/i.test(series?.fill || "")
    ? series.fill.toUpperCase()
    : PREVIEW_PALETTE[index % PREVIEW_PALETTE.length];
}

function lineAttributes(series, index, fallbackWidth = 2) {
  const fill = seriesColor(series, index);
  const line = normalizeSpreadsheetChartSeriesLine(series);
  const stroke = line?.fill || fill;
  const width = line?.width ?? fallbackWidth;
  const dash = spreadsheetChartLineDashArray(line?.style);
  return { fill, stroke, width, attributes: ` stroke="${stroke}" stroke-width="${width}"${dash ? ` stroke-dasharray="${dash}"` : ""}` };
}

function cartesianGeometry(chart, seriesItems, plot) {
  const lineOptions = normalizeSpreadsheetChartLineOptions(chart.lineOptions);
  const grouping = lineOptions?.grouping || "standard";
  const pointCount = Math.max(0, ...seriesItems.map((series) => series.values?.length || 0));
  const totals = Array.from({ length: pointCount }, (_, pointIndex) => seriesItems.reduce((total, series) => total + (Number(series.values?.[pointIndex]) || 0), 0));
  const values = seriesItems.map((series, seriesIndex) => (series.values || []).map((value, pointIndex) => {
    const raw = Number(value) || 0;
    if (chart.type !== "line" || grouping === "standard") return raw;
    const stacked = seriesItems.slice(0, seriesIndex + 1).reduce((total, item) => total + (Number(item.values?.[pointIndex]) || 0), 0);
    return grouping === "percentStacked" ? (totals[pointIndex] === 0 ? 0 : stacked / totals[pointIndex]) : stacked;
  }));
  const allValues = values.flat();
  const minimum = Math.min(0, ...allValues);
  const maximum = Math.max(1, ...allValues);
  const span = Math.max(1, maximum - minimum);
  const y = (value) => plot.top + plot.height - (Number(value) - minimum) / span * plot.height;
  return { lineOptions, values, minimum, maximum, y, baseline: y(0) };
}

function barMarks(chart, categories, seriesItems, dataLabels, plot, geometry) {
  const values = seriesItems[0]?.values || [];
  const barWidth = values.length ? plot.width / values.length * 0.65 : 0;
  const gap = values.length ? plot.width / values.length * 0.35 : 0;
  const style = lineAttributes(seriesItems[0], 0, 0);
  return values.map((value, index) => {
    const x = plot.left + index * (barWidth + gap) + gap / 2;
    const valueY = geometry.y(Number(value) || 0);
    const y = Math.min(valueY, geometry.baseline);
    const height = Math.abs(geometry.baseline - valueY);
    const label = spreadsheetChartDataLabelText(dataLabels, categories[index], value, { seriesName: seriesItems[0]?.name });
    const placement = spreadsheetChartDataLabelSvgPlacement(dataLabels, { x, y, width: barWidth, height, baseY: geometry.baseline, plotTop: plot.top });
    return `<rect x="${x}" y="${y}" width="${barWidth}" height="${height}" fill="${style.fill}"${normalizeSpreadsheetChartSeriesLine(seriesItems[0]) == null ? "" : style.attributes}/>${label ? `<text x="${placement.x}" y="${placement.y}" text-anchor="${placement.textAnchor}" font-family="Arial" font-size="10" fill="#334155" data-chart-label-position="${placement.position}" data-chart-label-index="${index}">${xmlEscape(label)}</text>` : ""}`;
  }).join("");
}

function lineMarks(chart, categories, seriesItems, dataLabels, plot, geometry) {
  return geometry.values.map((seriesValues, seriesIndex) => {
    const series = seriesItems[seriesIndex];
    const style = lineAttributes(series, seriesIndex);
    const points = seriesValues.map((value, index) => ({ x: plot.left + (index + 0.5) * plot.width / Math.max(1, seriesValues.length), y: geometry.y(value) }));
    const mark = geometry.lineOptions?.smooth === true
      ? `<path d="${spreadsheetChartSmoothLinePath(points)}" fill="none"${style.attributes} data-series-index="${seriesIndex}"/>`
      : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none"${style.attributes} data-series-index="${seriesIndex}"/>`;
    const labels = points.map((point, index) => {
      const label = spreadsheetChartDataLabelText(dataLabels, categories[index], series?.values?.[index], { seriesName: series?.name });
      const placement = spreadsheetChartDataLabelSvgPlacement(dataLabels, { x: point.x, y: point.y, kind: "point", plotTop: plot.top });
      return label ? `<text x="${placement.x}" y="${placement.y}" text-anchor="${placement.textAnchor}" font-family="Arial" font-size="10" fill="#334155" data-chart-label-position="${placement.position}" data-chart-label-series="${seriesIndex}" data-chart-label-index="${index}">${xmlEscape(label)}</text>` : "";
    }).join("");
    return `${mark}${points.map((point) => spreadsheetChartMarkerSvg(series?.marker, point.x, point.y, style.stroke)).join("")}${labels}`;
  }).join("");
}

function areaMarks(seriesItems, plot, geometry) {
  return geometry.values.map((seriesValues, seriesIndex) => {
    if (seriesValues.length === 0) return "";
    const style = lineAttributes(seriesItems[seriesIndex], seriesIndex);
    const points = seriesValues.map((value, index) => ({ x: plot.left + (index + 0.5) * plot.width / Math.max(1, seriesValues.length), y: geometry.y(value) }));
    const path = [`M ${points[0].x} ${geometry.baseline}`, ...points.map((point) => `L ${point.x} ${point.y}`), `L ${points.at(-1).x} ${geometry.baseline}`, "Z"].join(" ");
    return `<path d="${path}" fill="${style.fill}" fill-opacity="0.32"${style.attributes} data-series-index="${seriesIndex}"/>`;
  }).join("");
}

function polarPoint(cx, cy, radius, angle) {
  return { x: cx + Math.cos(angle) * radius, y: cy + Math.sin(angle) * radius };
}

function ringSegmentPath(cx, cy, outerRadius, innerRadius, start, end) {
  const outerStart = polarPoint(cx, cy, outerRadius, start);
  const outerEnd = polarPoint(cx, cy, outerRadius, end);
  const largeArc = end - start > Math.PI ? 1 : 0;
  if (innerRadius <= 0) return `M ${cx} ${cy} L ${outerStart.x} ${outerStart.y} A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${outerEnd.x} ${outerEnd.y} Z`;
  const innerEnd = polarPoint(cx, cy, innerRadius, end);
  const innerStart = polarPoint(cx, cy, innerRadius, start);
  return `M ${outerStart.x} ${outerStart.y} A ${outerRadius} ${outerRadius} 0 ${largeArc} 1 ${outerEnd.x} ${outerEnd.y} L ${innerEnd.x} ${innerEnd.y} A ${innerRadius} ${innerRadius} 0 ${largeArc} 0 ${innerStart.x} ${innerStart.y} Z`;
}

function circularMarks(chart, categories, seriesItems, dataLabels, plot) {
  const series = seriesItems[0];
  const values = (series?.values || []).map((value) => Math.abs(Number(value) || 0));
  const total = values.reduce((sum, value) => sum + value, 0);
  if (total <= 0) return "";
  const cx = plot.left + plot.width / 2;
  const cy = plot.top + plot.height / 2;
  const outerRadius = Math.max(1, Math.min(plot.width, plot.height) * 0.42);
  const innerRadius = chart.type === "doughnut" ? outerRadius * 0.5 : 0;
  let angle = -Math.PI / 2;
  return values.map((value, index) => {
    const sweep = value / total * Math.PI * 2;
    const end = angle + sweep;
    const fill = PREVIEW_PALETTE[index % PREVIEW_PALETTE.length];
    const fullCircle = sweep >= Math.PI * 2 - 1e-9;
    const mark = fullCircle
      ? innerRadius > 0
        ? `<circle cx="${cx}" cy="${cy}" r="${(outerRadius + innerRadius) / 2}" fill="none" stroke="${fill}" stroke-width="${outerRadius - innerRadius}" data-point-index="${index}"/>`
        : `<circle cx="${cx}" cy="${cy}" r="${outerRadius}" fill="${fill}" data-point-index="${index}"/>`
      : `<path d="${ringSegmentPath(cx, cy, outerRadius, innerRadius, angle, end)}" fill="${fill}" stroke="#ffffff" stroke-width="1" data-point-index="${index}"/>`;
    const label = spreadsheetChartDataLabelText(dataLabels, categories[index], series?.values?.[index], { seriesName: series?.name });
    const labelRadius = innerRadius > 0 ? (innerRadius + outerRadius) / 2 : outerRadius * 0.62;
    const labelPoint = polarPoint(cx, cy, labelRadius, angle + sweep / 2);
    angle = end;
    return `${mark}${label ? `<text x="${labelPoint.x}" y="${labelPoint.y}" text-anchor="middle" dominant-baseline="middle" font-family="Arial" font-size="10" fill="#0f172a" data-chart-label-index="${index}">${xmlEscape(label)}</text>` : ""}`;
  }).join("");
}

export function renderWorksheetChartSvg(chart) {
  const frame = chart.position;
  const categories = resolvedWorksheetChartCategories(chart);
  const seriesItems = chart.series.items.map((series) => ({ ...series, values: resolvedWorksheetChartSeriesValues(chart, series) }));
  const dataLabels = normalizeSpreadsheetChartDataLabels(chart.dataLabels);
  const plot = { left: frame.left + 28, top: frame.top + 36, width: Math.max(0, frame.width - 44), height: Math.max(0, frame.height - 62) };
  const circular = CIRCULAR_TYPES.has(chart.type);
  const geometry = circular ? null : cartesianGeometry(chart, seriesItems, plot);
  const marks = circular
    ? circularMarks(chart, categories, seriesItems, dataLabels, plot)
    : chart.type === "line"
      ? lineMarks(chart, categories, seriesItems, dataLabels, plot, geometry)
      : chart.type === "area"
        ? areaMarks(seriesItems, plot, geometry)
        : barMarks(chart, categories, seriesItems, dataLabels, plot, geometry);
  const pointCount = seriesItems[0]?.values?.length || 0;
  const xTickSize = Number(chart.xAxis?.textStyle?.fontSize);
  const xTicks = !circular && Number.isFinite(xTickSize) && xTickSize > 0 && pointCount ? categories.map((category, index) => `<text x="${plot.left + (index + 0.5) * plot.width / pointCount}" y="${plot.top + plot.height + xTickSize + 2}" text-anchor="middle" font-family="Arial" font-size="${xTickSize}" fill="#64748b">${xmlEscape(category)}</text>`).join("") : "";
  const yTickSize = Number(chart.yAxis?.textStyle?.fontSize);
  const yTicks = !circular && Number.isFinite(yTickSize) && yTickSize > 0 ? `<text x="${plot.left - 4}" y="${plot.top + yTickSize}" text-anchor="end" font-family="Arial" font-size="${yTickSize}" fill="#64748b">${geometry.maximum}</text><text x="${plot.left - 4}" y="${plot.top + plot.height}" text-anchor="end" font-family="Arial" font-size="${yTickSize}" fill="#64748b">${geometry.minimum}</text>` : "";
  const xTitle = !circular && chart.xAxis?.title?.text ? `<text x="${plot.left + plot.width / 2}" y="${frame.top + frame.height - 6}" text-anchor="middle" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(chart.xAxis.title.text)}</text>` : "";
  const yTitle = !circular && chart.yAxis?.title?.text ? `<text x="${frame.left + 10}" y="${plot.top + plot.height / 2}" text-anchor="middle" transform="rotate(-90 ${frame.left + 10} ${plot.top + plot.height / 2})" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(chart.yAxis.title.text)}</text>` : "";
  const titleSize = Number.isFinite(Number(chart.titleTextStyle?.fontSize)) && Number(chart.titleTextStyle.fontSize) > 0 ? Number(chart.titleTextStyle.fontSize) : 13;
  return `<rect x="${frame.left}" y="${frame.top}" width="${frame.width}" height="${frame.height}" fill="#ffffff" stroke="#94a3b8"/><text x="${frame.left + 8}" y="${frame.top + 22}" font-family="Arial" font-size="${titleSize}" font-weight="700" fill="#0f172a">${xmlEscape(chart.title || chart.name)}</text>${marks}${xTicks}${yTicks}${xTitle}${yTitle}`;
}
