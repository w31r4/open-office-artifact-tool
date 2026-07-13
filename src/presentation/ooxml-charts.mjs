import { resolveColorToken } from "../shared/colors.mjs";

const BAR_GROUPINGS = new Set(["clustered", "stacked", "percentStacked"]);
const LINE_GROUPINGS = new Set(["standard", "stacked", "percentStacked"]);
const MARKER_SYMBOLS = new Set(["auto", "circle", "dash", "diamond", "dot", "none", "plus", "square", "star", "triangle", "x"]);
const DATA_LABEL_POSITIONS = new Set(["bestFit", "b", "ctr", "inBase", "inEnd", "l", "outEnd", "r", "t"]);
const DATA_LABEL_POSITION_ALIASES = new Map([
  ["bottom", "b"], ["center", "ctr"], ["insideBase", "inBase"], ["insideEnd", "inEnd"],
  ["left", "l"], ["outsideEnd", "outEnd"], ["right", "r"], ["top", "t"],
]);
const TRENDLINE_TYPES = new Set(["exp", "linear", "log", "movingAvg", "poly", "power"]);
const TRENDLINE_TYPE_ALIASES = new Map([
  ["exponential", "exp"], ["logarithmic", "log"], ["movingAverage", "movingAvg"],
  ["polynomial", "poly"],
]);
const ERROR_BAR_DIRECTIONS = new Set(["x", "y"]);
const ERROR_BAR_TYPES = new Set(["both", "minus", "plus"]);
const ERROR_BAR_VALUE_TYPES = new Set(["cust", "fixedVal", "percentage", "stdDev", "stdErr"]);
const ERROR_BAR_VALUE_TYPE_ALIASES = new Map([["custom", "cust"], ["fixed", "fixedVal"], ["percent", "percentage"], ["standardDeviation", "stdDev"], ["standardError", "stdErr"]]);
const LINE_DASH_TO_OOXML = new Map([
  ["solid", "solid"], ["dot", "dot"], ["dash", "dash"], ["longDash", "lgDash"],
  ["dashDot", "dashDot"], ["longDashDot", "lgDashDot"], ["longDashDotDot", "lgDashDotDot"],
  ["systemDash", "sysDash"], ["systemDot", "sysDot"], ["systemDashDot", "sysDashDot"], ["systemDashDotDot", "sysDashDotDot"],
]);
const LINE_DASH_FROM_OOXML = new Map([...LINE_DASH_TO_OOXML].map(([publicName, ooxmlName]) => [ooxmlName, publicName]));
const SCHEME_COLORS = new Set(["tx1", "tx2", "bg1", "bg2", "dk1", "dk2", "lt1", "lt2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);

function xmlEscape(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

function attrEscape(value) {
  return xmlEscape(value).replaceAll('"', "&quot;");
}

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function localTag(name) {
  return `(?:[A-Za-z_][\\w.-]*:)?${name}`;
}

function tagValue(xml, name) {
  return new RegExp(`<${localTag(name)}\\b[^>]*\\bval="([^"]*)"`, "i").exec(String(xml || ""))?.[1];
}

function tagBlock(xml, name) {
  return new RegExp(`<${localTag(name)}\\b[^>]*>([\\s\\S]*?)<\\/${localTag(name)}>`, "i").exec(String(xml || ""))?.[1] || "";
}

function booleanTag(xml, name) {
  const value = tagValue(xml, name);
  if (value != null) return value === "1" || value === "true";
  return new RegExp(`<${localTag(name)}\\b`, "i").test(String(xml || "")) ? true : undefined;
}

function boundedInteger(value, { name, min, max, fallback, optional = false }) {
  if (value == null || value === "") return optional ? undefined : fallback;
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < min || parsed > max) throw new RangeError(`${name} must be an integer from ${min} to ${max}.`);
  return parsed;
}

function boundedNumber(value, { name, min, max, fallback, optional = false }) {
  if (value == null || value === "") return optional ? undefined : fallback;
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed < min || parsed > max) throw new RangeError(`${name} must be a number from ${min} to ${max}.`);
  return parsed;
}

function enumValue(value, allowed, fallback, name) {
  if (value == null || value === "") return fallback;
  if (!allowed.has(value)) throw new TypeError(`${name} must be one of: ${[...allowed].join(", ")}.`);
  return value;
}

export function normalizePresentationChartMarker(marker) {
  if (marker == null || marker === false) return undefined;
  const raw = typeof marker === "string" ? { symbol: marker } : marker;
  if (!raw || typeof raw !== "object") throw new TypeError("chart marker must be a symbol string or object.");
  return {
    symbol: enumValue(raw.symbol || raw.style, MARKER_SYMBOLS, "auto", "chart marker symbol"),
    size: boundedInteger(raw.size, { name: "chart marker size", min: 2, max: 72, fallback: 5 }),
  };
}

function normalizePresentationChartPaint(value) {
  if (typeof value === "string" && value) return value;
  if (!value || typeof value !== "object") return undefined;
  return [value.fill, value.color, value.rgb].find((candidate) => typeof candidate === "string" && candidate) || undefined;
}

function normalizePresentationChartLine(line) {
  if (line == null || line === false) return undefined;
  const raw = typeof line === "string" ? { fill: line } : line;
  if (!raw || typeof raw !== "object") throw new TypeError("chart line must be a color string or object.");
  const style = raw.style || raw.dash || "solid";
  if (!LINE_DASH_TO_OOXML.has(style)) throw new TypeError(`chart line style must be one of: ${[...LINE_DASH_TO_OOXML.keys()].join(", ")}.`);
  return {
    fill: normalizePresentationChartPaint(raw.fill ?? raw.color),
    width: boundedNumber(raw.width, { name: "chart line width", min: 0.1, max: 100, fallback: 1 }),
    style,
  };
}

function normalizePresentationChartPoints(points, valueCount) {
  const seen = new Set();
  return (points || []).map((point) => {
    if (!point || typeof point !== "object") throw new TypeError("chart points must be objects.");
    const rawIndex = point.idx ?? point.index;
    if (rawIndex == null) throw new TypeError("chart point idx is required.");
    const idx = boundedInteger(rawIndex, { name: "chart point idx", min: 0, max: 1_048_575 });
    if (valueCount != null && idx >= valueCount) throw new RangeError(`chart point idx ${idx} is outside the series value range.`);
    if (seen.has(idx)) throw new TypeError(`chart point idx ${idx} is duplicated.`);
    seen.add(idx);
    const fill = normalizePresentationChartPaint(point.fill ?? point.color);
    const line = normalizePresentationChartLine(point.line ?? point.stroke);
    return { idx, ...(fill ? { fill } : {}), ...(line ? { line } : {}) };
  });
}

export function normalizePresentationChartStyle(chartType, config = {}) {
  const type = String(chartType || config.chartType || "bar").toLowerCase();
  const style = config.style && typeof config.style === "object" ? config.style : {};
  const rawBar = config.barOptions || style.bar || {};
  const rawLine = config.lineOptions || style.line || {};
  const directionValue = rawBar.direction || rawBar.barDirection;
  const direction = directionValue === "horizontal" ? "bar" : directionValue === "vertical" ? "column" : directionValue;
  return {
    styleId: boundedInteger(config.styleId ?? config.styleIndex ?? style.id, { name: "chart styleId", min: 1, max: 48, optional: true }),
    varyColors: Boolean(config.varyColors ?? style.varyColors ?? type === "pie"),
    barOptions: {
      direction: enumValue(direction, new Set(["column", "bar"]), "column", "chart bar direction"),
      grouping: enumValue(rawBar.grouping, BAR_GROUPINGS, "clustered", "chart bar grouping"),
      gapWidth: boundedInteger(rawBar.gapWidth, { name: "chart gapWidth", min: 0, max: 500, fallback: 150 }),
      overlap: boundedInteger(rawBar.overlap, { name: "chart overlap", min: -100, max: 100, fallback: 0 }),
    },
    lineOptions: {
      grouping: enumValue(rawLine.grouping, LINE_GROUPINGS, "standard", "chart line grouping"),
      marker: normalizePresentationChartMarker(rawLine.marker),
      smooth: Boolean(rawLine.smooth),
    },
  };
}

export function normalizePresentationChartSeriesStyle(series = {}, valueCount) {
  return {
    color: normalizePresentationChartPaint(series.color ?? series.fill),
    line: normalizePresentationChartLine(series.line ?? series.stroke),
    points: normalizePresentationChartPoints(series.points, valueCount),
    marker: normalizePresentationChartMarker(series.marker),
    smooth: series.smooth == null ? undefined : Boolean(series.smooth),
  };
}

export function normalizePresentationChartDataLabels(value) {
  if (value === true) return { showValue: true, showCategoryName: false, position: "bestFit" };
  if (value === false || value == null) return { showValue: false, showCategoryName: false, position: "bestFit" };
  if (typeof value !== "object") throw new TypeError("chart dataLabels must be a boolean or object.");
  const rawPosition = value.position || "bestFit";
  const position = DATA_LABEL_POSITION_ALIASES.get(rawPosition) || rawPosition;
  if (!DATA_LABEL_POSITIONS.has(position)) throw new TypeError(`chart data-label position must be one of: ${[...DATA_LABEL_POSITIONS].join(", ")}.`);
  return {
    showValue: Boolean(value.showValue),
    showCategoryName: Boolean(value.showCategoryName ?? value.showCategory),
    position,
  };
}

function normalizePresentationChartTrendline(value, valueCount) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("chart trendlines must be objects.");
  const rawType = value.type || "linear";
  const type = TRENDLINE_TYPE_ALIASES.get(rawType) || rawType;
  if (!TRENDLINE_TYPES.has(type)) throw new TypeError(`chart trendline type must be one of: ${[...TRENDLINE_TYPES].join(", ")}.`);
  if (value.order != null && type !== "poly") throw new TypeError("chart trendline order is supported only for polynomial trendlines.");
  if (value.period != null && type !== "movingAvg") throw new TypeError("chart trendline period is supported only for moving-average trendlines.");
  const order = type === "poly"
    ? boundedInteger(value.order, { name: "polynomial chart trendline order", min: 2, max: 6, fallback: 2 })
    : undefined;
  const periodMax = valueCount == null ? 255 : Math.min(255, valueCount - 1);
  if (type === "movingAvg" && periodMax < 2) throw new RangeError("moving-average chart trendlines require at least three series values.");
  const period = type === "movingAvg"
    ? boundedInteger(value.period, { name: "moving-average chart trendline period", min: 2, max: periodMax, fallback: 2 })
    : undefined;
  const normalizeExtension = (candidate, name) => {
    const normalized = boundedNumber(candidate, { name, min: 0, max: 1_000_000, optional: true });
    if (normalized != null && normalized * 2 !== Math.round(normalized * 2)) throw new RangeError(`${name} must use 0.5 increments for category charts.`);
    return normalized;
  };
  const forward = normalizeExtension(value.forward, "chart trendline forward");
  const backward = normalizeExtension(value.backward, "chart trendline backward");
  const intercept = boundedNumber(value.intercept, { name: "chart trendline intercept", min: -Number.MAX_SAFE_INTEGER, max: Number.MAX_SAFE_INTEGER, optional: true });
  const name = value.name == null ? undefined : String(value.name);
  if (name != null && (name.length < 1 || name.length > 255)) throw new RangeError("chart trendline name must contain 1 to 255 characters.");
  const line = normalizePresentationChartLine(value.line ?? value.stroke);
  return {
    type,
    ...(name ? { name } : {}),
    ...(order == null ? {} : { order }),
    ...(period == null ? {} : { period }),
    ...(forward == null ? {} : { forward }),
    ...(backward == null ? {} : { backward }),
    ...(intercept == null ? {} : { intercept }),
    displayEquation: Boolean(value.displayEquation ?? value.showEquation),
    displayRSquared: Boolean(value.displayRSquared ?? value.showRSquared),
    ...(line ? { line } : {}),
  };
}

export function normalizePresentationChartTrendlines(value, valueCount, chartType) {
  if (value == null || value === false) return [];
  const items = Array.isArray(value) ? value : [value];
  if (items.length > 0 && chartType === "pie") throw new TypeError("chart trendlines are supported only for bar and line series.");
  if (items.length > 16) throw new RangeError("chart series support at most 16 trendlines.");
  return items.map((item) => normalizePresentationChartTrendline(item, valueCount));
}

function normalizePresentationErrorBarValues(values, valueCount, name) {
  if (!Array.isArray(values) || values.length === 0) throw new TypeError(`${name} must be a non-empty numeric array.`);
  if (valueCount != null && values.length !== valueCount) throw new RangeError(`${name} must contain exactly ${valueCount} values.`);
  return values.map((value) => boundedNumber(value, { name, min: 0, max: Number.MAX_SAFE_INTEGER }));
}

export function normalizePresentationChartErrorBars(value, chartType, valueCount) {
  if (value == null || value === false) return undefined;
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("chart errorBars must be an object.");
  if (chartType === "pie") throw new TypeError("chart errorBars are supported only for bar and line series.");
  const direction = value.direction || "y";
  if (!ERROR_BAR_DIRECTIONS.has(direction)) throw new TypeError(`chart error-bar direction must be one of: ${[...ERROR_BAR_DIRECTIONS].join(", ")}.`);
  const type = value.type || "both";
  if (!ERROR_BAR_TYPES.has(type)) throw new TypeError(`chart error-bar type must be one of: ${[...ERROR_BAR_TYPES].join(", ")}.`);
  const rawValueType = value.valueType || value.kind || "fixedVal";
  const valueType = ERROR_BAR_VALUE_TYPE_ALIASES.get(rawValueType) || rawValueType;
  if (!ERROR_BAR_VALUE_TYPES.has(valueType)) throw new TypeError(`chart error-bar valueType must be one of: ${[...ERROR_BAR_VALUE_TYPES].join(", ")}.`);
  const amount = ["cust", "stdErr"].includes(valueType) ? undefined : boundedNumber(value.value ?? value.amount, { name: "chart error-bar value", min: 0, max: Number.MAX_SAFE_INTEGER, fallback: valueType === "percentage" ? 5 : 1 });
  if (valueType === "stdErr" && (value.value ?? value.amount) != null) throw new TypeError("standard-error chart error bars do not accept a value.");
  const plusValues = valueType === "cust" && type !== "minus" ? normalizePresentationErrorBarValues(value.plusValues ?? value.plus, valueCount, "chart error-bar plusValues") : undefined;
  const minusValues = valueType === "cust" && type !== "plus" ? normalizePresentationErrorBarValues(value.minusValues ?? value.minus, valueCount, "chart error-bar minusValues") : undefined;
  const line = normalizePresentationChartLine(value.line ?? value.stroke);
  return {
    direction,
    type,
    valueType,
    ...(amount == null ? {} : { value: amount }),
    ...(plusValues ? { plusValues } : {}),
    ...(minusValues ? { minusValues } : {}),
    noEndCap: Boolean(value.noEndCap ?? value.endStyle === "none"),
    ...(line ? { line } : {}),
  };
}

function chartTextTitleXml(text = "") {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(text)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
}

function markerXml(marker) {
  return marker ? `<c:marker><c:symbol val="${attrEscape(marker.symbol)}"/><c:size val="${marker.size}"/></c:marker>` : "";
}

function chartColorXml(value, fallback = "#0ea5e9") {
  const raw = String(value || fallback);
  if (SCHEME_COLORS.has(raw)) return `<a:schemeClr val="${raw}"/>`;
  const resolved = resolveColorToken(raw, fallback);
  const hex = /^#[A-Fa-f0-9]{6}$/.test(resolved) ? resolved.slice(1) : /^[A-Fa-f0-9]{6}$/.test(resolved) ? resolved : String(fallback).replace(/^#/, "");
  return `<a:srgbClr val="${attrEscape(hex.toUpperCase())}"/>`;
}

function chartShapePropertiesXml(fill, line) {
  const fillXml = fill ? `<a:solidFill>${chartColorXml(fill)}</a:solidFill>` : "";
  const lineXml = line ? `<a:ln w="${Math.round(line.width * 12_700)}"><a:solidFill>${chartColorXml(line.fill || fill || "#0f172a")}</a:solidFill><a:prstDash val="${LINE_DASH_TO_OOXML.get(line.style)}"/></a:ln>` : "";
  return fillXml || lineXml ? `<c:spPr>${fillXml}${lineXml}</c:spPr>` : "";
}

function chartPointXml(point) {
  return `<c:dPt><c:idx val="${point.idx}"/>${chartShapePropertiesXml(point.fill, point.line)}</c:dPt>`;
}

function presentationDataLabelsXml(dataLabels, force = false) {
  const normalized = normalizePresentationChartDataLabels(dataLabels);
  if (!force && !normalized.showValue && !normalized.showCategoryName) return "";
  const positionXml = normalized.position === "bestFit" ? "" : `<c:dLblPos val="${normalized.position}"/>`;
  return `<c:dLbls>${positionXml}<c:showLegendKey val="0"/><c:showVal val="${normalized.showValue ? 1 : 0}"/><c:showCatName val="${normalized.showCategoryName ? 1 : 0}"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`;
}

function presentationTrendlineXml(trendline) {
  const nameXml = trendline.name ? `<c:name>${xmlEscape(trendline.name)}</c:name>` : "";
  const lineXml = trendline.line ? chartShapePropertiesXml(undefined, trendline.line) : "";
  const orderXml = trendline.order == null ? "" : `<c:order val="${trendline.order}"/>`;
  const periodXml = trendline.period == null ? "" : `<c:period val="${trendline.period}"/>`;
  const forwardXml = trendline.forward == null ? "" : `<c:forward val="${trendline.forward}"/>`;
  const backwardXml = trendline.backward == null ? "" : `<c:backward val="${trendline.backward}"/>`;
  const interceptXml = trendline.intercept == null ? "" : `<c:intercept val="${trendline.intercept}"/>`;
  return `<c:trendline>${nameXml}${lineXml}<c:trendlineType val="${trendline.type}"/>${orderXml}${periodXml}${forwardXml}${backwardXml}${interceptXml}<c:dispRSqr val="${trendline.displayRSquared ? 1 : 0}"/><c:dispEq val="${trendline.displayEquation ? 1 : 0}"/></c:trendline>`;
}

function presentationErrorBarsXml(errorBars) {
  if (!errorBars) return "";
  const valuesXml = (name, values) => values?.length
    ? `<c:${name}><c:numLit><c:formatCode>General</c:formatCode><c:ptCount val="${values.length}"/>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:numLit></c:${name}>`
    : "";
  const plusXml = valuesXml("plus", errorBars.plusValues);
  const minusXml = valuesXml("minus", errorBars.minusValues);
  const valueXml = errorBars.value == null ? "" : `<c:val val="${errorBars.value}"/>`;
  const lineXml = errorBars.line ? chartShapePropertiesXml(undefined, errorBars.line) : "";
  return `<c:errBars><c:errDir val="${errorBars.direction}"/><c:errBarType val="${errorBars.type}"/><c:errValType val="${errorBars.valueType}"/><c:noEndCap val="${errorBars.noEndCap ? 1 : 0}"/>${plusXml}${minusXml}${valueXml}${lineXml}</c:errBars>`;
}

export function presentationChartXml(chart) {
  const type = chart.chartType === "combo" ? "combo" : chart.chartType === "line" ? "line" : chart.chartType === "pie" ? "pie" : "bar";
  const style = normalizePresentationChartStyle(type, chart);
  const dataLabelsXml = presentationDataLabelsXml(chart.dataLabels || {});
  const chartSeries = chart.series?.length ? chart.series : [{ name: chart.title || "Series", values: [] }];
  const seriesXml = (series, index, seriesType) => {
    const values = series.values || [];
    const categories = series.categories || chart.categories || values.map((_, pointIndex) => String(pointIndex + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    const seriesStyle = normalizePresentationChartSeriesStyle(series, values.length);
    const color = seriesStyle.color || series.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4];
    const effectiveMarker = seriesStyle.marker || style.lineOptions.marker;
    const effectiveSmooth = seriesStyle.smooth ?? style.lineOptions.smooth;
    const seriesLine = seriesStyle.line || (seriesType === "line" ? { fill: color, width: 2, style: "solid" } : undefined);
    const pointsXml = seriesStyle.points.map(chartPointXml).join("");
    const seriesDataLabelsXml = series.dataLabels == null ? "" : presentationDataLabelsXml(series.dataLabels, true);
    const trendlinesXml = normalizePresentationChartTrendlines(series.trendlines ?? series.trendline, values.length, seriesType).map(presentationTrendlineXml).join("");
    const errorBarsXml = presentationErrorBarsXml(normalizePresentationChartErrorBars(series.errorBars, seriesType, values.length));
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx>${chartShapePropertiesXml(color, seriesLine)}${seriesType === "line" ? markerXml(effectiveMarker) : ""}${pointsXml}${seriesDataLabelsXml}${trendlinesXml}${errorBarsXml}<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val>${seriesType === "line" ? `<c:smooth val="${effectiveSmooth ? 1 : 0}"/>` : ""}</c:ser>`;
  };
  const categoryAxisTitle = chart.axes?.category?.title ? chartTextTitleXml(chart.axes.category.title) : "";
  const valueAxisTitle = chart.axes?.value?.title ? chartTextTitleXml(chart.axes.value.title) : "";
  const legendXml = chart.legend?.visible || chart.hasLegend ? `<c:legend><c:legendPos val="${attrEscape(chart.legend?.position || "r")}"/><c:layout/></c:legend>` : "";
  const varyColorsXml = `<c:varyColors val="${style.varyColors ? 1 : 0}"/>`;
  const axisXml = `<c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>${categoryAxisTitle}<c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>${valueAxisTitle}<c:crossAx val="1"/></c:valAx>`;
  const plotForType = (plotType, entries) => {
    const content = entries.map(({ series, index }) => seriesXml(series, index, plotType)).join("");
    if (plotType === "line") return `<c:lineChart><c:grouping val="${style.lineOptions.grouping}"/>${varyColorsXml}${content}${dataLabelsXml}<c:axId val="1"/><c:axId val="2"/></c:lineChart>`;
    return `<c:barChart><c:barDir val="${style.barOptions.direction === "bar" ? "bar" : "col"}"/><c:grouping val="${style.barOptions.grouping}"/>${varyColorsXml}${content}${dataLabelsXml}<c:gapWidth val="${style.barOptions.gapWidth}"/><c:overlap val="${style.barOptions.overlap}"/><c:axId val="1"/><c:axId val="2"/></c:barChart>`;
  };
  let plotXml;
  if (type === "pie") {
    plotXml = `<c:pieChart>${varyColorsXml}${chartSeries.map((series, index) => seriesXml(series, index, "pie")).join("")}${dataLabelsXml}</c:pieChart>`;
  } else if (type === "combo") {
    const entries = chartSeries.map((series, index) => ({ series, index }));
    plotXml = `${plotForType("bar", entries.filter(({ series }) => series.chartType === "bar"))}${plotForType("line", entries.filter(({ series }) => series.chartType === "line"))}${axisXml}`;
  } else {
    plotXml = `${plotForType(type, chartSeries.map((series, index) => ({ series, index })))}${axisXml}`;
  }
  const styleXml = style.styleId ? `<c:style val="${style.styleId}"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${styleXml}<c:chart>${chartTextTitleXml(chart.title || chart.chartType)}<c:plotArea><c:layout/>${plotXml}</c:plotArea>${legendXml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function parseChartTitle(xml, ownerName) {
  const owner = tagBlock(xml, ownerName);
  const title = tagBlock(owner, "title");
  return decodeXml(new RegExp(`<${localTag("t")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("t")}>`, "i").exec(title)?.[1] || "");
}

function parseChartColor(xml) {
  const srgb = new RegExp(`<${localTag("srgbClr")}\\b[^>]*\\bval="([A-Fa-f0-9]{6})"`, "i").exec(String(xml || ""))?.[1];
  if (srgb) return `#${srgb}`;
  const scheme = new RegExp(`<${localTag("schemeClr")}\\b[^>]*\\bval="([^"]+)"`, "i").exec(String(xml || ""))?.[1];
  return scheme || undefined;
}

function parseChartShapeProperties(xml) {
  const spPr = tagBlock(xml, "spPr");
  if (!spPr) return {};
  const linePattern = new RegExp(`<${localTag("ln")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("ln")}>`, "i");
  const lineMatch = linePattern.exec(spPr);
  const lineXml = lineMatch?.[0] || "";
  const fill = parseChartColor(lineMatch ? spPr.replace(lineMatch[0], "") : spPr);
  if (!lineXml) return { fill };
  const width = Number(new RegExp(`<${localTag("ln")}\\b[^>]*\\bw="(\\d+)"`, "i").exec(lineXml)?.[1]);
  const dash = tagValue(lineXml, "prstDash") || "solid";
  return {
    fill,
    line: {
      fill: parseChartColor(lineXml),
      width: Number.isFinite(width) && width > 0 ? width / 12_700 : 1,
      style: LINE_DASH_FROM_OOXML.get(dash) || "solid",
    },
  };
}

function parseChartPoints(xml) {
  const pattern = new RegExp(`<${localTag("dPt")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("dPt")}>`, "gi");
  return [...String(xml || "").matchAll(pattern)].map((match) => {
    const style = parseChartShapeProperties(match[0]);
    return { idx: Number(tagValue(match[0], "idx")), ...(style.fill ? { fill: style.fill } : {}), ...(style.line ? { line: style.line } : {}) };
  });
}

function parseDataLabels(xml) {
  const block = tagBlock(xml, "dLbls");
  if (!block) return undefined;
  return normalizePresentationChartDataLabels({
    showValue: Boolean(booleanTag(block, "showVal")),
    showCategoryName: Boolean(booleanTag(block, "showCatName")),
    position: tagValue(block, "dLblPos") || "bestFit",
  });
}

function parseTrendlines(xml, valueCount) {
  const pattern = new RegExp(`<${localTag("trendline")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("trendline")}>`, "gi");
  return [...String(xml || "").matchAll(pattern)].map((match) => {
    const block = match[0];
    const shape = parseChartShapeProperties(block);
    return normalizePresentationChartTrendline({
      type: tagValue(block, "trendlineType") || "linear",
      name: decodeXml(tagBlock(block, "name")) || undefined,
      order: tagValue(block, "order"),
      period: tagValue(block, "period"),
      forward: tagValue(block, "forward"),
      backward: tagValue(block, "backward"),
      intercept: tagValue(block, "intercept"),
      displayRSquared: Boolean(booleanTag(block, "dispRSqr")),
      displayEquation: Boolean(booleanTag(block, "dispEq")),
      line: shape.line,
    }, valueCount);
  });
}

function parseErrorBars(xml, chartType, valueCount) {
  const block = tagBlock(xml, "errBars");
  if (!block) return undefined;
  const shape = parseChartShapeProperties(block);
  const valuesFrom = (name) => [...tagBlock(block, name).matchAll(new RegExp(`<${localTag("v")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("v")}>`, "gi"))].map((item) => Number(decodeXml(item[1])) || 0);
  return normalizePresentationChartErrorBars({
    direction: tagValue(block, "errDir") || "y",
    type: tagValue(block, "errBarType") || "both",
    valueType: tagValue(block, "errValType") || "fixedVal",
    value: tagValue(block, "val"),
    plusValues: valuesFrom("plus"),
    minusValues: valuesFrom("minus"),
    noEndCap: Boolean(booleanTag(block, "noEndCap")),
    line: shape.line,
  }, chartType, valueCount);
}

function parseSeries(chartBlock, chartType) {
  const pattern = new RegExp(`<${localTag("ser")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("ser")}>`, "gi");
  return [...String(chartBlock || "").matchAll(pattern)].map((match, index) => {
    const xml = match[0];
    const xmlWithoutDecorations = xml
      .replace(new RegExp(`<${localTag("dPt")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("dPt")}>`, "gi"), "")
      .replace(new RegExp(`<${localTag("trendline")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("trendline")}>`, "gi"), "")
      .replace(new RegExp(`<${localTag("errBars")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("errBars")}>`, "gi"), "");
    const tx = tagBlock(xml, "tx");
    const name = decodeXml(new RegExp(`<${localTag("v")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("v")}>`, "i").exec(tx)?.[1] || `Series ${index + 1}`);
    const shapeStyle = parseChartShapeProperties(xmlWithoutDecorations);
    const valuesFrom = (name) => [...tagBlock(xmlWithoutDecorations, name).matchAll(new RegExp(`<${localTag("v")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("v")}>`, "gi"))].map((item) => decodeXml(item[1]));
    const markerBlock = tagBlock(xml, "marker");
    const markerSymbol = tagValue(markerBlock, "symbol");
    const markerSize = tagValue(markerBlock, "size");
    const marker = markerSymbol ? normalizePresentationChartMarker({ symbol: markerSymbol, size: markerSize == null ? undefined : Number(markerSize) }) : undefined;
    const dataLabels = parseDataLabels(xml);
    const values = valuesFrom("val").map((value) => Number(value) || 0);
    const trendlines = parseTrendlines(xml, values.length);
    const errorBars = parseErrorBars(xml, chartType, values.length);
    return {
      ...(chartType ? { chartType } : {}),
      order: Number(tagValue(xml, "order") ?? index),
      name,
      values,
      categories: valuesFrom("cat"),
      color: shapeStyle.fill || shapeStyle.line?.fill,
      line: shapeStyle.line,
      points: parseChartPoints(xml),
      marker,
      smooth: booleanTag(xml, "smooth"),
      ...(dataLabels ? { dataLabels } : {}),
      ...(trendlines.length ? { trendlines } : {}),
      ...(errorBars ? { errorBars } : {}),
    };
  });
}

export function parsePresentationChartXml(xml = "") {
  const text = String(xml || "");
  const hasPie = new RegExp(`<${localTag("pieChart")}\\b`, "i").test(text);
  const hasLine = new RegExp(`<${localTag("lineChart")}\\b`, "i").test(text);
  const hasBar = new RegExp(`<${localTag("barChart")}\\b`, "i").test(text);
  const chartType = hasBar && hasLine ? "combo" : hasPie ? "pie" : hasLine ? "line" : "bar";
  const chartBlock = chartType === "combo" ? "" : tagBlock(text, `${chartType}Chart`);
  const barBlock = hasBar ? tagBlock(text, "barChart") : "";
  const lineBlock = hasLine ? tagBlock(text, "lineChart") : "";
  const series = (chartType === "combo"
    ? [...parseSeries(barBlock, "bar"), ...parseSeries(lineBlock, "line")].sort((left, right) => left.order - right.order)
    : parseSeries(chartBlock)).map(({ order: _order, ...item }) => item);
  const legendBlock = tagBlock(text, "legend");
  const hasLegend = new RegExp(`<${localTag("legend")}\\b`, "i").test(text);
  const withoutSeries = (block) => String(block || "").replace(new RegExp(`<${localTag("ser")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("ser")}>`, "gi"), "");
  const labelsBlock = tagBlock(withoutSeries(chartBlock || barBlock), "dLbls") || tagBlock(withoutSeries(lineBlock), "dLbls");
  const styleId = tagValue(text.slice(0, text.search(new RegExp(`<${localTag("chart")}\\b`, "i"))), "style");
  const parsed = {
    chartType,
    title: parseChartTitle(text, "chart"),
    categories: series[0]?.categories || [],
    series,
    axes: {
      category: { title: parseChartTitle(text, "catAx") },
      value: { title: parseChartTitle(text, "valAx") },
    },
    legend: { visible: hasLegend, position: tagValue(legendBlock, "legendPos") || "r" },
    dataLabels: {
      showValue: Boolean(booleanTag(labelsBlock, "showVal")),
      showCategoryName: Boolean(booleanTag(labelsBlock, "showCatName")),
      position: tagValue(labelsBlock, "dLblPos") || "bestFit",
    },
    styleId: styleId == null ? undefined : Number(styleId),
    varyColors: Boolean(booleanTag(chartBlock || barBlock || lineBlock, "varyColors")),
  };
  if (chartType === "bar" || chartType === "combo") {
    const optionsBlock = chartType === "combo" ? barBlock : chartBlock;
    parsed.barOptions = {
      direction: tagValue(optionsBlock, "barDir") === "bar" ? "bar" : "column",
      grouping: tagValue(optionsBlock, "grouping") || "clustered",
      gapWidth: Number(tagValue(optionsBlock, "gapWidth") ?? 150),
      overlap: Number(tagValue(optionsBlock, "overlap") ?? 0),
    };
  }
  if (chartType === "line" || chartType === "combo") {
    const optionsBlock = chartType === "combo" ? lineBlock : chartBlock;
    const lineSeries = chartType === "combo" ? series.filter((item) => item.chartType === "line") : series;
    parsed.lineOptions = {
      grouping: tagValue(optionsBlock, "grouping") || "standard",
      marker: lineSeries.every((item) => JSON.stringify(item.marker) === JSON.stringify(lineSeries[0]?.marker)) ? lineSeries[0]?.marker : undefined,
      smooth: lineSeries.length > 0 && lineSeries.every((item) => item.smooth === true),
    };
  }
  return parsed;
}
