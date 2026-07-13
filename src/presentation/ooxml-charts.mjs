const BAR_GROUPINGS = new Set(["clustered", "stacked", "percentStacked"]);
const LINE_GROUPINGS = new Set(["standard", "stacked", "percentStacked"]);
const MARKER_SYMBOLS = new Set(["auto", "circle", "dash", "diamond", "dot", "none", "plus", "square", "star", "triangle", "x"]);

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

export function normalizePresentationChartStyle(chartType, config = {}) {
  const type = String(chartType || config.chartType || "bar").toLowerCase();
  const style = config.style && typeof config.style === "object" ? config.style : {};
  const rawBar = config.barOptions || style.bar || {};
  const rawLine = config.lineOptions || style.line || {};
  const directionValue = rawBar.direction || rawBar.barDirection;
  const direction = directionValue === "horizontal" ? "bar" : directionValue === "vertical" ? "column" : directionValue;
  return {
    styleId: boundedInteger(config.styleId ?? style.id, { name: "chart styleId", min: 1, max: 48, optional: true }),
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

export function normalizePresentationChartSeriesStyle(series = {}) {
  return {
    marker: normalizePresentationChartMarker(series.marker),
    smooth: series.smooth == null ? undefined : Boolean(series.smooth),
  };
}

function chartTextTitleXml(text = "") {
  return `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>${xmlEscape(text)}</a:t></a:r></a:p></c:rich></c:tx></c:title>`;
}

function markerXml(marker) {
  return marker ? `<c:marker><c:symbol val="${attrEscape(marker.symbol)}"/><c:size val="${marker.size}"/></c:marker>` : "";
}

export function presentationChartXml(chart) {
  const type = chart.chartType === "line" ? "line" : chart.chartType === "pie" ? "pie" : "bar";
  const style = normalizePresentationChartStyle(type, chart);
  const chartElementName = `${type}Chart`;
  const dataLabels = chart.dataLabels || {};
  const dataLabelsXml = dataLabels.showValue || dataLabels.showCategoryName
    ? `<c:dLbls><c:showLegendKey val="0"/><c:showVal val="${dataLabels.showValue ? 1 : 0}"/><c:showCatName val="${dataLabels.showCategoryName ? 1 : 0}"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>`
    : "";
  const seriesXml = (chart.series?.length ? chart.series : [{ name: chart.title || "Series", values: [] }]).map((series, index) => {
    const values = series.values || [];
    const categories = series.categories || chart.categories || values.map((_, pointIndex) => String(pointIndex + 1));
    const catPts = categories.map((category, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${xmlEscape(category)}</c:v></c:pt>`).join("");
    const valPts = values.map((value, pointIndex) => `<c:pt idx="${pointIndex}"><c:v>${Number(value) || 0}</c:v></c:pt>`).join("");
    const color = String(series.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4]).replace(/^#/, "").slice(0, 6).padEnd(6, "0");
    const seriesStyle = normalizePresentationChartSeriesStyle(series);
    const effectiveMarker = seriesStyle.marker || style.lineOptions.marker;
    const effectiveSmooth = seriesStyle.smooth ?? style.lineOptions.smooth;
    return `<c:ser><c:idx val="${index}"/><c:order val="${index}"/><c:tx><c:v>${xmlEscape(series.name || `Series ${index + 1}`)}</c:v></c:tx><c:spPr><a:solidFill><a:srgbClr val="${attrEscape(color)}"/></a:solidFill></c:spPr>${type === "line" ? markerXml(effectiveMarker) : ""}<c:cat><c:strLit><c:ptCount val="${categories.length}"/>${catPts}</c:strLit></c:cat><c:val><c:numLit><c:ptCount val="${values.length}"/>${valPts}</c:numLit></c:val>${type === "line" ? `<c:smooth val="${effectiveSmooth ? 1 : 0}"/>` : ""}</c:ser>`;
  }).join("");
  const categoryAxisTitle = chart.axes?.category?.title ? chartTextTitleXml(chart.axes.category.title) : "";
  const valueAxisTitle = chart.axes?.value?.title ? chartTextTitleXml(chart.axes.value.title) : "";
  const legendXml = chart.legend?.visible || chart.hasLegend ? `<c:legend><c:legendPos val="${attrEscape(chart.legend?.position || "r")}"/><c:layout/></c:legend>` : "";
  const varyColorsXml = `<c:varyColors val="${style.varyColors ? 1 : 0}"/>`;
  let plotXml;
  if (type === "pie") {
    plotXml = `<c:${chartElementName}>${varyColorsXml}${seriesXml}${dataLabelsXml}</c:${chartElementName}>`;
  } else {
    const optionsXml = type === "line"
      ? `<c:grouping val="${style.lineOptions.grouping}"/>${varyColorsXml}`
      : `<c:barDir val="${style.barOptions.direction === "bar" ? "bar" : "col"}"/><c:grouping val="${style.barOptions.grouping}"/>${varyColorsXml}`;
    const tailXml = type === "bar" ? `<c:gapWidth val="${style.barOptions.gapWidth}"/><c:overlap val="${style.barOptions.overlap}"/>` : "";
    plotXml = `<c:${chartElementName}>${optionsXml}${seriesXml}${dataLabelsXml}${tailXml}<c:axId val="1"/><c:axId val="2"/></c:${chartElementName}><c:catAx><c:axId val="1"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="b"/>${categoryAxisTitle}<c:crossAx val="2"/></c:catAx><c:valAx><c:axId val="2"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:axPos val="l"/>${valueAxisTitle}<c:crossAx val="1"/></c:valAx>`;
  }
  const styleXml = style.styleId ? `<c:style val="${style.styleId}"/>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">${styleXml}<c:chart>${chartTextTitleXml(chart.title || chart.chartType)}<c:plotArea><c:layout/>${plotXml}</c:plotArea>${legendXml}<c:plotVisOnly val="1"/></c:chart></c:chartSpace>`;
}

function parseChartTitle(xml, ownerName) {
  const owner = tagBlock(xml, ownerName);
  const title = tagBlock(owner, "title");
  return decodeXml(new RegExp(`<${localTag("t")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("t")}>`, "i").exec(title)?.[1] || "");
}

function parseSeries(chartBlock) {
  const pattern = new RegExp(`<${localTag("ser")}\\b[^>]*>[\\s\\S]*?<\\/${localTag("ser")}>`, "gi");
  return [...String(chartBlock || "").matchAll(pattern)].map((match, index) => {
    const xml = match[0];
    const tx = tagBlock(xml, "tx");
    const name = decodeXml(new RegExp(`<${localTag("v")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("v")}>`, "i").exec(tx)?.[1] || `Series ${index + 1}`);
    const color = new RegExp(`<${localTag("srgbClr")}\\b[^>]*\\bval="([A-Fa-f0-9]{6})"`, "i").exec(tagBlock(xml, "spPr"))?.[1];
    const valuesFrom = (name) => [...tagBlock(xml, name).matchAll(new RegExp(`<${localTag("v")}\\b[^>]*>([\\s\\S]*?)<\\/${localTag("v")}>`, "gi"))].map((item) => decodeXml(item[1]));
    const markerBlock = tagBlock(xml, "marker");
    const markerSymbol = tagValue(markerBlock, "symbol");
    const markerSize = tagValue(markerBlock, "size");
    const marker = markerSymbol ? normalizePresentationChartMarker({ symbol: markerSymbol, size: markerSize == null ? undefined : Number(markerSize) }) : undefined;
    return {
      name,
      values: valuesFrom("val").map((value) => Number(value) || 0),
      categories: valuesFrom("cat"),
      color: color ? `#${color}` : undefined,
      marker,
      smooth: booleanTag(xml, "smooth"),
    };
  });
}

export function parsePresentationChartXml(xml = "") {
  const text = String(xml || "");
  const chartType = new RegExp(`<${localTag("pieChart")}\\b`, "i").test(text) ? "pie" : new RegExp(`<${localTag("lineChart")}\\b`, "i").test(text) ? "line" : "bar";
  const chartBlock = tagBlock(text, `${chartType}Chart`);
  const series = parseSeries(chartBlock);
  const legendBlock = tagBlock(text, "legend");
  const hasLegend = new RegExp(`<${localTag("legend")}\\b`, "i").test(text);
  const labelsBlock = tagBlock(chartBlock, "dLbls");
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
    varyColors: Boolean(booleanTag(chartBlock, "varyColors")),
  };
  if (chartType === "bar") {
    parsed.barOptions = {
      direction: tagValue(chartBlock, "barDir") === "bar" ? "bar" : "column",
      grouping: tagValue(chartBlock, "grouping") || "clustered",
      gapWidth: Number(tagValue(chartBlock, "gapWidth") ?? 150),
      overlap: Number(tagValue(chartBlock, "overlap") ?? 0),
    };
  }
  if (chartType === "line") {
    parsed.lineOptions = {
      grouping: tagValue(chartBlock, "grouping") || "standard",
      marker: series.every((item) => JSON.stringify(item.marker) === JSON.stringify(series[0]?.marker)) ? series[0]?.marker : undefined,
      smooth: series.length > 0 && series.every((item) => item.smooth === true),
    };
  }
  return parsed;
}
