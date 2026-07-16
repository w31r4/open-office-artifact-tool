import { SpreadsheetChartDataLabelPosition, SpreadsheetChartLineDashStyle, SpreadsheetChartLineGrouping, SpreadsheetChartMarkerSymbol, SpreadsheetChartType } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizeSpreadsheetChartLineOptions } from "../spreadsheet/chart-line-options.mjs";
import { normalizeSpreadsheetChartSeriesLine, SPREADSHEET_CHART_LINE_MAX_WIDTH_POINTS } from "../spreadsheet/chart-line-style.mjs";
import { normalizeSpreadsheetChartSeriesMarker } from "../spreadsheet/chart-marker-style.mjs";
import { normalizeSpreadsheetChartDataLabels } from "../spreadsheet/chart-data-labels.mjs";
import { resolvedWorksheetChartCategories, resolvedWorksheetChartSeriesValues } from "../spreadsheet/chart-source-data.mjs";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

const EMU_PER_PIXEL = 9525;
const MAX_TEXT = 32_767;
const MAX_FORMULA = 8_192;
const MAX_SERIES = 256;
const MAX_POINTS = 1_048_576;
const MAX_FONT_SIZE_POINTS = 4_000;
const TYPES_TO_WIRE = new Map([
  ["bar", SpreadsheetChartType.BAR],
  ["line", SpreadsheetChartType.LINE],
  ["pie", SpreadsheetChartType.PIE],
]);
const TYPES_FROM_WIRE = new Map([...TYPES_TO_WIRE].map(([name, value]) => [value, name]));
const LINE_STYLES_TO_WIRE = new Map([
  ["solid", SpreadsheetChartLineDashStyle.SOLID],
  ["dashed", SpreadsheetChartLineDashStyle.DASHED],
  ["dotted", SpreadsheetChartLineDashStyle.DOTTED],
  ["dash-dot", SpreadsheetChartLineDashStyle.DASH_DOT],
  ["dash-dot-dot", SpreadsheetChartLineDashStyle.DASH_DOT_DOT],
]);
const LINE_STYLES_FROM_WIRE = new Map([...LINE_STYLES_TO_WIRE].map(([name, value]) => [value, name]));
const LINE_GROUPINGS_TO_WIRE = new Map([
  ["standard", SpreadsheetChartLineGrouping.STANDARD],
  ["stacked", SpreadsheetChartLineGrouping.STACKED],
  ["percentStacked", SpreadsheetChartLineGrouping.PERCENT_STACKED],
]);
const LINE_GROUPINGS_FROM_WIRE = new Map([...LINE_GROUPINGS_TO_WIRE].map(([name, value]) => [value, name]));
const DATA_LABEL_POSITIONS_TO_WIRE = new Map([
  ["bestFit", SpreadsheetChartDataLabelPosition.BEST_FIT],
  ["bottom", SpreadsheetChartDataLabelPosition.BOTTOM],
  ["center", SpreadsheetChartDataLabelPosition.CENTER],
  ["insideBase", SpreadsheetChartDataLabelPosition.INSIDE_BASE],
  ["insideEnd", SpreadsheetChartDataLabelPosition.INSIDE_END],
  ["left", SpreadsheetChartDataLabelPosition.LEFT],
  ["outsideEnd", SpreadsheetChartDataLabelPosition.OUTSIDE_END],
  ["right", SpreadsheetChartDataLabelPosition.RIGHT],
  ["top", SpreadsheetChartDataLabelPosition.TOP],
]);
const DATA_LABEL_POSITIONS_FROM_WIRE = new Map([...DATA_LABEL_POSITIONS_TO_WIRE].map(([name, value]) => [value, name]));
const MARKER_SYMBOLS_TO_WIRE = new Map([
  ["none", SpreadsheetChartMarkerSymbol.NONE],
  ["dot", SpreadsheetChartMarkerSymbol.DOT],
  ["circle", SpreadsheetChartMarkerSymbol.CIRCLE],
  ["square", SpreadsheetChartMarkerSymbol.SQUARE],
  ["diamond", SpreadsheetChartMarkerSymbol.DIAMOND],
  ["triangle", SpreadsheetChartMarkerSymbol.TRIANGLE],
  ["x", SpreadsheetChartMarkerSymbol.X],
  ["star", SpreadsheetChartMarkerSymbol.STAR],
  ["plus", SpreadsheetChartMarkerSymbol.PLUS],
  ["dash", SpreadsheetChartMarkerSymbol.DASH],
]);
const MARKER_SYMBOLS_FROM_WIRE = new Map([...MARKER_SYMBOLS_TO_WIRE].map(([name, value]) => [value, name]));

function fail(chart, message, code = "invalid_spreadsheet_chart") {
  throw new OpenChestnutCodecError(`Worksheet chart ${chart?.name || chart?.id || "(unnamed)"} ${message}`, [], { code });
}

function text(value, name, chart, maximum = MAX_TEXT) {
  const output = String(value ?? "");
  if (output.length > maximum || /\p{Cc}/u.test(output)) fail(chart, `${name} must contain at most ${maximum} characters without controls.`);
  return output;
}

function formula(value, name, chart) {
  const output = text(value, name, chart, MAX_FORMULA);
  if (output.startsWith("=")) fail(chart, `${name} must omit the leading equals sign.`);
  return output;
}

function finite(value, name, chart) {
  const output = Number(value);
  if (!Number.isFinite(output)) fail(chart, `${name} must be finite.`);
  return output;
}

function chartSeries(chart) {
  return Array.isArray(chart?.series?.items) ? chart.series.items : Array.isArray(chart?.series) ? chart.series : [];
}

function seriesFill(value, name, chart) {
  if (value == null) return null;
  if (typeof value !== "string" || !/^#[0-9a-f]{6}$/i.test(value)) fail(chart, `${name} must be a #RRGGBB solid color.`);
  return value.toUpperCase();
}

function seriesFillFromWire(value, name, chart) {
  if (value == null) return undefined;
  if (value.source?.case !== "rgb") fail(chart, `${name} has an unsupported non-RGB fill source.`, "unsupported_spreadsheet_chart");
  return seriesFill(`#${String(value.source.value)}`, name, chart);
}

function seriesLineSnapshot(series, chart) {
  try {
    return normalizeSpreadsheetChartSeriesLine(series);
  } catch (error) {
    const message = String(error?.message || error).replace(/^Worksheet chart\s+/, "");
    const unsupported = /supports only|aliases must describe the same/i.test(message);
    fail(chart, message, unsupported ? "unsupported_spreadsheet_chart" : "invalid_spreadsheet_chart");
  }
}

function seriesLineFromWire(value, name, chart) {
  if (value == null) return undefined;
  const output = {};
  if (value.color != null) {
    if (value.color.source?.case !== "rgb") fail(chart, `${name} has an unsupported non-RGB color source.`, "unsupported_spreadsheet_chart");
    output.fill = seriesFill(`#${String(value.color.source.value)}`, `${name}.fill`, chart);
  }
  const dashStyle = value.dashStyle ?? SpreadsheetChartLineDashStyle.UNSPECIFIED;
  if (dashStyle !== SpreadsheetChartLineDashStyle.UNSPECIFIED) {
    const style = LINE_STYLES_FROM_WIRE.get(dashStyle);
    if (!style) fail(chart, `${name} has unsupported dash style ${dashStyle}.`, "unsupported_spreadsheet_chart");
    output.style = style;
  }
  if (value.widthPoints != null) {
    const width = finite(value.widthPoints, `${name}.width`, chart);
    if (width < 0 || width > SPREADSHEET_CHART_LINE_MAX_WIDTH_POINTS) fail(chart, `${name}.width must be from 0 through ${SPREADSHEET_CHART_LINE_MAX_WIDTH_POINTS} points.`);
    output.width = width;
  }
  return output;
}

function seriesMarkerSnapshot(value, chart) {
  try {
    return normalizeSpreadsheetChartSeriesMarker(value);
  } catch (error) {
    const message = String(error?.message || error).replace(/^Worksheet chart\s+/, "");
    const unsupported = /supports only/i.test(message);
    fail(chart, message, unsupported ? "unsupported_spreadsheet_chart" : "invalid_spreadsheet_chart");
  }
}

function seriesMarkerFromWire(value, name, chart) {
  if (value == null) return undefined;
  const output = {};
  const symbolValue = value.symbol ?? SpreadsheetChartMarkerSymbol.UNSPECIFIED;
  if (symbolValue !== SpreadsheetChartMarkerSymbol.UNSPECIFIED) {
    const symbol = MARKER_SYMBOLS_FROM_WIRE.get(symbolValue);
    if (!symbol) fail(chart, `${name} has unsupported symbol ${symbolValue}.`, "unsupported_spreadsheet_chart");
    output.symbol = symbol;
  }
  if (value.size != null) output.size = Number(value.size);
  if (value.fill != null) output.fill = seriesFillFromWire(value.fill, `${name}.fill`, chart);
  if (value.line != null) output.line = seriesLineFromWire(value.line, `${name}.line`, chart);
  return seriesMarkerSnapshot(output, chart) || undefined;
}

function lineOptionsSnapshot(value, chart) {
  try { return normalizeSpreadsheetChartLineOptions(value); }
  catch (error) {
    const message = String(error?.message || error).replace(/^Worksheet chart\s+/, "");
    fail(chart, message, /supports only/i.test(message) ? "unsupported_spreadsheet_chart" : "invalid_spreadsheet_chart");
  }
}

function dataLabelsSnapshot(value, chart) {
  try { return normalizeSpreadsheetChartDataLabels(value); }
  catch (error) {
    const message = String(error?.message || error).replace(/^Worksheet chart\s+/, "");
    fail(chart, message, /supports only/i.test(message) ? "unsupported_spreadsheet_chart" : "invalid_spreadsheet_chart");
  }
}

function textStyleSnapshot(value, name, chart) {
  if (value == null) return null;
  if (typeof value !== "object" || Array.isArray(value)) fail(chart, `${name} must be an object.`);
  const unsupported = Object.keys(value).filter((key) => key !== "fontSize" && value[key] != null);
  if (unsupported.length) fail(chart, `${name} supports only fontSize; received ${unsupported.join(", ")}.`, "unsupported_spreadsheet_chart");
  if (value.fontSize == null) return null;
  const fontSize = finite(value.fontSize, `${name}.fontSize`, chart);
  if (fontSize < 1 || fontSize > MAX_FONT_SIZE_POINTS) fail(chart, `${name}.fontSize must be from 1 through ${MAX_FONT_SIZE_POINTS} points.`);
  return { fontSize };
}

function textStyleFromWire(value, name, chart) {
  if (value == null) return undefined;
  if (value.fontSizePoints == null) fail(chart, `${name} has no supported font size.`, "unsupported_spreadsheet_chart");
  return textStyleSnapshot({ fontSize: value.fontSizePoints }, name, chart) || undefined;
}

function axisSnapshot(axis, kind, chart) {
  if (axis == null) return null;
  const title = typeof axis.title === "string" ? axis.title : axis.title?.text;
  const aliasMajorUnit = kind === "y" && axis.majorUnit == null ? axis.tickLabelInterval : axis.majorUnit;
  return {
    axisType: String(axis.axisType || (kind === "x" ? "textAxis" : "valueAxis")),
    title: String(title ?? ""),
    textStyle: textStyleSnapshot(axis.textStyle, `${kind}Axis.textStyle`, chart),
    numberFormatCode: axis.numberFormatCode == null ? "" : String(axis.numberFormatCode),
    tickLabelInterval: kind === "x" && axis.tickLabelInterval != null ? Number(axis.tickLabelInterval) : null,
    minimum: axis.min == null ? null : Number(axis.min),
    maximum: axis.max == null ? null : Number(axis.max),
    majorUnit: aliasMajorUnit == null ? null : Number(aliasMajorUnit),
  };
}

export function spreadsheetChartSnapshot(chart, options = {}) {
  const position = chart?.position || {};
  const categories = options.resolveSourceData
    ? resolvedWorksheetChartCategories(chart)
    : [...(chart?.categories || [])].map((value) => String(value));
  return {
    id: String(chart?.id || ""),
    name: String(chart?.name || ""),
    title: String(chart?.title || ""),
    titleTextStyle: textStyleSnapshot(chart?.titleTextStyle, "titleTextStyle", chart),
    lineOptions: lineOptionsSnapshot(chart?.lineOptions, chart),
    dataLabels: dataLabelsSnapshot(chart?.dataLabels, chart),
    type: String(chart?.type || chart?.chartType || "bar").toLowerCase(),
    hasLegend: chart?.hasLegend !== false,
    categories,
    xAxis: axisSnapshot(chart?.xAxis, "x", chart),
    yAxis: axisSnapshot(chart?.yAxis, "y", chart),
    position: {
      left: Number(position.left ?? 420),
      top: Number(position.top ?? 40),
      width: Number(position.width ?? 360),
      height: Number(position.height ?? 220),
    },
    series: chartSeries(chart).map((series) => ({
      name: String(series?.name || ""),
      values: options.resolveSourceData ? resolvedWorksheetChartSeriesValues(chart, series) : [...(series?.values || [])].map(Number),
      categoryFormula: series?.categoryFormula == null ? "" : String(series.categoryFormula),
      formula: series?.formula == null ? "" : String(series.formula),
      fill: series?.fill == null ? null : String(series.fill),
      line: seriesLineSnapshot(series, chart),
      marker: seriesMarkerSnapshot(series?.marker, chart),
    })),
  };
}

function validateSnapshot(snapshot, chart) {
  if (!snapshot.id || snapshot.id.length > 512 || /\p{Cc}/u.test(snapshot.id)) fail(chart, "id must contain 1 through 512 characters without controls.");
  if (!snapshot.name || snapshot.name.length > 255 || /\p{Cc}/u.test(snapshot.name)) fail(chart, "name must contain 1 through 255 characters without controls.");
  text(snapshot.title, "title", chart);
  if (snapshot.titleTextStyle != null && snapshot.title.length === 0) fail(chart, "titleTextStyle requires a non-empty title.");
  const type = TYPES_TO_WIRE.get(snapshot.type);
  if (type == null) fail(chart, `type must be bar, line, or pie; received ${snapshot.type}.`, "unsupported_spreadsheet_chart");
  if (snapshot.lineOptions != null && snapshot.type !== "line") fail(chart, "lineOptions require a line chart.", "unsupported_spreadsheet_chart");
  if (snapshot.type === "pie") {
    if (snapshot.xAxis != null || snapshot.yAxis != null) fail(chart, "pie charts cannot carry category/value axes in the bounded native profile.", "unsupported_spreadsheet_chart");
  } else {
    if (snapshot.xAxis == null || snapshot.yAxis == null) fail(chart, "bar and line charts require primary xAxis and yAxis objects.");
    if ([chart?.xAxis, chart?.yAxis].some((axis) => axis?.title?.textStyle != null)) fail(chart, "axis-title text styling is outside the bounded native chart profile.", "unsupported_spreadsheet_chart");
    validateAxis(snapshot.xAxis, "x", chart);
    validateAxis(snapshot.yAxis, "y", chart);
    if (chart?.yAxis?.majorUnit != null && chart?.yAxis?.tickLabelInterval != null && Number(chart.yAxis.majorUnit) !== Number(chart.yAxis.tickLabelInterval)) {
      fail(chart, "yAxis.majorUnit and compatibility alias yAxis.tickLabelInterval must match when both are set.");
    }
  }
  if (snapshot.series.length < 1 || snapshot.series.length > MAX_SERIES) fail(chart, `must contain 1 through ${MAX_SERIES} series.`);
  if (snapshot.categories.length > MAX_POINTS) fail(chart, `exceeds the ${MAX_POINTS}-category budget.`);
  snapshot.categories.forEach((value, index) => text(value, `category ${index + 1}`, chart));
  let points = 0;
  snapshot.series.forEach((series, seriesIndex) => {
    text(series.name, `series ${seriesIndex + 1} name`, chart, 255);
    formula(series.categoryFormula, `series ${seriesIndex + 1} categoryFormula`, chart);
    formula(series.formula, `series ${seriesIndex + 1} formula`, chart);
    series.fill = seriesFill(series.fill, `series ${seriesIndex + 1} fill`, chart);
    if (series.marker != null && snapshot.type !== "line") fail(chart, `series ${seriesIndex + 1} markers require a line chart.`, "unsupported_spreadsheet_chart");
    points += series.values.length;
    if (points > MAX_POINTS) fail(chart, `exceeds the ${MAX_POINTS}-value budget.`);
    if (series.values.length !== snapshot.categories.length) fail(chart, `series ${seriesIndex + 1} has ${series.values.length} values for ${snapshot.categories.length} categories.`);
    series.values.forEach((value, pointIndex) => finite(value, `series ${seriesIndex + 1} value ${pointIndex + 1}`, chart));
  });
  for (const [name, value] of Object.entries(snapshot.position)) {
    const number = finite(value, `position.${name}`, chart);
    if (Math.abs(number) > 10_000_000 || (["width", "height"].includes(name) && number <= 0)) fail(chart, `position.${name} must be bounded${["width", "height"].includes(name) ? " and positive" : ""}.`);
  }
  return type;
}

function validateAxis(axis, kind, chart) {
  const expectedType = kind === "x" ? "textAxis" : "valueAxis";
  if (axis.axisType !== expectedType) fail(chart, `${kind}Axis.axisType must be ${expectedType}; received ${axis.axisType}.`, "unsupported_spreadsheet_chart");
  text(axis.title, `${kind}Axis.title.text`, chart);
  text(axis.numberFormatCode, `${kind}Axis.numberFormatCode`, chart, 255);
  if (kind === "x") {
    if (axis.minimum != null || axis.maximum != null || axis.majorUnit != null) fail(chart, "xAxis supports title, numberFormatCode, and tickLabelInterval only.", "unsupported_spreadsheet_chart");
    if (axis.tickLabelInterval != null && (!Number.isInteger(axis.tickLabelInterval) || axis.tickLabelInterval < 1 || axis.tickLabelInterval > MAX_POINTS)) fail(chart, `xAxis.tickLabelInterval must be an integer from 1 through ${MAX_POINTS}.`);
    return;
  }
  if (axis.tickLabelInterval != null) fail(chart, "yAxis.tickLabelInterval is normalized to majorUnit and cannot remain a category-axis interval.");
  for (const [name, value] of [["min", axis.minimum], ["max", axis.maximum], ["majorUnit", axis.majorUnit]]) if (value != null) finite(value, `yAxis.${name}`, chart);
  if (axis.minimum != null && axis.maximum != null && axis.minimum >= axis.maximum) fail(chart, "yAxis.min must be less than yAxis.max.");
  if (axis.majorUnit != null && axis.majorUnit <= 0) fail(chart, "yAxis.majorUnit must be positive.");
}

function wireAxis(axis) {
  if (axis == null) return undefined;
  return {
    title: axis.title,
    textStyle: axis.textStyle == null ? undefined : { fontSizePoints: axis.textStyle.fontSize },
    numberFormatCode: axis.numberFormatCode,
    tickLabelInterval: axis.tickLabelInterval == null ? undefined : axis.tickLabelInterval,
    minimum: axis.minimum == null ? undefined : axis.minimum,
    maximum: axis.maximum == null ? undefined : axis.maximum,
    majorUnit: axis.majorUnit == null ? undefined : axis.majorUnit,
  };
}

function wireChart(chart, original) {
  const snapshot = spreadsheetChartSnapshot(chart, { resolveSourceData: original == null });
  const type = validateSnapshot(snapshot, chart);
  const output = {
    id: snapshot.id,
    name: snapshot.name,
    title: snapshot.title,
    titleTextStyle: snapshot.titleTextStyle == null ? undefined : { fontSizePoints: snapshot.titleTextStyle.fontSize },
    lineOptions: snapshot.lineOptions == null ? undefined : {
      grouping: snapshot.lineOptions.grouping == null ? undefined : LINE_GROUPINGS_TO_WIRE.get(snapshot.lineOptions.grouping),
      smooth: snapshot.lineOptions.smooth,
      varyColors: snapshot.lineOptions.varyColors === true,
    },
    dataLabels: snapshot.dataLabels == null ? undefined : {
      showValue: snapshot.dataLabels.showValue,
      showCategoryName: snapshot.dataLabels.showCategoryName,
      position: snapshot.dataLabels.position == null ? undefined : DATA_LABEL_POSITIONS_TO_WIRE.get(snapshot.dataLabels.position),
      showSeriesName: snapshot.dataLabels.showSeriesName,
    },
    type,
    hasLegend: snapshot.hasLegend,
    categories: snapshot.categories,
    xAxis: wireAxis(snapshot.xAxis),
    yAxis: wireAxis(snapshot.yAxis),
    series: snapshot.series.map((series) => ({
      name: series.name,
      values: series.values,
      categoryFormula: series.categoryFormula,
      valueFormula: series.formula,
      fill: series.fill == null ? undefined : { source: { case: "rgb", value: series.fill.slice(1) } },
      line: series.line == null ? undefined : {
        color: series.line.fill == null ? undefined : { source: { case: "rgb", value: series.line.fill.slice(1) } },
        dashStyle: series.line.style == null ? SpreadsheetChartLineDashStyle.UNSPECIFIED : LINE_STYLES_TO_WIRE.get(series.line.style),
        widthPoints: series.line.width == null ? undefined : series.line.width,
      },
      marker: series.marker == null ? undefined : {
        symbol: series.marker.symbol == null ? SpreadsheetChartMarkerSymbol.UNSPECIFIED : MARKER_SYMBOLS_TO_WIRE.get(series.marker.symbol),
        size: series.marker.size == null ? undefined : series.marker.size,
        fill: series.marker.fill == null ? undefined : { source: { case: "rgb", value: series.marker.fill.slice(1) } },
        line: series.marker.line == null ? undefined : {
          color: series.marker.line.fill == null ? undefined : { source: { case: "rgb", value: series.marker.line.fill.slice(1) } },
          dashStyle: series.marker.line.style == null ? SpreadsheetChartLineDashStyle.UNSPECIFIED : LINE_STYLES_TO_WIRE.get(series.marker.line.style),
          widthPoints: series.marker.line.width == null ? undefined : series.marker.line.width,
        },
      },
    })),
    source: original?.source,
  };
  if (original) {
    if (original.anchor) output.anchor = original.anchor;
    else if (original.twoCellAnchor) output.twoCellAnchor = original.twoCellAnchor;
    else if (original.absoluteAnchor) output.absoluteAnchor = original.absoluteAnchor;
    else fail(chart, "imported source has no recognized anchor.", "invalid_spreadsheet_chart_topology");
    return output;
  }
  output.absoluteAnchor = {
    xEmu: BigInt(Math.round((snapshot.position.left - 40) * EMU_PER_PIXEL)),
    yEmu: BigInt(Math.round((snapshot.position.top - 40) * EMU_PER_PIXEL)),
    widthEmu: BigInt(Math.round(snapshot.position.width * EMU_PER_PIXEL)),
    heightEmu: BigInt(Math.round(snapshot.position.height * EMU_PER_PIXEL)),
  };
  return output;
}

export function wireWorksheetCharts(sheet, state) {
  const remaining = new Set(sheet.charts?.items || []);
  const output = [];
  for (const slot of state?.slots || []) {
    if (!remaining.delete(slot.chart)) {
      throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot remove imported chart ${slot.chart?.name || slot.wire.id} in the bounded OpenChestnut slice.`, [], { code: "invalid_spreadsheet_chart_topology" });
    }
    const current = spreadsheetChartSnapshot(slot.chart);
    if (JSON.stringify(current.position) !== JSON.stringify(slot.publicSnapshot.position)) {
      fail(slot.chart, "cannot change imported anchor geometry in the bounded OpenChestnut slice.", "unsupported_spreadsheet_chart_edit");
    }
    if (JSON.stringify(current) === JSON.stringify(slot.publicSnapshot)) output.push(slot.wire);
    else {
      if (slot.wire.source?.editable !== true) fail(slot.chart, "is read-only because its native chart profile is outside the editable subset.", "unsupported_spreadsheet_chart_edit");
      output.push(wireChart(slot.chart, slot.wire));
    }
  }
  if (state && remaining.size) {
    const chart = [...remaining][0];
    fail(chart, "cannot be added to an imported source package in the bounded OpenChestnut slice.", "invalid_spreadsheet_chart_topology");
  }
  output.push(...[...remaining].map((chart) => wireChart(chart)));
  return output;
}

function columnWidthPixels(sheet, column) {
  const dimension = sheet.columnDimensions?.get(column) || {};
  if (dimension.hidden) return 0;
  const width = dimension.width == null ? Math.trunc(((8.43 * 7 + 5) / 7) * 256) / 256 : Number(dimension.width);
  return Math.trunc(((256 * width + Math.trunc(128 / 7)) / 256) * 7);
}

function rowHeightPixels(sheet, row) {
  const dimension = sheet.rowDimensions?.get(row) || {};
  if (dimension.hidden) return 0;
  return Number(dimension.height ?? 15) * 96 / 72;
}

function axisOffset(sheet, axis, index) {
  let output = 0;
  for (let current = 0; current < index; current += 1) output += axis === "column" ? columnWidthPixels(sheet, current) : rowHeightPixels(sheet, current);
  return output;
}

function markerPosition(sheet, marker) {
  return {
    left: 40 + axisOffset(sheet, "column", Number(marker?.column || 0)) + Number(marker?.columnOffsetEmu || 0n) / EMU_PER_PIXEL,
    top: 40 + axisOffset(sheet, "row", Number(marker?.row || 0)) + Number(marker?.rowOffsetEmu || 0n) / EMU_PER_PIXEL,
  };
}

function positionFromWire(sheet, source) {
  if (source.absoluteAnchor) return {
    left: 40 + Number(source.absoluteAnchor.xEmu) / EMU_PER_PIXEL,
    top: 40 + Number(source.absoluteAnchor.yEmu) / EMU_PER_PIXEL,
    width: Number(source.absoluteAnchor.widthEmu) / EMU_PER_PIXEL,
    height: Number(source.absoluteAnchor.heightEmu) / EMU_PER_PIXEL,
  };
  if (source.twoCellAnchor) {
    const from = markerPosition(sheet, source.twoCellAnchor.from);
    const to = markerPosition(sheet, source.twoCellAnchor.to);
    return { left: from.left, top: from.top, width: Math.max(1, to.left - from.left), height: Math.max(1, to.top - from.top) };
  }
  const from = markerPosition(sheet, source.anchor);
  return {
    left: from.left,
    top: from.top,
    width: Number(source.anchor?.widthEmu || 0n) / EMU_PER_PIXEL,
    height: Number(source.anchor?.heightEmu || 0n) / EMU_PER_PIXEL,
  };
}

function axisFromWire(axis, kind) {
  if (axis == null) return undefined;
  return {
    axisType: kind === "x" ? "textAxis" : "valueAxis",
    title: { text: axis.title || "" },
    ...(axis.textStyle == null ? {} : { textStyle: textStyleFromWire(axis.textStyle, `${kind}Axis.textStyle`, axis) }),
    ...(axis.numberFormatCode ? { numberFormatCode: axis.numberFormatCode } : {}),
    ...(axis.tickLabelInterval == null ? {} : { tickLabelInterval: axis.tickLabelInterval }),
    ...(axis.minimum == null ? {} : { min: axis.minimum }),
    ...(axis.maximum == null ? {} : { max: axis.maximum }),
    ...(axis.majorUnit == null ? {} : { majorUnit: axis.majorUnit }),
  };
}

export function spreadsheetChartFromWire(sheet, source) {
  const type = TYPES_FROM_WIRE.get(source.type);
  if (!type) fail(source, `has unsupported wire type ${source.type}.`, "unsupported_spreadsheet_chart");
  const sourceSeries = source.series || [];
  const importedFills = sourceSeries.map((series, index) => seriesFillFromWire(series.fill, `series ${index + 1} fill`, source));
  const importedLines = sourceSeries.map((series, index) => seriesLineFromWire(series.line, `series ${index + 1} line`, source));
  const importedMarkers = sourceSeries.map((series, index) => seriesMarkerFromWire(series.marker, `series ${index + 1} marker`, source));
  if (type !== "line" && importedMarkers.some((marker) => marker != null)) fail(source, "series markers require a line chart.", "unsupported_spreadsheet_chart");
  const titleTextStyle = textStyleFromWire(source.titleTextStyle, "titleTextStyle", source);
  let lineOptionsInput = null;
  if (source.lineOptions != null) {
    lineOptionsInput = {};
    if (source.lineOptions.grouping != null) {
      const grouping = LINE_GROUPINGS_FROM_WIRE.get(source.lineOptions.grouping);
      if (!grouping) fail(source, `lineOptions has unsupported grouping ${source.lineOptions.grouping}.`, "unsupported_spreadsheet_chart");
      lineOptionsInput.grouping = grouping;
    }
    if (source.lineOptions.smooth != null) lineOptionsInput.smooth = source.lineOptions.smooth;
    if (source.lineOptions.varyColors === true) lineOptionsInput.varyColors = true;
    if (Object.keys(lineOptionsInput).length === 0) fail(source, "lineOptions must carry explicit grouping, smooth, or vary-colors semantics.");
  }
  const lineOptions = lineOptionsSnapshot(lineOptionsInput, source);
  if (type !== "line" && lineOptions != null) fail(source, "lineOptions require a line chart.", "unsupported_spreadsheet_chart");
  let dataLabelsInput;
  if (source.dataLabels != null) {
    dataLabelsInput = {
      showValue: source.dataLabels.showValue === true,
      showCategoryName: source.dataLabels.showCategoryName === true,
    };
    if (source.dataLabels.showSeriesName != null) dataLabelsInput.showSeriesName = source.dataLabels.showSeriesName;
    if (source.dataLabels.position != null) {
      const position = DATA_LABEL_POSITIONS_FROM_WIRE.get(source.dataLabels.position);
      if (!position) fail(source, `dataLabels has unsupported position ${source.dataLabels.position}.`, "unsupported_spreadsheet_chart");
      dataLabelsInput.position = position;
    }
  }
  const dataLabels = source.dataLabels == null ? undefined : dataLabelsSnapshot(dataLabelsInput, source);
  const chart = sheet.charts.add(type, {
    name: source.name,
    title: source.title,
    ...(titleTextStyle == null ? {} : { titleTextStyle }),
    ...(lineOptions == null ? {} : { lineOptions }),
    ...(dataLabels == null ? {} : { dataLabels }),
    hasLegend: source.hasLegend,
    categories: [...(source.categories || [])],
    xAxis: axisFromWire(source.xAxis, "x"),
    yAxis: axisFromWire(source.yAxis, "y"),
    position: positionFromWire(sheet, source),
    series: sourceSeries.map((series, index) => {
      return {
        name: series.name,
        values: [...series.values],
        ...(importedFills[index] == null ? {} : { fill: importedFills[index] }),
        ...(importedLines[index] == null ? {} : { line: importedLines[index] }),
        ...(importedMarkers[index] == null ? {} : { marker: importedMarkers[index] }),
      };
    }),
  });
  chart.id = source.id || chart.id;
  sourceSeries.forEach((series, index) => Object.assign(chart.series.items[index], {
    categoryFormula: series.categoryFormula || undefined,
    formula: series.valueFormula || undefined,
    fill: importedFills[index],
    ...(importedLines[index] == null ? {} : { line: importedLines[index] }),
    ...(importedMarkers[index] == null ? {} : { marker: importedMarkers[index] }),
  }));
  return chart;
}
