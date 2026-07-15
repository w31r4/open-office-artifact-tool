import { SpreadsheetChartType } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

const EMU_PER_PIXEL = 9525;
const MAX_TEXT = 32_767;
const MAX_FORMULA = 8_192;
const MAX_SERIES = 256;
const MAX_POINTS = 1_048_576;
const TYPES_TO_WIRE = new Map([
  ["bar", SpreadsheetChartType.BAR],
  ["line", SpreadsheetChartType.LINE],
  ["pie", SpreadsheetChartType.PIE],
]);
const TYPES_FROM_WIRE = new Map([...TYPES_TO_WIRE].map(([name, value]) => [value, name]));

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

export function spreadsheetChartSnapshot(chart) {
  const position = chart?.position || {};
  return {
    id: String(chart?.id || ""),
    name: String(chart?.name || ""),
    title: String(chart?.title || ""),
    type: String(chart?.type || chart?.chartType || "bar").toLowerCase(),
    hasLegend: chart?.hasLegend !== false,
    categories: [...(chart?.categories || [])].map((value) => String(value)),
    position: {
      left: Number(position.left ?? 420),
      top: Number(position.top ?? 40),
      width: Number(position.width ?? 360),
      height: Number(position.height ?? 220),
    },
    series: chartSeries(chart).map((series) => ({
      name: String(series?.name || ""),
      values: [...(series?.values || [])].map(Number),
      categoryFormula: series?.categoryFormula == null ? "" : String(series.categoryFormula),
      formula: series?.formula == null ? "" : String(series.formula),
    })),
  };
}

function validateSnapshot(snapshot, chart) {
  if (!snapshot.id || snapshot.id.length > 512 || /\p{Cc}/u.test(snapshot.id)) fail(chart, "id must contain 1 through 512 characters without controls.");
  if (!snapshot.name || snapshot.name.length > 255 || /\p{Cc}/u.test(snapshot.name)) fail(chart, "name must contain 1 through 255 characters without controls.");
  text(snapshot.title, "title", chart);
  const type = TYPES_TO_WIRE.get(snapshot.type);
  if (type == null) fail(chart, `type must be bar, line, or pie; received ${snapshot.type}.`, "unsupported_spreadsheet_chart");
  if (snapshot.series.length < 1 || snapshot.series.length > MAX_SERIES) fail(chart, `must contain 1 through ${MAX_SERIES} series.`);
  if (chartSeries(chart).some((series) => series?.fill != null)) fail(chart, "series fill styling is outside the bounded native chart profile.", "unsupported_spreadsheet_chart");
  if (snapshot.categories.length > MAX_POINTS) fail(chart, `exceeds the ${MAX_POINTS}-category budget.`);
  snapshot.categories.forEach((value, index) => text(value, `category ${index + 1}`, chart));
  let points = 0;
  snapshot.series.forEach((series, seriesIndex) => {
    text(series.name, `series ${seriesIndex + 1} name`, chart, 255);
    formula(series.categoryFormula, `series ${seriesIndex + 1} categoryFormula`, chart);
    formula(series.formula, `series ${seriesIndex + 1} formula`, chart);
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

function wireChart(chart, original) {
  const snapshot = spreadsheetChartSnapshot(chart);
  const type = validateSnapshot(snapshot, chart);
  const output = {
    id: snapshot.id,
    name: snapshot.name,
    title: snapshot.title,
    type,
    hasLegend: snapshot.hasLegend,
    categories: snapshot.categories,
    series: snapshot.series.map((series) => ({
      name: series.name,
      values: series.values,
      categoryFormula: series.categoryFormula,
      valueFormula: series.formula,
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

export function spreadsheetChartFromWire(sheet, source) {
  const type = TYPES_FROM_WIRE.get(source.type);
  if (!type) fail(source, `has unsupported wire type ${source.type}.`, "unsupported_spreadsheet_chart");
  const chart = sheet.charts.add(type, {
    name: source.name,
    title: source.title,
    hasLegend: source.hasLegend,
    categories: [...(source.categories || [])],
    position: positionFromWire(sheet, source),
    series: (source.series || []).map((series) => ({ name: series.name, values: [...series.values] })),
  });
  chart.id = source.id || chart.id;
  (source.series || []).forEach((series, index) => Object.assign(chart.series.items[index], {
    categoryFormula: series.categoryFormula || undefined,
    formula: series.valueFormula || undefined,
  }));
  return chart;
}
