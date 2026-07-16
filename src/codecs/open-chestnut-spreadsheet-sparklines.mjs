import {
  SpreadsheetSparklineAxisMode,
  SpreadsheetSparklineEmptyCells,
  SpreadsheetSparklineType,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { parseRangeAddress } from "../spreadsheet/range-addressing.mjs";
import { sparklineRangeMappings } from "../spreadsheet/sparklines.mjs";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

const TYPE_TO_WIRE = new Map([
  ["line", SpreadsheetSparklineType.LINE],
  ["column", SpreadsheetSparklineType.COLUMN],
  ["stacked", SpreadsheetSparklineType.STACKED],
]);
const TYPE_FROM_WIRE = new Map([...TYPE_TO_WIRE].map(([name, value]) => [value, name]));
const EMPTY_TO_WIRE = new Map([
  ["span", SpreadsheetSparklineEmptyCells.SPAN],
  ["connect", SpreadsheetSparklineEmptyCells.SPAN],
  ["gap", SpreadsheetSparklineEmptyCells.GAP],
  ["zero", SpreadsheetSparklineEmptyCells.ZERO],
]);
const EMPTY_FROM_REFERENCE_NUMBER = new Map([
  [1, SpreadsheetSparklineEmptyCells.SPAN],
  [2, SpreadsheetSparklineEmptyCells.GAP],
  [3, SpreadsheetSparklineEmptyCells.ZERO],
]);
const EMPTY_TO_REFERENCE_NUMBER = new Map([...EMPTY_FROM_REFERENCE_NUMBER].map(([number, value]) => [value, number]));
const AXIS_TO_WIRE = new Map([
  ["individual", SpreadsheetSparklineAxisMode.INDIVIDUAL],
  ["group", SpreadsheetSparklineAxisMode.GROUP],
  ["custom", SpreadsheetSparklineAxisMode.CUSTOM],
]);
const AXIS_FROM_REFERENCE_NUMBER = new Map([
  [0, SpreadsheetSparklineAxisMode.INDIVIDUAL],
  [1, SpreadsheetSparklineAxisMode.GROUP],
  [2, SpreadsheetSparklineAxisMode.CUSTOM],
]);
const AXIS_TO_REFERENCE_NUMBER = new Map([...AXIS_FROM_REFERENCE_NUMBER].map(([number, value]) => [value, number]));

function fail(group, sheet, message, code = "invalid_spreadsheet_sparkline") {
  throw new OpenChestnutCodecError(`Worksheet ${sheet.name} sparkline ${group?.id || "(unnamed)"} ${message}`, [], { code });
}

function rangeReference(ref, fallbackSheet, label, group, sheet, { target = false } = {}) {
  const address = String(ref?.address || ref || "");
  let bounds;
  try { bounds = parseRangeAddress(address); }
  catch { fail(group, sheet, `${label} must be a bounded A1 range.`); }
  const sheetName = ref?.sheetName || fallbackSheet;
  if (target && ref?.sheetName && ref.sheetName !== sheet.name) fail(group, sheet, `${label} must belong to its containing worksheet.`);
  if (!sheetName) fail(group, sheet, `${label} has no worksheet.`);
  if (!target && !sheet.workbook.worksheets.getItem(sheetName)) fail(group, sheet, `${label} refers to missing worksheet ${sheetName}.`);
  const quoted = /^[A-Za-z_][A-Za-z0-9_.]*$/.test(sheetName) ? sheetName : `'${String(sheetName).replaceAll("'", "''")}'`;
  return { bounds, text: target ? address : `${quoted}!${address}` };
}

function wireColor(value, group, sheet, label) {
  if (value == null) return undefined;
  let tint;
  if (typeof value === "string") {
    const rgb = value.replace(/^#/, "").slice(-6).toUpperCase();
    if (!/^[0-9A-F]{6}$/.test(rgb)) fail(group, sheet, `${label} must be a six/eight-digit RGB or supported symbolic color.`);
    return { source: { case: "rgb", value: rgb } };
  }
  if (value.tint != null) {
    tint = Number(value.tint);
    if (!Number.isFinite(tint) || tint < -1 || tint > 1) fail(group, sheet, `${label} tint must be between -1 and 1.`);
  }
  if (value.theme != null) {
    const theme = Number(value.theme);
    if (!Number.isInteger(theme) || theme < 0 || theme > 11) fail(group, sheet, `${label} theme index must be 0 through 11.`);
    return { source: { case: "theme", value: theme }, tint };
  }
  if (value.indexed != null) {
    const indexed = Number(value.indexed);
    if (!Number.isInteger(indexed) || indexed < 0 || indexed > 65) fail(group, sheet, `${label} indexed value must be 0 through 65.`);
    return { source: { case: "indexed", value: indexed }, tint };
  }
  if (value.auto === true) return { source: { case: "automatic", value: true }, tint };
  if (value.rgb != null) return { ...wireColor(value.rgb, group, sheet, label), tint };
  fail(group, sheet, `${label} has no supported color source.`);
}

function publicColor(value) {
  if (!value?.source?.case) return undefined;
  const tint = value.tint == null || value.tint === 0 ? {} : { tint: value.tint };
  if (value.source.case === "rgb") return value.tint == null || value.tint === 0
    ? `#${String(value.source.value).slice(-6).toUpperCase()}`
    : { rgb: `#${String(value.source.value).slice(-6).toUpperCase()}`, ...tint };
  if (value.source.case === "theme") return { theme: value.source.value, ...tint };
  if (value.source.case === "indexed") return { indexed: value.source.value, ...tint };
  if (value.source.case === "automatic") return { auto: true, ...tint };
  return undefined;
}

function displayEmptyCells(value, group, sheet) {
  const wire = typeof value === "number" ? EMPTY_FROM_REFERENCE_NUMBER.get(value) : EMPTY_TO_WIRE.get(String(value || "gap"));
  if (wire == null) fail(group, sheet, "displayEmptyCellsAs must be span/connect, gap, zero, or the compatible numeric value 1, 2, or 3.");
  return wire;
}

function axisMode(value, fallback, group, sheet, label) {
  if (value == null) return fallback;
  const wire = typeof value === "number" ? AXIS_FROM_REFERENCE_NUMBER.get(value) : AXIS_TO_WIRE.get(String(value));
  if (wire == null) fail(group, sheet, `${label} must be individual/group/custom or the compatible numeric value 0, 1, or 2.`);
  return wire;
}

function optionalBoolean(object, key) {
  return Object.hasOwn(object || {}, key) ? Boolean(object[key]) : undefined;
}

function optionalFinite(object, key, group, sheet) {
  if (!Object.hasOwn(object || {}, key) || object[key] == null) return undefined;
  const value = Number(object[key]);
  if (!Number.isFinite(value)) fail(group, sheet, `${key} must be finite.`);
  return value;
}

export function spreadsheetSparklineSnapshot(group) {
  return {
    id: String(group.id || ""),
    type: String(group.type || "line"),
    targetRange: { sheetName: group.targetRange?.sheetName, address: String(group.targetRange?.address || "") },
    sourceData: { sheetName: group.sourceData?.sheetName, address: String(group.sourceData?.address || "") },
    dateAxisRange: group.dateAxisRange ? { sheetName: group.dateAxisRange.sheetName, address: String(group.dateAxisRange.address || "") } : undefined,
    seriesColor: group.seriesColor,
    negativeColor: group.negativeColor,
    axisColor: group.axisColor,
    markersColor: group.markersColor,
    firstMarkerColor: group.firstMarkerColor,
    lastMarkerColor: group.lastMarkerColor,
    highMarkerColor: group.highMarkerColor,
    lowMarkerColor: group.lowMarkerColor,
    lineWeight: Number(group.lineWeight),
    displayHidden: Boolean(group.displayHidden),
    displayEmptyCellsAs: group.displayEmptyCellsAs,
    markers: { ...(group.markers || {}) },
    axis: { ...(group.axis || {}) },
  };
}

function wireSparkline(group, sheet, source) {
  const type = TYPE_TO_WIRE.get(String(group.type));
  if (type == null) fail(group, sheet, "type must be line, column, or stacked.");
  const target = rangeReference(group.targetRange, sheet.name, "targetRange", group, sheet, { target: true });
  const data = rangeReference(group.sourceData, group.sourceData?.sheetName || sheet.name, "sourceData", group, sheet);
  const mappings = sparklineRangeMappings(group);
  if (!mappings.length) fail(group, sheet, "requires a one-dimensional target range and a reversible row- or column-oriented source rectangle.");
  if (target.bounds.rowCount * target.bounds.colCount !== mappings.length) fail(group, sheet, "target/source ranges do not produce exactly one native sparkline per target cell.");
  let dateAxisRange = "";
  if (group.dateAxisRange) {
    const date = rangeReference(group.dateAxisRange, group.dateAxisRange.sheetName || sheet.name, "dateAxisRange", group, sheet);
    if (date.bounds.rowCount > 1 && date.bounds.colCount > 1) fail(group, sheet, "dateAxisRange must be one-dimensional.");
    const firstSource = parseRangeAddress(mappings[0].sourceAddress);
    const pointCount = firstSource.rowCount * firstSource.colCount;
    if (date.bounds.rowCount * date.bounds.colCount !== pointCount) fail(group, sheet, "dateAxisRange must contain one entry for every point in each sparkline.");
    dateAxisRange = date.text;
  }
  const lineWeight = Number(group.lineWeight);
  if (!Number.isFinite(lineWeight) || lineWeight <= 0 || lineWeight > 1584) fail(group, sheet, "lineWeight must be greater than 0 and no more than 1584 points.");
  const markers = group.markers || {};
  const axis = group.axis || {};
  const manualMin = optionalFinite(axis, "manualMin", group, sheet);
  const manualMax = optionalFinite(axis, "manualMax", group, sheet);
  if (manualMin != null && manualMax != null && manualMin >= manualMax) fail(group, sheet, "axis.manualMin must be less than axis.manualMax.");
  const minMode = axisMode(axis.minMode, manualMin == null ? SpreadsheetSparklineAxisMode.INDIVIDUAL : SpreadsheetSparklineAxisMode.CUSTOM, group, sheet, "axis.minMode");
  const maxMode = axisMode(axis.maxMode, manualMax == null ? SpreadsheetSparklineAxisMode.INDIVIDUAL : SpreadsheetSparklineAxisMode.CUSTOM, group, sheet, "axis.maxMode");
  if (manualMin != null && minMode !== SpreadsheetSparklineAxisMode.CUSTOM) fail(group, sheet, "axis.manualMin requires minMode custom/2.");
  if (manualMax != null && maxMode !== SpreadsheetSparklineAxisMode.CUSTOM) fail(group, sheet, "axis.manualMax requires maxMode custom/2.");
  return {
    id: String(group.id || ""),
    type,
    targetRange: target.text,
    sourceDataRange: data.text,
    dateAxisRange,
    seriesColor: wireColor(group.seriesColor, group, sheet, "seriesColor"),
    negativeColor: wireColor(group.negativeColor, group, sheet, "negativeColor"),
    axisColor: wireColor(group.axisColor, group, sheet, "axisColor"),
    markersColor: wireColor(group.markersColor, group, sheet, "markersColor"),
    firstMarkerColor: wireColor(group.firstMarkerColor, group, sheet, "firstMarkerColor"),
    lastMarkerColor: wireColor(group.lastMarkerColor, group, sheet, "lastMarkerColor"),
    highMarkerColor: wireColor(group.highMarkerColor, group, sheet, "highMarkerColor"),
    lowMarkerColor: wireColor(group.lowMarkerColor, group, sheet, "lowMarkerColor"),
    lineWeight,
    displayHidden: Boolean(group.displayHidden),
    displayEmptyCellsAs: displayEmptyCells(group.displayEmptyCellsAs, group, sheet),
    markers: {
      show: optionalBoolean(markers, "show"), high: optionalBoolean(markers, "high"), low: optionalBoolean(markers, "low"),
      first: optionalBoolean(markers, "first"), last: optionalBoolean(markers, "last"), negative: optionalBoolean(markers, "negative"),
    },
    axis: {
      manualMin, manualMax, minMode, maxMode,
      showAxis: optionalBoolean(axis, "showAxis"), rightToLeft: optionalBoolean(axis, "rightToLeft"),
    },
    source,
  };
}

export function wireWorksheetSparklines(sheet, state) {
  if (!state) return (sheet.sparklineGroups?.items || []).map((group) => wireSparkline(group, sheet));
  const remaining = new Set(sheet.sparklineGroups?.items || []);
  const output = [];
  for (const slot of state.slots || []) {
    if (!slot.group || !remaining.delete(slot.group)) fail(slot.group || slot.wire, sheet, "cannot remove or reorder an imported source-bound group.", "invalid_spreadsheet_sparkline_topology");
    output.push(JSON.stringify(spreadsheetSparklineSnapshot(slot.group)) === JSON.stringify(slot.publicSnapshot)
      ? slot.wire
      : wireSparkline(slot.group, sheet, slot.wire.source));
  }
  if (remaining.size) fail([...remaining][0], sheet, "cannot add a group to a source-bound worksheet.", "invalid_spreadsheet_sparkline_topology");
  return output;
}

function publicRange(value) {
  const text = String(value || "");
  const bang = text.lastIndexOf("!");
  if (bang < 0) return text;
  const sheetName = text.slice(0, bang).replace(/^'|'$/g, "").replaceAll("''", "'");
  return { sheetName, address: text.slice(bang + 1) };
}

export function spreadsheetSparklineFromWire(sheet, source) {
  const type = TYPE_FROM_WIRE.get(source.type);
  const displayEmptyCellsAs = EMPTY_TO_REFERENCE_NUMBER.get(source.displayEmptyCellsAs);
  if (!type || !displayEmptyCellsAs) fail(source, sheet, "contains an unsupported wire type or empty-cell mode.");
  const axis = source.axis || {};
  const config = {
    id: source.id || undefined,
    type,
    targetRange: source.targetRange,
    sourceData: publicRange(source.sourceDataRange),
    ...(source.dateAxisRange ? { dateAxisRange: publicRange(source.dateAxisRange) } : {}),
    seriesColor: publicColor(source.seriesColor) || "#0EA5E9",
    negativeColor: publicColor(source.negativeColor),
    axisColor: publicColor(source.axisColor),
    markersColor: publicColor(source.markersColor),
    firstMarkerColor: publicColor(source.firstMarkerColor),
    lastMarkerColor: publicColor(source.lastMarkerColor),
    highMarkerColor: publicColor(source.highMarkerColor),
    lowMarkerColor: publicColor(source.lowMarkerColor),
    lineWeight: source.lineWeight ?? 1,
    displayHidden: source.displayHidden ?? false,
    displayEmptyCellsAs,
    markers: {
      show: source.markers?.show ?? false, high: source.markers?.high ?? false, low: source.markers?.low ?? false,
      first: source.markers?.first ?? false, last: source.markers?.last ?? false, negative: source.markers?.negative ?? false,
    },
    axis: {
      ...(axis.manualMin !== undefined ? { manualMin: axis.manualMin } : {}),
      ...(axis.manualMax !== undefined ? { manualMax: axis.manualMax } : {}),
      minMode: AXIS_TO_REFERENCE_NUMBER.get(axis.minMode) ?? 0,
      maxMode: AXIS_TO_REFERENCE_NUMBER.get(axis.maxMode) ?? 0,
      showAxis: axis.showAxis ?? true,
      rightToLeft: axis.rightToLeft ?? false,
    },
  };
  const group = sheet.sparklineGroups.add(config);
  if (source.id) group.id = source.id;
  return group;
}
