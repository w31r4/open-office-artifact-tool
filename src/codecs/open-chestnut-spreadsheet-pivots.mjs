import { SpreadsheetPivotAggregation, SpreadsheetPivotItemFilterMode } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";
import { pivotItemVisible } from "../spreadsheet/pivot-filters.mjs";

const A1_RANGE = /^\$?([A-Z]{1,3})\$?([1-9]\d*)(?::\$?([A-Z]{1,3})\$?([1-9]\d*))?$/i;
const MAX_NATIVE_VALUE_FIELDS = 32;
const MAX_NATIVE_FILTER_ITEMS = 1_024;
const TO_WIRE_AGGREGATION = new Map([
  ["sum", SpreadsheetPivotAggregation.SUM],
  ["count", SpreadsheetPivotAggregation.COUNT],
  ["average", SpreadsheetPivotAggregation.AVERAGE],
  ["min", SpreadsheetPivotAggregation.MIN],
  ["max", SpreadsheetPivotAggregation.MAX],
]);
const FROM_WIRE_AGGREGATION = new Map([...TO_WIRE_AGGREGATION].map(([name, value]) => [value, name]));

function wireFilterItem(value, label) {
  if (value === null) return { value: { case: "blankValue", value: true } };
  if (typeof value === "string") return { value: { case: "stringValue", value } };
  if (typeof value === "number" && Number.isFinite(value)) return { value: { case: "numberValue", value } };
  if (typeof value === "boolean") return { value: { case: "boolValue", value } };
  throw invalid(`${label} supports only exact string, finite number, boolean, or null items in the native profile.`, "unsupported_spreadsheet_pivot_filter");
}

function publicFilterItem(item, label) {
  if (item?.value?.case === "stringValue" || item?.value?.case === "errorValue") return item.value.value;
  if (item?.value?.case === "numberValue" && Number.isFinite(item.value.value)) return item.value.value;
  if (item?.value?.case === "boolValue") return item.value.value;
  if (item?.value?.case === "blankValue" && item.value.value === true) return null;
  throw invalid(`${label} contains an unsupported cached item.`, "unsupported_spreadsheet_pivot_filter");
}

function wireItemFilters(pivot) {
  return pivot.filters.map((filter) => {
    const mode = filter.include ? "include" : "exclude";
    const items = filter[mode];
    if (!items?.length || items.length > MAX_NATIVE_FILTER_ITEMS) {
      throw invalid(`PivotTable ${pivot.name} filter ${filter.field} requires 1 through ${MAX_NATIVE_FILTER_ITEMS} items.`, "unsupported_spreadsheet_pivot_filter");
    }
    return {
      field: filter.field,
      mode: mode === "include" ? SpreadsheetPivotItemFilterMode.INCLUDE : SpreadsheetPivotItemFilterMode.EXCLUDE,
      items: items.map((item, index) => wireFilterItem(item, `PivotTable ${pivot.name} filter ${filter.field} item ${index}`)),
    };
  });
}

function publicItemFilters(filters = [], pivotName = "PivotTable") {
  if (filters.length > 2) throw invalid(`${pivotName} exceeds the two-axis item-filter budget.`, "unsupported_spreadsheet_pivot_filter");
  const seen = new Set();
  return filters.map((filter, filterIndex) => {
    const field = String(filter?.field || "").trim();
    if (!field || seen.has(field) || ![SpreadsheetPivotItemFilterMode.INCLUDE, SpreadsheetPivotItemFilterMode.EXCLUDE].includes(filter.mode) ||
        !filter.items?.length || filter.items.length > MAX_NATIVE_FILTER_ITEMS) {
      throw invalid(`${pivotName} item filter ${filterIndex} is outside the bounded native profile.`, "unsupported_spreadsheet_pivot_filter");
    }
    seen.add(field);
    const mode = filter.mode === SpreadsheetPivotItemFilterMode.INCLUDE ? "include" : "exclude";
    return { field, [mode]: filter.items.map((item, index) => publicFilterItem(item, `${pivotName} filter ${field} item ${index}`)) };
  });
}

function invalid(message, code = "invalid_spreadsheet_pivot") {
  return new OpenChestnutCodecError(message, [], { code });
}

function columnIndex(label) {
  let value = 0;
  for (const character of label.toUpperCase()) value = value * 26 + character.charCodeAt(0) - 64;
  return value - 1;
}

function columnLabel(index) {
  let value = Number(index) + 1;
  let output = "";
  while (value > 0) {
    value -= 1;
    output = String.fromCharCode(65 + value % 26) + output;
    value = Math.floor(value / 26);
  }
  return output;
}

function parseRange(reference, label) {
  const match = A1_RANGE.exec(String(reference || "").trim());
  if (!match) throw invalid(`${label} must be one bounded A1 cell or rectangle.`);
  const left = columnIndex(match[1]);
  const top = Number(match[2]) - 1;
  const right = match[3] ? columnIndex(match[3]) : left;
  const bottom = match[4] ? Number(match[4]) - 1 : top;
  if (left > right || top > bottom || right > 16_383 || bottom > 1_048_575) throw invalid(`${label} is outside the XLSX worksheet grid.`);
  return { left, top, right, bottom, rowCount: bottom - top + 1, colCount: right - left + 1 };
}

function rangeAddress(bounds) {
  const first = `${columnLabel(bounds.left)}${bounds.top + 1}`;
  const last = `${columnLabel(bounds.right)}${bounds.bottom + 1}`;
  return first === last ? first : `${first}:${last}`;
}

function snapshot(value) {
  return JSON.stringify(value, (_key, item) => item instanceof Date ? { $date: item.toISOString() } : item);
}

function pivotSnapshot(pivot) {
  return snapshot({
    id: pivot.id,
    name: pivot.name,
    sourceRange: pivot.sourceRange,
    targetRange: pivot.targetRange,
    rowFields: pivot.rowFields,
    columnFields: pivot.columnFields,
    valueFields: pivot.valueFields,
    groupFields: pivot.groupFields,
    calculatedFields: pivot.calculatedFields,
    filters: pivot.filters,
    refreshPolicy: pivot.refreshPolicy,
    rowGrandTotals: pivot.rowGrandTotals,
    columnGrandTotals: pivot.columnGrandTotals,
  });
}

function rangeMatrix(sheet, reference) {
  return sheet.getRange(reference).values.map((row) => [...row]);
}

function outputMatrix(sheet, reference) {
  return rangeMatrix(sheet, reference);
}

function sourceSheetFor(workbook, targetSheet, pivot) {
  const name = pivot.sourceRange?.sheetName || targetSheet.name;
  const sourceSheet = workbook.worksheets.getItem(name);
  if (!sourceSheet || sourceSheet.isNullObject) throw invalid(`PivotTable ${pivot.name} references missing source worksheet ${name}.`);
  return sourceSheet;
}

function targetReference(pivot, values) {
  if (!values.length || !values[0]?.length) throw invalid(`PivotTable ${pivot.name} produced no cached output.`);
  if (values.some((row) => row.length !== values[0].length)) throw invalid(`PivotTable ${pivot.name} produced a ragged cached output.`);
  const anchor = parseRange(pivot.targetRange?.address, `PivotTable ${pivot.name} targetRange`);
  const bounds = {
    left: anchor.left,
    top: anchor.top,
    right: anchor.left + values[0].length - 1,
    bottom: anchor.top + values.length - 1,
  };
  if (bounds.right > 16_383 || bounds.bottom > 1_048_575) throw invalid(`PivotTable ${pivot.name} cached output exceeds the XLSX worksheet grid.`);
  if ((anchor.colCount > 1 || anchor.rowCount > 1) && (anchor.right !== bounds.right || anchor.bottom !== bounds.bottom)) {
    throw invalid(`PivotTable ${pivot.name} targetRange must be a single anchor cell or exactly match its ${values.length}x${values[0].length} cached output.`);
  }
  return rangeAddress(bounds);
}

function assertBoundedProfile(workbook, targetSheet, pivot) {
  if (!String(pivot.name || "").trim() || String(pivot.name).length > 255) throw invalid("PivotTable names must contain 1 through 255 characters.");
  const normalizedName = String(pivot.name).toLocaleLowerCase("en-US");
  const duplicates = workbook.worksheets.items
    .flatMap((sheet) => sheet.pivotTables.items)
    .filter((item) => String(item.name).toLocaleLowerCase("en-US") === normalizedName);
  if (duplicates.length !== 1) throw invalid(`PivotTable name ${pivot.name} must be unique across the workbook.`);
  if (pivot.targetRange?.sheetName && pivot.targetRange.sheetName !== targetSheet.name) throw invalid(`PivotTable ${pivot.name} targetRange must belong to worksheet ${targetSheet.name}.`);
  const sourceSheet = sourceSheetFor(workbook, targetSheet, pivot);
  const sourceBounds = parseRange(pivot.sourceRange?.address, `PivotTable ${pivot.name} sourceRange`);
  if (sourceBounds.rowCount < 2 || sourceBounds.colCount < 2) throw invalid(`PivotTable ${pivot.name} sourceRange must include a header row and at least one data row across two or more columns.`);
  const matrix = pivot.sourceValues();
  const headers = (matrix[0] || []).map((value) => String(value ?? "").trim());
  if (headers.length !== sourceBounds.colCount || headers.some((header) => !header) || new Set(headers).size !== headers.length) {
    throw invalid(`PivotTable ${pivot.name} source headers must be non-empty and unique.`);
  }
  if (pivot.rowFields.length !== 1) throw invalid(`PivotTable ${pivot.name} requires exactly one row field in the first native profile.`, "unsupported_spreadsheet_pivot_profile");
  if (pivot.columnFields.length > 1) throw invalid(`PivotTable ${pivot.name} supports at most one column field in the first native profile.`, "unsupported_spreadsheet_pivot_profile");
  if (!pivot.valueFields.length || pivot.valueFields.length > MAX_NATIVE_VALUE_FIELDS) {
    throw invalid(`PivotTable ${pivot.name} requires 1 through ${MAX_NATIVE_VALUE_FIELDS} value fields in the bounded native profile.`, "unsupported_spreadsheet_pivot_profile");
  }
  if (pivot.groupFields.length || pivot.calculatedFields.length) {
    throw invalid(`PivotTable ${pivot.name} grouping and calculated fields remain model/preview-only and cannot yet be authored as native SpreadsheetML.`, "unsupported_spreadsheet_pivot_profile");
  }
  if (pivot.filters.length > pivot.rowFields.length + pivot.columnFields.length || pivot.filters.some((filter) => filter.type)) {
    throw invalid(`PivotTable ${pivot.name} supports only exact include/exclude item filters on its native row or column axis.`, "unsupported_spreadsheet_pivot_filter");
  }
  for (const value of pivot.valueFields) {
    if (!TO_WIRE_AGGREGATION.has(value.summarizeBy)) throw invalid(`PivotTable ${pivot.name} uses unsupported aggregation ${value.summarizeBy}.`, "unsupported_spreadsheet_pivot_profile");
  }
  for (const field of [...pivot.rowFields, ...pivot.columnFields, ...pivot.valueFields.map((value) => value.field)]) {
    if (!headers.includes(field)) throw invalid(`PivotTable ${pivot.name} field ${field} is not present in its source headers.`);
  }
  wireItemFilters(pivot);
  const headerIndexes = new Map(headers.map((header, index) => [header, index]));
  if (pivot.filters.length && !matrix.slice(1).some((row) => pivot.filters.every((filter) => pivotItemVisible(pivot.filters, filter.field, row[headerIndexes.get(filter.field)])))) {
    throw invalid(`PivotTable ${pivot.name} item filters hide every source row.`, "unsupported_spreadsheet_pivot_filter");
  }
  return { sourceSheet, sourceBounds, matrix };
}

function wireRefreshPolicy(policy = {}) {
  return {
    refreshOnLoad: policy.refreshOnLoad !== false,
    saveData: policy.saveData !== false,
    enableRefresh: policy.enableRefresh !== false,
    invalid: Boolean(policy.invalid),
    missingItemsLimit: Number(policy.missingItemsLimit || 0),
    refreshedBy: policy.refreshedBy || "",
    refreshedDateIso: policy.refreshedDateIso || "",
  };
}

function publicRefreshPolicy(policy = {}) {
  return {
    refreshOnLoad: policy.refreshOnLoad,
    saveData: policy.saveData,
    enableRefresh: policy.enableRefresh,
    invalid: policy.invalid,
    missingItemsLimit: policy.missingItemsLimit,
    refreshedBy: policy.refreshedBy || undefined,
    refreshedDateIso: policy.refreshedDateIso || undefined,
  };
}

function pivotCell(row, column, value) {
  const wireValue = typeof value === "number"
    ? { case: "numberValue", value }
    : typeof value === "boolean"
      ? { case: "boolValue", value }
      : { case: "stringValue", value: String(value ?? "") };
  return { row, column, formula: "", numberFormatCode: "", style: {}, value: wireValue };
}

function appendCachedOutput(cells, pivot, reference, values) {
  const bounds = parseRange(reference, `PivotTable ${pivot.name} cached output`);
  const existing = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  for (let rowOffset = 0; rowOffset < values.length; rowOffset++) {
    for (let columnOffset = 0; columnOffset < values[rowOffset].length; columnOffset++) {
      const row = bounds.top + rowOffset;
      const column = bounds.left + columnOffset;
      const key = `${row}:${column}`;
      const cached = pivotCell(row, column, values[rowOffset][columnOffset]);
      if (existing.has(key)) {
        const target = existing.get(key);
        if (target.formula || target.value?.case) throw invalid(`PivotTable ${pivot.name} cached output overlaps existing worksheet cell ${columnLabel(column)}${row + 1}.`, "spreadsheet_pivot_output_collision");
        target.value = cached.value;
      } else {
        existing.set(key, cached);
        cells.push(cached);
      }
    }
  }
}

function wireSourceFreePivot(workbook, sheet, pivot, cells) {
  const { sourceSheet } = assertBoundedProfile(workbook, sheet, pivot);
  const values = pivot.computedValues();
  const target = targetReference(pivot, values);
  appendCachedOutput(cells, pivot, target, values);
  return {
    id: pivot.id,
    name: pivot.name,
    sourceWorksheetId: sourceSheet.id,
    sourceReference: rangeAddress(parseRange(pivot.sourceRange.address, `PivotTable ${pivot.name} sourceRange`)),
    targetReference: target,
    rowFields: [...pivot.rowFields],
    columnFields: [...pivot.columnFields],
    valueFields: pivot.valueFields.map((value) => ({
      field: value.field,
      name: value.name || "",
      aggregation: TO_WIRE_AGGREGATION.get(value.summarizeBy),
    })),
    rowGrandTotals: Boolean(pivot.rowGrandTotals),
    columnGrandTotals: Boolean(pivot.columnGrandTotals),
    refreshPolicy: wireRefreshPolicy(pivot.refreshPolicy),
    itemFilters: wireItemFilters(pivot),
  };
}

export function wireWorksheetPivots(workbook, sheet, state, cells) {
  const remaining = new Set(sheet.pivotTables.items);
  const output = [];
  for (const slot of state?.slots || []) {
    if (!slot.pivot) {
      output.push(slot.wire);
      continue;
    }
    if (!remaining.delete(slot.pivot)) throw invalid(`Worksheet ${sheet.name} cannot remove imported PivotTable ${slot.pivot.name}.`, "unsupported_spreadsheet_pivot_edit");
    if (pivotSnapshot(slot.pivot) !== slot.publicSnapshot ||
        snapshot(slot.pivot.sourceValues()) !== slot.sourceValuesSnapshot ||
        snapshot(outputMatrix(sheet, slot.wire.targetReference)) !== slot.outputValuesSnapshot) {
      throw invalid(`Imported PivotTable ${slot.pivot.name}, its source data, and its cached output are read-only in the first OpenChestnut native profile.`, "unsupported_spreadsheet_pivot_edit");
    }
    output.push(slot.wire);
  }
  if (state && remaining.size) throw invalid(`Source-bound worksheet ${sheet.name} cannot add PivotTables in the first OpenChestnut native profile.`, "unsupported_spreadsheet_pivot_edit");
  for (const pivot of remaining) output.push(wireSourceFreePivot(workbook, sheet, pivot, cells));
  return output;
}

export function hydrateWorkbookPivots(workbook, sourceWorksheets) {
  const byId = new Map(workbook.worksheets.items.map((sheet) => [sheet.id, sheet]));
  const pivotsBySheet = new Map();
  for (const sourceSheet of sourceWorksheets) {
    const sheet = byId.get(sourceSheet.id);
    if (!sheet) continue;
    const slots = [];
    for (const wire of sourceSheet.pivotTables || []) {
      const sourceSheetModel = byId.get(wire.sourceWorksheetId);
      const valueFields = (wire.valueFields || []).map((value) => ({
        field: value.field,
        summarizeBy: FROM_WIRE_AGGREGATION.get(value.aggregation),
        name: value.name || undefined,
      }));
      let filters;
      try {
        filters = publicItemFilters(wire.itemFilters, `PivotTable ${wire.name || wire.id || "unknown"}`);
      } catch {
        slots.push({ wire });
        continue;
      }
      if (!sourceSheetModel || !wire.name || !wire.sourceReference || !wire.targetReference || wire.rowFields?.length !== 1 ||
          wire.columnFields?.length > 1 || !valueFields.length || valueFields.length > MAX_NATIVE_VALUE_FIELDS ||
          valueFields.some((value) => !value.field || !value.summarizeBy) || filters.some((filter) => ![...wire.rowFields, ...wire.columnFields].includes(filter.field))) {
        slots.push({ wire });
        continue;
      }
      const pivot = sheet.pivotTables.add({
        id: wire.id,
        name: wire.name,
        sourceRange: `${sourceSheetModel.name}!${wire.sourceReference}`,
        targetRange: wire.targetReference,
        rowFields: [...wire.rowFields],
        columnFields: [...wire.columnFields],
        valueFields,
        filters,
        rowGrandTotals: wire.rowGrandTotals,
        columnGrandTotals: wire.columnGrandTotals,
        refreshPolicy: publicRefreshPolicy(wire.refreshPolicy),
      });
      slots.push({
        wire,
        pivot,
        publicSnapshot: pivotSnapshot(pivot),
        sourceValuesSnapshot: snapshot(pivot.sourceValues()),
        outputValuesSnapshot: snapshot(outputMatrix(sheet, wire.targetReference)),
      });
    }
    pivotsBySheet.set(sheet.id, { slots });
  }
  return pivotsBySheet;
}

export function worksheetHasSourceBoundPivots(state) {
  return Boolean(state?.slots?.some((slot) => slot.wire?.source));
}
