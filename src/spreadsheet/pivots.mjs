import { evaluatePivotFormula, normalizeCalculatedFields } from "./pivot-formulas.mjs";
import { normalizePivotFilters, pivotItemKey, pivotItemVisible } from "./pivot-filters.mjs";
import { normalizePivotGroupFields, pivotGroupValue, projectPivotMatrix } from "./pivot-groups.mjs";

const AGGREGATIONS = new Set(["sum", "count", "average", "min", "max"]);

function stringFields(value, label) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError(`${label} must be an array.`);
  const fields = value.map((field) => String(field || "").trim());
  if (fields.some((field) => !field)) throw new TypeError(`${label} entries must not be empty.`);
  if (new Set(fields).size !== fields.length) throw new Error(`${label} must not contain duplicate fields.`);
  return fields;
}

function valueFields(value) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("PivotTable valueFields must be an array.");
  return value.map((entry, index) => {
    const field = typeof entry === "string" ? entry : entry?.field || entry?.name;
    if (!String(field || "").trim()) throw new TypeError(`PivotTable valueFields[${index}] requires field.`);
    const rawAggregation = String(typeof entry === "string" ? "sum" : entry.summarizeBy || entry.aggregation || "sum").toLowerCase();
    const summarizeBy = rawAggregation === "avg" ? "average" : rawAggregation;
    if (!AGGREGATIONS.has(summarizeBy)) throw new TypeError(`PivotTable valueFields[${index}] summarizeBy must be sum, count, average, min, or max.`);
    return { ...(typeof entry === "object" ? entry : {}), field: String(field), summarizeBy, name: typeof entry === "object" && entry.name ? String(entry.name) : undefined };
  });
}


function booleanOption(value, fallback, label) {
  if (value == null) return fallback;
  if (typeof value !== "boolean") throw new TypeError(`${label} must be a boolean.`);
  return value;
}

function refreshPolicy(config = {}) {
  const nested = config.refreshPolicy || config.refresh || {};
  if (nested == null || typeof nested !== "object" || Array.isArray(nested)) throw new TypeError("PivotTable refreshPolicy must be an object.");
  const read = (key) => Object.hasOwn(config, key) ? config[key] : nested[key];
  const missingItemsLimit = read("missingItemsLimit") ?? 0;
  if (!Number.isInteger(Number(missingItemsLimit)) || Number(missingItemsLimit) < 0 || Number(missingItemsLimit) > 4_294_967_295) throw new RangeError("PivotTable missingItemsLimit must be an unsigned 32-bit integer.");
  const refreshedBy = read("refreshedBy");
  const refreshedDateIso = read("refreshedDateIso");
  if (refreshedBy != null && !String(refreshedBy).trim()) throw new TypeError("PivotTable refreshedBy must not be empty.");
  if (refreshedDateIso != null && (!/^\d{4,}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?$/.test(String(refreshedDateIso)) || Number.isNaN(Date.parse(refreshedDateIso)))) throw new TypeError("PivotTable refreshedDateIso must be an XML date-time string.");
  return {
    refreshOnLoad: booleanOption(read("refreshOnLoad"), true, "PivotTable refreshOnLoad"),
    saveData: booleanOption(read("saveData"), true, "PivotTable saveData"),
    enableRefresh: booleanOption(read("enableRefresh"), true, "PivotTable enableRefresh"),
    invalid: booleanOption(read("invalid"), false, "PivotTable invalid"),
    missingItemsLimit: Number(missingItemsLimit),
    refreshedBy: refreshedBy == null ? undefined : String(refreshedBy),
    refreshedDateIso: refreshedDateIso == null ? undefined : String(refreshedDateIso),
  };
}

export function normalizePivotConfig(config = {}, headers = []) {
  const rowFields = stringFields(config.rowFields || config.rows, "PivotTable rowFields");
  const columnFields = stringFields(config.columnFields || config.columns, "PivotTable columnFields");
  const duplicateAxisField = rowFields.find((field) => columnFields.includes(field));
  if (duplicateAxisField) throw new Error(`PivotTable field ${duplicateAxisField} cannot be both a row and column field.`);
  const groupFields = normalizePivotGroupFields(config.groupFields || config.groups, headers.map(String), config.allowUnsupportedGroupFields === true);
  const normalizedCalculatedFields = normalizeCalculatedFields(config.calculatedFields, headers.map(String), config.allowUnsupportedCalculatedFields === true);
  const groupNames = new Set(groupFields.map((field) => field.name));
  const calculatedGroupCollision = normalizedCalculatedFields.find((field) => groupNames.has(field.name));
  if (calculatedGroupCollision) throw new Error(`PivotTable calculated field name ${calculatedGroupCollision.name} must not replace a group field.`);
  const normalizedValueFields = valueFields(config.valueFields || config.values);
  const filters = normalizePivotFilters(config.filters, new Set([...rowFields, ...columnFields]));
  const knownHeaders = new Set(headers.map(String));
  const knownAxisFields = new Set([...knownHeaders, ...groupNames]);
  const knownValueFields = new Set([...knownAxisFields, ...normalizedCalculatedFields.map((field) => field.name)]);
  if (config.validateSource !== false && knownHeaders.size) {
    for (const field of [...rowFields, ...columnFields]) if (!knownAxisFields.has(field)) throw new Error(`PivotTable field ${field} is not present in the source headers or group fields.`);
    for (const field of normalizedValueFields.map((entry) => entry.field)) if (!knownValueFields.has(field)) throw new Error(`PivotTable field ${field} is not present in the source or calculated fields.`);
    for (const filter of filters) {
      if (filter.type && groupNames.has(filter.field)) throw new Error(`PivotTable date filter ${filter.field} must target a source date field, not a grouped label.`);
      const group = groupFields.find((field) => field.name === filter.field);
      const sourceValues = group
        ? config.sourceValues?.[group.sourceField]?.map((value) => pivotGroupValue(group, value, config.dateSystem))
        : config.sourceValues?.[filter.field];
      if (!sourceValues) continue;
      const knownItems = new Set(sourceValues.map(pivotItemKey));
      for (const value of filter.include || filter.exclude || []) if (!knownItems.has(pivotItemKey(value))) throw new Error(`PivotTable filter ${filter.field} references unknown item ${String(value)}.`);
    }
  }
  return { rowFields, columnFields, valueFields: normalizedValueFields, groupFields, calculatedFields: normalizedCalculatedFields, filters, refreshPolicy: refreshPolicy(config) };
}

export function pivotValueLabel(valueField = {}) {
  const summarizeBy = valueField.summarizeBy || valueField.aggregation || "sum";
  return valueField.name || `${summarizeBy} of ${valueField.field || valueField.name || "Value"}`;
}

function summarize(values = [], summarizeBy = "sum") {
  const numbers = values.map((value) => Number(value)).filter(Number.isFinite);
  if (summarizeBy === "count") return values.filter((value) => value != null && value !== "").length;
  if (!numbers.length) return 0;
  if (summarizeBy === "average") return numbers.reduce((sum, value) => sum + value, 0) / numbers.length;
  if (summarizeBy === "min") return Math.min(...numbers);
  if (summarizeBy === "max") return Math.max(...numbers);
  return numbers.reduce((sum, value) => sum + value, 0);
}

export function computePivotValues(matrix = [], config = {}) {
  if (!matrix.length) return [];
  const sourceHeaders = config.sourceFields?.length ? config.sourceFields.map(String) : matrix[0].map((value) => String(value ?? ""));
  const projection = projectPivotMatrix(matrix, config.groupFields || [], config.dateSystem, sourceHeaders);
  const headers = projection.headers;
  const rowFields = config.rowFields || [];
  const columnFields = config.columnFields || [];
  const rowIndexes = rowFields.map((field) => headers.indexOf(field));
  const columnIndexes = columnFields.map((field) => headers.indexOf(field));
  const valueFields = config.valueFields?.length ? config.valueFields : sourceHeaders.slice(1, 2).map((field) => ({ field, summarizeBy: "sum" }));
  const calculated = new Map((config.calculatedFields || []).map((field) => [field.name, field]));
  const valueIndexes = valueFields.map((field) => ({ field, index: headers.indexOf(field.field), calculated: calculated.get(field.field) })).filter((entry) => entry.index >= 0 || entry.calculated);
  const rows = projection.rows.filter((row) => (config.filters || []).every((filter) => pivotItemVisible(config.filters, filter.field, row[headers.indexOf(filter.field)], config.dateSystem)));
  const rowGroups = new Map();
  const columnGroups = new Map();
  for (const row of rows) {
    const rowValues = rowIndexes.length ? rowIndexes.map((index) => row[index]) : ["(all)"];
    const columnValues = columnIndexes.length ? columnIndexes.map((index) => row[index]) : [];
    const rowKey = JSON.stringify(rowValues.map(pivotItemKey));
    const columnKey = JSON.stringify(columnValues.map(pivotItemKey));
    if (!rowGroups.has(rowKey)) rowGroups.set(rowKey, { values: rowValues, columns: new Map() });
    const rowGroup = rowGroups.get(rowKey);
    if (!rowGroup.columns.has(columnKey)) rowGroup.columns.set(columnKey, []);
    rowGroup.columns.get(columnKey).push(row);
    if (!columnGroups.has(columnKey)) columnGroups.set(columnKey, columnValues);
  }
  if (!columnIndexes.length) columnGroups.set("[]", []);
  const columnEntries = [...columnGroups.entries()];
  const columnHeaders = columnEntries.flatMap(([, values]) => valueIndexes.map(({ field }) => {
    const prefix = values.map((value) => String(value ?? "")).join(" / ");
    return prefix ? (valueIndexes.length === 1 ? prefix : `${prefix} — ${pivotValueLabel(field)}`) : pivotValueLabel(field);
  }));
  const header = [...(rowFields.length ? rowFields : ["Group"]), ...columnHeaders];
  const outputRows = [...rowGroups.values()].map((group) => [
    ...group.values,
    ...columnEntries.flatMap(([key]) => valueIndexes.map(({ field, index, calculated: calculatedField }) => {
      const groupedRows = group.columns.get(key) || [];
      if (!calculatedField) return summarize(groupedRows.map((row) => row[index]), field.summarizeBy);
      if (calculatedField.supported === false) return "#NAME?";
      const aggregates = Object.fromEntries(sourceHeaders.map((header, fieldIndex) => [header, summarize(groupedRows.map((row) => row[fieldIndex]), "sum")]));
      return evaluatePivotFormula(calculatedField.formula, aggregates, sourceHeaders);
    })),
  ]);
  return [header, ...outputRows];
}
