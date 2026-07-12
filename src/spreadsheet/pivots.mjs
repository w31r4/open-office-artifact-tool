const AGGREGATIONS = new Set(["sum", "count", "average", "min", "max"]);

function itemKey(value) {
  if (value instanceof Date) return `date:${value.toISOString()}`;
  return `${value === null ? "null" : typeof value}:${String(value)}`;
}

function uniqueItems(values = []) {
  const seen = new Set();
  return values.filter((value) => {
    const key = itemKey(value);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

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

function filterEntries(value) {
  if (value == null) return [];
  if (Array.isArray(value)) return value;
  if (typeof value !== "object") throw new TypeError("PivotTable filters must be an object or array.");
  return Object.entries(value).map(([field, filter]) => Array.isArray(filter) ? { field, include: filter } : { field, ...(filter || {}) });
}

function pivotFilters(value, axisFields) {
  const filters = filterEntries(value).map((filter, index) => {
    const field = String(filter?.field || filter?.name || "").trim();
    if (!field) throw new TypeError(`PivotTable filters[${index}] requires field.`);
    if (!axisFields.has(field)) throw new Error(`PivotTable filter field ${field} must also be a row or column field.`);
    const hasInclude = Object.hasOwn(filter, "include");
    const hasExclude = Object.hasOwn(filter, "exclude");
    if (hasInclude === hasExclude) throw new Error(`PivotTable filter ${field} requires exactly one of include or exclude.`);
    const mode = hasInclude ? "include" : "exclude";
    const rawItems = filter[mode];
    if (!Array.isArray(rawItems) || rawItems.length === 0) throw new TypeError(`PivotTable filter ${field} ${mode} must be a non-empty array.`);
    return { field, [mode]: uniqueItems(rawItems) };
  });
  if (new Set(filters.map((filter) => filter.field)).size !== filters.length) throw new Error("PivotTable filters must not contain duplicate fields.");
  return filters;
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
  const normalizedValueFields = valueFields(config.valueFields || config.values);
  const filters = pivotFilters(config.filters, new Set([...rowFields, ...columnFields]));
  const knownHeaders = new Set(headers.map(String));
  if (config.validateSource !== false && knownHeaders.size) {
    for (const field of [...rowFields, ...columnFields, ...normalizedValueFields.map((entry) => entry.field)]) if (!knownHeaders.has(field)) throw new Error(`PivotTable field ${field} is not present in the source headers.`);
    for (const filter of filters) {
      const sourceValues = config.sourceValues?.[filter.field];
      if (!sourceValues) continue;
      const knownItems = new Set(sourceValues.map(itemKey));
      for (const value of filter.include || filter.exclude || []) if (!knownItems.has(itemKey(value))) throw new Error(`PivotTable filter ${filter.field} references unknown item ${String(value)}.`);
    }
  }
  return { rowFields, columnFields, valueFields: normalizedValueFields, filters, refreshPolicy: refreshPolicy(config) };
}

export function pivotValueLabel(valueField = {}) {
  const summarizeBy = valueField.summarizeBy || valueField.aggregation || "sum";
  return valueField.name || `${summarizeBy} of ${valueField.field || valueField.name || "Value"}`;
}

export function pivotItemVisible(filters = [], field, value) {
  const filter = filters.find((entry) => entry.field === field);
  if (!filter) return true;
  const key = itemKey(value);
  if (filter.include) return filter.include.some((item) => itemKey(item) === key);
  return !filter.exclude.some((item) => itemKey(item) === key);
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
  const headers = config.sourceFields?.length ? config.sourceFields.map(String) : matrix[0].map((value) => String(value ?? ""));
  const rowFields = config.rowFields || [];
  const columnFields = config.columnFields || [];
  const rowIndexes = rowFields.map((field) => headers.indexOf(field));
  const columnIndexes = columnFields.map((field) => headers.indexOf(field));
  const valueFields = config.valueFields?.length ? config.valueFields : headers.slice(1, 2).map((field) => ({ field, summarizeBy: "sum" }));
  const valueIndexes = valueFields.map((field) => ({ field, index: headers.indexOf(field.field) })).filter((entry) => entry.index >= 0);
  const rows = matrix.slice(1).filter((row) => (config.filters || []).every((filter) => pivotItemVisible(config.filters, filter.field, row[headers.indexOf(filter.field)])));
  const rowGroups = new Map();
  const columnGroups = new Map();
  for (const row of rows) {
    const rowValues = rowIndexes.length ? rowIndexes.map((index) => row[index]) : ["(all)"];
    const columnValues = columnIndexes.length ? columnIndexes.map((index) => row[index]) : [];
    const rowKey = JSON.stringify(rowValues.map(itemKey));
    const columnKey = JSON.stringify(columnValues.map(itemKey));
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
    ...columnEntries.flatMap(([key]) => valueIndexes.map(({ field, index }) => summarize((group.columns.get(key) || []).map((row) => row[index]), field.summarizeBy))),
  ]);
  return [header, ...outputRows];
}
