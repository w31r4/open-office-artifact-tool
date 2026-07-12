import { normalizePivotDate, pivotDateKey } from "./pivot-dates.mjs";

export const PIVOT_DATE_FILTER_TYPES = new Set([
  "dateEqual", "dateNotEqual", "dateOlderThan", "dateOlderThanOrEqual",
  "dateNewerThan", "dateNewerThanOrEqual", "dateBetween", "dateNotBetween",
]);

export function pivotItemKey(value) {
  if (value instanceof Date) return `date:${value.toISOString()}`;
  return `${value === null ? "null" : typeof value}:${String(value)}`;
}

function uniqueItems(values = []) {
  const seen = new Set();
  return values.filter((value) => {
    const key = pivotItemKey(value);
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function filterEntries(value) {
  if (value == null) return [];
  if (Array.isArray(value)) return value;
  if (typeof value !== "object") throw new TypeError("PivotTable filters must be an object or array.");
  return Object.entries(value).map(([field, filter]) => Array.isArray(filter) ? { field, include: filter } : { field, ...(filter || {}) });
}

function canonicalDate(value, label) {
  return normalizePivotDate(value, label);
}

function dateFilter(filter, field) {
  const type = String(filter.type || "");
  if (!PIVOT_DATE_FILTER_TYPES.has(type)) throw new TypeError(`PivotTable filter ${field} has unsupported type ${type}.`);
  if (filter.useWholeDay === false) throw new TypeError(`PivotTable date filter ${field} currently requires useWholeDay=true.`);
  if (Object.hasOwn(filter, "include") || Object.hasOwn(filter, "exclude")) throw new Error(`PivotTable date filter ${field} cannot combine type with include or exclude.`);
  const value1 = canonicalDate(filter.value1 ?? filter.start ?? filter.value, `PivotTable date filter ${field} value1`);
  const between = type === "dateBetween" || type === "dateNotBetween";
  const value2 = between ? canonicalDate(filter.value2 ?? filter.end, `PivotTable date filter ${field} value2`) : undefined;
  if (between && value1 > value2) throw new RangeError(`PivotTable date filter ${field} value1 must not be after value2.`);
  return { field, type, value1, value2, useWholeDay: true };
}

export function normalizePivotFilters(value, axisFields) {
  const filters = filterEntries(value).map((filter, index) => {
    const field = String(filter?.field || filter?.name || "").trim();
    if (!field) throw new TypeError(`PivotTable filters[${index}] requires field.`);
    if (!axisFields.has(field)) throw new Error(`PivotTable filter field ${field} must also be a row or column field.`);
    if (filter.type != null) return dateFilter(filter, field);
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

export function pivotItemVisible(filters = [], field, value, dateSystem = "1900") {
  const filter = filters.find((entry) => entry.field === field);
  if (!filter) return true;
  if (PIVOT_DATE_FILTER_TYPES.has(filter.type)) {
    const current = pivotDateKey(value, dateSystem);
    if (!current) return false;
    if (filter.type === "dateEqual") return current === filter.value1;
    if (filter.type === "dateNotEqual") return current !== filter.value1;
    if (filter.type === "dateOlderThan") return current < filter.value1;
    if (filter.type === "dateOlderThanOrEqual") return current <= filter.value1;
    if (filter.type === "dateNewerThan") return current > filter.value1;
    if (filter.type === "dateNewerThanOrEqual") return current >= filter.value1;
    const between = current >= filter.value1 && current <= filter.value2;
    return filter.type === "dateBetween" ? between : !between;
  }
  const key = pivotItemKey(value);
  if (filter.include) return filter.include.some((item) => pivotItemKey(item) === key);
  return !filter.exclude.some((item) => pivotItemKey(item) === key);
}
