import { normalizePivotDate, normalizePivotDateTime, pivotDateKey, pivotDateTimeKey } from "./pivot-dates.mjs";

export const PIVOT_ABSOLUTE_DATE_FILTER_TYPES = new Set([
  "dateEqual", "dateNotEqual", "dateOlderThan", "dateOlderThanOrEqual",
  "dateNewerThan", "dateNewerThanOrEqual", "dateBetween", "dateNotBetween",
]);
export const PIVOT_RELATIVE_DATE_FILTER_TYPES = new Set([
  "yesterday", "today", "tomorrow",
  "lastWeek", "thisWeek", "nextWeek",
  "lastMonth", "thisMonth", "nextMonth",
  "lastQuarter", "thisQuarter", "nextQuarter",
  "lastYear", "thisYear", "nextYear", "yearToDate",
]);
export const PIVOT_DATE_FILTER_TYPES = new Set([...PIVOT_ABSOLUTE_DATE_FILTER_TYPES, ...PIVOT_RELATIVE_DATE_FILTER_TYPES]);

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

function canonicalDate(value, label, useWholeDay = true) {
  return useWholeDay ? normalizePivotDate(value, label) : normalizePivotDateTime(value, label);
}

function dateFilter(filter, field) {
  const type = String(filter.type || "");
  if (!PIVOT_DATE_FILTER_TYPES.has(type)) throw new TypeError(`PivotTable filter ${field} has unsupported type ${type}.`);
  const useWholeDay = filter.useWholeDay !== false;
  if (Object.hasOwn(filter, "include") || Object.hasOwn(filter, "exclude")) throw new Error(`PivotTable date filter ${field} cannot combine type with include or exclude.`);
  if (PIVOT_RELATIVE_DATE_FILTER_TYPES.has(type)) {
    if (!useWholeDay) throw new TypeError(`PivotTable relative date filter ${field} requires useWholeDay=true; sub-day thresholds apply only to absolute date filters.`);
    if (["value", "value1", "value2", "start", "end"].some((key) => Object.hasOwn(filter, key))) throw new Error(`PivotTable relative date filter ${field} cannot define absolute date values.`);
    const asOf = canonicalDate(filter.asOf ?? new Date(), `PivotTable relative date filter ${field} asOf`);
    return { field, type, asOf, useWholeDay: true };
  }
  const value1 = canonicalDate(filter.value1 ?? filter.start ?? filter.value, `PivotTable date filter ${field} value1`, useWholeDay);
  const between = type === "dateBetween" || type === "dateNotBetween";
  const value2 = between ? canonicalDate(filter.value2 ?? filter.end, `PivotTable date filter ${field} value2`, useWholeDay) : undefined;
  if (between && value1 > value2) throw new RangeError(`PivotTable date filter ${field} value1 must not be after value2.`);
  return { field, type, value1, value2, useWholeDay };
}

function dateKey(date) {
  return date.toISOString().slice(0, 10);
}

function relativeDateBounds(type, asOf = dateKey(new Date())) {
  const current = new Date(`${asOf}T00:00:00Z`);
  const day = (offset) => new Date(current.valueOf() + offset * 86_400_000);
  const month = (offset) => new Date(Date.UTC(current.getUTCFullYear(), current.getUTCMonth() + offset, 1));
  const monthBounds = (offset) => [month(offset), new Date(month(offset + 1).valueOf() - 86_400_000)];
  const weekStart = day(-((current.getUTCDay() + 6) % 7));
  const weekBounds = (offset) => {
    const start = new Date(weekStart.valueOf() + offset * 7 * 86_400_000);
    return [start, new Date(start.valueOf() + 6 * 86_400_000)];
  };
  const quarterStart = Math.floor(current.getUTCMonth() / 3) * 3;
  const quarterBounds = (offset) => {
    const start = new Date(Date.UTC(current.getUTCFullYear(), quarterStart + offset * 3, 1));
    const end = new Date(Date.UTC(current.getUTCFullYear(), quarterStart + (offset + 1) * 3, 1) - 86_400_000);
    return [start, end];
  };
  const yearBounds = (offset) => [new Date(Date.UTC(current.getUTCFullYear() + offset, 0, 1)), new Date(Date.UTC(current.getUTCFullYear() + offset, 11, 31))];
  let bounds;
  if (type === "yesterday") bounds = [day(-1), day(-1)];
  else if (type === "today") bounds = [current, current];
  else if (type === "tomorrow") bounds = [day(1), day(1)];
  else if (type === "lastWeek") bounds = weekBounds(-1);
  else if (type === "thisWeek") bounds = weekBounds(0);
  else if (type === "nextWeek") bounds = weekBounds(1);
  else if (type === "lastMonth") bounds = monthBounds(-1);
  else if (type === "thisMonth") bounds = monthBounds(0);
  else if (type === "nextMonth") bounds = monthBounds(1);
  else if (type === "lastQuarter") bounds = quarterBounds(-1);
  else if (type === "thisQuarter") bounds = quarterBounds(0);
  else if (type === "nextQuarter") bounds = quarterBounds(1);
  else if (type === "lastYear") bounds = yearBounds(-1);
  else if (type === "thisYear") bounds = yearBounds(0);
  else if (type === "nextYear") bounds = yearBounds(1);
  else bounds = [yearBounds(0)[0], current];
  return bounds.map(dateKey);
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
    const current = filter.useWholeDay === false ? pivotDateTimeKey(value, dateSystem) : pivotDateKey(value, dateSystem);
    if (!current) return false;
    if (PIVOT_RELATIVE_DATE_FILTER_TYPES.has(filter.type)) {
      const [start, end] = relativeDateBounds(filter.type, filter.asOf);
      return current >= start && current <= end;
    }
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
