import { pivotDateParts } from "./pivot-dates.mjs";

export const PIVOT_CALENDAR_GROUP_TYPES = new Set(["years", "quarters", "months", "days", "hours", "minutes", "seconds"]);
export const PIVOT_GROUP_TYPES = new Set([...PIVOT_CALENDAR_GROUP_TYPES, "range", "discrete"]);

const GROUP_ORDER = new Map([["seconds", 0], ["minutes", 1], ["hours", 2], ["days", 3], ["months", 4], ["quarters", 5], ["years", 6]]);
const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function groupEntries(value) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("PivotTable groupFields must be an array.");
  if (value.length > 128) throw new RangeError("PivotTable groupFields exceeds 128 fields.");
  return value;
}

function itemKey(value) {
  return `${value === null ? "null" : typeof value}:${String(value)}`;
}

function normalizeNumericRange(entry, name, sourceValues = []) {
  const range = entry?.range && typeof entry.range === "object" && !Array.isArray(entry.range) ? entry.range : {};
  const numbers = sourceValues.filter((value) => value != null && String(value).trim() !== "").map(Number).filter(Number.isFinite);
  const groupInterval = Number(range.groupInterval ?? entry?.groupInterval);
  if (!Number.isFinite(groupInterval) || groupInterval <= 0) throw new RangeError(`PivotTable numeric group field ${name} requires a positive groupInterval.`);
  const autoStart = range.autoStart == null ? range.startNum == null : range.autoStart !== false;
  const autoEnd = range.autoEnd == null ? range.endNum == null : range.autoEnd !== false;
  const startNum = range.startNum == null && autoStart ? Math.min(...numbers) : Number(range.startNum);
  const endNum = range.endNum == null && autoEnd ? Math.max(...numbers) : Number(range.endNum);
  if (!Number.isFinite(startNum) || !Number.isFinite(endNum)) throw new TypeError(`PivotTable numeric group field ${name} requires finite startNum and endNum values or numeric source items.`);
  if (startNum > endNum) throw new RangeError(`PivotTable numeric group field ${name} startNum must not exceed endNum.`);
  const bucketCount = Math.floor((endNum - startNum) / groupInterval) + 1;
  if (bucketCount > 10_000) throw new RangeError(`PivotTable numeric group field ${name} exceeds 10000 buckets.`);
  return { autoStart, autoEnd, startNum, endNum, groupInterval };
}

function normalizeTimeRange(entry, name) {
  const range = entry?.range && typeof entry.range === "object" && !Array.isArray(entry.range) ? entry.range : {};
  const groupInterval = Number(range.groupInterval ?? entry?.groupInterval ?? 1);
  if (!Number.isInteger(groupInterval) || groupInterval < 1 || groupInterval > 32_767) throw new RangeError(`PivotTable time group field ${name} groupInterval must be an integer from 1 to 32767.`);
  return { ...range, groupInterval };
}

function normalizeDiscreteGroups(entry, name, sourceValues = [], allowEmpty = false) {
  if (!Array.isArray(entry?.groups) || (!entry.groups.length && !allowEmpty)) throw new TypeError(`PivotTable discrete group field ${name} requires a non-empty groups array.`);
  if (entry.groups.length > 256) throw new RangeError(`PivotTable discrete group field ${name} exceeds 256 groups.`);
  const sourceKeys = new Set(sourceValues.map(itemKey));
  const usedNames = new Set();
  const usedItems = new Set();
  const groups = entry.groups.map((group, index) => {
    const groupName = String(group?.name || "").trim();
    if (!groupName) throw new TypeError(`PivotTable discrete group field ${name} groups[${index}] requires name.`);
    if (usedNames.has(groupName)) throw new Error(`PivotTable discrete group field ${name} group name ${groupName} must be unique.`);
    usedNames.add(groupName);
    if (!Array.isArray(group?.items) || !group.items.length) throw new TypeError(`PivotTable discrete group field ${name} group ${groupName} requires a non-empty items array.`);
    const items = group.items.map((item) => item);
    for (const item of items) {
      const key = itemKey(item);
      if (usedItems.has(key)) throw new Error(`PivotTable discrete group field ${name} item ${String(item)} must belong to only one group.`);
      if (sourceKeys.size && !sourceKeys.has(key)) throw new Error(`PivotTable discrete group field ${name} references unknown source item ${String(item)}.`);
      usedItems.add(key);
    }
    return { name: groupName, items };
  });
  const ungroupedLabels = new Set(sourceValues.filter((item) => !usedItems.has(itemKey(item))).map(String));
  for (const group of groups) if (ungroupedLabels.has(group.name)) throw new Error(`PivotTable discrete group field ${name} group name ${group.name} conflicts with an ungrouped source item.`);
  return groups;
}

export function normalizePivotGroupFields(value, sourceFields, allowUnsupported = false, sourceValues = {}) {
  const sourceNames = new Set(sourceFields.map(String));
  const names = new Set(sourceNames);
  const groups = groupEntries(value).map((entry, index) => {
    const name = String(entry?.name || "").trim();
    const sourceField = String(entry?.sourceField || entry?.field || "").trim();
    const groupBy = String(entry?.groupBy || "").trim();
    if (!name) throw new TypeError(`PivotTable groupFields[${index}] requires name.`);
    if (!sourceNames.has(sourceField)) throw new Error(`PivotTable group field ${name} references unknown source field ${sourceField}.`);
    const supported = PIVOT_GROUP_TYPES.has(groupBy);
    if (!supported && !allowUnsupported) throw new TypeError(`PivotTable group field ${name} groupBy must be a supported calendar/time level, range, or discrete.`);
    if (names.has(name)) throw new Error(`PivotTable group field name ${name} must be unique and must not replace a source field.`);
    names.add(name);
    const normalized = { name, sourceField, groupBy, ...(entry?.parent ? { parent: String(entry.parent) } : {}), ...(entry?.items ? { items: [...entry.items] } : {}), ...(entry?.range ? { range: { ...entry.range } } : {}) };
    if (groupBy === "range") normalized.range = normalizeNumericRange(entry, name, sourceValues[sourceField] || []);
    if (groupBy === "discrete") normalized.groups = normalizeDiscreteGroups(entry, name, sourceValues[sourceField] || [], allowUnsupported);
    if (["days", "hours", "minutes", "seconds"].includes(groupBy)) normalized.range = normalizeTimeRange(entry, name);
    return { ...normalized, ...(supported ? {} : { supported: false, error: `Unsupported PivotTable groupBy ${groupBy}.` }) };
  });
  for (const sourceField of sourceNames) {
    const hierarchy = groups.filter((group) => group.sourceField === sourceField && PIVOT_CALENDAR_GROUP_TYPES.has(group.groupBy)).sort((left, right) => GROUP_ORDER.get(left.groupBy) - GROUP_ORDER.get(right.groupBy));
    if (new Set(hierarchy.map((group) => group.groupBy)).size !== hierarchy.length) throw new Error(`PivotTable groupFields must not repeat a groupBy level for source field ${sourceField}.`);
    hierarchy.forEach((group, index) => {
      const parent = hierarchy[index + 1]?.name;
      if (parent) group.parent = parent;
      else delete group.parent;
    });
  }
  return groups;
}

export function pivotGroupValue(group, value, dateSystem = "1900") {
  if (group.supported === false) return "#NAME?";
  if (group.groupBy === "discrete") {
    const key = itemKey(value);
    return group.groups.find((entry) => entry.items.some((item) => itemKey(item) === key))?.name ?? value;
  }
  if (group.groupBy === "range") {
    if (value == null || String(value).trim() === "") return null;
    const number = Number(value);
    if (!Number.isFinite(number)) return null;
    const { startNum, endNum, groupInterval } = group.range;
    if (number < startNum) return `<${startNum}`;
    if (number > endNum) return `>${endNum}`;
    const lower = startNum + Math.floor((number - startNum) / groupInterval) * groupInterval;
    const upper = Math.min(endNum, lower + groupInterval - (Number.isInteger(startNum) && Number.isInteger(groupInterval) ? 1 : 0));
    return `${lower}-${upper}`;
  }
  const parts = pivotDateParts(value, dateSystem);
  if (!parts) return null;
  const interval = group.range?.groupInterval || 1;
  const bucket = (current, minimum, maximum, pad = false) => {
    const lower = minimum + Math.floor((current - minimum) / interval) * interval;
    const upper = Math.min(maximum, lower + interval - 1);
    const format = (item) => pad ? String(item).padStart(2, "0") : String(item);
    return interval === 1 ? format(current) : `${format(lower)}-${format(upper)}`;
  };
  if (group.groupBy === "years") return String(parts.year);
  if (group.groupBy === "quarters") return `Q${Math.ceil(parts.month / 3)}`;
  if (group.groupBy === "months") return MONTH_LABELS[parts.month - 1];
  if (group.groupBy === "days") return bucket(parts.day, 1, 31);
  if (group.groupBy === "hours") return bucket(parts.hour, 0, 23, true);
  if (group.groupBy === "minutes") return bucket(parts.minute, 0, 59, true);
  return bucket(parts.second, 0, 59, true);
}

export function projectPivotMatrix(matrix = [], groupFields = [], dateSystem = "1900", sourceFields) {
  if (!matrix.length) return { headers: [], rows: [] };
  const headers = sourceFields?.length ? sourceFields.map(String) : matrix[0].map((value) => String(value ?? ""));
  const indexes = groupFields.map((group) => headers.indexOf(group.sourceField));
  const rows = matrix.slice(1).map((row) => [...row, ...groupFields.map((group, index) => pivotGroupValue(group, row[indexes[index]], dateSystem))]);
  return { headers: [...headers, ...groupFields.map((group) => group.name)], rows };
}

export function pivotGroupItems(matrix, group, dateSystem = "1900", sourceFields) {
  if (group.supported === false) return [...(group.items || [])];
  const projection = projectPivotMatrix(matrix, [group], dateSystem, sourceFields);
  const values = projection.rows.map((row) => row[row.length - 1]).filter((value) => value != null);
  const unique = [...new Set(values)];
  if (group.groupBy === "range") return unique.sort((left, right) => {
    const rank = (value) => String(value).startsWith("<") ? -Infinity : String(value).startsWith(">") ? Infinity : Number.parseFloat(value);
    return rank(left) - rank(right);
  });
  if (group.groupBy === "discrete") return unique;
  if (group.groupBy === "months") return MONTH_LABELS.filter((value) => unique.includes(value));
  if (group.groupBy === "quarters") return ["Q1", "Q2", "Q3", "Q4"].filter((value) => unique.includes(value));
  return unique.sort((left, right) => Number(left) - Number(right));
}
