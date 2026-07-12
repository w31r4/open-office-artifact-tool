import { pivotDateParts } from "./pivot-dates.mjs";

export const PIVOT_CALENDAR_GROUP_TYPES = new Set(["years", "quarters", "months"]);

const GROUP_ORDER = new Map([["months", 0], ["quarters", 1], ["years", 2]]);
const MONTH_LABELS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

function groupEntries(value) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("PivotTable groupFields must be an array.");
  if (value.length > 128) throw new RangeError("PivotTable groupFields exceeds 128 fields.");
  return value;
}

export function normalizePivotGroupFields(value, sourceFields, allowUnsupported = false) {
  const sourceNames = new Set(sourceFields.map(String));
  const names = new Set(sourceNames);
  const groups = groupEntries(value).map((entry, index) => {
    const name = String(entry?.name || "").trim();
    const sourceField = String(entry?.sourceField || entry?.field || "").trim();
    const groupBy = String(entry?.groupBy || "").trim();
    if (!name) throw new TypeError(`PivotTable groupFields[${index}] requires name.`);
    if (!sourceNames.has(sourceField)) throw new Error(`PivotTable group field ${name} references unknown source field ${sourceField}.`);
    const supported = PIVOT_CALENDAR_GROUP_TYPES.has(groupBy);
    if (!supported && !allowUnsupported) throw new TypeError(`PivotTable group field ${name} groupBy must be years, quarters, or months.`);
    if (names.has(name)) throw new Error(`PivotTable group field name ${name} must be unique and must not replace a source field.`);
    names.add(name);
    return { name, sourceField, groupBy, ...(entry?.parent ? { parent: String(entry.parent) } : {}), ...(entry?.items ? { items: [...entry.items] } : {}), ...(entry?.range ? { range: { ...entry.range } } : {}), ...(supported ? {} : { supported: false, error: `Unsupported PivotTable groupBy ${groupBy}.` }) };
  });
  for (const sourceField of sourceNames) {
    const hierarchy = groups.filter((group) => group.sourceField === sourceField && group.supported !== false).sort((left, right) => GROUP_ORDER.get(left.groupBy) - GROUP_ORDER.get(right.groupBy));
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
  const parts = pivotDateParts(value, dateSystem);
  if (!parts) return null;
  if (group.groupBy === "years") return String(parts.year);
  if (group.groupBy === "quarters") return `Q${Math.ceil(parts.month / 3)}`;
  return MONTH_LABELS[parts.month - 1];
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
  if (group.groupBy === "months") return MONTH_LABELS.filter((value) => unique.includes(value));
  if (group.groupBy === "quarters") return ["Q1", "Q2", "Q3", "Q4"].filter((value) => unique.includes(value));
  return unique.sort((left, right) => Number(left) - Number(right));
}
