import { pivotFormulaToOoxml } from "./pivot-formulas.mjs";
import { PIVOT_DATE_FILTER_TYPES, pivotItemVisible } from "./pivot-filters.mjs";
import { pivotDateKey, pivotDateTimeKey } from "./pivot-dates.mjs";
import { PIVOT_CALENDAR_GROUP_TYPES, pivotGroupItems, pivotGroupValue } from "./pivot-groups.mjs";
import { pivotValueLabel } from "./pivots.mjs";

function decodeXml(value) {
  return String(value ?? "")
    .replaceAll("&lt;", "<")
    .replaceAll("&gt;", ">")
    .replaceAll("&quot;", '"')
    .replaceAll("&apos;", "'")
    .replaceAll("&amp;", "&");
}

function attributes(tag = "") {
  return Object.fromEntries([...String(tag).matchAll(/([A-Za-z_][\w:.-]*)\s*=\s*(["'])(.*?)\2/g)].map((match) => [match[1], decodeXml(match[3])]));
}

function tag(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml))?.[0];
}

function body(xml, localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>([\\s\\S]*?)<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml))?.[1] || "";
}

function relationshipId(attrs = {}) {
  return Object.entries(attrs).find(([name]) => /:id$/.test(name))?.[1];
}

function xmlEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function attrEscape(value) {
  return xmlEscape(value).replaceAll('"', "&quot;");
}

function booleanAttribute(value, fallback = false) {
  if (value == null) return fallback;
  return ["1", "true", "on"].includes(String(value).toLowerCase());
}

function elements(xml, localName) {
  const prefix = "(?:[A-Za-z_][\\w.-]*:)?";
  return [...String(xml || "").matchAll(new RegExp(`<${prefix}${localName}\\b[^>]*\\/\\s*>|<${prefix}${localName}\\b[^>]*>[\\s\\S]*?<\\/${prefix}${localName}>`, "g"))].map((match) => ({ xml: match[0], opening: tag(match[0], localName), body: body(match[0], localName) }));
}

function parseSharedItem(itemXml) {
  const opening = /^<[^>]+>/.exec(itemXml)?.[0] || itemXml;
  const attrs = attributes(opening);
  const localName = /^<(?:[A-Za-z_][\w.-]*:)?([A-Za-z]+)/.exec(opening)?.[1];
  if (localName === "m") return null;
  if (localName === "n") return Number(attrs.v);
  if (localName === "b") return booleanAttribute(attrs.v);
  return attrs.v == null ? "" : attrs.v;
}

function sharedItems(cacheFieldXml) {
  const shared = body(cacheFieldXml, "sharedItems");
  return [...shared.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:s|n|b|d|e|m)\b[^>]*\/?\s*>/g)].map((match) => parseSharedItem(match[0]));
}

function cacheFieldItems(cacheFieldXml) {
  const shared = sharedItems(cacheFieldXml);
  if (shared.length) return shared;
  return [...body(cacheFieldXml, "groupItems").matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:s|n|b|d|e|m)\b[^>]*\/?\s*>/g)].map((match) => parseSharedItem(match[0]));
}

export function parseWorkbookPivotCaches(xml = "") {
  const block = body(xml, "pivotCaches");
  return [...block.matchAll(/<(?:[A-Za-z_][\w.-]*:)?pivotCache\b[^>]*\/?>/g)].map((match) => {
    const attrs = attributes(match[0]);
    return { cacheId: Number(attrs.cacheId), relationshipId: relationshipId(attrs) };
  }).filter((item) => Number.isFinite(item.cacheId) && item.relationshipId);
}

export function parsePivotCacheDefinition(xml = "") {
  const rootAttrs = attributes(tag(xml, "pivotCacheDefinition"));
  const sourceAttrs = attributes(tag(body(xml, "cacheSource"), "worksheetSource"));
  const cacheFields = body(xml, "cacheFields");
  const fieldEntries = elements(cacheFields, "cacheField");
  const fieldAttributes = fieldEntries.map((entry, index) => ({ ...attributes(entry.opening), name: attributes(entry.opening).name || `Field${index + 1}` }));
  const fields = fieldAttributes.map((entry) => entry.name);
  const items = fieldEntries.map((entry) => cacheFieldItems(entry.xml));
  const groupFields = fieldEntries.flatMap((entry, index) => {
    const groupOpening = tag(entry.xml, "fieldGroup");
    const rangeOpening = tag(entry.xml, "rangePr");
    const discreteOpening = tag(entry.xml, "discretePr");
    if (!groupOpening || (!rangeOpening && !discreteOpening)) return [];
    const groupAttrs = attributes(groupOpening);
    const base = Number(groupAttrs.base);
    const parent = Number(groupAttrs.par);
    if (!Number.isInteger(base) || base < 0 || base >= fields.length) return [];
    if (discreteOpening) {
      const mappings = elements(body(entry.xml, "discretePr"), "x").map((item) => Number(attributes(item.opening).v));
      const grouped = new Map();
      items[base].forEach((sourceItem, sourceIndex) => {
        const groupIndex = mappings[sourceIndex];
        const label = Number.isInteger(groupIndex) && groupIndex >= 0 && groupIndex < items[index].length ? items[index][groupIndex] : sourceItem;
        if (String(label) === String(sourceItem)) return;
        const key = String(label);
        if (!grouped.has(key)) grouped.set(key, { name: key, items: [] });
        grouped.get(key).items.push(sourceItem);
      });
      return [{ name: fields[index], sourceField: fields[base], groupBy: "discrete", groups: [...grouped.values()], items: items[index] }];
    }
    const rangeAttrs = attributes(rangeOpening);
    if (!rangeAttrs.groupBy) return [];
    return [{
      name: fields[index],
      sourceField: fields[base],
      groupBy: rangeAttrs.groupBy,
      parent: Number.isInteger(parent) && parent >= 0 && parent < fields.length ? fields[parent] : undefined,
      items: items[index],
      range: {
        autoStart: booleanAttribute(rangeAttrs.autoStart, true),
        autoEnd: booleanAttribute(rangeAttrs.autoEnd, true),
        startDate: rangeAttrs.startDate,
        endDate: rangeAttrs.endDate,
        startNum: rangeAttrs.startNum == null ? undefined : Number(rangeAttrs.startNum),
        endNum: rangeAttrs.endNum == null ? undefined : Number(rangeAttrs.endNum),
        groupInterval: rangeAttrs.groupInterval == null ? undefined : Number(rangeAttrs.groupInterval),
      },
    }];
  });
  const groupFieldNames = new Set(groupFields.map((field) => field.name));
  return {
    source: {
      sheet: sourceAttrs.sheet,
      ref: sourceAttrs.ref,
      name: sourceAttrs.name,
      relationshipId: relationshipId(sourceAttrs),
    },
    fields,
    sourceFields: fieldAttributes.filter((entry) => entry.formula == null && !groupFieldNames.has(entry.name) && booleanAttribute(entry.databaseField, true)).map((entry) => entry.name),
    groupFields,
    calculatedFields: fieldAttributes.filter((entry) => entry.formula != null).map((entry) => ({ name: entry.name, formula: entry.formula, numFmtId: Number(entry.numFmtId || 0) })),
    items,
    refreshPolicy: {
      refreshOnLoad: booleanAttribute(rootAttrs.refreshOnLoad, false),
      saveData: booleanAttribute(rootAttrs.saveData, true),
      enableRefresh: booleanAttribute(rootAttrs.enableRefresh, true),
      invalid: booleanAttribute(rootAttrs.invalid, false),
      missingItemsLimit: Number(rootAttrs.missingItemsLimit || 0),
      refreshedBy: rootAttrs.refreshedBy,
      refreshedDateIso: rootAttrs.refreshedDateIso,
    },
  };
}

function indexedFields(xml, containerName, fieldNames) {
  const container = body(xml, containerName);
  return [...container.matchAll(/<(?:[A-Za-z_][\w.-]*:)?field\b[^>]*\/?>/g)]
    .map((match) => Number(attributes(match[0]).x))
    .filter((index) => Number.isInteger(index) && index >= 0 && index < fieldNames.length)
    .map((index) => fieldNames[index]);
}

export function parsePivotTableDefinition(xml = "", cache = {}) {
  const rootAttrs = attributes(tag(xml, "pivotTableDefinition"));
  const locationAttrs = attributes(tag(xml, "location"));
  const fields = cache.fields || [];
  const dataFields = body(xml, "dataFields");
  const valueFields = [...dataFields.matchAll(/<(?:[A-Za-z_][\w.-]*:)?dataField\b[^>]*\/?>/g)].map((match) => {
    const attrs = attributes(match[0]);
    const index = Number(attrs.fld);
    return {
      field: Number.isInteger(index) && index >= 0 && index < fields.length ? fields[index] : attrs.name,
      name: attrs.name,
      summarizeBy: attrs.subtotal || "sum",
    };
  }).filter((field) => field.field);
  const pivotFieldEntries = elements(body(xml, "pivotFields"), "pivotField");
  const dateFilters = elements(body(xml, "filters"), "filter").flatMap((entry) => {
    const attrs = attributes(entry.opening);
    if (!PIVOT_DATE_FILTER_TYPES.has(attrs.type)) return [];
    const fieldIndex = Number(attrs.fld);
    const field = Number.isInteger(fieldIndex) && fieldIndex >= 0 && fieldIndex < fields.length ? fields[fieldIndex] : undefined;
    if (!field) return [];
    const customValues = elements(body(entry.xml, "customFilters"), "customFilter").map((item) => attributes(item.opening).val).filter((value) => value != null);
    const value1 = attrs.stringValue1 || customValues[0];
    const between = attrs.type === "dateBetween" || attrs.type === "dateNotBetween";
    const value2 = attrs.stringValue2 || customValues[1];
    return value1 && (!between || value2) ? [{ field, type: attrs.type, value1, value2: between ? value2 : undefined, useWholeDay: true }] : [];
  });
  const dateFilterFields = new Set(dateFilters.map((filter) => filter.field));
  const filters = pivotFieldEntries.flatMap((entry, fieldIndex) => {
    const field = fields[fieldIndex];
    if (!field || dateFilterFields.has(field)) return [];
    const hidden = elements(body(entry.xml, "items"), "item").flatMap((item) => {
      const attrs = attributes(item.opening);
      const index = Number(attrs.x);
      return booleanAttribute(attrs.h, false) && Number.isInteger(index) && index >= 0 && index < (cache.items?.[fieldIndex]?.length || 0) ? [cache.items[fieldIndex][index]] : [];
    });
    return hidden.length ? [{ field, exclude: hidden }] : [];
  });
  return {
    name: rootAttrs.name,
    cacheId: Number(rootAttrs.cacheId),
    targetRange: locationAttrs.ref,
    rowFields: indexedFields(xml, "rowFields", fields),
    columnFields: indexedFields(xml, "colFields", fields),
    valueFields,
    filters: [...filters, ...dateFilters],
  };
}

function uniqueValues(values = []) {
  const seen = new Set();
  return values.filter((value) => {
    const key = `${value === null ? "null" : typeof value}:${String(value)}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function cacheItemXml(value, dateSystem = "1900", asDate = false) {
  if (value == null || value === "") return "<m/>";
  const date = asDate ? pivotDateTimeKey(value, dateSystem) : undefined;
  if (date) return `<d v="${date}"/>`;
  if (typeof value === "number" && Number.isFinite(value)) return `<n v="${value}"/>`;
  if (typeof value === "boolean") return `<b v="${value ? 1 : 0}"/>`;
  return `<s v="${attrEscape(value)}"/>`;
}

function pivotSourceHeaders(pivot) {
  return pivot.sourceFields?.length ? [...pivot.sourceFields] : (pivot.sourceValues()[0] || []).map((value) => String(value ?? ""));
}

function pivotAllFields(pivot) {
  return [...pivotSourceHeaders(pivot), ...(pivot.groupFields || []).map((field) => field.name), ...(pivot.calculatedFields || []).map((field) => field.name)];
}

function pivotFieldValues(pivot, fieldIndex) {
  const sourceFields = pivotSourceHeaders(pivot);
  if (fieldIndex < sourceFields.length) return uniqueValues(pivot.sourceValues().slice(1).map((row) => row[fieldIndex]).filter((value) => value != null && value !== ""));
  const group = (pivot.groupFields || [])[fieldIndex - sourceFields.length];
  return group ? pivotGroupItems(pivot.sourceValues(), group, pivot.dateSystem, sourceFields) : [];
}

function rangePropertiesXml(group) {
  const range = group.range || {};
  const attrs = [`groupBy="${attrEscape(group.groupBy)}"`, `autoStart="${range.autoStart === false ? 0 : 1}"`, `autoEnd="${range.autoEnd === false ? 0 : 1}"`];
  for (const [name, value] of [["startDate", range.startDate], ["endDate", range.endDate], ["startNum", range.startNum], ["endNum", range.endNum], ["groupInterval", range.groupInterval]]) if (value != null && value !== "") attrs.push(`${name}="${attrEscape(value)}"`);
  return `<rangePr ${attrs.join(" ")}/>`;
}

function pivotGroupCacheFieldsXml(pivot, sourceFields) {
  const groups = pivot.groupFields || [];
  const allFields = [...sourceFields, ...groups.map((group) => group.name)];
  return groups.map((group) => {
    const base = sourceFields.indexOf(group.sourceField);
    const parent = group.parent ? allFields.indexOf(group.parent) : -1;
    const items = pivotGroupItems(pivot.sourceValues(), group, pivot.dateSystem, sourceFields);
    const groupItems = `<groupItems count="${items.length}">${items.map((value) => `<s v="${attrEscape(value)}"/>`).join("")}</groupItems>`;
    if (group.groupBy === "discrete") {
      const sourceItems = pivotFieldValues(pivot, base);
      const mapping = sourceItems.map((value) => items.findIndex((item) => String(item) === String(pivotGroupValue(group, value, pivot.dateSystem))));
      const discrete = `<discretePr count="${mapping.length}">${mapping.map((value) => `<x v="${Math.max(0, value)}"/>`).join("")}</discretePr>`;
      return `<cacheField name="${attrEscape(group.name)}" databaseField="0" numFmtId="0"><fieldGroup base="${base}">${discrete}${groupItems}</fieldGroup></cacheField>`;
    }
    return `<cacheField name="${attrEscape(group.name)}" databaseField="0" numFmtId="0"><fieldGroup base="${base}"${parent >= 0 ? ` par="${parent}"` : ""}>${rangePropertiesXml(group)}${groupItems}</fieldGroup></cacheField>`;
  }).join("");
}

export function spreadsheetPivotCacheDefinitionXml(part) {
  const { pivot } = part;
  const sourceSheet = pivot.sourceRange.sheetName || pivot.worksheet.name;
  const sourceFields = pivotSourceHeaders(pivot);
  const dateGroupSources = new Set((pivot.groupFields || []).filter((group) => PIVOT_CALENDAR_GROUP_TYPES.has(group.groupBy)).map((group) => group.sourceField));
  const sourceCacheFields = sourceFields.map((header, index) => {
    const values = pivotFieldValues(pivot, index);
    const dateField = dateGroupSources.has(header);
    const isDate = (value) => dateField && Boolean(pivotDateKey(value, pivot.dateSystem));
    const containsDate = values.some(isDate);
    const containsNumber = values.some((value) => !isDate(value) && typeof value === "number" && Number.isFinite(value));
    const containsString = values.some((value) => !isDate(value) && typeof value !== "number" && typeof value !== "boolean");
    const containsNonDate = dateField && values.some((value) => !isDate(value));
    const containsBlank = pivot.sourceValues().slice(1).some((row) => row[index] == null || row[index] === "");
    const containsMixedTypes = [containsDate, containsNumber, containsString, values.some((value) => typeof value === "boolean")].filter(Boolean).length > 1;
    const flags = `${containsDate ? ' containsDate="1"' : ""}${containsNonDate ? ' containsNonDate="1"' : ""}${containsNumber ? ' containsNumber="1"' : ""}${containsString ? ' containsString="1"' : ""}${containsBlank ? ' containsBlank="1"' : ""}${containsMixedTypes ? ' containsMixedTypes="1"' : ""}`;
    return `<cacheField name="${attrEscape(header || `Field${index + 1}`)}" numFmtId="0"><sharedItems${flags} count="${values.length}">${values.map((value) => cacheItemXml(value, pivot.dateSystem, dateField)).join("")}</sharedItems></cacheField>`;
  }).join("");
  const groupedCacheFields = pivotGroupCacheFieldsXml(pivot, sourceFields);
  const calculatedCacheFields = (pivot.calculatedFields || []).map((field) => `<cacheField name="${attrEscape(field.name)}" formula="${attrEscape(field.supported === false ? field.formula.replace(/^=/, "") : pivotFormulaToOoxml(field.formula, sourceFields))}" databaseField="0" numFmtId="${Number(field.numFmtId || 0)}"><sharedItems containsNumber="1" count="0"/></cacheField>`).join("");
  const fieldCount = sourceFields.length + (pivot.groupFields?.length || 0) + (pivot.calculatedFields?.length || 0);
  const policy = pivot.refreshPolicy || {};
  const recordsRelationship = policy.saveData === false ? "" : ` r:id="${attrEscape(part.recordsRelId || "rId1")}"`;
  const refreshedBy = policy.refreshedBy ? ` refreshedBy="${attrEscape(policy.refreshedBy)}"` : "";
  const refreshedDateIso = policy.refreshedDateIso ? ` refreshedDateIso="${attrEscape(policy.refreshedDateIso)}"` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"${recordsRelationship} refreshOnLoad="${policy.refreshOnLoad === false ? 0 : 1}" saveData="${policy.saveData === false ? 0 : 1}" enableRefresh="${policy.enableRefresh === false ? 0 : 1}" invalid="${policy.invalid ? 1 : 0}" missingItemsLimit="${Number(policy.missingItemsLimit || 0)}"${refreshedBy}${refreshedDateIso} recordCount="${Math.max(0, pivot.sourceValues().length - 1)}"><cacheSource type="worksheet"><worksheetSource ref="${attrEscape(pivot.sourceRange.address)}" sheet="${attrEscape(sourceSheet)}"/></cacheSource><cacheFields count="${fieldCount}">${sourceCacheFields}${groupedCacheFields}${calculatedCacheFields}</cacheFields></pivotCacheDefinition>`;
}

export function spreadsheetPivotCacheRecordsXml(part) {
  const sourceFields = pivotSourceHeaders(part.pivot);
  const dateGroupSources = new Set((part.pivot.groupFields || []).filter((group) => PIVOT_CALENDAR_GROUP_TYPES.has(group.groupBy)).map((group) => group.sourceField));
  const rows = part.pivot.sourceValues().slice(1);
  const records = rows.map((row) => `<r>${row.map((value, index) => cacheItemXml(value, part.pivot.dateSystem, dateGroupSources.has(sourceFields[index]))).join("")}</r>`).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotCacheRecords xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${rows.length}">${records}</pivotCacheRecords>`;
}

function columnNumberToLabel(index) {
  let value = index + 1;
  let label = "";
  while (value > 0) { label = String.fromCharCode(65 + ((value - 1) % 26)) + label; value = Math.floor((value - 1) / 26); }
  return label;
}

function cellAddress(row, column) {
  return `${columnNumberToLabel(column)}${row + 1}`;
}

function targetStart(address = "A1") {
  const match = /(?:^|:)(?:\$?)([A-Za-z]+)(?:\$?)(\d+)/.exec(String(address));
  if (!match) return { row: 0, column: 0 };
  const column = [...match[1].toUpperCase()].reduce((sum, char) => sum * 26 + char.charCodeAt(0) - 64, 0) - 1;
  return { row: Number(match[2]) - 1, column };
}

function dateFilterXml(filter, fieldIndex, id) {
  const operators = {
    dateEqual: ["equal"], dateNotEqual: ["notEqual"],
    dateOlderThan: ["lessThan"], dateOlderThanOrEqual: ["lessThanOrEqual"],
    dateNewerThan: ["greaterThan"], dateNewerThanOrEqual: ["greaterThanOrEqual"],
    dateBetween: ["greaterThanOrEqual", "lessThanOrEqual"],
    dateNotBetween: ["lessThan", "greaterThan"],
  };
  const values = [filter.value1, filter.value2].filter((value) => value != null);
  const custom = operators[filter.type].map((operator, index) => `<customFilter operator="${operator}" val="${attrEscape(values[index])}"/>`).join("");
  const join = values.length > 1 ? ` and="${filter.type === "dateBetween" ? 1 : 0}"` : "";
  const second = filter.value2 ? ` stringValue2="${attrEscape(filter.value2)}"` : "";
  return `<filter fld="${fieldIndex}" type="${filter.type}" id="${id}" stringValue1="${attrEscape(filter.value1)}"${second}><autoFilter ref="A1"><filterColumn colId="0"><customFilters${join}>${custom}</customFilters></filterColumn></autoFilter></filter>`;
}

export function spreadsheetPivotTableDefinitionXml(part) {
  const { pivot } = part;
  const headers = pivotAllFields(pivot);
  const start = targetStart(pivot.targetRange.address);
  const values = pivot.computedValues();
  const targetEnd = cellAddress(start.row + Math.max(0, values.length - 1), start.column + Math.max(0, (values[0]?.length || 1) - 1));
  const ref = `${cellAddress(start.row, start.column)}:${targetEnd}`;
  const rowIndexes = pivot.rowFields.map((field) => headers.indexOf(String(field))).filter((index) => index >= 0);
  const columnIndexes = pivot.columnFields.map((field) => headers.indexOf(String(field))).filter((index) => index >= 0);
  const valueIndexes = pivot.valueFields.map((field) => headers.indexOf(String(field.field || field.name))).filter((index) => index >= 0);
  const pivotFields = headers.map((header, index) => {
    const fieldValues = pivotFieldValues(pivot, index);
    const filter = pivot.filters.find((entry) => entry.field === header);
    const items = fieldValues.map((value, itemIndex) => `<item x="${itemIndex}"${pivotItemVisible(pivot.filters, header, value, pivot.dateSystem) ? "" : ' h="1"'}/>`).join("") + '<item t="default"/>';
    const axis = rowIndexes.includes(index) ? ' axis="axisRow"' : columnIndexes.includes(index) ? ' axis="axisCol"' : "";
    return `<pivotField${axis}${valueIndexes.includes(index) ? ' dataField="1"' : ""}${filter ? ' multipleItemSelectionAllowed="1"' : ""} showAll="0"><items count="${fieldValues.length + 1}">${items}</items></pivotField>`;
  }).join("");
  const rowFields = rowIndexes.length ? `<rowFields count="${rowIndexes.length}">${rowIndexes.map((index) => `<field x="${index}"/>`).join("")}</rowFields>` : "";
  const columnFields = columnIndexes.length ? `<colFields count="${columnIndexes.length}">${columnIndexes.map((index) => `<field x="${index}"/>`).join("")}</colFields>` : "";
  const dataFields = pivot.valueFields.length ? `<dataFields count="${pivot.valueFields.length}">${pivot.valueFields.map((field) => `<dataField name="${attrEscape(pivotValueLabel(field))}" fld="${Math.max(0, headers.indexOf(String(field.field || field.name)))}" subtotal="${attrEscape(field.summarizeBy || "sum")}"/>`).join("")}</dataFields>` : "";
  const dateFilters = pivot.filters.filter((filter) => PIVOT_DATE_FILTER_TYPES.has(filter.type));
  const filters = dateFilters.length ? `<filters count="${dateFilters.length}">${dateFilters.map((filter, index) => dateFilterXml(filter, headers.indexOf(filter.field), index + 1)).join("")}</filters>` : "";
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><pivotTableDefinition xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" name="${attrEscape(pivot.name)}" cacheId="${part.cacheId}" dataCaption="Values" updatedVersion="7" minRefreshableVersion="3" multipleFieldFilters="1"><location ref="${attrEscape(ref)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/><pivotFields count="${headers.length}">${pivotFields}</pivotFields>${rowFields}${columnFields}${dataFields}${filters}</pivotTableDefinition>`;
}
