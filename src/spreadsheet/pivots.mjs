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

function formulaTokens(formula, fields) {
  const text = String(formula || "").trim().replace(/^=/, "");
  if (!text) throw new TypeError("PivotTable calculated field formula must not be empty.");
  if (text.length > 4096) throw new RangeError("PivotTable calculated field formula exceeds 4096 characters.");
  const known = new Set(fields);
  const tokens = [];
  for (let index = 0; index < text.length;) {
    if (/\s/.test(text[index])) { index += 1; continue; }
    const rest = text.slice(index);
    const number = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[Ee][+-]?\d+)?/.exec(rest)?.[0];
    if (number) { tokens.push({ type: "number", value: Number(number) }); index += number.length; continue; }
    if ("+-*/^()%".includes(text[index])) { tokens.push({ type: "operator", value: text[index++] }); continue; }
    let field;
    if (text[index] === "[") {
      const end = text.indexOf("]", index + 1);
      if (end < 0) throw new SyntaxError("PivotTable calculated field formula has an unterminated [field] reference.");
      field = text.slice(index + 1, end).trim();
      index = end + 1;
    } else if (text[index] === "'") {
      let value = "";
      index += 1;
      while (index < text.length) {
        if (text[index] !== "'") { value += text[index++]; continue; }
        if (text[index + 1] === "'") { value += "'"; index += 2; continue; }
        index += 1;
        field = value;
        break;
      }
      if (field == null) throw new SyntaxError("PivotTable calculated field formula has an unterminated quoted field reference.");
    } else {
      const identifier = /^[A-Za-z_][A-Za-z0-9_.]*/.exec(rest)?.[0];
      if (!identifier) throw new SyntaxError(`PivotTable calculated field formula has unsupported token near ${rest.slice(0, 12)}.`);
      field = identifier;
      index += identifier.length;
    }
    if (!known.has(field)) throw new Error(`PivotTable calculated field formula references unknown source field ${field}.`);
    tokens.push({ type: "field", value: field });
  }
  if (tokens.length > 512) throw new RangeError("PivotTable calculated field formula exceeds 512 tokens.");
  return tokens;
}

function renderPivotFormula(tokens) {
  return tokens.map((token) => token.type === "field" ? `'${token.value.replaceAll("'", "''")}'` : String(token.value)).join("");
}

export function pivotFormulaToOoxml(formula, fields) {
  return renderPivotFormula(formulaTokens(formula, fields));
}

function calculatedFields(value, sourceFields, allowUnsupported = false) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("PivotTable calculatedFields must be an array.");
  if (value.length > 128) throw new RangeError("PivotTable calculatedFields exceeds 128 fields.");
  const names = new Set(sourceFields);
  return value.map((entry, index) => {
    const name = String(entry?.name || entry?.field || "").trim();
    if (!name) throw new TypeError(`PivotTable calculatedFields[${index}] requires name.`);
    if (names.has(name)) throw new Error(`PivotTable calculated field name ${name} must be unique and must not replace a source field.`);
    names.add(name);
    const formula = String(entry?.formula || "").trim();
    const numFmtId = entry?.numFmtId == null ? 0 : Number(entry.numFmtId);
    if (!Number.isInteger(numFmtId) || numFmtId < 0) throw new TypeError(`PivotTable calculatedFields[${index}] numFmtId must be a non-negative integer.`);
    try {
      const tokens = formulaTokens(formula, sourceFields);
      evaluatePivotFormula(formula, Object.fromEntries(sourceFields.map((field) => [field, 0])), sourceFields);
      return { name, formula: `=${renderPivotFormula(tokens)}`, numFmtId, references: uniqueItems(tokens.filter((token) => token.type === "field").map((token) => token.value)) };
    } catch (error) {
      if (!allowUnsupported || error instanceof RangeError) throw error;
      return { name, formula: formula.startsWith("=") ? formula : `=${formula}`, numFmtId, references: [], supported: false, error: error.message };
    }
  });
}

function formulaBinary(left, operator, right) {
  if (typeof left === "string" && left.startsWith("#")) return left;
  if (typeof right === "string" && right.startsWith("#")) return right;
  if (operator === "/" && right === 0) return "#DIV/0!";
  if (operator === "+") return left + right;
  if (operator === "-") return left - right;
  if (operator === "*") return left * right;
  if (operator === "/") return left / right;
  return left ** right;
}

function formulaUnary(value, transform) {
  return typeof value === "string" && value.startsWith("#") ? value : transform(Number(value));
}

export function evaluatePivotFormula(formula, aggregates = {}, fields = Object.keys(aggregates)) {
  const tokens = formulaTokens(formula, fields);
  let index = 0;
  const primary = () => {
    const token = tokens[index++];
    if (!token) throw new SyntaxError("PivotTable calculated field formula ended unexpectedly.");
    if (token.type === "number") return token.value;
    if (token.type === "field") return Number(aggregates[token.value]) || 0;
    if (token.value === "(") { const value = expression(); if (tokens[index++]?.value !== ")") throw new SyntaxError("PivotTable calculated field formula requires a closing parenthesis."); return value; }
    throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${token.value}.`);
  };
  const unary = () => tokens[index]?.value === "+" ? (index++, formulaUnary(unary(), (value) => value)) : tokens[index]?.value === "-" ? (index++, formulaUnary(unary(), (value) => -value)) : primary();
  const power = () => { let left = unary(); if (tokens[index]?.value === "^") { index += 1; left = formulaBinary(left, "^", power()); } return left; };
  const percent = () => { let value = power(); while (tokens[index]?.value === "%") { index += 1; value = formulaUnary(value, (number) => number / 100); } return value; };
  const term = () => { let left = percent(); while (["*", "/"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = formulaBinary(left, operator, percent()); } return left; };
  const expression = () => { let left = term(); while (["+", "-"].includes(tokens[index]?.value)) { const operator = tokens[index++].value; left = formulaBinary(left, operator, term()); } return left; };
  const result = expression();
  if (index !== tokens.length) throw new SyntaxError(`PivotTable calculated field formula has unexpected token ${tokens[index].value}.`);
  return Number.isFinite(result) || (typeof result === "string" && result.startsWith("#")) ? result : "#NUM!";
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
  const normalizedCalculatedFields = calculatedFields(config.calculatedFields, headers.map(String), config.allowUnsupportedCalculatedFields === true);
  const normalizedValueFields = valueFields(config.valueFields || config.values);
  const filters = pivotFilters(config.filters, new Set([...rowFields, ...columnFields]));
  const knownHeaders = new Set(headers.map(String));
  const knownValueFields = new Set([...knownHeaders, ...normalizedCalculatedFields.map((field) => field.name)]);
  if (config.validateSource !== false && knownHeaders.size) {
    for (const field of [...rowFields, ...columnFields]) if (!knownHeaders.has(field)) throw new Error(`PivotTable field ${field} is not present in the source headers.`);
    for (const field of normalizedValueFields.map((entry) => entry.field)) if (!knownValueFields.has(field)) throw new Error(`PivotTable field ${field} is not present in the source or calculated fields.`);
    for (const filter of filters) {
      const sourceValues = config.sourceValues?.[filter.field];
      if (!sourceValues) continue;
      const knownItems = new Set(sourceValues.map(itemKey));
      for (const value of filter.include || filter.exclude || []) if (!knownItems.has(itemKey(value))) throw new Error(`PivotTable filter ${filter.field} references unknown item ${String(value)}.`);
    }
  }
  return { rowFields, columnFields, valueFields: normalizedValueFields, calculatedFields: normalizedCalculatedFields, filters, refreshPolicy: refreshPolicy(config) };
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
  const calculated = new Map((config.calculatedFields || []).map((field) => [field.name, field]));
  const valueIndexes = valueFields.map((field) => ({ field, index: headers.indexOf(field.field), calculated: calculated.get(field.field) })).filter((entry) => entry.index >= 0 || entry.calculated);
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
    ...columnEntries.flatMap(([key]) => valueIndexes.map(({ field, index, calculated: calculatedField }) => {
      const groupedRows = group.columns.get(key) || [];
      if (!calculatedField) return summarize(groupedRows.map((row) => row[index]), field.summarizeBy);
      if (calculatedField.supported === false) return "#NAME?";
      const aggregates = Object.fromEntries(headers.map((header, fieldIndex) => [header, summarize(groupedRows.map((row) => row[fieldIndex]), "sum")]));
      return evaluatePivotFormula(calculatedField.formula, aggregates, headers);
    })),
  ]);
  return [header, ...outputRows];
}
