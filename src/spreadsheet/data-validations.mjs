const DATA_VALIDATION_TYPES = new Set(["list", "whole", "decimal", "date", "time", "textLength", "custom"]);
const DATA_VALIDATION_OPERATORS = new Set(["between", "notBetween", "equal", "notEqual", "lessThan", "lessThanOrEqual", "greaterThan", "greaterThanOrEqual"]);
const DATA_VALIDATION_ERROR_STYLES = new Set(["stop", "warning", "information"]);
const MAX_FORMULA_LENGTH = 8_192;
const MAX_INLINE_LIST_LENGTH = 255;
const MAX_INLINE_LIST_ITEMS = 256;
const MAX_TITLE_LENGTH = 32;
const MAX_MESSAGE_LENGTH = 255;
const RULE_KEYS = new Set([
  "type", "operator", "formula1", "formula2", "values", "allowBlank",
  "showInputMessage", "promptTitle", "prompt", "showErrorMessage",
  "errorTitle", "error", "errorStyle", "showDropdown",
]);
const RECORD_KEYS = new Set(["id", "kind", "sheet", "range", "source"]);

function hasOwn(value, key) {
  return Object.prototype.hasOwnProperty.call(value, key);
}

function isUnsupportedControl(character, allowLineBreaks) {
  const code = character.codePointAt(0);
  if (code > 0x1f && (code < 0x7f || code > 0x9f)) return false;
  return !(allowLineBreaks && (code === 0x09 || code === 0x0a || code === 0x0d));
}

function optionalBoolean(source, target, key) {
  if (!hasOwn(source, key)) return;
  if (typeof source[key] !== "boolean") throw new TypeError(`Spreadsheet data validation ${key} must be a boolean.`);
  target[key] = source[key];
}

function optionalText(source, target, key, maximum, allowLineBreaks) {
  if (!hasOwn(source, key)) return;
  const value = String(source[key] ?? "");
  if (value.length > maximum) throw new RangeError(`Spreadsheet data validation ${key} must be at most ${maximum} characters.`);
  if ([...value].some((character) => isUnsupportedControl(character, allowLineBreaks)))
    throw new TypeError(`Spreadsheet data validation ${key} contains an unsupported control character.`);
  target[key] = value;
}

function formulaText(value, key) {
  const text = String(value ?? "");
  if (text.length > MAX_FORMULA_LENGTH) throw new RangeError(`Spreadsheet data validation ${key} must be at most ${MAX_FORMULA_LENGTH} characters.`);
  if ([...text].some((character) => isUnsupportedControl(character, true)))
    throw new TypeError(`Spreadsheet data validation ${key} contains an unsupported control character.`);
  return text;
}

function inlineList(values) {
  if (!Array.isArray(values)) throw new TypeError("Spreadsheet list data validation values must be an array.");
  if (values.length > MAX_INLINE_LIST_ITEMS) throw new RangeError(`Spreadsheet list data validation supports at most ${MAX_INLINE_LIST_ITEMS} inline values.`);
  const normalized = values.map((value) => String(value));
  if (normalized.some((value) => value.length === 0 || value.includes(",") || [...value].some((character) => isUnsupportedControl(character, false))))
    throw new TypeError("Spreadsheet list data validation inline values must be non-empty and cannot contain commas or control characters.");
  if (`"${normalized.map((value) => value.replaceAll('"', '""')).join(",")}"`.length > MAX_INLINE_LIST_LENGTH)
    throw new RangeError(`Spreadsheet list data validation inline values must fit within ${MAX_INLINE_LIST_LENGTH} characters.`);
  return normalized;
}

export function normalizeSpreadsheetDataValidationRule(input = {}) {
  const source = input?.rule && typeof input.rule === "object" ? input.rule : input;
  if (!source || typeof source !== "object" || Array.isArray(source)) throw new TypeError("Spreadsheet data validation rule must be an object.");
  const unexpected = Object.keys(source).filter((key) => !RULE_KEYS.has(key) && !(source === input && RECORD_KEYS.has(key)));
  if (unexpected.length) throw new TypeError(`Spreadsheet data validation contains unsupported field${unexpected.length === 1 ? "" : "s"}: ${unexpected.join(", ")}.`);
  const type = String(source.type ?? "");
  if (!DATA_VALIDATION_TYPES.has(type)) throw new TypeError(`Spreadsheet data validation type must be one of ${[...DATA_VALIDATION_TYPES].join(", ")}.`);

  const normalized = { type };
  if (hasOwn(source, "operator") && source.operator != null && String(source.operator).length > 0) {
    const operator = String(source.operator);
    if (!DATA_VALIDATION_OPERATORS.has(operator)) throw new TypeError(`Spreadsheet data validation operator ${operator} is unsupported.`);
    normalized.operator = operator;
  }
  if (hasOwn(source, "formula1") && source.formula1 != null && String(source.formula1).length > 0) normalized.formula1 = formulaText(source.formula1, "formula1");
  if (hasOwn(source, "formula2") && source.formula2 != null && String(source.formula2).length > 0) normalized.formula2 = formulaText(source.formula2, "formula2");
  if (hasOwn(source, "values")) normalized.values = inlineList(source.values);

  if (normalized.values?.length && normalized.formula1) throw new TypeError("Spreadsheet data validation cannot combine inline values and formula1.");
  if (normalized.values && type !== "list") throw new TypeError("Spreadsheet data validation inline values require type list.");
  if ((type === "list" || type === "custom") && !normalized.formula1 && !normalized.values?.length)
    throw new TypeError(`Spreadsheet ${type} data validation requires formula1 or inline values.`);
  const between = normalized.operator === "between" || normalized.operator === "notBetween";
  if (between && !normalized.formula2) throw new TypeError(`Spreadsheet data validation operator ${normalized.operator} requires formula2.`);
  if (!between && normalized.formula2) throw new TypeError("Spreadsheet data validation formula2 is valid only with between or notBetween.");

  optionalBoolean(source, normalized, "allowBlank");
  optionalBoolean(source, normalized, "showInputMessage");
  optionalBoolean(source, normalized, "showErrorMessage");
  optionalBoolean(source, normalized, "showDropdown");
  if (hasOwn(normalized, "showDropdown") && type !== "list") throw new TypeError("Spreadsheet data validation showDropdown is valid only for list rules.");
  optionalText(source, normalized, "promptTitle", MAX_TITLE_LENGTH, false);
  optionalText(source, normalized, "prompt", MAX_MESSAGE_LENGTH, true);
  optionalText(source, normalized, "errorTitle", MAX_TITLE_LENGTH, false);
  optionalText(source, normalized, "error", MAX_MESSAGE_LENGTH, true);
  if (hasOwn(source, "errorStyle") && source.errorStyle != null && String(source.errorStyle).length > 0) {
    const errorStyle = String(source.errorStyle);
    if (!DATA_VALIDATION_ERROR_STYLES.has(errorStyle)) throw new TypeError(`Spreadsheet data validation errorStyle must be one of ${[...DATA_VALIDATION_ERROR_STYLES].join(", ")}.`);
    normalized.errorStyle = errorStyle;
  }
  return normalized;
}

export function normalizeSpreadsheetDataValidationRecord(config = {}, defaults = {}) {
  if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Spreadsheet data validation must be an object.");
  return {
    ...defaults,
    ...config,
    rule: normalizeSpreadsheetDataValidationRule(config),
  };
}

export function spreadsheetDataValidationIssue(record) {
  try {
    normalizeSpreadsheetDataValidationRecord(record);
    return undefined;
  } catch (error) {
    return String(error?.message || error);
  }
}
