function dataLabelsError(message) {
  throw new TypeError(`Worksheet chart dataLabels ${message}`);
}

export function normalizeSpreadsheetChartDataLabels(value) {
  if (value == null) return undefined;
  if (typeof value === "boolean") return { showValue: value, showCategoryName: false };
  if (typeof value !== "object" || Array.isArray(value)) dataLabelsError("must be a boolean or object.");
  const supported = new Set(["showValue", "showCategoryName"]);
  const unsupported = Object.keys(value).filter((key) => !supported.has(key) && value[key] != null);
  if (unsupported.length) dataLabelsError(`supports only showValue and showCategoryName; received ${unsupported.join(", ")}.`);
  const present = [...supported].filter((key) => value[key] != null);
  if (present.length === 0) dataLabelsError("must define showValue or showCategoryName.");
  for (const key of present) if (typeof value[key] !== "boolean") dataLabelsError(`${key} must be a boolean.`);
  return {
    showValue: value.showValue === true,
    showCategoryName: value.showCategoryName === true,
  };
}

export function spreadsheetChartDataLabelText(dataLabels, category, value) {
  const normalized = normalizeSpreadsheetChartDataLabels(dataLabels);
  if (!normalized?.showValue && !normalized?.showCategoryName) return "";
  if (normalized.showValue && normalized.showCategoryName) return `${category ?? ""}: ${value ?? ""}`;
  return normalized.showCategoryName ? String(category ?? "") : String(value ?? "");
}
