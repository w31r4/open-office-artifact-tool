function chartRangeTarget(worksheet, reference) {
  const raw = String(reference || "").trim().replace(/^=/, "");
  if (!raw) return undefined;
  const separator = raw.lastIndexOf("!");
  let targetSheet = worksheet;
  let address = raw;
  if (separator >= 0) {
    let sheetName = raw.slice(0, separator);
    address = raw.slice(separator + 1);
    if (sheetName.startsWith("'") && sheetName.endsWith("'")) sheetName = sheetName.slice(1, -1).replaceAll("''", "'");
    if (!sheetName || sheetName.startsWith("[")) return undefined;
    targetSheet = worksheet?.workbook?.worksheets?.getItem?.(sheetName);
  }
  address = address.replaceAll("$", "");
  if (!targetSheet || !/^[A-Za-z]{1,3}[1-9]\d*(?::[A-Za-z]{1,3}[1-9]\d*)?$/.test(address)) return undefined;
  return { targetSheet, address };
}

export function worksheetChartRangeValues(worksheet, reference) {
  const target = chartRangeTarget(worksheet, reference);
  if (!target) return [];
  try {
    return target.targetSheet.getRange(target.address).values.flat();
  } catch {
    return [];
  }
}

export function resolvedWorksheetChartCategories(chart) {
  const formula = chart?.series?.items?.find((series) => series?.categoryFormula)?.categoryFormula;
  const referenced = worksheetChartRangeValues(chart?.worksheet, formula);
  return referenced.length ? referenced.map((value) => String(value ?? "")) : [...(chart?.categories || [])].map((value) => String(value ?? ""));
}

export function resolvedWorksheetChartSeriesValues(chart, series) {
  const referenced = worksheetChartRangeValues(chart?.worksheet, series?.formula);
  if (referenced.length) {
    const numeric = referenced.map(Number);
    if (numeric.every(Number.isFinite)) return numeric;
  }
  return [...(series?.values || [])].map(Number);
}

export function resolvedWorksheetChartSeriesXValues(chart, series) {
  const referenced = worksheetChartRangeValues(chart?.worksheet, series?.xFormula);
  if (referenced.length) {
    const numeric = referenced.map(Number);
    if (numeric.every(Number.isFinite)) return numeric;
  }
  return [...(series?.xValues || [])].map(Number);
}

export function resolvedWorksheetChartSeriesBubbleSizes(chart, series) {
  const referenced = worksheetChartRangeValues(chart?.worksheet, series?.bubbleSizeFormula);
  if (referenced.length) {
    const numeric = referenced.map(Number);
    if (numeric.every((value) => Number.isFinite(value) && value > 0)) return numeric;
  }
  return [...(series?.bubbleSizes || [])].map(Number);
}
