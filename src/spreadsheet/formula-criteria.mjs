// Shared Excel-style criteria matching for conditional aggregation functions.

const WILDCARD_ESCAPES = new Set(["?", "*", "~"]);

function criteriaText(value) {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function wildcardPattern(value = "") {
  let source = "";
  let wildcard = false;
  for (let index = 0; index < value.length; index += 1) {
    const character = value[index];
    if (character === "~" && WILDCARD_ESCAPES.has(value[index + 1])) {
      wildcard = true;
      source += value[++index].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    } else if (character === "?") {
      wildcard = true;
      source += ".";
    } else if (character === "*") {
      wildcard = true;
      source += ".*";
    } else source += character.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return { wildcard, regex: new RegExp(`^${source}$`, "iu") };
}

export function matchesFormulaCriteria(value, criteria) {
  const raw = criteriaText(criteria);
  const match = /^(>=|<=|<>|=|>|<)([\s\S]*)$/.exec(raw);
  const operator = match?.[1] || "=";
  const expectedText = match?.[2] ?? raw;
  const blank = value == null || value === "";
  if (expectedText === "") {
    if (operator === "=") return blank;
    if (operator === "<>") return !blank;
  }

  const pattern = wildcardPattern(expectedText);
  if ((operator === "=" || operator === "<>") && pattern.wildcard) {
    const matched = !blank && typeof value === "string" && pattern.regex.test(value);
    return operator === "=" ? matched : !blank && !matched;
  }

  const actualText = criteriaText(value);
  const actualNumber = Number(value);
  const expectedNumber = Number(expectedText);
  const numeric = !blank && expectedText.trim() !== "" && Number.isFinite(actualNumber) && Number.isFinite(expectedNumber);
  const actual = numeric ? actualNumber : actualText.toLocaleLowerCase("en-US");
  const expected = numeric ? expectedNumber : expectedText.toLocaleLowerCase("en-US");
  switch (operator) {
    case ">=": return actual >= expected;
    case "<=": return actual <= expected;
    case "<>": return actual !== expected;
    case ">": return actual > expected;
    case "<": return actual < expected;
    default: return actual === expected;
  }
}
