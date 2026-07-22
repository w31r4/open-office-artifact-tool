export const WORKSHEET_PROTECTION_OPERATIONS = Object.freeze([
  "selectLockedCells",
  "selectUnlockedCells",
  "formatCells",
  "formatColumns",
  "formatRows",
  "insertColumns",
  "insertRows",
  "insertHyperlinks",
  "deleteColumns",
  "deleteRows",
  "sort",
  "autoFilter",
  "pivotTables",
  "editObjects",
  "editScenarios",
]);

const OPERATION_ORDER = new Map(WORKSHEET_PROTECTION_OPERATIONS.map((name, index) => [name, index]));
const DEFAULT_ALLOWED_OPERATIONS = Object.freeze(["selectLockedCells", "selectUnlockedCells"]);

function assertPlainObject(value) {
  if (!value || typeof value !== "object" || Array.isArray(value))
    throw new TypeError("Worksheet protection must be an object, false, or null.");
}

export function normalizeWorksheetProtection(value) {
  if (value == null || value === false) return undefined;
  assertPlainObject(value);
  const unsupported = Object.keys(value).filter((key) => key !== "enabled" && key !== "allow");
  if (unsupported.length)
    throw new TypeError(`Worksheet protection contains unsupported field${unsupported.length === 1 ? "" : "s"}: ${unsupported.join(", ")}. Password and verifier fields are intentionally not accepted.`);
  if (value.enabled !== undefined && typeof value.enabled !== "boolean")
    throw new TypeError("Worksheet protection enabled must be a boolean when provided.");
  if (value.enabled === false) {
    if (value.allow !== undefined && (!Array.isArray(value.allow) || value.allow.length))
      throw new TypeError("Disabled worksheet protection cannot declare allowed operations.");
    return undefined;
  }
  const allow = value.allow === undefined ? DEFAULT_ALLOWED_OPERATIONS : value.allow;
  if (!Array.isArray(allow)) throw new TypeError("Worksheet protection allow must be an array.");
  const seen = new Set();
  for (const operation of allow) {
    if (typeof operation !== "string" || !OPERATION_ORDER.has(operation))
      throw new TypeError(`Unsupported worksheet protection operation ${String(operation)}; expected one of ${WORKSHEET_PROTECTION_OPERATIONS.join(", ")}.`);
    if (seen.has(operation)) throw new TypeError(`Worksheet protection operation ${operation} is duplicated.`);
    seen.add(operation);
  }
  return Object.freeze({
    enabled: true,
    allow: Object.freeze([...seen].sort((left, right) => OPERATION_ORDER.get(left) - OPERATION_ORDER.get(right))),
  });
}

export function publicWorksheetProtection(value) {
  const normalized = normalizeWorksheetProtection(value);
  return normalized ? { enabled: true, allow: [...normalized.allow] } : undefined;
}

export function worksheetProtectionSnapshot(value) {
  const normalized = normalizeWorksheetProtection(value);
  return normalized ? { enabled: true, allow: [...normalized.allow] } : null;
}
