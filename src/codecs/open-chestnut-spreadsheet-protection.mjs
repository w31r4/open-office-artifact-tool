import { SpreadsheetWorksheetProtectionOperation } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizeWorksheetProtection, worksheetProtectionSnapshot } from "../spreadsheet/worksheet-protection.mjs";

const OPERATION_WIRE = Object.freeze({
  selectLockedCells: SpreadsheetWorksheetProtectionOperation.SELECT_LOCKED_CELLS,
  selectUnlockedCells: SpreadsheetWorksheetProtectionOperation.SELECT_UNLOCKED_CELLS,
  formatCells: SpreadsheetWorksheetProtectionOperation.FORMAT_CELLS,
  formatColumns: SpreadsheetWorksheetProtectionOperation.FORMAT_COLUMNS,
  formatRows: SpreadsheetWorksheetProtectionOperation.FORMAT_ROWS,
  insertColumns: SpreadsheetWorksheetProtectionOperation.INSERT_COLUMNS,
  insertRows: SpreadsheetWorksheetProtectionOperation.INSERT_ROWS,
  insertHyperlinks: SpreadsheetWorksheetProtectionOperation.INSERT_HYPERLINKS,
  deleteColumns: SpreadsheetWorksheetProtectionOperation.DELETE_COLUMNS,
  deleteRows: SpreadsheetWorksheetProtectionOperation.DELETE_ROWS,
  sort: SpreadsheetWorksheetProtectionOperation.SORT,
  autoFilter: SpreadsheetWorksheetProtectionOperation.AUTO_FILTER,
  pivotTables: SpreadsheetWorksheetProtectionOperation.PIVOT_TABLES,
  editObjects: SpreadsheetWorksheetProtectionOperation.EDIT_OBJECTS,
  editScenarios: SpreadsheetWorksheetProtectionOperation.EDIT_SCENARIOS,
});
const WIRE_OPERATION = new Map(Object.entries(OPERATION_WIRE).map(([name, value]) => [value, name]));

export function publicWorksheetProtectionFromWire(wire) {
  if (!wire?.enabled) return undefined;
  const allow = (wire.allowedOperations || []).map((operation) => {
    const name = WIRE_OPERATION.get(operation);
    if (!name) throw new Error(`OpenChestnut returned unsupported worksheet protection operation ${operation}.`);
    return name;
  });
  return normalizeWorksheetProtection({ allow });
}

export function wireWorksheetProtection(sheet, slot) {
  const snapshot = worksheetProtectionSnapshot(sheet.protection);
  const sourceWire = slot?.wire?.protection;
  if (slot && JSON.stringify(snapshot) === JSON.stringify(slot.publicProtectionSnapshot)) return sourceWire;
  if (!snapshot) {
    return sourceWire ? { enabled: false, allowedOperations: [], source: sourceWire.source } : undefined;
  }
  return {
    enabled: true,
    allowedOperations: snapshot.allow.map((operation) => OPERATION_WIRE[operation]),
    source: sourceWire?.source,
  };
}

export function worksheetProtectionPublicSnapshot(sheet) {
  return worksheetProtectionSnapshot(sheet.protection);
}
