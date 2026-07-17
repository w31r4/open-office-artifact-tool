import { CellFormulaKind } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";

function snapshot(definition) {
  return {
    reference: definition.formulaRef,
    rowInput: definition.rowInput || "",
    columnInput: definition.columnInput || "",
    rowOriented: Boolean(definition.rowOriented),
    twoVariable: Boolean(definition.twoVariable),
  };
}

function sameSnapshot(left, right) {
  return JSON.stringify(snapshot(left)) === JSON.stringify(right);
}

function metadata(definition) {
  return {
    kind: CellFormulaKind.DATA_TABLE,
    sharedIndex: 0,
    reference: definition.formulaRef,
    rowInput: definition.rowInput || "",
    columnInput: definition.columnInput || "",
    rowOriented: Boolean(definition.rowOriented),
    twoVariable: Boolean(definition.twoVariable),
    editable: true,
  };
}

function emptyCell(row, column) {
  return {
    row,
    column,
    formula: "",
    numberFormatCode: "",
    value: { case: undefined },
  };
}

export function wireWorksheetDataTables(sheet, state, cells) {
  const byCoordinate = new Map(cells.map((cell) => [`${cell.row}:${cell.column}`, cell]));
  const remaining = new Set(sheet.dataTables?._definitions || []);
  const ordered = [];
  for (const slot of state?.slots || []) {
    if (!slot.definition || !remaining.delete(slot.definition)) {
      throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot remove or reorder imported data table ${slot.wire.reference}.`, [], { code: "invalid_spreadsheet_data_table_topology" });
    }
    if (!sameSnapshot(slot.definition, slot.publicSnapshot)) {
      throw new OpenChestnutCodecError(`Imported data table ${sheet.name}!${slot.wire.reference} is source-bound and read-only.`, [], { code: "unsupported_spreadsheet_data_table_edit" });
    }
    ordered.push({ definition: slot.definition, wire: slot.wire });
  }
  if (state && remaining.size) {
    const added = [...remaining][0];
    throw new OpenChestnutCodecError(`Worksheet ${sheet.name} cannot add data table ${added.formulaRef} to a source-bound package.`, [], { code: "invalid_spreadsheet_data_table_topology" });
  }
  if (!state) ordered.push(...[...remaining].map((definition) => ({ definition })));

  for (const item of ordered) {
    const { row, col } = item.definition.anchor;
    const key = `${row}:${col}`;
    const cell = byCoordinate.get(key) || emptyCell(row, col);
    if (cell.formula || (cell.formulaMetadata && cell.formulaMetadata.kind !== CellFormulaKind.DATA_TABLE)) {
      throw new OpenChestnutCodecError(`Data table ${sheet.name}!${item.definition.formulaRef} conflicts with another formula at ${item.definition.anchor.address}.`, [], { code: "invalid_spreadsheet_data_table_topology" });
    }
    cell.formula = "";
    cell.formulaMetadata = item.wire || metadata(item.definition);
    if (!byCoordinate.has(key)) {
      byCoordinate.set(key, cell);
      cells.push(cell);
    }
  }
  cells.sort((left, right) => left.row - right.row || left.column - right.column);
  return cells;
}

export function hydrateWorksheetDataTable(sheet, sourceCell) {
  const wire = sourceCell.formulaMetadata;
  const definition = sheet.dataTables._hydrate(wire);
  return { wire, definition, publicSnapshot: snapshot(definition) };
}
