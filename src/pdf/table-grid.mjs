const VALID_ROLES = new Set(["TH", "TD"]);
const VALID_SCOPES = new Set(["Row", "Column", "Both"]);

function integer(value, fallback) {
  const number = Number(value);
  return Number.isInteger(number) ? number : fallback;
}

function cellId(tableId, row, column) {
  return `${tableId}/cell/${row + 1}/${column + 1}`;
}

export function normalizePdfTableGrid(table = {}) {
  const values = Array.isArray(table.values) ? table.values : [];
  const rows = values.length;
  const columns = Math.max(0, ...values.map((row) => Array.isArray(row) ? row.length : 0));
  const errors = [];
  const overrides = new Map();
  for (const [index, candidate] of (Array.isArray(table.cellConfigs) ? table.cellConfigs : Array.isArray(table.cells) ? table.cells : []).entries()) {
    const row = integer(candidate?.row, -1);
    const column = integer(candidate?.column ?? candidate?.col, -1);
    if (row < 0 || column < 0 || row >= rows || column >= columns) {
      errors.push({ code: "cellOutOfRange", message: `Cell override ${index + 1} targets (${row}, ${column}) outside the ${rows}x${columns} table.`, row, column });
      continue;
    }
    overrides.set(`${row}:${column}`, { ...candidate, row, column });
  }

  const occupied = Array.from({ length: rows }, () => Array(columns));
  const cells = [];
  for (let row = 0; row < rows; row += 1) {
    for (let column = 0; column < columns; column += 1) {
      if (occupied[row][column]) continue;
      const override = overrides.get(`${row}:${column}`) || {};
      const rowSpan = integer(override.rowSpan, 1);
      const columnSpan = integer(override.columnSpan ?? override.colSpan, 1);
      const role = String(override.role || (row === 0 ? "TH" : "TD")).toUpperCase();
      const scope = override.scope == null ? (role === "TH" ? "Column" : undefined) : String(override.scope);
      const id = String(override.id || cellId(table.id || "table", row, column));
      const cell = {
        id,
        row,
        column,
        rowSpan,
        columnSpan,
        role,
        scope,
        headers: Array.isArray(override.headers) ? [...new Set(override.headers.map(String))] : [],
        value: override.value ?? values[row]?.[column] ?? "",
      };
      if (rowSpan < 1 || columnSpan < 1) errors.push({ code: "invalidSpan", message: `Cell ${id} must have positive integer rowSpan and columnSpan.`, id, row, column });
      if (row + rowSpan > rows || column + columnSpan > columns) errors.push({ code: "spanOutOfRange", message: `Cell ${id} span exceeds the ${rows}x${columns} table.`, id, row, column });
      if (!VALID_ROLES.has(role)) errors.push({ code: "invalidRole", message: `Cell ${id} role must be TH or TD.`, id, row, column });
      if (scope != null && (!VALID_SCOPES.has(scope) || role !== "TH")) errors.push({ code: "invalidScope", message: `Cell ${id} scope is only valid for TH and must be Row, Column, or Both.`, id, row, column });
      const safeRowSpan = Math.max(1, Math.min(rows - row, rowSpan || 1));
      const safeColumnSpan = Math.max(1, Math.min(columns - column, columnSpan || 1));
      let overlap = false;
      for (let coveredRow = row; coveredRow < row + safeRowSpan; coveredRow += 1) {
        for (let coveredColumn = column; coveredColumn < column + safeColumnSpan; coveredColumn += 1) {
          if (occupied[coveredRow][coveredColumn]) overlap = true;
        }
      }
      if (overlap) {
        errors.push({ code: "overlappingSpan", message: `Cell ${id} overlaps another spanning cell.`, id, row, column });
        cell.rowSpan = 1;
        cell.columnSpan = 1;
      } else { cell.rowSpan = safeRowSpan; cell.columnSpan = safeColumnSpan; }
      cells.push(cell);
      for (let coveredRow = row; coveredRow < row + Math.max(1, Math.min(rows - row, cell.rowSpan)); coveredRow += 1) {
        for (let coveredColumn = column; coveredColumn < column + Math.max(1, Math.min(columns - column, cell.columnSpan)); coveredColumn += 1) occupied[coveredRow][coveredColumn] = cell;
      }
    }
  }

  for (const override of overrides.values()) {
    const owner = occupied[override.row]?.[override.column];
    if (owner && (owner.row !== override.row || owner.column !== override.column)) errors.push({ code: "overlappingSpan", message: `Cell override at (${override.row}, ${override.column}) is covered by spanning cell ${owner.id}.`, id: owner.id, row: override.row, column: override.column });
  }

  const byId = new Map();
  for (const cell of cells) {
    if (byId.has(cell.id)) errors.push({ code: "duplicateCellId", message: `Table cell ID ${cell.id} is duplicated.`, id: cell.id, row: cell.row, column: cell.column });
    else byId.set(cell.id, cell);
  }
  for (const cell of cells) {
    const inferred = [];
    if (cell.role === "TD" || cell.headers.length) {
      for (const header of cells) {
        if (header.role !== "TH" || header.id === cell.id) continue;
        const coversColumn = cell.column < header.column + header.columnSpan && cell.column + cell.columnSpan > header.column;
        const coversRow = cell.row < header.row + header.rowSpan && cell.row + cell.rowSpan > header.row;
        if ((header.scope === "Column" || header.scope === "Both") && header.row <= cell.row && coversColumn) inferred.push(header.id);
        if ((header.scope === "Row" || header.scope === "Both") && header.column <= cell.column && coversRow) inferred.push(header.id);
      }
    }
    cell.effectiveHeaders = cell.headers.length ? [...cell.headers] : [...new Set(inferred)];
    for (const headerId of cell.headers) {
      const header = byId.get(headerId);
      if (headerId === cell.id) errors.push({ code: "invalidHeader", message: `Cell ${cell.id} cannot reference itself as a header.`, id: cell.id, headerId });
      else if (!header) errors.push({ code: "missingHeader", message: `Cell ${cell.id} references missing header ${headerId}.`, id: cell.id, headerId });
      else if (header.role !== "TH") errors.push({ code: "invalidHeader", message: `Cell ${cell.id} references ${headerId}, which is not a TH cell.`, id: cell.id, headerId });
    }
  }

  return { rows, columns, cells, occupied, byId, errors };
}

export function pdfTableCellBBox(table, cell) {
  const grid = normalizePdfTableGrid(table);
  const [left, top, width, height] = (table.bbox || [0, 0, 0, 0]).map(Number);
  const cellWidth = width / Math.max(1, grid.columns);
  const cellHeight = height / Math.max(1, grid.rows);
  return [left + cell.column * cellWidth, top + cell.row * cellHeight, cell.columnSpan * cellWidth, cell.rowSpan * cellHeight];
}

export function serializePdfTableCells(table) {
  return normalizePdfTableGrid(table).cells.map((cell) => ({
    id: cell.id,
    row: cell.row,
    column: cell.column,
    rowSpan: cell.rowSpan,
    columnSpan: cell.columnSpan,
    role: cell.role,
    scope: cell.scope,
    headers: cell.headers,
    effectiveHeaders: cell.effectiveHeaders,
    value: cell.value,
    bbox: pdfTableCellBBox(table, cell),
  }));
}
