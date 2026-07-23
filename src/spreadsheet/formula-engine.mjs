/**
 * Internal Spreadsheet formula execution boundary.
 *
 * This module intentionally works against the worksheet/workbook data shape
 * instead of importing Spreadsheet model classes. Keeping formula parsing,
 * dependency analysis, spill management, and evaluation here prevents a
 * model-to-engine back-edge while preserving the public Spreadsheet API.
 */
import { matchesFormulaCriteria } from "./formula-criteria.mjs";
import {
  formulaTimeParts,
  formulaTimeSerial,
  parseFormulaDateText,
  parseFormulaNumberText,
  parseFormulaTimeText,
} from "./formula-coercion.mjs";
import {
  calculateCumipmt,
  calculateCumprinc,
  calculateDb,
  calculateDdb,
  calculateFv,
  calculateIpmt,
  calculateIrr,
  calculateMirr,
  calculateNper,
  calculateNpv,
  calculatePmt,
  calculatePpmt,
  calculatePv,
  calculateRate,
  calculateSln,
  calculateXirr,
  calculateXnpv,
} from "./financial-formulas.mjs";
import {
  makeCellAddress,
  parseCellAddress,
  parseRangeAddress,
  rangeToAddress,
} from "./range-addressing.mjs";
import {
  parseStructuredReference,
  scanStructuredReferenceIntersections,
  scanStructuredReferences,
  splitReferenceIntersectionOperands,
} from "./structured-references.mjs";
import { normalizeKinds } from "../shared/inspection.mjs";

function parseWorkbookReference(workbook, reference) {
  const raw = String(reference || "").trim();
  const bang = raw.lastIndexOf("!");
  let sheetName;
  let address = raw;
  if (bang !== -1) {
    sheetName = raw.slice(0, bang).replace(/^'|'$/g, "");
    address = raw.slice(bang + 1);
  }
  const sheet = sheetName ? workbook.worksheets.getItem(sheetName) : workbook.worksheets.getActiveWorksheet();
  if (!sheet) throw new Error(`Unknown worksheet in trace reference: ${reference}`);
  return { sheet, address: address.replaceAll("$", "").toUpperCase() };
}

function findWorkbookTable(sheet, tableName) {
  const workbook = sheet?.workbook;
  for (const candidateSheet of workbook?.worksheets || [sheet].filter(Boolean)) {
    const table = candidateSheet.tables.items.find((item) => item.name === tableName || item.id === tableName);
    if (table) return { sheet: candidateSheet, table };
  }
  return undefined;
}

function findContainingWorkbookTable(sheet, address) {
  if (!sheet || !address) return undefined;
  const cell = parseCellAddress(String(address).replaceAll("$", "").toUpperCase());
  const table = sheet.tables.items.find((item) => {
    const bounds = parseRangeAddress(item.range);
    return cell.row >= bounds.top && cell.row <= bounds.bottom && cell.col >= bounds.left && cell.col <= bounds.right;
  });
  return table ? { sheet, table } : undefined;
}

function formulaDefinedNameRange(sheet, refText = "", seen = new Set()) {
  const raw = String(refText || "").trim();
  if (!/^[A-Za-z_][A-Za-z0-9_.]*$/.test(raw)) return undefined;
  const workbook = sheet?.workbook;
  const item = workbook?.definedNames.getItem(raw, sheet?.name) || workbook?.definedNames.getItem(raw);
  if (!item) return undefined;
  if (seen.has(item.name)) return { missing: true, name: item.name, refersTo: item.refersTo };
  const target = String(item.refersTo || "").replace(/^=/, "").trim();
  const ref = formulaRefParts(target);
  if (ref) return { ...ref, name: item.name, id: item.id, refersTo: item.refersTo };
  return { missing: true, name: item.name, refersTo: item.refersTo };
}

function structuredColumnIndex(headers, columnName) {
  return headers.findIndex((header) => String(header ?? "").trim() === String(columnName ?? "").trim());
}

function structuredColumnIndexes(selectors, headers) {
  const columnCount = headers.length;
  if (!selectors.length) return { indexes: Array.from({ length: columnCount }, (_, index) => index) };
  const indexes = [];
  const push = (index) => { if (!indexes.includes(index)) indexes.push(index); };
  for (const selector of selectors) {
    if (selector.start && selector.end) {
      const start = structuredColumnIndex(headers, selector.start);
      const end = structuredColumnIndex(headers, selector.end);
      if (start < 0) return { missing: selector.start };
      if (end < 0) return { missing: selector.end };
      const step = start <= end ? 1 : -1;
      for (let index = start; step > 0 ? index <= end : index >= end; index += step) push(index);
    } else {
      const index = structuredColumnIndex(headers, selector.name);
      if (index < 0) return { missing: selector.name };
      push(index);
    }
  }
  return { indexes };
}

function formulaStructuredRefRange(sheet, refText = "", context = {}) {
  const parsed = parseStructuredReference(refText);
  if (!parsed) return undefined;
  const tableName = parsed.tableName;
  const found = tableName ? findWorkbookTable(sheet, tableName) : findContainingWorkbookTable(sheet, context.formulaAddress);
  const tokens = parsed.tokens;
  if (!found) return { missing: true, tableName, columnName: tokens.join(",") };
  const { table, sheet: tableSheet } = found;
  const bounds = parseRangeAddress(table.range);
  const columnCount = bounds.right - bounds.left + 1;
  const headers = Array.from({ length: columnCount }, (_, index) => table.columnNames?.[index] ?? (table.hasHeaders ? (table.values[0]?.[index] ?? `Column${index + 1}`) : `Column${index + 1}`));
  const currentRow = parsed.currentRow || (!parsed.qualified && parsed.sectionTokens.length === 0);
  const section = currentRow ? "#This Row" : parsed.sectionTokens.at(-1) || "#Data";
  const firstDataRow = bounds.top + (table.showHeaders ? 1 : 0);
  const totalsRow = table.showTotals ? bounds.bottom : undefined;
  const lastDataRow = bounds.bottom - (table.showTotals ? 1 : 0);
  let top = firstDataRow;
  let bottom = lastDataRow;
  if (currentRow) {
    const formulaCell = context.formulaAddress ? parseCellAddress(String(context.formulaAddress).replaceAll("$", "").toUpperCase()) : undefined;
    if (!formulaCell || formulaCell.row < firstDataRow || formulaCell.row > lastDataRow || formulaCell.col < bounds.left || formulaCell.col > bounds.right) return { error: "#VALUE!", tableName: table.name, sheetName: tableSheet.name, table, section };
    top = bottom = formulaCell.row;
  } else if (/^#Headers$/i.test(section)) {
    if (!table.showHeaders) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.top, bounds.left), end: makeCellAddress(bounds.top - 1, bounds.left), empty: true, tableName, columnName: parsed.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = bounds.top;
  } else if (/^#Totals$/i.test(section)) {
    if (totalsRow == null) return { sheetName: tableSheet.name, start: makeCellAddress(bounds.bottom + 1, bounds.left), end: makeCellAddress(bounds.bottom, bounds.left), empty: true, tableName, columnName: parsed.columnSelectors.map((item) => item.name || `${item.start}:${item.end}`).join(","), table };
    top = bottom = totalsRow;
  } else if (/^#All$/i.test(section)) {
    top = bounds.top;
    bottom = bounds.bottom;
  } else if (/^#Data$/i.test(section)) {
    top = firstDataRow;
    bottom = lastDataRow;
  } else if (parsed.sectionTokens.length) {
    return { missing: true, tableName, columnName: section, sheetName: tableSheet.name };
  }
  const selected = structuredColumnIndexes(parsed.columnSelectors, headers);
  if (selected.missing) return { missing: true, tableName, columnName: selected.missing, sheetName: tableSheet.name };
  const columns = selected.indexes;
  if (!columns.length) return { missing: true, tableName, columnName: tokens.join(","), sheetName: tableSheet.name };
  const left = bounds.left + Math.min(...columns);
  const right = bounds.left + Math.max(...columns);
  const columnNames = columns.map((index) => String(headers[index] ?? `Column${index + 1}`));
  const columnName = columnNames.join(",");
  const columnIndex = columns.length === 1 ? columns[0] : undefined;
  if (top > bottom) return { sheetName: tableSheet.name, start: makeCellAddress(top, left), end: makeCellAddress(bottom, right), empty: true, tableName, columnName, table, columnIndex, columns, columnNames, section };
  return { sheetName: tableSheet.name, start: makeCellAddress(top, left), end: makeCellAddress(bottom, right), tableName, columnName, table, columnIndex, columns, columnNames, section };
}

function structuredRangeGeometry(structured) {
  const start = parseCellAddress(structured.start);
  const end = parseCellAddress(structured.end);
  const tableBounds = structured.table ? parseRangeAddress(structured.table.range) : undefined;
  const columns = structured.absoluteColumns || (tableBounds && structured.columns?.length
    ? structured.columns.map((index) => tableBounds.left + index)
    : Array.from({ length: Math.abs(end.col - start.col) + 1 }, (_, index) => Math.min(start.col, end.col) + index));
  return {
    top: Math.min(start.row, end.row),
    bottom: Math.max(start.row, end.row),
    columns,
  };
}

function formulaStructuredRefIntersection(sheet, refText = "", context = {}) {
  const text = String(refText || "").trim();
  const groups = scanStructuredReferenceIntersections(text);
  if (groups.length !== 1 || groups[0].start !== 0 || groups[0].end !== text.length) return undefined;
  const group = groups[0];
  const ranges = group.references.map((reference) => formulaStructuredRefRange(sheet, reference.text, context));
  const failure = ranges.find((range) => !range || range.error || range.missing);
  if (failure) return failure || { missing: true };
  if (ranges.some((range) => range.empty)) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  const sheetNames = new Set(ranges.map((range) => range.sheetName || sheet.name));
  if (sheetNames.size !== 1) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  const geometries = ranges.map(structuredRangeGeometry);
  const top = Math.max(...geometries.map((geometry) => geometry.top));
  const bottom = Math.min(...geometries.map((geometry) => geometry.bottom));
  const commonColumns = geometries[0].columns.filter((column) => geometries.slice(1).every((geometry) => geometry.columns.includes(column)));
  if (top > bottom || commonColumns.length === 0) return { error: "#NULL!", intersectionReferences: group.references.map((reference) => reference.text) };
  commonColumns.sort((left, right) => left - right);
  const sharedTable = ranges.every((range) => range.table === ranges[0].table) ? ranges[0].table : undefined;
  const sharedTableBounds = sharedTable ? parseRangeAddress(sharedTable.range) : undefined;
  const columns = sharedTableBounds ? commonColumns.map((column) => column - sharedTableBounds.left) : undefined;
  const columnNames = columns?.map((column) => sharedTable.columnNames?.[column] ?? `Column${column + 1}`);
  return {
    sheetName: [...sheetNames][0],
    start: makeCellAddress(top, commonColumns[0]),
    end: makeCellAddress(bottom, commonColumns.at(-1)),
    absoluteColumns: commonColumns,
    table: sharedTable,
    tableName: sharedTable?.name,
    columns,
    columnNames,
    columnName: columnNames?.join(","),
    section: "intersection",
    intersectionReferences: group.references.map((reference) => reference.text),
  };
}

function formulaA1RefIntersection(sheet, refText = "") {
  const parts = splitReferenceIntersectionOperands(refText);
  if (!parts) return undefined;
  const references = parts.map(formulaRefParts);
  if (references.some((reference) => !reference)) return undefined;
  const sheetNames = new Set(references.map((reference) => reference.sheetName || sheet.name));
  if (sheetNames.size !== 1) return { error: "#NULL!", intersectionReferences: parts };
  const bounds = references.map((reference) => {
    const start = parseCellAddress(reference.start);
    const end = parseCellAddress(reference.end);
    return {
      top: Math.min(start.row, end.row),
      bottom: Math.max(start.row, end.row),
      left: Math.min(start.col, end.col),
      right: Math.max(start.col, end.col),
    };
  });
  const top = Math.max(...bounds.map((item) => item.top));
  const bottom = Math.min(...bounds.map((item) => item.bottom));
  const left = Math.max(...bounds.map((item) => item.left));
  const right = Math.min(...bounds.map((item) => item.right));
  if (top > bottom || left > right) return { error: "#NULL!", intersectionReferences: parts };
  return {
    sheetName: [...sheetNames][0],
    start: makeCellAddress(top, left),
    end: makeCellAddress(bottom, right),
    absoluteColumns: Array.from({ length: right - left + 1 }, (_, index) => left + index),
    section: "intersection",
    intersectionReferences: parts,
  };
}

function scanA1ReferenceIntersections(formula = "") {
  const text = String(formula || "");
  const structured = scanStructuredReferences(text);
  const referenceRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?\$?[A-Za-z]+\$?\d+(?::\$?[A-Za-z]+\$?\d+)?/g;
  const references = [...text.matchAll(referenceRegex)]
    .map((match) => ({ text: match[0], start: match.index, end: match.index + match[0].length }))
    .filter((reference) => !structured.some((item) => reference.start >= item.start && reference.end <= item.end));
  const groups = [];
  for (let index = 0; index < references.length;) {
    const group = [references[index]];
    let cursor = index + 1;
    while (cursor < references.length && /^\s+$/.test(text.slice(group.at(-1).end, references[cursor].start))) {
      group.push(references[cursor]);
      cursor += 1;
    }
    if (group.length > 1) groups.push({ text: text.slice(group[0].start, group.at(-1).end), start: group[0].start, end: group.at(-1).end, references: group });
    index = cursor;
  }
  return groups;
}

function formulaReferences(formula, sheet, formulaAddress) {
  const raw = String(formula || "");
  const refs = [];
  const consumed = [];
  const intersectionGroups = scanStructuredReferenceIntersections(raw);
  const appendStructuredReferences = (structured, refText) => {
    if (!structured || structured.missing || structured.empty || structured.error) return;
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const cols = structuredRangeGeometry(structured).columns;
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (const col of cols) {
        refs.push({ sheetName: structured.sheetName, address: makeCellAddress(row, col), structuredRef: refText, tableName: structured.tableName, columnName: structured.columnName, columnNames: structured.columnNames });
      }
    }
  };
  for (const group of intersectionGroups) {
    consumed.push([group.start, group.end]);
    appendStructuredReferences(formulaStructuredRefIntersection(sheet, group.text, { formulaAddress }), group.text);
  }
  for (const group of scanA1ReferenceIntersections(raw)) {
    consumed.push([group.start, group.end]);
    appendStructuredReferences(formulaA1RefIntersection(sheet, group.text), group.text);
  }
  for (const match of scanStructuredReferences(raw)) {
    if (intersectionGroups.some((group) => match.start >= group.start && match.end <= group.end)) continue;
    consumed.push([match.start, match.end]);
    const structured = formulaStructuredRefRange(sheet, match.text, { formulaAddress });
    appendStructuredReferences(structured, match.text);
  }
  const rangeRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)\s*:\s*(\$?[A-Za-z]+\$?\d+)/g;
  for (const match of raw.matchAll(rangeRegex)) {
    if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
    consumed.push([match.index, match.index + match[0].length]);
    const sheetName = match[1] || match[2] || undefined;
    const start = parseCellAddress(match[3].replaceAll("$", ""));
    const end = parseCellAddress(match[4].replaceAll("$", ""));
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
        refs.push({ sheetName, address: makeCellAddress(row, col) });
      }
    }
  }
  const cellRegex = /(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)/g;
  for (const match of raw.matchAll(cellRegex)) {
    if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
    refs.push({ sheetName: match[1] || match[2] || undefined, address: match[3].replaceAll("$", "").toUpperCase() });
  }
  for (const definedName of sheet?.workbook?.definedNames.items || []) {
    const escaped = definedName.name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    const re = new RegExp(`\\b${escaped}\\b`, "g");
    for (const match of raw.matchAll(re)) {
      if (consumed.some(([start, end]) => match.index >= start && match.index < end)) continue;
      const resolved = formulaDefinedNameRange(sheet, definedName.name);
      if (!resolved || resolved.missing) continue;
      const start = parseCellAddress(resolved.start);
      const end = parseCellAddress(resolved.end);
      for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
        for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) refs.push({ sheetName: resolved.sheetName, address: makeCellAddress(row, col), definedName: definedName.name });
      }
    }
  }
  const seen = new Set();
  return refs.filter((ref) => {
    const key = `${ref.sheetName || ""}!${ref.address}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function formulaCellKey(sheetName, address) {
  return `${String(sheetName || "").replaceAll("'", "''")}!${String(address || "").replaceAll("$", "").toUpperCase()}`;
}

function formulaCellId(sheetName, address) {
  return `fx/${encodeURIComponent(formulaCellKey(sheetName, address))}`;
}

function displayFormulaRef(sheetName, address) {
  const safeSheet = String(sheetName || "").includes(" ") ? `'${String(sheetName).replaceAll("'", "''")}'` : String(sheetName || "");
  return `${safeSheet}!${String(address || "").replaceAll("$", "").toUpperCase()}`;
}

function formulaErrorCode(value) {
  const text = String(value ?? "");
  return /^#(NULL!|DIV\/0!|VALUE!|REF!|NAME\?|NUM!|N\/A|GETTING_DATA|SPILL!|CALC!|CYCLE!|FIELD!|DATA!|CONNECT!|BLOCKED!|UNKNOWN!|BUSY!)/.test(text)
    ? text.match(/^#[A-Z0-9\/?!_]+/)?.[0]
    : undefined;
}

function publicFormulaNode(node) {
  return {
    kind: "formulaNode",
    id: node.id,
    key: node.key,
    sheet: node.sheet,
    address: node.address,
    formula: node.formula,
    value: node.cell?.value,
    precedents: node.precedents.map((ref) => ({ ...ref })),
    dependents: [...node.dependents],
    circular: node.circular || undefined,
    error: formulaErrorCode(node.cell?.value) || undefined,
  };
}

function buildWorkbookFormulaGraph(workbook) {
  const nodes = [];
  const nodeByKey = new Map();
  const errors = [];
  for (const sheet of workbook.worksheets) {
    for (const [address, cell] of sheet.store.entries()) {
      if (!cell.formula) continue;
      const key = formulaCellKey(sheet.name, address);
      const node = { kind: "formulaNode", id: formulaCellId(sheet.name, address), key, sheet: sheet.name, sheetObject: sheet, address, cell, formula: cell.formula, precedents: [], dependents: [], circular: false };
      nodes.push(node);
      nodeByKey.set(key, node);
    }
  }

  const edges = [];
  for (const node of nodes) {
    for (const ref of formulaReferences(node.formula, node.sheetObject, node.address)) {
      const sheetName = ref.sheetName || node.sheet;
      const targetSheet = workbook.worksheets.getItem(sheetName);
      const targetAddress = String(ref.address || "").replaceAll("$", "").toUpperCase();
      const targetKey = formulaCellKey(sheetName, targetAddress);
      const precedent = {
        sheet: sheetName,
        address: targetAddress,
        key: targetKey,
        missing: !targetSheet || undefined,
        hasFormula: nodeByKey.has(targetKey) || undefined,
      };
      node.precedents.push(precedent);
      edges.push({ kind: "formulaEdge", from: node.key, to: targetKey, fromSheet: node.sheet, fromAddress: node.address, toSheet: sheetName, toAddress: targetAddress, missing: !targetSheet || undefined });
      const targetNode = nodeByKey.get(targetKey);
      if (targetNode) targetNode.dependents.push(node.key);
      if (!targetSheet) errors.push({ kind: "formulaGraphError", type: "missingSheet", from: node.key, sheet: sheetName, address: targetAddress, ref: displayFormulaRef(sheetName, targetAddress) });
    }
  }

  const cycles = detectFormulaCycles(nodes, nodeByKey);
  const cycleKeys = new Set(cycles.flatMap((cycle) => cycle.keys));
  for (const node of nodes) if (cycleKeys.has(node.key)) node.circular = true;
  for (const node of nodes) {
    const error = formulaErrorCode(node.cell?.value);
    if (error) errors.push({ kind: "formulaGraphError", type: "formulaError", key: node.key, sheet: node.sheet, address: node.address, value: error });
  }
  return { kind: "formulaGraph", nodes, edges, cycles, errors };
}

function detectFormulaCycles(nodes, nodeByKey) {
  const cycles = [];
  const emitted = new Set();
  const visiting = new Set();
  const visited = new Set();
  const stack = [];

  const emitCycle = (keys) => {
    if (keys.length === 0) return;
    const canonical = [...keys].sort().join("|");
    if (emitted.has(canonical)) return;
    emitted.add(canonical);
    cycles.push({ kind: "formulaCycle", keys, path: [...keys, keys[0]] });
  };

  const visit = (node) => {
    if (visiting.has(node.key)) {
      const start = stack.indexOf(node.key);
      emitCycle(stack.slice(start));
      return;
    }
    if (visited.has(node.key)) return;
    visiting.add(node.key);
    stack.push(node.key);
    for (const ref of node.precedents) {
      const target = nodeByKey.get(ref.key);
      if (target) visit(target);
    }
    stack.pop();
    visiting.delete(node.key);
    visited.add(node.key);
  };

  for (const node of nodes) visit(node);
  return cycles;
}

function formulaGraphRecords(graph, options = {}) {
  const kinds = options.kinds || normalizeKinds(options.kind, ["formulaGraph", "formulaNode", "formulaEdge", "formulaCycle"]);
  const records = [];
  if (kinds.has("formulaGraph")) {
    records.push({ kind: "formulaGraph", formulas: graph.nodes.length, edges: graph.edges.length, cycles: graph.cycles.length, errors: graph.errors.length });
  }
  if (kinds.has("formulaNode")) records.push(...graph.nodes.map(publicFormulaNode));
  if (kinds.has("formulaEdge")) records.push(...graph.edges.map((edge) => ({ ...edge })));
  if (kinds.has("formulaCycle")) records.push(...graph.cycles.map((cycle) => ({ ...cycle })));
  return records;
}

function splitFormulaArgs(text = "") {
  const args = [];
  let current = "";
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  for (let i = 0; i < String(text).length; i++) {
    const ch = String(text)[i];
    if (ch === '"') {
      current += ch;
      if (String(text)[i + 1] === '"') { current += '"'; i += 1; continue; }
      inString = !inString;
      continue;
    }
    if (!inString && ch === "(") depth += 1;
    if (!inString && ch === ")") depth -= 1;
    if (!inString && ch === "[") bracketDepth += 1;
    if (!inString && ch === "]") bracketDepth -= 1;
    if (!inString && ch === "," && depth === 0 && bracketDepth === 0) { args.push(current.trim()); current = ""; continue; }
    current += ch;
  }
  if (current.trim() || text === "") args.push(current.trim());
  return args;
}

function formulaUnquote(value) {
  const text = String(value ?? "").trim();
  if (/^"(?:[^"]|"")*"$/.test(text)) return text.slice(1, -1).replaceAll('""', '"');
  return undefined;
}

function formulaRefParts(refText = "") {
  const raw = String(refText || "").trim();
  const match = /^(?:(?:'([^']+)'|([A-Za-z_][A-Za-z0-9_ ]*))!)?(\$?[A-Za-z]+\$?\d+)(?::(\$?[A-Za-z]+\$?\d+))?$/.exec(raw);
  if (!match) return undefined;
  return { sheetName: match[1] || match[2] || undefined, start: match[3].replaceAll("$", "").toUpperCase(), end: (match[4] || match[3]).replaceAll("$", "").toUpperCase() };
}

function formulaRangeMatrix(sheet, refText, context = {}) {
  const structured = formulaStructuredRefIntersection(sheet, refText, context) || formulaA1RefIntersection(sheet, refText) || formulaStructuredRefRange(sheet, refText, context);
  if (structured) {
    if (structured.error) return [[structured.error]];
    if (structured.missing) return [["#REF!"]];
    if (structured.empty) return [];
    const targetSheet = structured.sheetName ? sheet.workbook?.worksheets.getItem(structured.sheetName) : sheet;
    if (!targetSheet) return [["#REF!"]];
    const start = parseCellAddress(structured.start);
    const end = parseCellAddress(structured.end);
    const cols = structuredRangeGeometry(structured).columns;
    const rows = [];
    for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
      const values = [];
      for (const col of cols) {
        const address = makeCellAddress(row, col);
        let value = context.getValue ? context.getValue({ sheetName: structured.sheetName, address, structuredRef: refText, tableName: structured.tableName, columnName: structured.columnName, columnNames: structured.columnNames }) : targetSheet.store.get(address).value;
        if ((value == null || value === "") && structured.table) {
          const tableBounds = parseRangeAddress(structured.table.range);
          const tableRow = row - tableBounds.top;
          const tableCol = col - tableBounds.left;
          value = structured.table.values[tableRow]?.[tableCol] ?? value;
        }
        values.push(value);
      }
      rows.push(values);
    }
    return rows;
  }
  const defined = formulaDefinedNameRange(sheet, refText);
  if (defined) {
    if (defined.missing) return [["#REF!"]];
    return formulaRangeMatrix(sheet, `${defined.sheetName ? `${defined.sheetName}!` : ""}${defined.start}${defined.end && defined.end !== defined.start ? `:${defined.end}` : ""}`, context);
  }
  const ref = formulaRefParts(refText);
  if (!ref) return undefined;
  const targetSheet = ref.sheetName ? sheet.workbook?.worksheets.getItem(ref.sheetName) : sheet;
  if (!targetSheet) return [["#REF!"]];
  const start = parseCellAddress(ref.start);
  const end = parseCellAddress(ref.end);
  const rows = [];
  for (let row = Math.min(start.row, end.row); row <= Math.max(start.row, end.row); row++) {
    const values = [];
    for (let col = Math.min(start.col, end.col); col <= Math.max(start.col, end.col); col++) {
      const address = makeCellAddress(row, col);
      values.push(context.getValue ? context.getValue({ sheetName: ref.sheetName, address }) : targetSheet.store.get(address).value);
    }
    rows.push(values);
  }
  return rows;
}

function formulaReferenceValues(sheet, refText, context = {}) {
  const matrix = formulaRangeMatrix(sheet, refText, context);
  if (matrix) return matrix.flat();
  const scalar = formulaScalar(sheet, refText, context);
  return scalar === undefined ? [] : [scalar];
}

const FORMULA_NUMERIC_LITERAL = /^-?(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?$/;

function formulaAtomicScalar(sheet, expr, context = {}) {
  const text = String(expr ?? "").trim();
  if (text === "") return undefined;
  const quoted = formulaUnquote(text);
  if (quoted !== undefined) return quoted;
  const error = formulaErrorCode(text);
  if (error) return error;
  if (FORMULA_NUMERIC_LITERAL.test(text)) return Number(text);
  if (/^TRUE$/i.test(text)) return true;
  if (/^FALSE$/i.test(text)) return false;
  const matrix = formulaRangeMatrix(sheet, text, context);
  if (matrix) return matrix[0]?.[0] ?? null;
  return undefined;
}

function formulaScanTopLevel(text, visit) {
  let depth = 0;
  let bracketDepth = 0;
  let inString = false;
  let inSheetName = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (inString) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1;
        else inString = false;
      }
      continue;
    }
    if (inSheetName) {
      if (character === "'") {
        if (text[index + 1] === "'") index += 1;
        else inSheetName = false;
      }
      continue;
    }
    if (character === '"') { inString = true; continue; }
    if (character === "'") { inSheetName = true; continue; }
    if (character === "[") { bracketDepth += 1; continue; }
    if (character === "]") { bracketDepth = Math.max(0, bracketDepth - 1); continue; }
    if (bracketDepth > 0) continue;
    if (character === "(") { depth += 1; continue; }
    if (character === ")") { depth -= 1; continue; }
    if (depth === 0) visit(index, character);
  }
}

function formulaMatchingParenthesis(text, openIndex) {
  if (text[openIndex] !== "(") return undefined;
  let depth = 0;
  let inString = false;
  let inSheetName = false;
  let bracketDepth = 0;
  for (let index = openIndex; index < text.length; index += 1) {
    const character = text[index];
    if (inString) {
      if (character === '"') {
        if (text[index + 1] === '"') index += 1;
        else inString = false;
      }
      continue;
    }
    if (inSheetName) {
      if (character === "'") {
        if (text[index + 1] === "'") index += 1;
        else inSheetName = false;
      }
      continue;
    }
    if (character === '"') { inString = true; continue; }
    if (character === "'") { inSheetName = true; continue; }
    if (character === "[") { bracketDepth += 1; continue; }
    if (character === "]") { bracketDepth = Math.max(0, bracketDepth - 1); continue; }
    if (bracketDepth > 0) continue;
    if (character === "(") depth += 1;
    else if (character === ")") {
      depth -= 1;
      if (depth === 0) return index;
      if (depth < 0) return undefined;
    }
  }
  return undefined;
}

function formulaOuterParentheses(text) {
  if (!text.startsWith("(") || !text.endsWith(")")) return undefined;
  const closeIndex = formulaMatchingParenthesis(text, 0);
  return closeIndex === text.length - 1 ? text.slice(1, -1).trim() : undefined;
}

function formulaTopLevelOperators(text, operators, { binarySigns = false } = {}) {
  const candidates = [];
  formulaScanTopLevel(text, (index, character) => {
    const operator = operators.find((value) => text.startsWith(value, index));
    if (!operator) return;
    if (binarySigns && (operator === "+" || operator === "-")) {
      let previousIndex = index - 1;
      while (previousIndex >= 0 && /\s/.test(text[previousIndex])) previousIndex -= 1;
      const previous = text[previousIndex];
      const exponentSign = (previous === "e" || previous === "E") && /[0-9.]/.test(text[previousIndex - 1] || "");
      if (previousIndex < 0 || "+-*/^&=<>(".includes(previous) || exponentSign) return;
    }
    candidates.push({ index, operator });
  });
  const nonOverlapping = [];
  let consumedThrough = -1;
  for (const candidate of candidates) {
    if (candidate.index < consumedThrough) continue;
    nonOverlapping.push(candidate);
    consumedThrough = candidate.index + candidate.operator.length;
  }
  return nonOverlapping;
}

function formulaWholeFunctionCall(text) {
  const match = /^([A-Z][A-Z0-9.]*)\(/i.exec(text);
  if (!match) return undefined;
  const openIndex = match[0].length - 1;
  const closeIndex = formulaMatchingParenthesis(text, openIndex);
  return closeIndex === text.length - 1 ? { name: match[1].toUpperCase(), args: text.slice(openIndex + 1, -1) } : undefined;
}

function formulaApplyBinaryOperator(left, operator, right) {
  const leftError = formulaErrorCode(left);
  if (leftError) return leftError;
  const rightError = formulaErrorCode(right);
  if (rightError) return rightError;
  if (["=", "<>", ">", "<", ">=", "<="].includes(operator)) return compareFormulaValues(left, operator, right);
  if (operator === "&") return `${formulaText(left)}${formulaText(right)}`;
  const leftNumber = formulaNumber(left);
  const rightNumber = formulaNumber(right);
  if (formulaErrorCode(leftNumber)) return leftNumber;
  if (formulaErrorCode(rightNumber)) return rightNumber;
  if (operator === "+") return leftNumber + rightNumber;
  if (operator === "-") return leftNumber - rightNumber;
  if (operator === "*") return leftNumber * rightNumber;
  if (operator === "/") return rightNumber === 0 ? "#DIV/0!" : leftNumber / rightNumber;
  if (operator === "^") {
    const value = leftNumber ** rightNumber;
    return Number.isFinite(value) ? value : "#NUM!";
  }
  return "#NAME?";
}

function formulaExpressionBinary(sheet, leftText, operator, rightText, context = {}) {
  const left = evaluateFormulaExpression(sheet, leftText, context);
  const leftError = formulaErrorCode(left);
  if (leftError) return leftError;
  const right = evaluateFormulaExpression(sheet, rightText, context);
  return formulaApplyBinaryOperator(left, operator, right);
}

function formulaApplyUnarySign(value, sign) {
  const error = formulaErrorCode(value);
  if (error) return error;
  const number = formulaNumber(value);
  return sign === "-" ? -number : number;
}

const FORMULA_VECTOR_MAX_CELLS = 10_000;

function formulaMatrixGeometry(matrix) {
  if (!Array.isArray(matrix)) return undefined;
  if (matrix.length === 0) return { rows: 0, cols: 0, cells: 0 };
  if (!Array.isArray(matrix[0])) return undefined;
  const rows = matrix.length;
  const cols = matrix[0].length;
  if (!matrix.every((row) => Array.isArray(row) && row.length === cols)) return undefined;
  return { rows, cols, cells: rows * cols };
}

function formulaBoundedVectorMatrix(matrix) {
  const geometry = formulaMatrixGeometry(matrix);
  if (!geometry || geometry.cells > FORMULA_VECTOR_MAX_CELLS) return "#VALUE!";
  return matrix;
}

function formulaReferenceCellCount(reference) {
  const start = parseCellAddress(reference.start);
  const end = parseCellAddress(reference.end);
  return (Math.abs(end.row - start.row) + 1) * (Math.abs(end.col - start.col) + 1);
}

function formulaVectorOperandAt(value, isMatrix, row, col) {
  return isMatrix ? value[row][col] : value;
}

function formulaVectorBinary(sheet, leftText, operator, rightText, context = {}) {
  const left = evaluateFormulaVectorExpression(sheet, leftText, context);
  const right = evaluateFormulaVectorExpression(sheet, rightText, context);
  const leftIsMatrix = isFormulaMatrix(left);
  const rightIsMatrix = isFormulaMatrix(right);
  if (!leftIsMatrix && !rightIsMatrix) return formulaApplyBinaryOperator(left, operator, right);

  const leftGeometry = leftIsMatrix ? formulaMatrixGeometry(left) : undefined;
  const rightGeometry = rightIsMatrix ? formulaMatrixGeometry(right) : undefined;
  if ((leftIsMatrix && (!leftGeometry || leftGeometry.cells > FORMULA_VECTOR_MAX_CELLS))
    || (rightIsMatrix && (!rightGeometry || rightGeometry.cells > FORMULA_VECTOR_MAX_CELLS))) return "#VALUE!";
  const geometry = leftGeometry || rightGeometry;
  if (leftGeometry && rightGeometry && (leftGeometry.rows !== rightGeometry.rows || leftGeometry.cols !== rightGeometry.cols)) return "#VALUE!";
  return Array.from({ length: geometry.rows }, (_, row) => Array.from({ length: geometry.cols }, (_, col) =>
    formulaApplyBinaryOperator(formulaVectorOperandAt(left, leftIsMatrix, row, col), operator, formulaVectorOperandAt(right, rightIsMatrix, row, col))));
}

function formulaVectorUnarySign(sheet, valueText, sign, context = {}) {
  const value = evaluateFormulaVectorExpression(sheet, valueText, context);
  if (!isFormulaMatrix(value)) return formulaApplyUnarySign(value, sign);
  const geometry = formulaMatrixGeometry(value);
  if (!geometry || geometry.cells > FORMULA_VECTOR_MAX_CELLS) return "#VALUE!";
  return value.map((row) => row.map((item) => formulaApplyUnarySign(item, sign)));
}

function evaluateFormulaVectorExpression(sheet, expr, context = {}) {
  let text = String(expr ?? "").trim();
  if (text === "") return undefined;
  let outer;
  while ((outer = formulaOuterParentheses(text)) !== undefined) text = outer;

  const comparison = formulaTopLevelOperators(text, [">=", "<=", "<>", "=", ">", "<"]);
  if (comparison.length > 1) return "#VALUE!";
  if (comparison.length === 1) {
    const { index, operator } = comparison[0];
    return formulaVectorBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const concatenation = formulaTopLevelOperators(text, ["&"]);
  if (concatenation.length) {
    const { index, operator } = concatenation.at(-1);
    return formulaVectorBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const addition = formulaTopLevelOperators(text, ["+", "-"], { binarySigns: true });
  if (addition.length) {
    const { index, operator } = addition.at(-1);
    return formulaVectorBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const multiplication = formulaTopLevelOperators(text, ["*", "/"]);
  if (multiplication.length) {
    const { index, operator } = multiplication.at(-1);
    return formulaVectorBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const exponentiation = formulaTopLevelOperators(text, ["^"]);
  if (exponentiation.length) {
    const { index, operator } = exponentiation[0];
    return formulaVectorBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  if (text.startsWith("+") || text.startsWith("-")) return formulaVectorUnarySign(sheet, text.slice(1), text[0], context);

  const directReference = formulaRefParts(text);
  if (directReference && directReference.start !== directReference.end) {
    if (formulaReferenceCellCount(directReference) > FORMULA_VECTOR_MAX_CELLS) return "#VALUE!";
    return formulaBoundedVectorMatrix(formulaRangeMatrix(sheet, text, context));
  }
  const range = formulaRangeMatrix(sheet, text, context);
  if (range && !directReference) return formulaBoundedVectorMatrix(range);
  const functionCall = formulaWholeFunctionCall(text);
  if (functionCall) return evaluateFormulaFunction(sheet, functionCall.name, splitFormulaArgs(functionCall.args), context);
  const atomic = formulaAtomicScalar(sheet, text, context);
  return atomic === undefined ? "#NAME?" : atomic;
}

function evaluateFormulaExpression(sheet, expr, context = {}) {
  let text = String(expr ?? "").trim();
  if (text === "") return undefined;
  let outer;
  while ((outer = formulaOuterParentheses(text)) !== undefined) text = outer;

  const comparison = formulaTopLevelOperators(text, [">=", "<=", "<>", "=", ">", "<"]);
  if (comparison.length > 1) return "#VALUE!";
  if (comparison.length === 1) {
    const { index, operator } = comparison[0];
    return formulaExpressionBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const concatenation = formulaTopLevelOperators(text, ["&"]);
  if (concatenation.length) {
    const { index, operator } = concatenation.at(-1);
    return formulaExpressionBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const addition = formulaTopLevelOperators(text, ["+", "-"], { binarySigns: true });
  if (addition.length) {
    const { index, operator } = addition.at(-1);
    return formulaExpressionBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const multiplication = formulaTopLevelOperators(text, ["*", "/"]);
  if (multiplication.length) {
    const { index, operator } = multiplication.at(-1);
    return formulaExpressionBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  const exponentiation = formulaTopLevelOperators(text, ["^"]);
  if (exponentiation.length) {
    const { index, operator } = exponentiation[0];
    return formulaExpressionBinary(sheet, text.slice(0, index), operator, text.slice(index + operator.length), context);
  }
  if (text.startsWith("+") || text.startsWith("-")) {
    const value = evaluateFormulaExpression(sheet, text.slice(1), context);
    return formulaApplyUnarySign(value, text[0]);
  }
  const functionCall = formulaWholeFunctionCall(text);
  if (functionCall) return evaluateFormulaFunction(sheet, functionCall.name, splitFormulaArgs(functionCall.args), context);
  const atomic = formulaAtomicScalar(sheet, text, context);
  return atomic === undefined ? "#NAME?" : atomic;
}

function formulaScalar(sheet, expr, context = {}) {
  const atomic = formulaAtomicScalar(sheet, expr, context);
  return atomic === undefined ? evaluateFormulaExpression(sheet, expr, context) : atomic;
}

function formulaNumber(value) {
  if (formulaErrorCode(value)) return value;
  if (value === true) return 1;
  if (value === false || value == null || value === "") return 0;
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function formulaText(value) {
  if (value == null) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  return String(value);
}

function formulaTruthy(value) {
  if (formulaErrorCode(value)) return false;
  if (typeof value === "boolean") return value;
  if (typeof value === "number") return value !== 0;
  const text = String(value ?? "").trim();
  if (/^TRUE$/i.test(text)) return true;
  if (/^FALSE$/i.test(text) || text === "") return false;
  return Boolean(Number(text));
}

function isFormulaMatrix(value) {
  return Array.isArray(value) && (Array.isArray(value[0]) || value.length === 0);
}

function normalizeFormulaMatrix(value) {
  if (!Array.isArray(value)) return [[value]];
  if (value.length === 0) return [];
  return Array.isArray(value[0]) ? value.map((row) => Array.isArray(row) ? row : [row]) : value.map((item) => [item]);
}

function clearFormulaSpills(workbook) {
  for (const sheet of workbook.worksheets) {
    for (const [address, cell] of sheet.store.entries()) {
      if (cell.spillParent) sheet.store.cells.delete(address);
      else {
        delete cell.spillRange;
        delete cell.spillValues;
        delete cell.spillError;
      }
    }
  }
}

function hydrateDeclaredDynamicArraySpills(workbook) {
  for (const sheet of workbook.worksheets) {
    const cells = sheet.store.entries();
    for (const [anchorAddress, anchorCell] of cells) {
      if (anchorCell.formulaType !== "dynamicArray" || !anchorCell.formula || !anchorCell.dynamicArrayRef) continue;
      let bounds;
      try {
        bounds = parseRangeAddress(anchorCell.dynamicArrayRef);
      } catch {
        continue;
      }
      if (makeCellAddress(bounds.top, bounds.left) !== anchorAddress) continue;
      markDeclaredDynamicArrayChildren(sheet, anchorAddress, anchorCell.dynamicArrayRef, bounds, cells);
    }
  }
}

function markDeclaredDynamicArrayChildren(sheet, anchorAddress, reference, bounds, cells = sheet.store.entries()) {
  const parentKey = formulaCellKey(sheet.name, anchorAddress);
  for (const [address, cell] of cells) {
    if (address === anchorAddress || cell.formula) continue;
    const position = parseCellAddress(address);
    if (position.row < bounds.top || position.row > bounds.bottom || position.col < bounds.left || position.col > bounds.right) continue;
    cell.spillParent = parentKey;
    cell.spillAnchor = anchorAddress;
    cell.spillRange = reference;
  }
}

function writeFormulaSpill(sheet, anchorAddress, matrixValue, parentKey) {
  const matrix = normalizeFormulaMatrix(matrixValue);
  const rows = matrix.length;
  const cols = Math.max(0, ...matrix.map((row) => row.length));
  if (!rows || !cols) return { value: null, range: anchorAddress, values: matrix };
  const anchor = parseCellAddress(anchorAddress);
  const spillRange = rangeToAddress({ top: anchor.row, left: anchor.col, bottom: anchor.row + rows - 1, right: anchor.col + cols - 1 });
  const collisions = [];
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (r === 0 && c === 0) continue;
      const address = makeCellAddress(anchor.row + r, anchor.col + c);
      const cell = sheet.store.get(address);
      if (!cell.spillParent && (cell.formula || cell.value != null)) collisions.push(address);
    }
  }
  const anchorCell = sheet.store.get(anchorAddress);
  if (collisions.length) {
    anchorCell.value = "#SPILL!";
    anchorCell.spillRange = spillRange;
    anchorCell.spillValues = matrix;
    anchorCell.spillError = { type: "blocked", addresses: collisions };
    return { value: "#SPILL!", range: spillRange, values: matrix, blocked: collisions };
  }
  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const address = makeCellAddress(anchor.row + r, anchor.col + c);
      const cell = sheet.store.get(address);
      cell.value = matrix[r]?.[c] ?? null;
      cell.formula = r === 0 && c === 0 ? anchorCell.formula : null;
      cell.spillParent = r === 0 && c === 0 ? undefined : parentKey;
      cell.spillAnchor = anchorAddress;
      cell.spillRange = spillRange;
      cell.spillValues = matrix;
      delete cell.spillError;
    }
  }
  return { value: matrix[0]?.[0] ?? null, range: spillRange, values: matrix };
}

function evaluateFormulaCondition(sheet, expr, context = {}) {
  return formulaTruthy(evaluateFormulaExpression(sheet, String(expr || "").trim(), context));
}

function aggregateFormulaValues(values, fnName) {
  const errors = values.map(formulaErrorCode).filter(Boolean);
  if (errors.length) return errors[0];
  const nums = values.map(formulaNumber).filter((value) => Number.isFinite(value));
  if (fnName === "COUNT") return nums.length;
  if (nums.length === 0) return 0;
  if (fnName === "AVERAGE") return nums.reduce((acc, value) => acc + value, 0) / nums.length;
  if (fnName === "MIN") return Math.min(...nums);
  if (fnName === "MAX") return Math.max(...nums);
  return nums.reduce((acc, value) => acc + value, 0);
}

function statisticalFormulaNumbers(values) {
  const error = values.map(formulaErrorCode).find(Boolean);
  if (error) return { error, numbers: [] };
  return { numbers: values.filter((value) => typeof value === "number" && Number.isFinite(value)) };
}

function roundFormulaNumber(value, digits = 0, mode = "nearest") {
  const number = formulaNumber(value);
  if (formulaErrorCode(number)) return number;
  const places = Math.trunc(Number(digits) || 0);
  const factor = 10 ** Math.min(308, Math.abs(places));
  const scaled = places >= 0 ? Math.abs(number) * factor : Math.abs(number) / factor;
  const rounded = mode === "up" ? Math.ceil(scaled) : mode === "down" ? Math.floor(scaled) : Math.round(scaled + Number.EPSILON * Math.max(1, scaled));
  const result = places >= 0 ? rounded / factor : rounded * factor;
  return Object.is(number, -0) || number < 0 ? -result : result;
}

const EXCEL_1900_DATE_EPOCH_UTC = Date.UTC(1899, 11, 31);
const EXCEL_1904_DATE_EPOCH_UTC = Date.UTC(1904, 0, 1);
const EXCEL_MAX_DATE_SERIALS = { "1900": 2_958_465, "1904": 2_957_003 };

function excelFormulaDateSystem(sheet) {
  return sheet?.workbook?.dateSystem === "1904" ? "1904" : "1900";
}

function excelMaxDateSerial(dateSystem = "1900") {
  return EXCEL_MAX_DATE_SERIALS[dateSystem === "1904" ? "1904" : "1900"];
}

function excelFormulaDateNumber(value) {
  const error = formulaErrorCode(value);
  if (error) return error;
  if (value == null || value === "" || value === false) return 0;
  if (value === true) return 1;
  const number = Number(value);
  return Number.isFinite(number) ? number : "#VALUE!";
}

function excelGregorianSerial(year, month, day = 1, dateSystem = "1900") {
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month - 1, day);
  const epoch = dateSystem === "1904" ? EXCEL_1904_DATE_EPOCH_UTC : EXCEL_1900_DATE_EPOCH_UTC;
  const days = Math.round((date.getTime() - epoch) / 86_400_000);
  return dateSystem === "1904" ? days : days + (date.getTime() >= Date.UTC(1900, 2, 1) ? 1 : 0);
}

function excelDateSerial(yearValue, monthValue, dayValue, dateSystem = "1900") {
  let year = Math.trunc(yearValue);
  const month = Math.trunc(monthValue);
  const day = Math.trunc(dayValue);
  if (year >= 0 && year <= 1899) year += 1900;
  if (!Number.isFinite(year) || year < 0 || year > 9999 || !Number.isFinite(month) || !Number.isFinite(day)) return "#NUM!";
  const normalized = new Date(0);
  normalized.setUTCHours(0, 0, 0, 0);
  normalized.setUTCFullYear(year, month - 1, 1);
  const serial = excelGregorianSerial(normalized.getUTCFullYear(), normalized.getUTCMonth() + 1, 1, dateSystem) + day - 1;
  return serial < 0 || serial > excelMaxDateSerial(dateSystem) ? "#NUM!" : serial;
}

function excelDateParts(serialValue, dateSystem = "1900") {
  const serial = Math.floor(serialValue);
  if (!Number.isFinite(serial) || serial < 0 || serial > excelMaxDateSerial(dateSystem)) return undefined;
  if (dateSystem === "1904") {
    const date = new Date(EXCEL_1904_DATE_EPOCH_UTC + serial * 86_400_000);
    return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
  }
  if (serial === 0) return { year: 1900, month: 1, day: 0 };
  if (serial === 60) return { year: 1900, month: 2, day: 29 };
  const date = new Date(EXCEL_1900_DATE_EPOCH_UTC + (serial > 60 ? serial - 1 : serial) * 86_400_000);
  return { year: date.getUTCFullYear(), month: date.getUTCMonth() + 1, day: date.getUTCDate() };
}

const EXCEL_TEXT_MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"];
const EXCEL_TEXT_DATE_TOKENS = /yyyy|yy|mmmm|mmm|mm|m|dd|d/gi;

function excelTextDateSerial(value, dateSystem = "1900") {
  if (!(value instanceof Date)) return excelFormulaDateNumber(value);
  if (!Number.isFinite(value.getTime())) return "#VALUE!";
  return excelGregorianSerial(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate(), dateSystem);
}

function excelTextDateFormat(value, formatValue, dateSystem = "1900") {
  const valueError = formulaErrorCode(value);
  const formatError = formulaErrorCode(formatValue);
  if (valueError) return valueError;
  if (formatError) return formatError;
  const serial = excelTextDateSerial(value, dateSystem);
  if (formulaErrorCode(serial)) return serial;
  const parts = excelDateParts(serial, dateSystem);
  if (!parts) return "#VALUE!";
  const format = formulaText(formatValue);
  if (!format) return "#VALUE!";
  const tokens = {
    yyyy: String(parts.year).padStart(4, "0"),
    yy: String(parts.year % 100).padStart(2, "0"),
    mmmm: EXCEL_TEXT_MONTHS[parts.month - 1],
    mmm: EXCEL_TEXT_MONTHS[parts.month - 1].slice(0, 3),
    mm: String(parts.month).padStart(2, "0"),
    m: String(parts.month),
    dd: String(parts.day).padStart(2, "0"),
    d: String(parts.day),
  };
  let output = "";
  let cursor = 0;
  for (const match of format.matchAll(EXCEL_TEXT_DATE_TOKENS)) {
    const literal = format.slice(cursor, match.index);
    if (!/^[\s/.,:_-]*$/.test(literal)) return "#VALUE!";
    output += literal + tokens[match[0].toLowerCase()];
    cursor = match.index + match[0].length;
  }
  const tail = format.slice(cursor);
  if (!/^[\s/.,:_-]*$/.test(tail) || !output) return "#VALUE!";
  return output + tail;
}

function excelDateValue(value, dateSystem = "1900") {
  const parsed = parseFormulaDateText(value);
  if (!parsed) return "#VALUE!";
  if (dateSystem === "1900" && parsed.year === 1900 && parsed.month === 2 && parsed.day === 29) return 60;
  const serial = excelGregorianSerial(parsed.year, parsed.month, parsed.day, dateSystem);
  const restored = excelDateParts(serial, dateSystem);
  if (!restored || restored.year !== parsed.year || restored.month !== parsed.month || restored.day !== parsed.day) return "#VALUE!";
  return serial;
}

function excelTimeValue(value, dateSystem = "1900") {
  const time = parseFormulaTimeText(value);
  if (time) return time.dateText && formulaErrorCode(excelDateValue(time.dateText, dateSystem)) ? "#VALUE!" : time.serial;
  return formulaErrorCode(excelDateValue(value, dateSystem)) ? "#VALUE!" : 0;
}

function excelTimePart(value, part, dateSystem = "1900") {
  if (typeof value === "string") {
    const time = parseFormulaTimeText(value);
    if (time) return time.dateText && formulaErrorCode(excelDateValue(time.dateText, dateSystem)) ? "#VALUE!" : time[part];
    return formulaErrorCode(excelDateValue(value, dateSystem)) ? "#VALUE!" : 0;
  }
  const parsed = formulaTimeParts(value);
  if (parsed) return parsed[part];
  return "#VALUE!";
}

function excelDaysInMonth(year, month, dateSystem = "1900") {
  if (dateSystem === "1900" && year === 1900 && month === 2) return 29;
  const date = new Date(0);
  date.setUTCHours(0, 0, 0, 0);
  date.setUTCFullYear(year, month, 0);
  return date.getUTCDate();
}

function excelShiftMonth(serialValue, monthsValue, endOfMonth = false, dateSystem = "1900") {
  const serial = excelFormulaDateNumber(serialValue);
  const months = excelFormulaDateNumber(monthsValue);
  if (formulaErrorCode(serial)) return serial;
  if (formulaErrorCode(months)) return months;
  const parts = excelDateParts(serial, dateSystem);
  if (!parts) return "#NUM!";
  const first = new Date(0);
  first.setUTCHours(0, 0, 0, 0);
  first.setUTCFullYear(parts.year, parts.month - 1 + Math.trunc(months), 1);
  const year = first.getUTCFullYear();
  const month = first.getUTCMonth() + 1;
  const day = endOfMonth ? excelDaysInMonth(year, month, dateSystem) : Math.min(Math.max(1, parts.day), excelDaysInMonth(year, month, dateSystem));
  return excelDateSerial(year, month, day, dateSystem);
}

function excelWeekdayIndex(serial, dateSystem = "1900") {
  const day = Math.floor(serial);
  if (dateSystem === "1904") return ((day + 5) % 7 + 7) % 7;
  const adjusted = day > 60 ? day - 1 : day;
  return ((adjusted % 7) + 7) % 7;
}

function excelHolidaySet(values = [], dateSystem = "1900") {
  const error = values.map(formulaErrorCode).find(Boolean);
  if (error) return { error, holidays: new Set() };
  const holidays = new Set();
  for (const value of values) {
    if (value == null || value === "") continue;
    const serial = excelFormulaDateNumber(value);
    if (formulaErrorCode(serial)) return { error: serial, holidays: new Set() };
    const day = Math.floor(serial);
    if (day < 0 || day > excelMaxDateSerial(dateSystem)) return { error: "#NUM!", holidays: new Set() };
    holidays.add(day);
  }
  return { holidays };
}

function excelWeekendDays(value = 1, allowAllWeekend = false) {
  const error = formulaErrorCode(value);
  if (error) return { error, weekends: new Set() };
  if (typeof value === "string") {
    const weekend = value.trim();
    if (/^[01]{7}$/.test(weekend)) {
      if (weekend === "1111111" && !allowAllWeekend) return { error: "#VALUE!", weekends: new Set() };
      const weekends = new Set();
      for (let index = 0; index < 7; index += 1) if (weekend[index] === "1") weekends.add((index + 1) % 7);
      return { weekends };
    }
    if (!/^\d+(?:\.0+)?$/.test(weekend)) return { error: "#VALUE!", weekends: new Set() };
  }
  const weekendNumber = Number(value);
  if (!Number.isInteger(weekendNumber)) return { error: "#NUM!", weekends: new Set() };
  if (weekendNumber >= 1 && weekendNumber <= 7) {
    const first = weekendNumber === 1 ? 6 : weekendNumber - 2;
    return { weekends: new Set([first, (first + 1) % 7]) };
  }
  if (weekendNumber >= 11 && weekendNumber <= 17) return { weekends: new Set([weekendNumber - 11]) };
  return { error: "#NUM!", weekends: new Set() };
}

function excelBusinessDay(serial, holidays, dateSystem = "1900", weekends = new Set([0, 6])) {
  const weekday = excelWeekdayIndex(serial, dateSystem);
  return !weekends.has(weekday) && !holidays.has(serial);
}

function excelNetworkDays(startValue, endValue, holidayValues = [], dateSystem = "1900", weekendValue = 1, allowAllWeekend = false) {
  const startNumber = excelFormulaDateNumber(startValue);
  const endNumber = excelFormulaDateNumber(endValue);
  if (formulaErrorCode(startNumber)) return startNumber;
  if (formulaErrorCode(endNumber)) return endNumber;
  const start = Math.floor(startNumber), end = Math.floor(endNumber);
  if (!excelDateParts(start, dateSystem) || !excelDateParts(end, dateSystem)) return "#NUM!";
  const holidayResult = excelHolidaySet(holidayValues, dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = excelWeekendDays(weekendValue, allowAllWeekend);
  if (weekendResult.error) return weekendResult.error;
  const direction = start <= end ? 1 : -1;
  const low = Math.min(start, end), high = Math.max(start, end);
  const total = high - low + 1;
  const fullWeeks = Math.floor(total / 7);
  let weekdays = fullWeeks * (7 - weekendResult.weekends.size);
  for (let serial = low + fullWeeks * 7; serial <= high; serial += 1) if (excelBusinessDay(serial, new Set(), dateSystem, weekendResult.weekends)) weekdays += 1;
  for (const holiday of holidayResult.holidays) if (holiday >= low && holiday <= high && !weekendResult.weekends.has(excelWeekdayIndex(holiday, dateSystem))) weekdays -= 1;
  return weekdays * direction;
}

function excelWorkday(startValue, daysValue, holidayValues = [], dateSystem = "1900", weekendValue = 1) {
  const startNumber = excelFormulaDateNumber(startValue);
  const daysNumber = excelFormulaDateNumber(daysValue);
  if (formulaErrorCode(startNumber)) return startNumber;
  if (formulaErrorCode(daysNumber)) return daysNumber;
  let serial = Math.floor(startNumber);
  const days = Math.trunc(daysNumber);
  if (!excelDateParts(serial, dateSystem) || Math.abs(days) > excelMaxDateSerial(dateSystem)) return "#NUM!";
  const holidayResult = excelHolidaySet(holidayValues, dateSystem);
  if (holidayResult.error) return holidayResult.error;
  const weekendResult = excelWeekendDays(weekendValue);
  if (weekendResult.error) return weekendResult.error;
  const direction = days < 0 ? -1 : 1;
  let remaining = Math.abs(days);
  while (remaining > 0) {
    serial += direction;
    if (!excelDateParts(serial, dateSystem)) return "#NUM!";
    if (excelBusinessDay(serial, holidayResult.holidays, dateSystem, weekendResult.weekends)) remaining -= 1;
  }
  return serial;
}

function compareFormulaValues(left, op, right) {
  const leftNum = Number(left), rightNum = Number(right);
  const numeric = Number.isFinite(leftNum) && Number.isFinite(rightNum) && String(left ?? "").trim() !== "" && String(right ?? "").trim() !== "";
  const a = numeric ? leftNum : formulaText(left);
  const b = numeric ? rightNum : formulaText(right);
  switch (op) {
    case ">=": return a >= b;
    case "<=": return a <= b;
    case "<>": return a !== b;
    case ">": return a > b;
    case "<": return a < b;
    default: return a === b;
  }
}

function formulaCriteriaArray(sheet, expr, context = {}) {
  const text = String(expr || "").trim();
  const comparison = /^(.*?)\s*(>=|<=|<>|=|>|<)\s*(.*?)$/.exec(text);
  if (comparison) {
    const leftMatrix = formulaRangeMatrix(sheet, comparison[1], context);
    const rightMatrix = formulaRangeMatrix(sheet, comparison[3], context);
    const leftValues = leftMatrix ? leftMatrix.flat() : [formulaScalar(sheet, comparison[1], context)];
    const rightValues = rightMatrix ? rightMatrix.flat() : [formulaScalar(sheet, comparison[3], context)];
    const length = Math.max(leftValues.length, rightValues.length);
    return Array.from({ length }, (_, index) => compareFormulaValues(leftValues[leftValues.length === 1 ? 0 : index], comparison[2], rightValues[rightValues.length === 1 ? 0 : index]));
  }
  const matrix = formulaRangeMatrix(sheet, text, context);
  if (matrix) return matrix.flat().map(formulaTruthy);
  return [formulaTruthy(formulaScalar(sheet, text, context))];
}

function formulaSortCompare(left, right) {
  const leftNum = Number(left), rightNum = Number(right);
  if (Number.isFinite(leftNum) && Number.isFinite(rightNum)) return leftNum - rightNum;
  return formulaText(left).localeCompare(formulaText(right));
}

function formulaMatchIndex(lookup, lookupValues = [], matchType = 0) {
  const values = lookupValues.flat ? lookupValues.flat() : lookupValues;
  const same = (value) => formulaText(value) === formulaText(lookup) || Number(value) === Number(lookup);
  const exact = values.findIndex(same);
  if (matchType === 0 || exact >= 0) return exact >= 0 ? exact + 1 : "#N/A";
  const lookupNum = Number(lookup);
  if (Number.isFinite(lookupNum)) {
    if (matchType < 0) {
      const index = values.findIndex((value) => Number(value) <= lookupNum);
      return index >= 0 ? index + 1 : "#N/A";
    }
    let best = -1;
    for (let i = 0; i < values.length; i++) if (Number(values[i]) <= lookupNum) best = i;
    return best >= 0 ? best + 1 : "#N/A";
  }
  return "#N/A";
}

function formulaWildcardRegex(pattern, { anchored = true, flags = "i" } = {}) {
  let source = "";
  const characters = Array.from(formulaText(pattern));
  for (let index = 0; index < characters.length; index++) {
    const char = characters[index];
    if (char === "~" && index + 1 < characters.length) source += characters[++index].replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    else if (char === "*") source += ".*";
    else if (char === "?") source += ".";
    else source += char.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  }
  return new RegExp(`${anchored ? "^" : ""}${source}${anchored ? "$" : ""}`, flags);
}

function formulaTextSearchPosition(findValue, withinValue, startValue = 1, { caseSensitive = false, wildcard = false } = {}) {
  for (const value of [findValue, withinValue, startValue]) {
    const error = formulaErrorCode(value);
    if (error) return error;
  }
  const findText = formulaText(findValue);
  const characters = Array.from(formulaText(withinValue));
  const start = formulaNumber(startValue);
  if (formulaErrorCode(start)) return start;
  if (!Number.isInteger(start) || start < 1 || start > characters.length) return "#VALUE!";
  const tail = characters.slice(start - 1).join("");
  if (wildcard) {
    const match = formulaWildcardRegex(findText, { anchored: false, flags: "iu" }).exec(tail);
    if (!match) return "#VALUE!";
    return start + Array.from(tail.slice(0, match.index)).length;
  }
  const needle = caseSensitive ? findText : findText.toLocaleLowerCase();
  const haystack = caseSensitive ? tail : tail.toLocaleLowerCase();
  const index = haystack.indexOf(needle);
  return index < 0 ? "#VALUE!" : start + Array.from(tail.slice(0, index)).length;
}

function formulaXmatchIndex(lookup, lookupValues = [], matchMode = 0, searchMode = 1) {
  const values = lookupValues.flat ? lookupValues.flat() : lookupValues;
  if (![0, -1, 1, 2].includes(matchMode) || ![1, -1, 2, -2].includes(searchMode)) return "#VALUE!";
  const indexes = Array.from({ length: values.length }, (_, index) => index);
  if (searchMode === -1 || searchMode === -2) indexes.reverse();
  const lookupNumber = Number(lookup);
  const numericLookup = Number.isFinite(lookupNumber) && formulaText(lookup).trim() !== "";
  const exact = (value) => {
    const valueNumber = Number(value);
    if (numericLookup && Number.isFinite(valueNumber) && formulaText(value).trim() !== "") return valueNumber === lookupNumber;
    return formulaText(value).toLocaleLowerCase() === formulaText(lookup).toLocaleLowerCase();
  };
  const wildcard = matchMode === 2 ? formulaWildcardRegex(lookup) : undefined;
  const found = indexes.find((index) => wildcard ? wildcard.test(formulaText(values[index])) : exact(values[index]));
  if (found != null) return found + 1;
  if (matchMode === 0 || matchMode === 2) return "#N/A";
  const comparable = indexes.map((index) => ({ index, value: values[index], number: Number(values[index]) }))
    .filter((item) => numericLookup ? Number.isFinite(item.number) : true)
    .filter((item) => matchMode < 0
      ? (numericLookup ? item.number < lookupNumber : formulaText(item.value).localeCompare(formulaText(lookup), undefined, { sensitivity: "base" }) < 0)
      : (numericLookup ? item.number > lookupNumber : formulaText(item.value).localeCompare(formulaText(lookup), undefined, { sensitivity: "base" }) > 0));
  if (!comparable.length) return "#N/A";
  comparable.sort((left, right) => {
    const delta = numericLookup ? left.number - right.number : formulaText(left.value).localeCompare(formulaText(right.value), undefined, { sensitivity: "base" });
    return matchMode < 0 ? -delta : delta;
  });
  return comparable[0].index + 1;
}

function uniqueFormulaRows(matrix) {
  const seen = new Set();
  const rows = [];
  for (const row of normalizeFormulaMatrix(matrix)) {
    const key = JSON.stringify(row.map((value) => [typeof value, value]));
    if (seen.has(key)) continue;
    seen.add(key);
    rows.push(row);
  }
  return rows;
}

function evaluateFormulaFunction(sheet, fnName, args, context = {}) {
  const values = (parts = args) => parts.flatMap((part) => formulaReferenceValues(sheet, part, context));
  const dateSystem = excelFormulaDateSystem(sheet);
  const financialHelpers = {
    errorCode: formulaErrorCode,
    dateNumber: (value) => {
      if (!(value instanceof Date)) return excelFormulaDateNumber(value);
      if (!Number.isFinite(value.getTime())) return "#VALUE!";
      return excelGregorianSerial(value.getUTCFullYear(), value.getUTCMonth() + 1, value.getUTCDate(), dateSystem);
    },
    isValidDate: (serial) => Boolean(excelDateParts(serial, dateSystem)),
  };
  const scalar = (index, fallback = undefined) => {
    const value = formulaScalar(sheet, args[index], context);
    return value === undefined ? fallback : value;
  };
  const criteriaRange = (part) => {
    const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, part, context) || [[formulaScalar(sheet, part, context)]]);
    const columns = Math.max(0, ...matrix.map((row) => row.length));
    return { values: matrix.flat(), rows: matrix.length, columns, rectangular: matrix.every((row) => row.length === columns) };
  };
  const sameCriteriaShape = (left, right) => left.rectangular && right.rectangular && left.rows === right.rows && left.columns === right.columns;
  switch (fnName) {
    case "SUM":
    case "AVERAGE":
    case "MIN":
    case "MAX":
    case "COUNT":
      return aggregateFormulaValues(values(), fnName);
    case "COUNTA": return values().filter((value) => value !== null && value !== undefined).length;
    case "COUNTBLANK": {
      if (args.length !== 1) return "#VALUE!";
      const matrix = formulaRangeMatrix(sheet, args[0], context);
      if (!matrix) return "#VALUE!";
      return matrix.flat().filter((value) => value == null || value === "").length;
    }
    case "ABS": return Math.abs(formulaNumber(scalar(0, 0)));
    case "ROUND": return roundFormulaNumber(scalar(0, 0), scalar(1, 0));
    case "ROUNDUP": return roundFormulaNumber(scalar(0, 0), scalar(1, 0), "up");
    case "ROUNDDOWN": return roundFormulaNumber(scalar(0, 0), scalar(1, 0), "down");
    case "MEDIAN": {
      const stats = statisticalFormulaNumbers(values());
      if (stats.error) return stats.error;
      if (!stats.numbers.length) return "#NUM!";
      stats.numbers.sort((left, right) => left - right);
      const middle = Math.floor(stats.numbers.length / 2);
      return stats.numbers.length % 2 ? stats.numbers[middle] : (stats.numbers[middle - 1] + stats.numbers[middle]) / 2;
    }
    case "LARGE":
    case "SMALL": {
      const stats = statisticalFormulaNumbers(values([args[0]]));
      if (stats.error) return stats.error;
      const rank = Math.trunc(formulaNumber(scalar(1, 0)));
      if (rank < 1 || rank > stats.numbers.length) return "#NUM!";
      stats.numbers.sort((left, right) => fnName === "LARGE" ? right - left : left - right);
      return stats.numbers[rank - 1];
    }
    case "RANK":
    case "RANK.EQ": {
      const number = formulaNumber(scalar(0, 0));
      if (formulaErrorCode(number)) return number;
      const stats = statisticalFormulaNumbers(values([args[1]]));
      if (stats.error) return stats.error;
      if (!stats.numbers.includes(number)) return "#N/A";
      const ascending = formulaTruthy(scalar(2, false));
      return 1 + stats.numbers.filter((value) => ascending ? value < number : value > number).length;
    }
    case "MODE":
    case "MODE.SNGL": {
      const stats = statisticalFormulaNumbers(values());
      if (stats.error) return stats.error;
      const counts = new Map();
      for (const number of stats.numbers) counts.set(number, (counts.get(number) || 0) + 1);
      const count = Math.max(0, ...counts.values());
      if (count <= 1) return "#N/A";
      return Math.min(...[...counts].filter(([, occurrences]) => occurrences === count).map(([number]) => number));
    }
    case "INT": return Math.floor(formulaNumber(scalar(0, 0)));
    case "CEILING": return Math.ceil(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "FLOOR": return Math.floor(formulaNumber(scalar(0, 0)) / Math.max(1, formulaNumber(scalar(1, 1)))) * Math.max(1, formulaNumber(scalar(1, 1)));
    case "SLN": return args.length === 3
      ? calculateSln({ cost: scalar(0), salvage: scalar(1), life: scalar(2) }, financialHelpers)
      : "#VALUE!";
    case "DB": return args.length >= 4 && args.length <= 5
      ? calculateDb({ cost: scalar(0), salvage: scalar(1), life: scalar(2), period: scalar(3), month: scalar(4, 12) }, financialHelpers)
      : "#VALUE!";
    case "DDB": return args.length >= 4 && args.length <= 5
      ? calculateDdb({ cost: scalar(0), salvage: scalar(1), life: scalar(2), period: scalar(3), factor: scalar(4, 2) }, financialHelpers)
      : "#VALUE!";
    case "PMT": return args.length >= 3 && args.length <= 5
      ? calculatePmt({ rate: scalar(0), nper: scalar(1), pv: scalar(2), fv: scalar(3, 0), type: scalar(4, 0) }, financialHelpers)
      : "#VALUE!";
    case "IPMT": return args.length >= 4 && args.length <= 6
      ? calculateIpmt({ rate: scalar(0), per: scalar(1), nper: scalar(2), pv: scalar(3), fv: scalar(4, 0), type: scalar(5, 0) }, financialHelpers)
      : "#VALUE!";
    case "PPMT": return args.length >= 4 && args.length <= 6
      ? calculatePpmt({ rate: scalar(0), per: scalar(1), nper: scalar(2), pv: scalar(3), fv: scalar(4, 0), type: scalar(5, 0) }, financialHelpers)
      : "#VALUE!";
    case "CUMIPMT": return args.length === 6
      ? calculateCumipmt({ rate: scalar(0), nper: scalar(1), pv: scalar(2), startPeriod: scalar(3), endPeriod: scalar(4), type: scalar(5) }, financialHelpers)
      : "#VALUE!";
    case "CUMPRINC": return args.length === 6
      ? calculateCumprinc({ rate: scalar(0), nper: scalar(1), pv: scalar(2), startPeriod: scalar(3), endPeriod: scalar(4), type: scalar(5) }, financialHelpers)
      : "#VALUE!";
    case "PV": return args.length >= 3 && args.length <= 5
      ? calculatePv({ rate: scalar(0), nper: scalar(1), pmt: scalar(2), fv: scalar(3, 0), type: scalar(4, 0) }, financialHelpers)
      : "#VALUE!";
    case "FV": return args.length >= 3 && args.length <= 5
      ? calculateFv({ rate: scalar(0), nper: scalar(1), pmt: scalar(2), pv: scalar(3, 0), type: scalar(4, 0) }, financialHelpers)
      : "#VALUE!";
    case "NPER": return args.length >= 3 && args.length <= 5
      ? calculateNper({ rate: scalar(0), pmt: scalar(1), pv: scalar(2), fv: scalar(3, 0), type: scalar(4, 0) }, financialHelpers)
      : "#VALUE!";
    case "RATE": return args.length >= 3 && args.length <= 6
      ? calculateRate({ nper: scalar(0), pmt: scalar(1), pv: scalar(2), fv: scalar(3, 0), type: scalar(4, 0), guess: args.length === 6 ? scalar(5) : undefined }, financialHelpers)
      : "#VALUE!";
    case "NPV": return args.length >= 2
      ? calculateNpv({ rate: scalar(0), cashFlows: values(args.slice(1)) }, financialHelpers)
      : "#VALUE!";
    case "MIRR": return args.length === 3
      ? calculateMirr({ cashFlows: values([args[0]]), financeRate: scalar(1), reinvestRate: scalar(2) }, financialHelpers)
      : "#VALUE!";
    case "XNPV": return args.length === 3
      ? calculateXnpv({ rate: scalar(0), cashFlows: values([args[1]]), dates: values([args[2]]) }, financialHelpers)
      : "#VALUE!";
    case "IRR": return args.length >= 1 && args.length <= 2
      ? calculateIrr({ cashFlows: values([args[0]]), guess: args.length === 2 ? scalar(1) : undefined }, financialHelpers)
      : "#VALUE!";
    case "XIRR": return args.length >= 2 && args.length <= 3
      ? calculateXirr({ cashFlows: values([args[0]]), dates: values([args[1]]), guess: args.length === 3 ? scalar(2) : undefined }, financialHelpers)
      : "#VALUE!";
    case "DATE": {
      const parts = [0, 1, 2].map((index) => excelFormulaDateNumber(scalar(index, 0)));
      return parts.find(formulaErrorCode) || excelDateSerial(parts[0], parts[1], parts[2], dateSystem);
    }
    case "DATEVALUE": return args.length === 1 ? excelDateValue(scalar(0), dateSystem) : "#VALUE!";
    case "TIMEVALUE": return args.length === 1 ? excelTimeValue(scalar(0), dateSystem) : "#VALUE!";
    case "TIME": {
      if (args.length !== 3) return "#VALUE!";
      const parts = [0, 1, 2].map((index) => excelFormulaDateNumber(scalar(index)));
      const error = parts.find(formulaErrorCode);
      if (error) return error;
      return formulaTimeSerial(...parts) ?? "#NUM!";
    }
    case "YEAR":
    case "MONTH":
    case "DAY": {
      const serial = excelFormulaDateNumber(scalar(0, 0));
      if (formulaErrorCode(serial)) return serial;
      const parts = excelDateParts(serial, dateSystem);
      return parts ? parts[fnName.toLowerCase()] : "#NUM!";
    }
    case "HOUR": return args.length === 1 ? excelTimePart(scalar(0), "hour", dateSystem) : "#VALUE!";
    case "MINUTE": return args.length === 1 ? excelTimePart(scalar(0), "minute", dateSystem) : "#VALUE!";
    case "SECOND": return args.length === 1 ? excelTimePart(scalar(0), "second", dateSystem) : "#VALUE!";
    case "EDATE": return excelShiftMonth(scalar(0, 0), scalar(1, 0), false, dateSystem);
    case "EOMONTH": return excelShiftMonth(scalar(0, 0), scalar(1, 0), true, dateSystem);
    case "DAYS": {
      const end = excelFormulaDateNumber(scalar(0, 0));
      const start = excelFormulaDateNumber(scalar(1, 0));
      if (formulaErrorCode(end)) return end;
      if (formulaErrorCode(start)) return start;
      return excelDateParts(end, dateSystem) && excelDateParts(start, dateSystem) ? Math.floor(end) - Math.floor(start) : "#NUM!";
    }
    case "WEEKDAY": {
      const serial = excelFormulaDateNumber(scalar(0, 0));
      const returnTypeValue = excelFormulaDateNumber(scalar(1, 1));
      if (formulaErrorCode(serial)) return serial;
      if (formulaErrorCode(returnTypeValue)) return returnTypeValue;
      const returnType = Math.trunc(returnTypeValue);
      if (!excelDateParts(serial, dateSystem)) return "#NUM!";
      const weekday = excelWeekdayIndex(serial, dateSystem);
      if (returnType === 1) return weekday + 1;
      if (returnType === 2 || returnType === 11) return (weekday + 6) % 7 + 1;
      if (returnType === 3) return (weekday + 6) % 7;
      if (returnType >= 12 && returnType <= 17) return (weekday - (returnType - 10) + 7) % 7 + 1;
      return "#NUM!";
    }
    case "NETWORKDAYS": return excelNetworkDays(scalar(0, 0), scalar(1, 0), args[2] == null ? [] : values([args[2]]), dateSystem);
    case "WORKDAY": return excelWorkday(scalar(0, 0), scalar(1, 0), args[2] == null ? [] : values([args[2]]), dateSystem);
    case "NETWORKDAYS.INTL": return excelNetworkDays(scalar(0, 0), scalar(1, 0), args[3] == null ? [] : values([args[3]]), dateSystem, scalar(2, 1), true);
    case "WORKDAY.INTL": return excelWorkday(scalar(0, 0), scalar(1, 0), args[3] == null ? [] : values([args[3]]), dateSystem, scalar(2, 1));
    case "IF": return evaluateFormulaCondition(sheet, args[0], context) ? scalar(1, true) : scalar(2, false);
    case "IFS": {
      if (args.length < 2 || args.length % 2 !== 0) return "#VALUE!";
      for (let index = 0; index < args.length; index += 2) {
        if (evaluateFormulaCondition(sheet, args[index], context)) return scalar(index + 1);
      }
      return "#N/A";
    }
    case "SWITCH": {
      if (args.length < 3) return "#VALUE!";
      const expression = scalar(0);
      const expressionError = formulaErrorCode(expression);
      if (expressionError) return expressionError;
      const hasDefault = args.length % 2 === 0;
      const pairEnd = hasDefault ? args.length - 1 : args.length;
      for (let index = 1; index < pairEnd; index += 2) {
        const candidate = scalar(index);
        const candidateError = formulaErrorCode(candidate);
        if (candidateError) return candidateError;
        if (compareFormulaValues(expression, "=", candidate)) return scalar(index + 1);
      }
      return hasDefault ? scalar(args.length - 1) : "#N/A";
    }
    case "IFERROR": { const value = scalar(0); return formulaErrorCode(value) ? scalar(1, "") : value; }
    case "IFNA": { const value = scalar(0); return formulaErrorCode(value) === "#N/A" ? scalar(1, "") : value; }
    case "AND": return args.every((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "OR": return args.some((arg) => evaluateFormulaCondition(sheet, arg, context));
    case "NOT": return !evaluateFormulaCondition(sheet, args[0], context);
    case "ISNUMBER": { const value = scalar(0); return typeof value === "number" && Number.isFinite(value); }
    case "ISTEXT": { const value = scalar(0); return typeof value === "string" && !formulaErrorCode(value); }
    case "ISBLANK": { const value = scalar(0); return value == null; }
    case "ISERROR": return Boolean(formulaErrorCode(scalar(0)));
    case "ISNA": return formulaErrorCode(scalar(0)) === "#N/A";
    case "ISERR": { const error = formulaErrorCode(scalar(0)); return Boolean(error && error !== "#N/A"); }
    case "NA": return "#N/A";
    case "CONCAT":
    case "CONCATENATE": return values().map(formulaText).join("");
    case "TEXTJOIN": {
      const delimiter = formulaText(scalar(0, ""));
      const ignoreEmpty = formulaTruthy(scalar(1, false));
      const joined = values(args.slice(2)).map(formulaText).filter((value) => !ignoreEmpty || value !== "");
      return joined.join(delimiter);
    }
    case "LEFT": return formulaText(scalar(0, "")).slice(0, Number(scalar(1, 1)) || 1);
    case "RIGHT": { const text = formulaText(scalar(0, "")); const count = Number(scalar(1, 1)) || 1; return text.slice(Math.max(0, text.length - count)); }
    case "MID": { const text = formulaText(scalar(0, "")); const start = Math.max(1, Number(scalar(1, 1)) || 1); const count = Math.max(0, Number(scalar(2, 1)) || 1); return text.slice(start - 1, start - 1 + count); }
    case "LEN": return formulaText(scalar(0, "")).length;
    case "SEARCH": {
      if (args.length < 2 || args.length > 3) return "#VALUE!";
      return formulaTextSearchPosition(scalar(0), scalar(1), scalar(2, 1), { wildcard: true });
    }
    case "FIND": {
      if (args.length < 2 || args.length > 3) return "#VALUE!";
      return formulaTextSearchPosition(scalar(0), scalar(1), scalar(2, 1), { caseSensitive: true });
    }
    case "UPPER": return formulaText(scalar(0, "")).toUpperCase();
    case "LOWER": return formulaText(scalar(0, "")).toLowerCase();
    case "TRIM": return formulaText(scalar(0, "")).trim().replace(/\s+/g, " ");
    case "VALUE": {
      if (args.length !== 1) return "#VALUE!";
      const value = scalar(0);
      const error = formulaErrorCode(value);
      if (error) return error;
      return parseFormulaNumberText(value) ?? "#VALUE!";
    }
    case "TEXT": return args.length === 2 ? excelTextDateFormat(scalar(0), scalar(1), dateSystem) : "#VALUE!";
    case "COUNTIF": { if (args.length < 2) return "#VALUE!"; const range = values([args[0]]); const criteria = scalar(1, ""); return range.filter((value) => matchesFormulaCriteria(value, criteria)).length; }
    case "COUNTIFS": {
      if (args.length < 2 || args.length % 2 !== 0) return "#VALUE!";
      const pairs = [];
      for (let i = 0; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      const firstRange = pairs[0]?.range;
      if (!firstRange || pairs.some((pair) => !sameCriteriaShape(pair.range, firstRange))) return "#VALUE!";
      const length = firstRange.values.length;
      let count = 0;
      for (let index = 0; index < length; index++) if (pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria))) count += 1;
      return count;
    }
    case "SEQUENCE": {
      const rows = Math.max(1, Math.floor(formulaNumber(scalar(0, 1))) || 1);
      const cols = Math.max(1, Math.floor(formulaNumber(scalar(1, 1))) || 1);
      const start = formulaNumber(scalar(2, 1));
      const step = formulaNumber(scalar(3, 1));
      return Array.from({ length: rows }, (_, row) => Array.from({ length: cols }, (_, col) => start + (row * cols + col) * step));
    }
    case "TRANSPOSE": {
      const matrix = formulaRangeMatrix(sheet, args[0], context) || [];
      const rows = matrix.length;
      const cols = Math.max(0, ...matrix.map((row) => row.length));
      return Array.from({ length: cols }, (_, col) => Array.from({ length: rows }, (_, row) => matrix[row]?.[col] ?? null));
    }
    case "FILTER": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const include = formulaCriteriaArray(sheet, args[1], context);
      const rows = matrix.filter((_, index) => formulaTruthy(include[index]));
      if (rows.length) return rows;
      return [[args[2] ? scalar(2, "") : "#CALC!"]];
    }
    case "UNIQUE": {
      return uniqueFormulaRows(formulaRangeMatrix(sheet, args[0], context) || []);
    }
    case "SORT": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const sortIndex = Math.max(0, Math.floor(formulaNumber(scalar(1, 1))) - 1);
      const sortOrder = formulaNumber(scalar(2, 1)) < 0 ? -1 : 1;
      return [...matrix].sort((a, b) => formulaSortCompare(a[sortIndex], b[sortIndex]) * sortOrder);
    }
    case "TAKE": {
      let matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const rows = String(args[1] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(1, 0)));
      const columns = String(args[2] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(2, 0)));
      if (!matrix.length || rows === 0 || columns === 0 || (rows == null && columns == null)) return "#CALC!";
      if (rows != null) matrix = rows > 0 ? matrix.slice(0, rows) : matrix.slice(Math.max(0, matrix.length + rows));
      if (columns != null) matrix = matrix.map((row) => columns > 0 ? row.slice(0, columns) : row.slice(Math.max(0, row.length + columns)));
      return matrix.length && matrix.some((row) => row.length) ? matrix : "#CALC!";
    }
    case "DROP": {
      let matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const rows = String(args[1] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(1, 0)));
      const columns = String(args[2] ?? "").trim() === "" ? undefined : Math.trunc(formulaNumber(scalar(2, 0)));
      if (!matrix.length || rows === 0 || columns === 0 || (rows == null && columns == null)) return "#CALC!";
      if (rows != null) matrix = rows >= 0 ? matrix.slice(rows) : matrix.slice(0, Math.max(0, matrix.length + rows));
      if (columns != null) matrix = matrix.map((row) => columns >= 0 ? row.slice(columns) : row.slice(0, Math.max(0, row.length + columns)));
      return matrix.length && matrix.some((row) => row.length) ? matrix : "#CALC!";
    }
    case "CHOOSECOLS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const indexes = args.slice(1).map((_, index) => Math.trunc(formulaNumber(scalar(index + 1, 0)))).map((value) => value > 0 ? value - 1 : width + value);
      if (!matrix.length || !indexes.length || indexes.some((index) => index < 0 || index >= width)) return "#VALUE!";
      return matrix.map((row) => indexes.map((index) => row[index] ?? null));
    }
    case "CHOOSEROWS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const indexes = args.slice(1).map((_, index) => Math.trunc(formulaNumber(scalar(index + 1, 0)))).map((value) => value > 0 ? value - 1 : matrix.length + value);
      if (!matrix.length || !indexes.length || indexes.some((index) => index < 0 || index >= matrix.length)) return "#VALUE!";
      return indexes.map((index) => [...matrix[index]]);
    }
    case "TOCOL":
    case "TOROW": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const ignore = Number(scalar(1, 0));
      if (!matrix.length || !Number.isInteger(ignore) || ignore < 0 || ignore > 3) return "#VALUE!";
      const height = matrix.length;
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const scanByColumn = formulaTruthy(scalar(2, false));
      const flattened = [];
      for (let outer = 0; outer < (scanByColumn ? width : height); outer++) {
        for (let inner = 0; inner < (scanByColumn ? height : width); inner++) {
          const value = scanByColumn ? matrix[inner]?.[outer] : matrix[outer]?.[inner];
          const blank = value == null || value === "";
          const error = Boolean(formulaErrorCode(value));
          if ((ignore & 1) && blank) continue;
          if ((ignore & 2) && error) continue;
          flattened.push(blank ? 0 : value);
        }
      }
      if (!flattened.length) return "#CALC!";
      return fnName === "TOCOL" ? flattened.map((value) => [value]) : [flattened];
    }
    case "WRAPROWS":
    case "WRAPCOLS": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      const width = Math.max(0, ...matrix.map((row) => row.length));
      if (!matrix.length || (matrix.length > 1 && width > 1)) return "#VALUE!";
      const count = Number(scalar(1, 0));
      if (!Number.isInteger(count)) return "#VALUE!";
      if (count < 1) return "#NUM!";
      const vector = matrix.length === 1 ? [...matrix[0]] : matrix.map((row) => row[0]);
      const pad = String(args[2] ?? "").trim() === "" ? "#N/A" : scalar(2, "#N/A");
      if (fnName === "WRAPROWS") {
        const rows = Math.ceil(vector.length / count);
        return Array.from({ length: rows }, (_, row) => Array.from({ length: count }, (_, col) => {
          const index = row * count + col;
          return index < vector.length ? vector[index] : pad;
        }));
      }
      const columns = Math.ceil(vector.length / count);
      return Array.from({ length: count }, (_, row) => Array.from({ length: columns }, (_, col) => {
        const index = col * count + row;
        return index < vector.length ? vector[index] : pad;
      }));
    }
    case "HSTACK":
    case "VSTACK": {
      if (!args.length) return "#VALUE!";
      const matrices = args.map((arg, index) => normalizeFormulaMatrix(formulaRangeMatrix(sheet, arg, context) || [[scalar(index)]]));
      const normalizedValue = (value) => value == null || value === "" ? 0 : value;
      if (fnName === "HSTACK") {
        const height = Math.max(0, ...matrices.map((matrix) => matrix.length));
        return Array.from({ length: height }, (_, row) => matrices.flatMap((matrix) => {
          const width = Math.max(0, ...matrix.map((values) => values.length));
          return Array.from({ length: width }, (_, col) => row < matrix.length ? normalizedValue(matrix[row]?.[col]) : "#N/A");
        }));
      }
      const width = Math.max(0, ...matrices.flatMap((matrix) => matrix.map((row) => row.length)));
      return matrices.flatMap((matrix) => matrix.map((row) => Array.from({ length: width }, (_, col) => col < row.length ? normalizedValue(row[col]) : "#N/A")));
    }
    case "EXPAND": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      if (!matrix.length) return "#VALUE!";
      const height = matrix.length;
      const width = Math.max(0, ...matrix.map((row) => row.length));
      const rows = String(args[1] ?? "").trim() === "" ? height : Number(scalar(1, height));
      const columns = String(args[2] ?? "").trim() === "" ? width : Number(scalar(2, width));
      if (!Number.isInteger(rows) || !Number.isInteger(columns)) return "#VALUE!";
      if (rows < height || columns < width) return "#VALUE!";
      const pad = String(args[3] ?? "").trim() === "" ? "#N/A" : scalar(3, "#N/A");
      return Array.from({ length: rows }, (_, row) => Array.from({ length: columns }, (_, col) => {
        if (row >= height || col >= width) return pad;
        const value = matrix[row]?.[col];
        return value == null || value === "" ? 0 : value;
      }));
    }
    case "SUMIF": {
      if (args.length < 2) return "#VALUE!";
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const sumRange = args[2] ? values([args[2]]) : range;
      let sum = 0;
      for (let index = 0; index < range.length; index += 1) {
        if (!matchesFormulaCriteria(range[index], criteria)) continue;
        const number = formulaNumber(sumRange[index]);
        if (formulaErrorCode(number)) return number;
        sum += number;
      }
      return sum;
    }
    case "SUMIFS": {
      if (args.length < 3 || args.length % 2 === 0) return "#VALUE!";
      const sumRange = criteriaRange(args[0]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => !sameCriteriaShape(pair.range, sumRange))) return "#VALUE!";
      let sum = 0;
      for (let index = 0; index < sumRange.values.length; index += 1) {
        if (!pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria))) continue;
        const number = formulaNumber(sumRange.values[index]);
        if (formulaErrorCode(number)) return number;
        sum += number;
      }
      return sum;
    }
    case "AVERAGEIF": {
      if (args.length < 2) return "#VALUE!";
      const range = values([args[0]]);
      const criteria = scalar(1, "");
      const averageRange = args[2] ? values([args[2]]) : range;
      if (averageRange.length !== range.length) return "#VALUE!";
      const matched = averageRange.filter((_, index) => matchesFormulaCriteria(range[index], criteria));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => value !== "" && value != null && Number.isFinite(Number(value))).map(Number);
      return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
    }
    case "AVERAGEIFS": {
      if (args.length < 3 || args.length % 2 === 0) return "#VALUE!";
      const averageRange = criteriaRange(args[0]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => !sameCriteriaShape(pair.range, averageRange))) return "#VALUE!";
      const matched = averageRange.values.filter((_, index) => pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria)));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => value !== "" && value != null && Number.isFinite(Number(value))).map(Number);
      return numbers.length ? numbers.reduce((sum, value) => sum + value, 0) / numbers.length : "#DIV/0!";
    }
    case "MINIFS":
    case "MAXIFS": {
      if (args.length < 3 || args.length % 2 === 0) return "#VALUE!";
      const valueRange = criteriaRange(args[0]);
      const pairs = [];
      for (let i = 1; i < args.length; i += 2) pairs.push({ range: criteriaRange(args[i]), criteria: scalar(i + 1, "") });
      if (pairs.some((pair) => !sameCriteriaShape(pair.range, valueRange))) return "#VALUE!";
      const matched = valueRange.values.filter((_, index) => pairs.every((pair) => matchesFormulaCriteria(pair.range.values[index], pair.criteria)));
      const error = matched.map(formulaErrorCode).find(Boolean);
      if (error) return error;
      const numbers = matched.filter((value) => typeof value === "number" && Number.isFinite(value));
      if (!numbers.length) return 0;
      return fnName === "MINIFS" ? Math.min(...numbers) : Math.max(...numbers);
    }
    case "SUMPRODUCT": {
      const matrices = args.map((arg) => {
        const value = evaluateFormulaVectorExpression(sheet, arg, context);
        return isFormulaMatrix(value) ? value : [[value]];
      });
      if (!matrices.length) return 0;
      const shape = (matrix) => [matrix.length, matrix[0]?.length || 0];
      const [rows, cols] = shape(matrices[0]);
      if (matrices.some((matrix) => {
        const [matrixRows, matrixCols] = shape(matrix);
        return matrixRows !== rows || matrixCols !== cols || matrix.some((row) => row.length !== cols);
      })) return "#VALUE!";
      const arrays = matrices.map((matrix) => matrix.flat());
      const length = rows * cols;
      const error = arrays.flat().map(formulaErrorCode).find(Boolean);
      if (error) return error;
      return Array.from({ length }, (_, index) => arrays.reduce((product, array) => product * formulaNumber(array[index]), 1)).reduce((sum, value) => sum + value, 0);
    }
    case "INDEX": {
      const matrix = normalizeFormulaMatrix(formulaRangeMatrix(sheet, args[0], context) || []);
      if (!matrix.length) return "#REF!";
      const rowValue = scalar(1, 1);
      const rowError = formulaErrorCode(rowValue);
      if (rowError) return rowError;
      const columnValue = scalar(2, 1);
      const columnError = formulaErrorCode(columnValue);
      if (columnError) return columnError;
      const rowIndex = Math.max(1, Math.floor(formulaNumber(rowValue)) || 1) - 1;
      const colIndex = Math.max(1, Math.floor(formulaNumber(columnValue)) || 1) - 1;
      return matrix[rowIndex]?.[colIndex] ?? "#REF!";
    }
    case "MATCH": {
      const lookup = scalar(0, "");
      const matchValues = formulaRangeMatrix(sheet, args[1], context) || [];
      const matchType = Math.sign(formulaNumber(scalar(2, 0)) || 0);
      return formulaMatchIndex(lookup, matchValues.flat(), matchType);
    }
    case "XMATCH": {
      const lookup = scalar(0, "");
      const matchValues = formulaRangeMatrix(sheet, args[1], context) || [];
      return formulaXmatchIndex(lookup, matchValues.flat(), formulaNumber(scalar(2, 0)), formulaNumber(scalar(3, 1)));
    }
    case "VLOOKUP": {
      const lookup = scalar(0, "");
      const matrix = formulaRangeMatrix(sheet, args[1], context) || [];
      const colIndex = Math.max(1, Number(scalar(2, 1)) || 1) - 1;
      const row = matrix.find((item) => formulaText(item[0]) === formulaText(lookup) || Number(item[0]) === Number(lookup));
      return row ? row[colIndex] ?? "#N/A" : "#N/A";
    }
    case "HLOOKUP": {
      const lookup = scalar(0, "");
      const matrix = formulaRangeMatrix(sheet, args[1], context) || [];
      const rowIndex = Math.floor(Number(scalar(2, 1)) || 1) - 1;
      if (rowIndex < 0 || rowIndex >= matrix.length) return "#REF!";
      const header = matrix[0] || [];
      const equals = (value) => formulaText(value).toLowerCase() === formulaText(lookup).toLowerCase()
        || (formulaText(value).trim() !== "" && formulaText(lookup).trim() !== "" && Number.isFinite(Number(value)) && Number(value) === Number(lookup));
      let columnIndex = header.findIndex(equals);
      if (columnIndex < 0 && formulaTruthy(scalar(3, true))) {
        for (let index = 0; index < header.length; index++) if (compareFormulaValues(header[index], "<=", lookup)) columnIndex = index;
      }
      return columnIndex >= 0 ? matrix[rowIndex]?.[columnIndex] ?? "#N/A" : "#N/A";
    }
    case "XLOOKUP": {
      const lookup = scalar(0, "");
      const lookupValues = values([args[1]]);
      const returnValues = values([args[2]]);
      const index = lookupValues.findIndex((value) => formulaText(value) === formulaText(lookup) || Number(value) === Number(lookup));
      return index >= 0 ? returnValues[index] : scalar(3, "#N/A");
    }
    default:
      return "#NAME?";
  }
}

function evaluateFormula(sheet, formula, address, context = {}) {
  const raw = String(formula || "").trim();
  if (!raw.startsWith("=")) return raw;
  // Excel persists post-2010 worksheet functions with compatibility prefixes.
  // They are package syntax, not part of the agent-facing formula language.
  const expr = raw.slice(1).trim().replace(/_xlfn\.(?:_xlws\.)?/gi, "");
  const evaluationContext = address && !context.formulaAddress ? { ...context, formulaAddress: address } : context;
  return evaluateFormulaExpression(sheet, expr, evaluationContext);
}

export {
  buildWorkbookFormulaGraph,
  clearFormulaSpills,
  evaluateFormula,
  evaluateFormulaCondition,
  formulaCellKey,
  formulaErrorCode,
  formulaGraphRecords,
  formulaRefParts,
  formulaReferences,
  formulaScalar,
  formulaText,
  hydrateDeclaredDynamicArraySpills,
  isFormulaMatrix,
  markDeclaredDynamicArrayChildren,
  publicFormulaNode,
  writeFormulaSpill,
};
