import { inspectOoxmlPackage, patchOoxmlPackage } from "../ooxml/package.mjs";
import { normalizeDocxSettings } from "../ooxml/docx-settings.mjs";
import { effectiveDocxRunStyle, mergeDocxRunStyleCascade, normalizeDocxRunStyle, normalizeDocxThemeConfig } from "../ooxml/docx-run-styles.mjs";
import { validateDocxCommentPackageSemantics } from "../ooxml/docx-comments.mjs";
import { validateDocxLinkPackageSemantics } from "../ooxml/docx-links.mjs";
import { normalizeDocxBibliographySource, validateDocxBibliographyPackageSemantics } from "../ooxml/docx-bibliography.mjs";
import { normalizeDocxSectionSettings, planDocxHeaderFooterSections, resolveDocxPageHeaderFooter } from "../ooxml/docx-sections.mjs";
import { normalizeDocumentPictureBullet } from "../ooxml/docx-numbering.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { officeFontFamilies } from "../shared/font-design-metrics.mjs";
import { aid } from "../shared/ids.mjs";
import { imageDataFromDataUrl } from "../shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "../shared/inspection.mjs";
import { fileBlobFromRenderOutput, LAYOUT_MIME, renderTypeForOptions } from "../shared/render-output.mjs";
import { attrEscape, isXmlSafeText, xmlEscape } from "../shared/xml.mjs";
import { createTextRange, textRangeRecord } from "../shared/text-range.mjs";
import { queryHelpRecords } from "../help/index.mjs";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

const DOCX_PACKAGE_CONFIG = {
  family: "DOCX",
  packageKind: "docxPackage",
  partKind: "docxPart",
  semanticIssues: (context) => [
    ...validateDocxCommentPackageSemantics(context),
    ...validateDocxLinkPackageSemantics(context),
    ...validateDocxBibliographyPackageSemantics(context),
  ],
};

function documentHelp(query = "*", options = {}) {
  return ndjson(queryHelpRecords("document", query, options), options.maxChars ?? Infinity);
}

class DocumentStyleCollection {
  constructor(styles = {}) {
    this.items = new Map(Object.entries({
      Normal: { id: "Normal", name: "Normal", type: "paragraph", fontSize: 22, fontFamily: "Aptos" },
      Title: { id: "Title", name: "Title", type: "paragraph", fontSize: 48, bold: true, fontFamily: "Aptos Display" },
      Heading1: { id: "Heading1", name: "Heading 1", type: "paragraph", fontSize: 32, bold: true, fontFamily: "Aptos Display" },
      Heading2: { id: "Heading2", name: "Heading 2", type: "paragraph", fontSize: 26, bold: true, fontFamily: "Aptos" },
      ...styles,
    }));
  }

  add(id, config = {}) { const style = { id, name: config.name || id, type: config.type || "paragraph", ...config }; this.items.set(id, style); return style; }
  get(id) { return this.items.get(id); }
  effective(id, seen = new Set()) {
    const style = this.get(id) || this.get("Normal");
    if (!style) return undefined;
    const parentId = style.basedOn || style.parent || style.extends;
    if (!parentId || seen.has(style.id)) return { ...style };
    seen.add(style.id);
    const parent = this.effective(parentId, seen) || {};
    return { ...parent, ...style, basedOn: parentId };
  }
  cascade(id, seen = new Set()) {
    const style = this.get(id);
    if (!style || seen.has(style.id)) return [];
    seen.add(style.id);
    const parentId = style.basedOn || style.parent || style.extends;
    return [...(parentId ? this.cascade(parentId, seen) : []), style];
  }
  values() { return [...this.items.values()]; }
}

class DocumentTableCell {
  constructor(table, row, column) {
    this.table = table;
    this.kind = "tableCell";
    this.tableId = table.id;
    this.row = Number(row);
    this.column = Number(column);
    this.id = `${table.id}/cell/${this.row}/${this.column}`;
  }
  _record() { return this.table.cells?.find((cell) => cell.row === this.row && cell.column === this.column); }
  get value() {
    let value = String(this.table.values[this.row]?.[this.column] ?? "");
    for (const patch of this.table.textPatches.filter((item) => item.row === this.row && item.column === this.column)) {
      value = value.replace(patch.search, patch.replacement);
    }
    return value;
  }
  set value(value) {
    if (!this.editable) throw new Error(`Document table cell ${this.id} is source-bound and does not support whole-cell replacement; use its text range replace() operation when textPatchable is true.`);
    this.table.ensureCell(this.row, this.column);
    this.table.textPatches = this.table.textPatches.filter((item) => item.row !== this.row || item.column !== this.column);
    this.table.values[this.row][this.column] = String(value ?? "");
  }
  replaceText(search, replacement) {
    if (!this.textPatchable) throw new Error(`Document table cell ${this.id} does not advertise source-bound text replacement capability.`);
    if (search instanceof RegExp) throw new TypeError("Document table cell text replacement requires a literal search string.");
    const expected = String(search ?? "");
    const next = String(replacement ?? "");
    if (!expected) throw new TypeError("Document table cell text replacement requires a non-empty search string.");
    if (expected.length > 1_000_000 || next.length > 1_000_000 || !isXmlSafeText(expected) || !isXmlSafeText(next)) {
      throw new TypeError("Document table cell text replacement must use XML-safe strings of at most 1,000,000 characters.");
    }
    const value = this.value;
    const first = value.indexOf(expected);
    if (first < 0 || value.indexOf(expected, first + 1) >= 0) {
      throw new Error(`Document table cell ${this.id} text replacement requires exactly one visible match.`);
    }
    this.table.textPatches.push({ row: this.row, column: this.column, search: expected, replacement: next });
    return this;
  }
  get gridColumn() { return this._record()?.gridColumn ?? this.column; }
  get columnSpan() { return this._record()?.columnSpan ?? 1; }
  get rowSpan() { return this._record()?.rowSpan ?? 1; }
  get verticalMerge() { return this._record()?.verticalMerge ?? "none"; }
  get editable() { return this._record()?.editable ?? true; }
  get textPatchable() { return this._record()?.textPatchable === true; }
  inspectRecord() { return { kind: this.kind, id: this.id, textRangeId: this.textPatchable || this.editable ? `${this.id}/text` : undefined, tableId: this.tableId, row: this.row, column: this.column, gridColumn: this.gridColumn, columnSpan: this.columnSpan, rowSpan: this.rowSpan, verticalMerge: this.verticalMerge, editable: this.editable, textPatchable: this.textPatchable, pendingTextPatches: this.table.textPatches.filter((item) => item.row === this.row && item.column === this.column).length, value: this.value }; }
}

function documentTableDefaultColumnWidths(columns, widthDxa) {
  const count = Math.max(1, Number(columns) || 1);
  const total = Math.max(count, Math.round(Number(widthDxa) || 9360));
  const base = Math.floor(total / count);
  return Array.from({ length: count }, (_, index) => base + (index < total - base * count ? 1 : 0));
}

class DocumentTableBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "table";
    this.id = config.id || aid("dtb");
    this.name = config.name || "";
    this.styleId = config.styleId || config.style || "TableGrid";
    this.values = (config.values || Array.from({ length: config.rows || 1 }, () => Array.from({ length: config.columns || 1 }, () => ""))).map((row) => [...row]);
    this.rows = this.values.length;
    this.columns = Math.max(0, ...this.values.map((row) => row.length));
    this.cells = Array.isArray(config.cells) ? config.cells.map((cell) => {
      const verticalMerge = String(cell.verticalMerge || "none");
      return {
        row: Math.max(0, Math.round(Number(cell.row) || 0)),
        column: Math.max(0, Math.round(Number(cell.column) || 0)),
        gridColumn: Math.max(0, Math.round(Number(cell.gridColumn) || 0)),
        columnSpan: Math.max(1, Math.round(Number(cell.columnSpan) || 1)),
        rowSpan: Math.max(0, Math.round(Number(cell.rowSpan ?? (verticalMerge === "continue" ? 0 : 1)) || 0)),
        verticalMerge,
        editable: verticalMerge === "continue" ? false : cell.editable !== false,
        textPatchable: verticalMerge === "continue" ? false : cell.textPatchable === true,
      };
    }) : undefined;
    this.textPatches = Array.isArray(config.textPatches) ? config.textPatches.map((patch) => ({
      row: Math.max(0, Math.round(Number(patch.row) || 0)),
      column: Math.max(0, Math.round(Number(patch.column) || 0)),
      search: String(patch.search ?? ""),
      replacement: String(patch.replacement ?? ""),
    })) : [];
    const derivedGridColumns = this.cells?.reduce((maximum, cell) => Math.max(maximum, cell.gridColumn + cell.columnSpan), 0) || 0;
    this.gridColumns = Math.max(0, Math.round(Number(config.gridColumns ?? Math.max(this.columns, derivedGridColumns))));
    const formattingColumns = this.cells?.length ? this.gridColumns : this.columns;
    this.widthDxa = Math.round(Number(config.widthDxa ?? 9360));
    this.indentDxa = Math.round(Number(config.indentDxa ?? 120));
    this.columnWidthsDxa = Array.isArray(config.columnWidthsDxa)
      ? config.columnWidthsDxa.map((value) => Math.round(Number(value)))
      : documentTableDefaultColumnWidths(formattingColumns, this.widthDxa);
    this.cellMarginsDxa = {
      top: Math.round(Number(config.cellMarginsDxa?.top ?? 80)),
      bottom: Math.round(Number(config.cellMarginsDxa?.bottom ?? 80)),
      start: Math.round(Number(config.cellMarginsDxa?.start ?? config.cellMarginsDxa?.left ?? 120)),
      end: Math.round(Number(config.cellMarginsDxa?.end ?? config.cellMarginsDxa?.right ?? 120)),
    };
    this.borderColor = String(config.borderColor || "D9D9D9").replace(/^#/, "").toUpperCase();
    this.borderSize = Math.round(Number(config.borderSize ?? 4));
    this.headerFill = String(config.headerFill || "F2F4F7").replace(/^#/, "").toUpperCase();
  }

  ensureCell(row, column) { while (this.values.length <= row) this.values.push([]); while (this.values[row].length <= column) this.values[row].push(""); this.rows = this.values.length; this.columns = Math.max(this.columns, column + 1); }
  getCell(row, column) { return new DocumentTableCell(this, row, column); }
  inspectRecord(index) { return { kind: "table", id: this.id, index, name: this.name || undefined, rows: this.rows, cols: this.columns, gridColumns: this.gridColumns, cells: this.cells, pendingTextPatches: this.textPatches.length, styleId: this.styleId, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values.map((row, rowIndex) => row.map((_, columnIndex) => this.getCell(rowIndex, columnIndex).value)) }; }
  toProto() { return { kind: "table", id: this.id, name: this.name, styleId: this.styleId, gridColumns: this.gridColumns, cells: this.cells, textPatches: this.textPatches, widthDxa: this.widthDxa, indentDxa: this.indentDxa, columnWidthsDxa: this.columnWidthsDxa, cellMarginsDxa: this.cellMarginsDxa, borderColor: this.borderColor, borderSize: this.borderSize, headerFill: this.headerFill, values: this.values }; }
}

const DOCUMENT_CHECKBOX_GLYPHS = Object.freeze({ checked: "☒", unchecked: "☐" });
const DOCUMENT_DROPDOWN_MAX_CHOICES = 256;
const DOCUMENT_DROPDOWN_MAX_TEXT_LENGTH = 255;

function documentCheckboxGlyph(checked) {
  return checked ? DOCUMENT_CHECKBOX_GLYPHS.checked : DOCUMENT_CHECKBOX_GLYPHS.unchecked;
}

function normalizeDocumentDropdownChoices(value) {
  if (!Array.isArray(value) || value.length < 1 || value.length > DOCUMENT_DROPDOWN_MAX_CHOICES) {
    throw new TypeError(`Document drop-down content control requires 1 through ${DOCUMENT_DROPDOWN_MAX_CHOICES} choices.`);
  }
  const values = new Set();
  const displayTexts = new Set();
  return value.map((choice, index) => {
    const source = typeof choice === "string" ? { displayText: choice, value: choice } : choice;
    if (!source || typeof source !== "object") throw new TypeError(`Document drop-down choice ${index + 1} must be a string or object.`);
    const displayText = source.displayText ?? source.text ?? source.label ?? "";
    const itemValue = source.value ?? source.id ?? "";
    if (typeof displayText !== "string" || typeof itemValue !== "string") {
      throw new TypeError(`Document drop-down choice ${index + 1} displayText and value must be strings.`);
    }
    if (!displayText || !itemValue || displayText.length > DOCUMENT_DROPDOWN_MAX_TEXT_LENGTH || itemValue.length > DOCUMENT_DROPDOWN_MAX_TEXT_LENGTH || !isXmlSafeText(displayText) || !isXmlSafeText(itemValue) || /[\u0000-\u001f\u007f]/.test(displayText + itemValue)) {
      throw new TypeError(`Document drop-down choice ${index + 1} requires XML-safe displayText and value strings of 1 through ${DOCUMENT_DROPDOWN_MAX_TEXT_LENGTH} characters.`);
    }
    if (values.has(itemValue) || displayTexts.has(displayText)) {
      throw new TypeError("Document drop-down content-control choice values and displayText strings must be unique.");
    }
    values.add(itemValue);
    displayTexts.add(displayText);
    return { displayText, value: itemValue };
  });
}

function documentDropdownChoice(control, selectedValue = control?.selectedValue) {
  return control?.choices?.find((choice) => choice.value === selectedValue);
}

function normalizeDocumentContentControl(value) {
  if (value == null || value === false) return undefined;
  const source = typeof value === "string" ? { tag: value } : value;
  if (!source || typeof source !== "object") throw new TypeError("Document content control must be an object or tag string.");
  const tag = String(source.tag ?? source.name ?? "").trim();
  if (!tag) throw new TypeError("Document content control requires a non-empty tag.");
  const nativeId = source.nativeId == null ? undefined : Number(source.nativeId);
  if (nativeId !== undefined && (!Number.isInteger(nativeId) || nativeId < 1 || nativeId > 0x7fffffff)) {
    throw new TypeError("Document content control nativeId must be an integer from 1 through 2147483647.");
  }
  const inferredType = source.checked !== undefined
    ? "checkbox"
    : source.choices !== undefined || source.items !== undefined || source.options !== undefined || source.selectedValue !== undefined ? "dropdown" : "text";
  const rawType = String(source.controlType ?? source.type ?? inferredType).trim().toLowerCase();
  const controlType = rawType === "text" || rawType === "plain-text" || rawType === "plaintext"
    ? "text"
    : rawType === "checkbox" || rawType === "check-box" ? "checkbox"
      : rawType === "dropdown" || rawType === "drop-down" || rawType === "drop_down" ? "dropdown" : undefined;
  if (!controlType) throw new TypeError("Document content control type must be text, checkbox, or dropdown.");
  if (controlType === "checkbox" && source.checked !== undefined && typeof source.checked !== "boolean") {
    throw new TypeError("Document checkbox content control checked state must be boolean.");
  }
  const choices = controlType === "dropdown"
    ? normalizeDocumentDropdownChoices(source.choices ?? source.items ?? source.options)
    : undefined;
  const requestedSelectedValue = source.selectedValue ?? source.value;
  if (controlType === "dropdown" && requestedSelectedValue !== undefined && typeof requestedSelectedValue !== "string") {
    throw new TypeError("Document drop-down content-control selectedValue must be a string.");
  }
  const selectedValue = controlType === "dropdown" ? requestedSelectedValue ?? choices[0].value : undefined;
  if (controlType === "dropdown" && !choices.some((choice) => choice.value === selectedValue)) {
    throw new TypeError(`Document drop-down content-control selectedValue ${selectedValue} does not match a choice value.`);
  }
  return {
    id: String(source.id || aid("dcc")),
    tag,
    alias: String(source.alias ?? source.title ?? tag),
    nativeId,
    controlType,
    ...(controlType === "checkbox" ? { checked: source.checked === true } : {}),
    ...(controlType === "dropdown" ? { choices, selectedValue } : {}),
  };
}

function normalizeDocumentInlineField(value) {
  if (value == null || value === false) return undefined;
  const source = typeof value === "string" ? { instruction: value } : value;
  if (!source || typeof source !== "object") throw new TypeError("Document inline field must be an instruction string or object.");
  const instruction = String(source.instruction ?? source.code ?? "").trim();
  if (!instruction) throw new TypeError("Document inline field requires a non-empty instruction.");
  const bookmarkName = String(source.bookmarkName ?? source.bookmark ?? "").trim();
  const bookmarkNativeId = source.bookmarkNativeId == null ? undefined : Number(source.bookmarkNativeId);
  if (bookmarkNativeId !== undefined && (!Number.isInteger(bookmarkNativeId) || bookmarkNativeId < 0 || bookmarkNativeId > 0xffffffff)) {
    throw new TypeError("Document inline field bookmarkNativeId must be an unsigned 32-bit integer.");
  }
  return {
    instruction,
    ...(bookmarkName ? { bookmarkName } : {}),
    ...(bookmarkNativeId !== undefined ? { bookmarkNativeId } : {}),
  };
}

function normalizeDocumentRun(run = {}, theme = {}) {
  const contentControl = normalizeDocumentContentControl(run.contentControl ?? run.textContentControl ?? run.control);
  const inlineField = normalizeDocumentInlineField(run.inlineField ?? run.field);
  if (contentControl && inlineField) throw new TypeError("Document run cannot be both a content control and an inline field.");
  const requestedText = String(run.text ?? run.value ?? "");
  const text = contentControl?.controlType === "checkbox"
    ? documentCheckboxGlyph(contentControl.checked)
    : contentControl?.controlType === "dropdown" ? documentDropdownChoice(contentControl).displayText : requestedText;
  if (contentControl?.controlType !== "text" && contentControl && requestedText && requestedText !== text) {
    throw new TypeError(`Document ${contentControl.controlType} content-control text is codec-owned; set its typed value instead of supplying visible text.`);
  }
  return {
    text,
    style: normalizeDocxRunStyle(run.style || run.textStyle || {}, theme),
    ...(contentControl ? { contentControl } : {}),
    ...(inlineField ? { inlineField } : {}),
  };
}

function normalizeDocumentRuns(text, config = {}, theme = {}) {
  const runs = (config.runs || config.textRuns || []).map((run) => normalizeDocumentRun(run, theme)).filter((run) => run.text.length > 0 || run.contentControl || run.inlineField);
  if (runs.length) return runs;
  const rawText = String(text ?? "");
  return rawText ? [{ text: rawText, style: normalizeDocxRunStyle({}, theme) }] : [];
}

function documentEffectiveRunStyle(document, block, run) {
  const characterStyleId = run.style?.runStyleId;
  const cascade = [document.defaultRunStyle, ...document.styles.cascade(block.styleId), ...(characterStyleId ? document.styles.cascade(characterStyleId) : []), run.style || {}];
  return effectiveDocxRunStyle(mergeDocxRunStyleCascade(cascade, document.theme), run.text, document.theme);
}

function documentRunsNeedSerialization(runs = []) {
  return runs.length > 1 || runs.some((run) => Object.keys(run.style || {}).length > 0 || run.contentControl || run.inlineField);
}

class DocumentParagraphBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "paragraph";
    this.id = config.id || aid("dp");
    this.runs = normalizeDocumentRuns(text, config, document.theme);
    this.text = this.runs.map((run) => run.text).join("") || String(text ?? "");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
    this.paragraphFormat = { ...(config.paragraphFormat || config.formatting || {}) };
    this.textEditable = config.textEditable !== false;
    this.textPatchable = config.textPatchable === true;
    this.textPatches = Array.isArray(config.textPatches) ? config.textPatches.map((patch) => ({ search: String(patch.search ?? ""), replacement: String(patch.replacement ?? "") })) : [];
  }

  _syncText() { this.text = this.runs.map((run) => String(run.text ?? "")).join(""); return this.text; }
  addRun(text, config = {}) { const run = normalizeDocumentRun({ ...config, text }, this.document.theme); this.runs.push(run); this._syncText(); return run; }
  addTextContentControl(text, config = {}) { return this.addRun(text, { ...config, contentControl: config.contentControl || config }); }
  addCheckboxContentControl(checked = false, config = {}) {
    if (typeof checked !== "boolean") throw new TypeError("Document checkbox content control checked state must be boolean.");
    return this.addRun("", { ...config, contentControl: { ...(config.contentControl || config), controlType: "checkbox", checked } });
  }
  addDropdownContentControl(choices, config = {}) {
    return this.addRun("", {
      ...config,
      contentControl: {
        ...(config.contentControl || config),
        controlType: "dropdown",
        choices,
        selectedValue: config.selectedValue ?? config.value,
      },
    });
  }
  addField(instruction, display = "0", config = {}) {
    return this.addRun(display, {
      ...config,
      inlineField: {
        instruction,
        bookmarkName: config.bookmarkName ?? config.bookmark,
        bookmarkNativeId: config.bookmarkNativeId,
      },
    });
  }
  replaceText(search, replacement) {
    if (this.textEditable) {
      const value = String(this.text).replace(search, replacement);
      if (this.runs.length > 1 && value !== this.text) throw new Error(`Document paragraph ${this.id} contains multiple source runs; edit the intended run(s) explicitly so formatting boundaries are not flattened.`);
      this.text = value;
      if (this.runs.length === 1) this.runs[0].text = value;
      else if (!this.runs.length && value) this.runs = [{ text: value, style: {} }];
      return this;
    }
    if (!this.textPatchable) throw new Error(`Document paragraph ${this.id} does not advertise source-bound text replacement capability.`);
    if (search instanceof RegExp) throw new TypeError("Document paragraph source-bound text replacement requires a literal search string.");
    const expected = String(search ?? "");
    const next = String(replacement ?? "");
    if (!expected) throw new TypeError("Document paragraph text replacement requires a non-empty search string.");
    if (expected.length > 1_000_000 || next.length > 1_000_000 || !isXmlSafeText(expected) || !isXmlSafeText(next)) {
      throw new TypeError("Document paragraph text replacement must use XML-safe strings of at most 1,000,000 characters.");
    }
    const first = this.text.indexOf(expected);
    if (first < 0 || this.text.indexOf(expected, first + 1) >= 0) throw new Error(`Document paragraph ${this.id} text replacement requires exactly one visible match.`);
    this.textPatches.push({ search: expected, replacement: next });
    this.text = this.text.replace(expected, next);
    if (this.runs.length === 1) this.runs[0].text = this.text;
    return this;
  }
  inspectRecord(index) { return { kind: "paragraph", id: this.id, index, name: this.name || undefined, styleId: this.styleId, textEditable: this.textEditable, textPatchable: this.textPatchable, pendingTextPatches: this.textPatches.length, textRangeId: this.textEditable || this.textPatchable ? `${this.id}/text` : undefined, paragraphFormat: Object.keys(this.paragraphFormat).length ? this.paragraphFormat : undefined, text: this.text, textChars: this.text.length, runs: documentRunsNeedSerialization(this.runs) ? this.runs : undefined }; }
  toProto() { return { kind: "paragraph", id: this.id, name: this.name, styleId: this.styleId, textEditable: this.textEditable, textPatchable: this.textPatchable, textPatches: this.textPatches, paragraphFormat: Object.keys(this.paragraphFormat).length ? this.paragraphFormat : undefined, text: this.text, runs: documentRunsNeedSerialization(this.runs) ? this.runs : undefined }; }
}

class DocumentContentControlHandle {
  constructor(block, runIndex) { this.document = block.document; this.block = block; this.runIndex = runIndex; this.kind = "contentControl"; }
  get run() { return this.block.runs[this.runIndex]; }
  get control() { return this.run?.contentControl; }
  get id() { return this.control?.id; }
  get targetId() { return this.block.id; }
  get tag() { return this.control?.tag || ""; }
  set tag(value) { this.control.tag = String(value ?? "").trim(); }
  get alias() { return this.control?.alias || ""; }
  set alias(value) { this.control.alias = String(value ?? ""); }
  get nativeId() { return this.control?.nativeId; }
  get controlType() { return this.control?.controlType || "text"; }
  get text() { return String(this.run?.text ?? ""); }
  set text(value) {
    if (this.controlType === "checkbox") throw new TypeError("Checkbox content-control text is codec-owned; set checked instead.");
    if (this.controlType === "dropdown") throw new TypeError("Drop-down content-control text is codec-owned; set selectedValue instead.");
    if (this.controlType !== "text") throw new TypeError(`Unsupported ${this.controlType} content-control text mutation.`);
    this.run.text = String(value ?? "");
    this.block._syncText();
  }
  get checked() { return this.controlType === "checkbox" ? this.control?.checked === true : undefined; }
  set checked(value) {
    if (this.controlType !== "checkbox") throw new TypeError("Only checkbox content controls have checked state.");
    if (typeof value !== "boolean") throw new TypeError("Document checkbox content control checked state must be boolean.");
    this.control.checked = value;
    this.run.text = documentCheckboxGlyph(value);
    this.block._syncText();
  }
  get choices() { return this.controlType === "dropdown" ? this.control.choices.map((choice) => ({ ...choice })) : undefined; }
  get selectedValue() { return this.controlType === "dropdown" ? this.control?.selectedValue : undefined; }
  set selectedValue(value) {
    if (this.controlType !== "dropdown") throw new TypeError("Only drop-down content controls have selectedValue state.");
    if (typeof value !== "string") throw new TypeError("Document drop-down content-control selectedValue must be a string.");
    const selectedValue = value;
    const choice = documentDropdownChoice(this.control, selectedValue);
    if (!choice) throw new TypeError(`Document drop-down content-control selectedValue ${selectedValue} does not match a choice value.`);
    this.control.selectedValue = selectedValue;
    this.run.text = choice.displayText;
    this.block._syncText();
  }
  inspectRecord() {
    return {
      kind: this.kind,
      id: this.id,
      targetId: this.targetId,
      runIndex: this.runIndex,
      tag: this.tag,
      alias: this.alias,
      nativeId: this.nativeId,
      controlType: this.controlType,
      ...(this.controlType === "checkbox"
        ? { checked: this.checked, visibleText: this.text }
        : this.controlType === "dropdown"
          ? { choices: this.choices, selectedValue: this.selectedValue, visibleText: this.text }
          : { text: this.text, textChars: this.text.length }),
    };
  }
}

function documentContentControls(document) {
  return document.blocks.flatMap((block) => block.kind === "paragraph"
    ? block.runs.flatMap((run, runIndex) => run.contentControl ? [new DocumentContentControlHandle(block, runIndex)] : [])
    : []);
}

class DocumentChangeBlock {
  constructor(document, changeType, text, config = {}) {
    this.document = document;
    this.kind = "change";
    this.id = config.id || aid("dchg");
    const rawType = String(changeType ?? config.changeType ?? config.type ?? "insert").toLowerCase();
    this.changeType = rawType === "delete" || rawType === "deletion" || rawType === "del" ? "delete" : "insert";
    this.text = String(text ?? "");
    this.author = config.author === undefined && config._restore ? "" : (config.author || "User");
    this.date = config.date === undefined ? (config._restore ? undefined : new Date().toISOString()) : config.date;
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "change", id: this.id, index, name: this.name || undefined, changeType: this.changeType, author: this.author, date: this.date, styleId: this.styleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "change", id: this.id, name: this.name, changeType: this.changeType, author: this.author, date: this.date, styleId: this.styleId, text: this.text }; }
}

class DocumentListItemBlock {
  constructor(document, text, config = {}) {
    this.document = document;
    this.kind = "listItem";
    this.id = config.id || aid("dli");
    this.text = String(text ?? "");
    this.level = Number(config.level ?? 0);
    const configuredFormat = config.numberFormat || config.numFmt;
    this.listType = config.listType || config.type || (configuredFormat === "bullet" ? "bullet" : configuredFormat ? "number" : "bullet");
    this.numberFormat = configuredFormat || (this.listType === "number" ? "decimal" : "bullet");
    this.start = Number(config.start ?? 1);
    this.levelText = config.levelText || config.lvlText || (this.listType === "number" ? `%${this.level + 1}.` : "•");
    this.pictureBullet = normalizeDocumentPictureBullet(config.pictureBullet ?? config.bulletImage ?? config.bullet?.image);
    if (this.pictureBullet && (this.listType !== "bullet" || this.numberFormat !== "bullet")) throw new TypeError("DOCX picture bullet requires listType and numberFormat to be bullet.");
    this.numberingId = config.numberingId ?? config.numId;
    this.abstractNumberingId = config.abstractNumberingId ?? config.abstractNumId;
    this.numberingStyleId = config.numberingStyleId ?? config.numStyleId;
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "listItem", id: this.id, index, name: this.name || undefined, styleId: this.styleId, listType: this.listType, level: this.level, numberFormat: this.numberFormat, start: this.start, levelText: this.levelText, pictureBullet: this.pictureBullet, numberingId: this.numberingId, abstractNumberingId: this.abstractNumberingId, numberingStyleId: this.numberingStyleId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: "listItem", id: this.id, name: this.name, styleId: this.styleId, listType: this.listType, level: this.level, numberFormat: this.numberFormat, start: this.start, levelText: this.levelText, pictureBullet: this.pictureBullet, numberingId: this.numberingId, abstractNumberingId: this.abstractNumberingId, numberingStyleId: this.numberingStyleId, text: this.text }; }
}

class DocumentHyperlinkBlock {
  constructor(document, text, url, config = {}) {
    this.document = document;
    this.kind = "hyperlink";
    this.id = config.id || aid("dhl");
    this.text = String(text ?? "");
    const target = typeof url === "object" && url ? url : config.anchor || config.bookmark;
    const rawUrl = typeof url === "string" ? url : config.url;
    this.anchor = String(target?.name || target || config.anchor?.name || config.anchor || config.bookmarkName || (String(rawUrl || "").startsWith("#") ? String(rawUrl).slice(1) : "")).trim() || undefined;
    this.url = this.anchor ? "" : String(rawUrl || "");
    this.relationshipId = config.relationshipId || config.relId;
    this.tooltip = config.tooltip;
    this.history = config.history !== false;
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "hyperlink", id: this.id, index, name: this.name || undefined, styleId: this.styleId, relationshipId: this.relationshipId, text: this.text, url: this.url || undefined, anchor: this.anchor, tooltip: this.tooltip, history: this.history, textChars: this.text.length }; }
  toProto() { return { kind: "hyperlink", id: this.id, name: this.name, styleId: this.styleId, relationshipId: this.relationshipId, text: this.text, url: this.url || undefined, anchor: this.anchor, tooltip: this.tooltip, history: this.history }; }
}

function documentBookmarkEndpoint(value) {
  if (!value) return undefined;
  const raw = typeof value === "object" ? value : undefined;
  const id = typeof value === "string" ? value : raw?.id;
  const cellMatch = /^(.+)\/cell\/(\d+)\/(\d+)$/.exec(String(id || ""));
  if (cellMatch || raw?.kind === "tableCell" || raw?.type === "tableCell" || raw?.tableId !== undefined) {
    const tableId = String(raw?.tableId || raw?.table?.id || cellMatch?.[1] || "");
    const row = Number(raw?.row ?? raw?.rowIndex ?? cellMatch?.[2]);
    const column = Number(raw?.column ?? raw?.columnIndex ?? cellMatch?.[3]);
    return { type: "tableCell", id: `${tableId}/cell/${row}/${column}`, tableId, row, column };
  }
  return { type: "block", id: String(id || raw?.blockId || value), blockId: String(raw?.blockId || id || value) };
}

function documentBookmarkEndpointBlockId(endpoint) {
  return endpoint?.type === "tableCell" ? endpoint.tableId : endpoint?.blockId;
}

function documentBookmarkEndpointOrder(endpoint, blockIndexes) {
  const blockIndex = blockIndexes.get(documentBookmarkEndpointBlockId(endpoint));
  return endpoint?.type === "tableCell" ? [blockIndex, 1, endpoint.row, endpoint.column] : [blockIndex, 0, 0, 0];
}

function compareDocumentBookmarkEndpoints(left, right, blockIndexes) {
  const a = documentBookmarkEndpointOrder(left, blockIndexes);
  const b = documentBookmarkEndpointOrder(right, blockIndexes);
  for (let index = 0; index < a.length; index += 1) if (a[index] !== b[index]) return a[index] - b[index];
  return 0;
}

class DocumentBookmark {
  constructor(document, target, name, config = {}) {
    this.document = document;
    this.kind = "bookmark";
    this.id = config.id || aid("dbm");
    this.name = String(name ?? config.name ?? "");
    this.target = documentBookmarkEndpoint(target ?? config.target ?? config.targetId);
    const endTarget = config.endTarget ?? config.end ?? config.endTargetId ?? this.target;
    this.endTarget = documentBookmarkEndpoint(endTarget);
    this.targetId = this.target?.id;
    this.endTargetId = this.endTarget?.id || this.targetId;
    this.nativeId = config.nativeId === undefined ? undefined : Number(config.nativeId);
  }

  inspectRecord() { return { kind: "bookmark", id: this.id, name: this.name, targetId: this.targetId, endTargetId: this.endTargetId, target: this.target, endTarget: this.endTarget, nativeId: this.nativeId }; }
  toProto() { return { kind: "bookmark", id: this.id, name: this.name, targetId: this.targetId, endTargetId: this.endTargetId, target: this.target, endTarget: this.endTarget, nativeId: this.nativeId }; }
}

class DocumentNote {
  constructor(document, kind, target, text, config = {}) {
    this.document = document;
    const rawKind = String(kind ?? config.kind ?? config.noteKind ?? config.type ?? "footnote").toLowerCase();
    this.kind = rawKind === "endnote" || rawKind === "end" ? "endnote" : "footnote";
    this.id = config.id || aid(this.kind === "endnote" ? "den" : "dfn");
    this.name = config.name || "";
    this.targetId = String(typeof target === "string" ? target : target?.id || config.targetId || "");
    this.text = String(text ?? config.text ?? "");
    this.nativeId = config.nativeId === undefined ? undefined : Number(config.nativeId);
  }

  inspectRecord(index) { return { kind: this.kind, id: this.id, index, name: this.name || undefined, targetId: this.targetId, nativeId: this.nativeId, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: this.kind, id: this.id, name: this.name, targetId: this.targetId, nativeId: this.nativeId, text: this.text }; }
}

class DocumentFieldBlock {
  constructor(document, instruction, display, config = {}) {
    this.document = document;
    this.kind = "field";
    this.id = config.id || aid("dfld");
    this.instruction = String(instruction ?? config.instruction ?? "PAGE");
    this.display = String(display ?? config.display ?? "1");
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
    this.complex = Boolean(config.complex);
  }

  get text() { return this.display; }
  inspectRecord(index) { return { kind: "field", id: this.id, index, name: this.name || undefined, styleId: this.styleId, instruction: this.instruction, display: this.display, complex: this.complex }; }
  toProto() { return { kind: "field", id: this.id, name: this.name, styleId: this.styleId, instruction: this.instruction, display: this.display, complex: this.complex }; }
}

class DocumentBibliographySource {
  constructor(document, config = {}) {
    this.document = document;
    Object.assign(this, normalizeDocxBibliographySource(config, document.bibliographySources.length));
  }

  inspectRecord(index) { return { ...this.toProto(), index }; }
  toProto() { return Object.fromEntries(Object.entries(this).filter(([key, value]) => key !== "document" && value !== undefined)); }
}

class DocumentCitationBlock {
  constructor(document, text, metadata = {}, config = {}) {
    this.document = document;
    this.kind = "citation";
    this.id = config.id || aid("dct");
    this.text = String(text ?? "");
    this.metadata = { ...metadata };
    this.styleId = config.styleId || config.style || "Normal";
    this.name = config.name || "";
  }

  inspectRecord(index) { return { kind: "citation", id: this.id, index, name: this.name || undefined, styleId: this.styleId, text: this.text, metadata: this.metadata, textChars: this.text.length }; }
  toProto() { return { kind: "citation", id: this.id, name: this.name, styleId: this.styleId, text: this.text, metadata: this.metadata }; }
}

class DocumentImageBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "image";
    this.id = config.id || aid("dim");
    this.name = config.name || "";
    this.dataUrl = config.dataUrl;
    this.uri = config.uri;
    this.prompt = config.prompt;
    this.alt = config.alt || config.altText || config.name || "image";
    this.widthPx = Number(config.widthPx || config.width || 240);
    this.heightPx = Number(config.heightPx || config.height || 160);
    this.styleId = config.styleId || config.style || "Normal";
  }

  inspectRecord(index) { return { kind: "image", id: this.id, index, name: this.name || undefined, styleId: this.styleId, alt: this.alt, uri: this.uri, prompt: this.prompt, widthPx: this.widthPx, heightPx: this.heightPx, hasDataUrl: Boolean(this.dataUrl) }; }
  toProto() { return { kind: "image", id: this.id, name: this.name, styleId: this.styleId, dataUrl: this.dataUrl, uri: this.uri, prompt: this.prompt, alt: this.alt, widthPx: this.widthPx, heightPx: this.heightPx }; }
}

class DocumentSectionBlock {
  constructor(document, config = {}) {
    this.document = document;
    this.kind = "section";
    this.id = config.id || aid("dsec");
    this.name = config.name || "";
    this.breakType = config.breakType || config.type || "nextPage";
    this.orientation = config.orientation || "portrait";
    const pageSize = config.pageSize || {};
    this.pageSize = {
      widthTwips: Number(config.widthTwips || pageSize.widthTwips || (this.orientation === "landscape" ? 15840 : 12240)),
      heightTwips: Number(config.heightTwips || pageSize.heightTwips || (this.orientation === "landscape" ? 12240 : 15840)),
    };
    const margins = config.margins || {};
    this.margins = {
      top: Number(margins.top ?? config.marginTop ?? 1440),
      right: Number(margins.right ?? config.marginRight ?? 1440),
      bottom: Number(margins.bottom ?? config.marginBottom ?? 1440),
      left: Number(margins.left ?? config.marginLeft ?? 1440),
    };
  }

  inspectRecord(index) { return { kind: "section", id: this.id, index, name: this.name || undefined, breakType: this.breakType, orientation: this.orientation, pageSize: this.pageSize, margins: this.margins }; }
  toProto() { return { kind: "section", id: this.id, name: this.name, breakType: this.breakType, orientation: this.orientation, pageSize: this.pageSize, margins: this.margins }; }
}

class DocumentHeaderFooterBlock {
  constructor(document, kind, text, config = {}) {
    this.document = document;
    this.kind = kind;
    this.id = config.id || aid(kind === "header" ? "dh" : "df");
    this.text = String(text ?? "");
    this.name = config.name || kind;
    this.styleId = config.styleId || "Normal";
    this.referenceType = ["default", "first", "even"].includes(config.referenceType || config.type) ? (config.referenceType || config.type) : "default";
    this.relationshipId = config.relationshipId || config.relId;
    this.partPath = config.partPath;
    this.sectionIndex = config.sectionIndex === undefined || config.sectionIndex === null ? undefined : Number(config.sectionIndex);
    this.variantActive = config.variantActive ?? config.activateVariant ?? config.active;
    this.fieldInstruction = config.fieldInstruction || config.field;
  }

  inspectRecord(index) { return { kind: this.kind, id: this.id, index, name: this.name || undefined, styleId: this.styleId, referenceType: this.referenceType, variantActive: this.variantActive, relationshipId: this.relationshipId, partPath: this.partPath, sectionIndex: this.sectionIndex, fieldInstruction: this.fieldInstruction, text: this.text, textChars: this.text.length }; }
  toProto() { return { kind: this.kind, id: this.id, name: this.name, styleId: this.styleId, referenceType: this.referenceType, variantActive: this.variantActive, relationshipId: this.relationshipId, partPath: this.partPath, sectionIndex: this.sectionIndex, fieldInstruction: this.fieldInstruction, text: this.text }; }
}

function documentCommentInitials(author) {
  const words = String(author || "User").trim().split(/\s+/).filter(Boolean);
  return (words.length > 1 ? words.map((word) => word[0]).join("") : (words[0] || "U").slice(0, 2)).slice(0, 4).toUpperCase();
}

class DocumentComment {
  constructor(document, targetId, text, config = {}) {
    this.document = document;
    this.kind = "comment";
    this.id = config.id || aid("dc");
    this.targetId = targetId;
    this.author = config.author || "User";
    this.initials = config.initials || documentCommentInitials(this.author);
    this.date = config.date || undefined;
    this.text = String(text ?? "");
    this.resolved = Boolean(config.resolved);
    Object.defineProperty(this, "_resolvedSpecified", { value: config.resolved !== undefined, writable: true });
    this.parentId = typeof (config.parentId ?? config.replyToId ?? config.replyTo) === "object" ? (config.parentId ?? config.replyToId ?? config.replyTo)?.id : (config.parentId ?? config.replyToId ?? config.replyTo);
    this.paraId = config.paraId ? String(config.paraId).toUpperCase() : undefined;
    this.durableId = config.durableId ? String(config.durableId).toUpperCase() : undefined;
    this.dateUtc = config.dateUtc;
    this.person = config.person
      ? { providerId: String(config.person.providerId ?? ""), userId: String(config.person.userId ?? "") }
      : (config.providerId || config.userId ? { providerId: String(config.providerId ?? ""), userId: String(config.userId ?? "") } : undefined);
    this.intelligentPlaceholder = Boolean(config.intelligentPlaceholder);
  }

  inspectRecord() { return { kind: "comment", id: this.id, targetId: this.targetId, parentId: this.parentId, paraId: this.paraId, durableId: this.durableId, author: this.author, initials: this.initials, date: this.date, dateUtc: this.dateUtc, person: this.person, intelligentPlaceholder: this.intelligentPlaceholder || undefined, resolved: this.resolved, textPreview: this.text.slice(0, 300) }; }
  toProto() { return { kind: "comment", id: this.id, targetId: this.targetId, parentId: this.parentId, paraId: this.paraId, durableId: this.durableId, author: this.author, initials: this.initials, date: this.date, dateUtc: this.dateUtc, person: this.person, intelligentPlaceholder: this.intelligentPlaceholder || undefined, text: this.text, resolved: this._resolvedSpecified ? this.resolved : undefined }; }
  resolve() { this.resolved = true; this._resolvedSpecified = true; return this; }
  reopen() { this.resolved = false; this._resolvedSpecified = true; return this; }
}

function documentBlockHeight(document, block, pageWidth = 612, margin = 72) {
  if (block.kind === "table") return Math.max(24, block.rows * 24 + 16);
  if (block.kind === "image") return Math.max(32, Math.min(360, Number(block.heightPx) || 160)) + 20;
  if (block.kind === "section") return 34;
  if (block.kind === "change") return 22;
  const style = document.styles.effective(block.styleId) || document.styles.get("Normal") || {};
  const runSizes = block.kind === "paragraph" ? (block.runs || []).map((run) => documentEffectiveRunStyle(document, block, run).effectiveFontSize).filter(Number.isFinite) : [];
  const fontSize = Math.max(10, Math.max(style.fontSize || 22, ...runSizes) / 2);
  const text = block.text || block.display || "";
  const charsPerLine = Math.max(8, Math.floor((pageWidth - margin * 2) / (fontSize * 0.55)));
  const lines = String(text).split(/\r?\n/).reduce((sum, line) => sum + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
  return Math.max(20, lines * fontSize * 1.6);
}

function documentBlockLayoutText(block) {
  if (block.kind === "table") return block.values.map((row) => row.map((cell) => String(cell ?? "")).join(" ")).join(" ");
  if (block.kind === "section") return `section break: ${block.breakType || ""} ${block.orientation || ""}`.trim();
  return String(block.text ?? block.display ?? block.alt ?? block.prompt ?? block.uri ?? block.name ?? "");
}

function documentLayoutSlice(document, layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsDocument = targets.some((target) => target === document.id || target === document.name);
  const matchedPages = new Set(layout.pages.filter((pageRecord) => inspectRecordMatchesTarget(pageRecord, targets)).map((pageRecord) => pageRecord.page));
  const matches = [];
  layout.elements.forEach((element, index) => {
    const matchesSearch = !search || JSON.stringify(element).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsDocument || matchedPages.has(element.page) || inspectRecordMatchesTarget(element, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(layout.elements.length - 1, index + after); i += 1) keep.add(i);
  }
  const elements = layout.elements.filter((_, index) => keep.has(index));
  const referencedPages = new Set(elements.map((element) => element.page));
  const pages = layout.pages.filter((pageRecord) => referencedPages.has(pageRecord.page));
  return { ...layout, pages, elements, slice: { targets, search: search || undefined, before, after, matchedElements: matches.length, returnedElements: elements.length } };
}

function documentLayoutJson(document, options = {}) {
  const pageWidth = Number(options.pageWidth || 612);
  const pageHeight = Number(options.pageHeight || 792);
  const margin = Number(options.margin || 72);
  const headerFooterPlan = planDocxHeaderFooterSections(document);
  const pages = [];
  const elements = [];
  let page = 1;
  let sectionIndex = 0;
  let pageInSection = 1;
  let y = margin;
  const ensurePage = () => {
    if (!pages.find((item) => item.page === page)) {
      const headerFooter = resolveDocxPageHeaderFooter(headerFooterPlan, sectionIndex, pageInSection);
      pages.push({ id: `${document.id}/page/${page}`, page, width: pageWidth, height: pageHeight, margin, sectionIndex, sectionIndexes: [sectionIndex], ...headerFooter });
    }
  };
  ensurePage();
  for (const block of document.blocks) {
    const height = documentBlockHeight(document, block, pageWidth, margin);
    if (y + height > pageHeight - margin && y > margin) { page += 1; pageInSection += 1; y = margin; ensurePage(); }
    const textPreview = documentBlockLayoutText(block).slice(0, 120);
    const comments = document.comments.filter((comment) => comment.targetId === block.id).map((comment) => comment.id);
    const effectiveStyle = block.styleId ? document.styles.effective(block.styleId) : undefined;
    const runs = block.kind === "paragraph" && block.runs?.length ? block.runs.map((run) => ({ text: run.text, style: documentEffectiveRunStyle(document, block, run) })) : undefined;
    elements.push({ kind: "layoutElement", id: block.id, layoutId: `${block.id}/layout`, blockKind: block.kind, name: block.name || undefined, textRangeId: ("text" in block || "display" in block) ? `${block.id}/text` : undefined, commentIds: comments.length ? comments : undefined, page, bbox: [margin, y, pageWidth - margin * 2, height], styleId: block.styleId, effectiveStyle, runs, textPreview });
    y += height;
    if (block.kind === "section") {
      sectionIndex += 1;
      pageInSection = 1;
      if (block.breakType === "nextPage") { page += 1; y = margin; ensurePage(); }
      else {
        const currentPage = pages.find((item) => item.page === page);
        if (currentPage && !currentPage.sectionIndexes.includes(sectionIndex)) currentPage.sectionIndexes.push(sectionIndex);
        if (currentPage) currentPage.continuousSectionBoundary = true;
      }
    }
  }
  return documentLayoutSlice(document, { schema: "open-office-artifact.document-layout/v1", unit: "px", document: { id: document.id, name: document.name, designPreset: document.designPreset, theme: document.theme }, pages, elements }, options);
}

function documentLayoutRecords(document, options = {}) {
  const layout = documentLayoutJson(document, options);
  return [
    { kind: "layout", id: `${document.id}/layout`, pages: layout.pages.length, elements: layout.elements.length, designPreset: document.designPreset },
    ...layout.pages.map((page) => ({ kind: "page", id: `${document.id}/page/${page.page}`, ...page })),
    ...layout.elements,
  ];
}

function documentTextParent(document, parentId) {
  return [...document.blocks, ...document.headers, ...document.footers, ...document.comments, ...document.notes].find((item) => item.id === parentId);
}

function documentTextRange(document, id) {
  const parentId = String(id || "").endsWith("/text") ? String(id).slice(0, -5) : undefined;
  const cellMatch = /^(.+)\/cell\/(\d+)\/(\d+)$/.exec(parentId || "");
  if (cellMatch) {
    const table = document.blocks.find((block) => block.kind === "table" && block.id === cellMatch[1]);
    const cell = table?.getCell(Number(cellMatch[2]), Number(cellMatch[3]));
    if (!cell || (!cell.editable && !cell.textPatchable)) return undefined;
    return createTextRange(cell, id, {
      parentKind: cell.kind,
      getText: () => cell.value,
      setText: (value) => { cell.value = value; },
      replace: (search, replacement) => {
        if (cell.textPatchable) cell.replaceText(search, replacement);
        else cell.value = String(cell.value).replace(search, replacement);
      },
    });
  }
  const parent = parentId ? documentTextParent(document, parentId) : undefined;
  if (!parent) return undefined;
  if (parent.kind === "paragraph" && parent.textEditable === false && parent.textPatchable === false) return undefined;
  const setText = (value) => {
    if (parent.kind === "field" || ("display" in parent && !("text" in parent))) parent.display = String(value ?? "");
    else {
      const text = String(value ?? "");
      if (parent.kind === "paragraph" && parent.textEditable === false) throw new Error(`Document paragraph ${parent.id} is source-bound; use replace() for a bounded literal text patch.`);
      if (parent.kind === "paragraph" && parent.runs?.length > 1 && text !== parent.text) {
        throw new Error(`Document paragraph ${parent.id} contains multiple source runs; edit the intended run(s) explicitly so formatting boundaries are not flattened.`);
      }
      parent.text = text;
      if (parent.kind === "paragraph") {
        if (parent.runs?.length === 1) parent.runs[0].text = text;
        else if (!parent.runs?.length) parent.runs = text ? [{ text, style: {} }] : [];
      }
    }
  };
  return createTextRange(parent, id, {
    parentKind: parent.kind,
    getText: () => parent.text ?? parent.display ?? "",
    setText,
    replace: (search, replacement) => {
      if (parent.kind === "paragraph") parent.replaceText(search, replacement);
      else setText(String(parent.text ?? parent.display ?? "").replace(search, replacement));
    },
  });
}

function documentInspectRecord(document, block, index) {
  const record = block.inspectRecord(index);
  if (record.styleId) record.effectiveStyle = document.styles.effective(record.styleId);
  return record;
}

function documentTextRangeRecords(document) {
  const parents = [...document.blocks, ...document.headers, ...document.footers, ...document.comments, ...document.notes].filter((item) => item && ("text" in item || "display" in item));
  const records = parents.map((parent, index) => textRangeRecord(parent, {
    parentKind: parent.kind,
    getText: () => parent.text ?? parent.display ?? "",
    record: { index, styleId: parent.styleId, targetId: parent.targetId, textEditable: parent.textEditable, textPatchable: parent.textPatchable },
  }));
  for (const table of document.blocks.filter((block) => block.kind === "table")) {
    for (let row = 0; row < table.rows; row += 1) for (let column = 0; column < table.columns; column += 1) {
      const cell = table.getCell(row, column);
      if (!cell.editable && !cell.textPatchable) continue;
      records.push(textRangeRecord(cell, {
        parentKind: cell.kind,
        getText: () => cell.value,
        record: { tableId: table.id, row, column, editable: cell.editable, textPatchable: cell.textPatchable },
      }));
    }
  }
  return records;
}

function documentRenderUsesDocxSource(options = {}) {
  const source = String(options.source || options.inputFormat || options.renderSource || "").trim().toLowerCase();
  return source === "docx" || options.docx === true || options.useDocx === true || options.native === true || options.nativeOffice === true || options.office === true;
}

async function renderDocumentFromDocx(document, options = {}) {
  const docx = await DocumentFile.exportDocx(document);
  docx.metadata = { ...(docx.metadata || {}), artifactKind: "document", format: "docx", renderSource: "docx" };
  const desiredType = renderTypeForOptions(options, docx.type);
  const format = String(options.format || "docx").trim().toLowerCase();
  if (!format || format === "docx" || desiredType === docx.type) return docx;
  const adapter = options.renderer || options.rasterRenderer || options.renderAdapter;
  if (typeof adapter !== "function") return docx;
  const converted = await adapter({ input: docx, source: docx, inputType: docx.type, outputType: desiredType, format: options.format, artifactKind: "document", options: { ...options, source: "docx" } });
  const blob = await fileBlobFromRenderOutput(converted, desiredType, { artifactKind: "document", format: options.format, renderedFrom: docx.type, renderSource: "docx" });
  if (!blob.type || blob.type === "application/octet-stream") blob.type = desiredType;
  return blob;
}

const DOCUMENT_MATERIALIZE_SEQ = /^SEQ ([A-Za-z][A-Za-z0-9_]{0,39}) \\[*] ARABIC$/;
const DOCUMENT_MATERIALIZE_REF = /^REF ([A-Za-z][A-Za-z0-9_]{0,39}) \\h$/;
const DOCUMENT_MATERIALIZE_PAGEREF = /^PAGEREF ([A-Za-z][A-Za-z0-9_]{0,39}) \\h$/;

function materializeDocumentFields(document, options = {}) {
  const requestedTypes = options.types == null
    ? new Set(["SEQ", "REF"])
    : new Set((Array.isArray(options.types) ? options.types : [options.types]).map((value) => String(value).trim().toUpperCase()));
  const unsupportedTypes = [...requestedTypes].filter((type) => !new Set(["SEQ", "REF"]).has(type));
  if (unsupportedTypes.length) throw new TypeError(`Document field materialization supports only SEQ and REF cached results; unsupported type(s): ${unsupportedTypes.join(", ")}. PAGEREF requires a real pagination host.`);

  const fields = document.blocks.flatMap((block) => block.kind === "paragraph"
    ? block.runs.flatMap((run, runIndex) => run.inlineField ? [{ block, run, runIndex, instruction: String(run.inlineField.instruction || "").trim() }] : [])
    : []);
  const counters = new Map();
  const plannedValues = new Map();
  const seqFields = [];
  const refFields = [];
  let skippedPageReferences = 0;
  for (const field of fields) {
    const seq = DOCUMENT_MATERIALIZE_SEQ.exec(field.instruction);
    const ref = DOCUMENT_MATERIALIZE_REF.exec(field.instruction);
    if (seq) {
      if (!requestedTypes.has("SEQ")) continue;
      const counterKey = seq[1].toLowerCase();
      const value = (counters.get(counterKey) || 0) + 1;
      counters.set(counterKey, value);
      plannedValues.set(field.run, String(value));
      seqFields.push(field);
      continue;
    }
    if (ref) {
      if (requestedTypes.has("REF")) refFields.push({ ...field, bookmarkName: ref[1] });
      continue;
    }
    if (DOCUMENT_MATERIALIZE_PAGEREF.test(field.instruction)) {
      skippedPageReferences += 1;
      continue;
    }
    throw new TypeError(`Document field materialization encountered unsupported inline instruction: ${field.instruction || "(empty)"}.`);
  }

  const targetValues = new Map();
  const registerTarget = (name, value, source) => {
    const key = String(name || "").toLowerCase();
    if (!key) return;
    if (targetValues.has(key)) throw new Error(`Document field materialization found duplicate bookmark target ${name}.`);
    targetValues.set(key, { name, value: String(value ?? ""), source });
  };
  for (const field of fields) {
    if (!field.run.inlineField?.bookmarkName) continue;
    registerTarget(field.run.inlineField.bookmarkName, plannedValues.get(field.run) ?? field.run.text, `${field.block.id}/run/${field.runIndex}`);
  }
  for (const bookmark of document.bookmarks) {
    const target = document.blocks.find((block) => block.id === bookmark.targetId);
    const value = target?.kind === "paragraph"
      ? target.runs.map((run) => plannedValues.get(run) ?? String(run.text ?? "")).join("")
      : target?.text ?? target?.display;
    if (value === undefined) {
      if (options.strict !== false) throw new Error(`Document field materialization cannot resolve bookmark ${bookmark.name} target ${bookmark.targetId}.`);
      continue;
    }
    registerTarget(bookmark.name, value, bookmark.id);
  }

  const missingBookmarks = [];
  for (const field of refFields) {
    const target = targetValues.get(field.bookmarkName.toLowerCase());
    if (!target) {
      missingBookmarks.push(field.bookmarkName);
      continue;
    }
    plannedValues.set(field.run, target.value);
  }
  const uniqueMissingBookmarks = [...new Set(missingBookmarks)];
  if (options.strict !== false && uniqueMissingBookmarks.length) throw new Error(`Document field materialization cannot resolve bookmark(s): ${uniqueMissingBookmarks.join(", ")}.`);

  const changes = fields.flatMap((field) => {
    if (!plannedValues.has(field.run)) return [];
    const value = plannedValues.get(field.run);
    if (String(field.run.text ?? "") === value) return [];
    return [{ blockId: field.block.id, runIndex: field.runIndex, instruction: field.instruction, from: String(field.run.text ?? ""), to: value }];
  });
  if (!options.dryRun) {
    const changedBlocks = new Set();
    for (const change of changes) {
      const block = document.blocks.find((item) => item.id === change.blockId);
      block.runs[change.runIndex].text = change.to;
      changedBlocks.add(block);
    }
    for (const block of changedBlocks) block._syncText();
  }
  return {
    dryRun: Boolean(options.dryRun),
    updated: options.dryRun ? 0 : changes.length,
    wouldUpdate: changes.length,
    seqFields: seqFields.length,
    refFields: refFields.length,
    skippedPageReferences,
    missingBookmarks: uniqueMissingBookmarks,
    changes,
  };
}

export class DocumentModel {
  constructor(options = {}) {
    this.id = aid("doc");
    this.name = options.name || "New document";
    this.designPreset = options.designPreset || "default";
    this.theme = normalizeDocxThemeConfig(options.theme || {});
    this.defaultRunStyle = normalizeDocxRunStyle(options.defaultRunStyle || {}, this.theme);
    this.settings = normalizeDocxSettings(options.settings || {});
    this.styles = new DocumentStyleCollection(options.styles || {});
    this.blocks = [];
    this.bookmarks = [];
    this.notes = [];
    this.comments = [];
    this.bibliography = {
      selectedStyle: String(options.bibliography?.selectedStyle || ""),
      styleName: String(options.bibliography?.styleName || ""),
      uri: String(options.bibliography?.uri || ""),
    };
    this.bibliographySources = [];
    this.headers = [];
    this.footers = [];
    this.sectionSettings = [];
    const preservesEvenOddActivation = Object.prototype.hasOwnProperty.call(options.settings || {}, "evenAndOddHeaders");
    for (const source of options.bibliographySources || options.bibliography?.sources || []) this.addBibliographySource(source);
    const sourceBlocks = options.blocks || (options.paragraphs ? options.paragraphs.map((text, index) => ({ kind: "paragraph", text, styleId: index === 0 ? "Title" : "Normal" })) : [{ kind: "paragraph", text: "Start writing here...", styleId: "Normal" }]);
    for (const block of sourceBlocks) {
      if (block.kind === "table") this.addTable(block);
      else if (block.kind === "listItem") this.addListItem(block.text ?? "", block);
      else if (block.kind === "hyperlink") this.addHyperlink(block.text ?? "", block.url, block);
      else if (block.kind === "field") this.addField(block.instruction, block.display, block);
      else if (block.kind === "citation") this.addCitation(block.text ?? "", block.metadata || {}, block);
      else if (block.kind === "image") this.addImage(block);
      else if (block.kind === "section") this.addSection(block);
      else if (block.kind === "change") this.addChange(block.changeType || block.type, block.text ?? "", block);
      else if (!block.kind || block.kind === "paragraph") this.addParagraph(block.text ?? "", block);
      else throw new TypeError(`Unsupported document block kind ${block.kind}.`);
    }
    this.sectionSettings = normalizeDocxSectionSettings(options.sectionSettings || [], this.blocks);
    for (const header of options.headers || []) this.addHeader(header.text, { ...header, _restore: true });
    for (const footer of options.footers || []) this.addFooter(footer.text, { ...footer, _restore: true });
    if (!preservesEvenOddActivation && [...this.headers, ...this.footers].some((block) => block.referenceType === "even" && block.variantActive !== false)) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true });
    for (const bookmark of options.bookmarks || []) this.addBookmark(bookmark.targetId, bookmark.name, bookmark);
    for (const note of options.notes || []) this.addNote(note.kind || note.noteKind, note.targetId, note.text, note);
    for (const comment of options.comments || []) this.addComment(comment.targetId, comment.text, comment);
  }

  static create(options = {}) { return new DocumentModel(options); }
  get paragraphs() { return this.blocks.filter((block) => block.kind === "paragraph").map((block) => block.text); }
  get contentControls() { return documentContentControls(this); }
  get fontFamilies() {
    return officeFontFamilies(
      [this.toProto()],
      Object.values(this.theme?.fonts || {}),
    );
  }

  applyDesignPreset(name = "report", options = {}) {
    const presetName = String(name || "report");
    const presets = {
      report: {
        Title: { fontSize: 52, bold: true, fontFamily: "Aptos Display" },
        Heading1: { fontSize: 34, bold: true, fontFamily: "Aptos Display" },
        Heading2: { fontSize: 28, bold: true, fontFamily: "Aptos" },
        Normal: { fontSize: 22, fontFamily: "Aptos" },
        Caption: { name: "Caption", fontSize: 18, italic: true, fontFamily: "Aptos" },
        Callout: { name: "Callout", fontSize: 24, bold: true, fontFamily: "Aptos" },
        TableGrid: { name: "Table Grid", type: "table", fontSize: 20, fontFamily: "Aptos" },
      },
      memo: {
        Title: { fontSize: 44, bold: true, fontFamily: "Aptos Display" },
        Heading1: { fontSize: 30, bold: true, fontFamily: "Aptos" },
        Normal: { fontSize: 21, fontFamily: "Aptos" },
        Metadata: { name: "Metadata", fontSize: 18, fontFamily: "Aptos" },
        TableGrid: { name: "Table Grid", type: "table", fontSize: 20, fontFamily: "Aptos" },
      },
    };
    const preset = presets[presetName] || presets.report;
    for (const [id, style] of Object.entries(preset)) this.styles.add(id, { ...(this.styles.get(id) || {}), ...style, ...(options.styles?.[id] || {}) });
    this.designPreset = presetName;
    return this;
  }

  layoutJson(options = {}) {
    return documentLayoutJson(this, options);
  }

  addParagraph(text, config = {}) { const block = new DocumentParagraphBlock(this, text, config); this.blocks.push(block); return block; }
  fillContentControls(values = {}, options = {}) {
    const entries = values instanceof Map ? [...values.entries()] : Object.entries(values || {});
    const requested = new Map(entries.map(([tag, value]) => [String(tag), String(value ?? "")]));
    const controls = this.contentControls.filter((control) => control.controlType === "text");
    const matched = new Set(controls.filter((control) => requested.has(control.tag)).map((control) => control.tag));
    const missingTags = [...requested.keys()].filter((tag) => !matched.has(tag));
    if (options.strict !== false && missingTags.length) throw new Error(`Unknown document content-control tag(s): ${missingTags.join(", ")}`);
    let updated = 0;
    for (const control of controls) {
      if (!requested.has(control.tag)) continue;
      control.text = requested.get(control.tag);
      updated += 1;
    }
    return { updated, matchedTags: [...matched], missingTags };
  }
  setCheckboxContentControls(values = {}, options = {}) {
    const entries = values instanceof Map ? [...values.entries()] : Object.entries(values || {});
    const requested = new Map(entries.map(([tag, value]) => {
      if (typeof value !== "boolean") throw new TypeError(`Document checkbox content-control ${String(tag)} state must be boolean.`);
      return [String(tag), value];
    }));
    const controls = this.contentControls.filter((control) => control.controlType === "checkbox");
    const matched = new Set(controls.filter((control) => requested.has(control.tag)).map((control) => control.tag));
    const missingTags = [...requested.keys()].filter((tag) => !matched.has(tag));
    if (options.strict !== false && missingTags.length) throw new Error(`Unknown document checkbox content-control tag(s): ${missingTags.join(", ")}`);
    let updated = 0;
    for (const control of controls) {
      if (!requested.has(control.tag)) continue;
      control.checked = requested.get(control.tag);
      updated += 1;
    }
    return { updated, matchedTags: [...matched], missingTags };
  }
  setDropdownContentControls(values = {}, options = {}) {
    const entries = values instanceof Map ? [...values.entries()] : Object.entries(values || {});
    const requested = new Map(entries.map(([tag, value]) => {
      if (typeof value !== "string") throw new TypeError(`Document drop-down content-control ${String(tag)} selectedValue must be a string.`);
      return [String(tag), value];
    }));
    const controls = this.contentControls.filter((control) => control.controlType === "dropdown");
    const matched = new Set(controls.filter((control) => requested.has(control.tag)).map((control) => control.tag));
    const missingTags = [...requested.keys()].filter((tag) => !matched.has(tag));
    if (options.strict !== false && missingTags.length) throw new Error(`Unknown document drop-down content-control tag(s): ${missingTags.join(", ")}`);
    for (const control of controls) {
      if (!requested.has(control.tag)) continue;
      const selectedValue = requested.get(control.tag);
      if (!control.choices.some((choice) => choice.value === selectedValue)) {
        throw new TypeError(`Document drop-down content-control ${control.tag} selectedValue ${selectedValue} does not match a choice value.`);
      }
    }
    let updated = 0;
    for (const control of controls) {
      if (!requested.has(control.tag)) continue;
      control.selectedValue = requested.get(control.tag);
      updated += 1;
    }
    return { updated, matchedTags: [...matched], missingTags };
  }
  materializeFields(options = {}) { return materializeDocumentFields(this, options); }
  addListItem(text, config = {}) { const block = new DocumentListItemBlock(this, text, config); this.blocks.push(block); return block; }
  addList(items = [], config = {}) { return items.map((item) => this.addListItem(typeof item === "string" ? item : item.text, { ...config, ...(typeof item === "string" ? {} : item) })); }
  addHyperlink(text, url, config = {}) { const block = new DocumentHyperlinkBlock(this, text, url, config); this.blocks.push(block); return block; }
  addField(instruction, display, config = {}) { const block = new DocumentFieldBlock(this, instruction, display, config); this.blocks.push(block); return block; }
  addTableOfContents(config = {}) {
    const levels = String(config.levels ?? `${config.minLevel ?? 1}-${config.maxLevel ?? 3}`);
    const match = /^([1-9])-([1-9])$/.exec(levels);
    if (!match || Number(match[1]) > Number(match[2])) throw new TypeError("Document TOC levels must be an ascending range from 1-1 through 1-9.");
    const switches = [`\\o "${levels}"`];
    if (config.hyperlinks !== false) switches.push("\\h");
    if (config.hidePageNumbersInWeb !== false) switches.push("\\z");
    if (config.useOutlineLevels !== false) switches.push("\\u");
    const block = this.addField(
      `TOC ${switches.join(" ")}`,
      config.display ?? "(Table of contents will populate after fields are updated)",
      { ...config, complex: true },
    );
    if (config.updateFields !== false) this.settings = normalizeDocxSettings({ ...this.settings, updateFields: true });
    return block;
  }
  addBibliographySource(config = {}) { const source = new DocumentBibliographySource(this, config); this.bibliographySources.push(source); return source; }
  addCitation(text, metadata = {}, config = {}) {
    const block = new DocumentCitationBlock(this, text, metadata, config);
    const explicitTag = block.metadata.tag ?? block.metadata.bibliographyTag;
    const seed = explicitTag || block.metadata.source || block.metadata.title || block.name || `Source${this.bibliographySources.length + 1}`;
    const baseTag = explicitTag ? String(explicitTag).trim() : String(seed).replace(/[^A-Za-z0-9_.:-]+/g, "").slice(0, 255) || `Source${this.bibliographySources.length + 1}`;
    let tag = baseTag;
    if (!explicitTag) for (let suffix = 2; this.bibliographySources.some((source) => source.tag === tag && source.title !== (block.metadata.title || block.metadata.source || block.text)); suffix += 1) tag = `${baseTag.slice(0, 250)}${suffix}`;
    block.metadata.tag = tag;
    if (!this.bibliographySources.some((source) => source.tag === tag)) {
      const sourceConfig = { ...block.metadata, ...(block.metadata.bibliography || block.metadata.sourceData || {}), tag, title: block.metadata.title || block.metadata.source || block.text, sourceType: block.metadata.sourceType || (block.metadata.url ? "InternetSite" : "Misc") };
      this.addBibliographySource(sourceConfig);
    }
    this.blocks.push(block);
    if (!config._restore) {
      const bookmarkName = block.metadata.bookmark || `OpenOfficeCitation_${block.id.replace(/[^A-Za-z0-9_]/g, "_")}`.slice(0, 40);
      const bookmark = this.addBookmark(block, bookmarkName, { id: block.metadata.bookmarkId || `${block.id}/bookmark`, nativeId: block.metadata.bookmarkNativeId });
      block.metadata.bookmark = bookmark.name;
      block.metadata.bookmarkId = bookmark.id;
    }
    return block;
  }
  addImage(config = {}) { const block = new DocumentImageBlock(this, config); this.blocks.push(block); return block; }
  addSection(config = {}) { const block = new DocumentSectionBlock(this, config); this.blocks.push(block); return block; }
  addPageBreakSection(config = {}) { return this.addSection({ ...config, breakType: "nextPage" }); }
  addChange(changeType, text, config = {}) { const block = new DocumentChangeBlock(this, changeType, text, config); this.blocks.push(block); return block; }
  addInsertion(text, config = {}) { return this.addChange("insert", text, config); }
  addDeletion(text, config = {}) { return this.addChange("delete", text, config); }
  addTable(config = {}) { const block = new DocumentTableBlock(this, config); this.blocks.push(block); return block; }
  addHeader(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "header", text, config); this.headers.push(block); if (!config._restore && block.referenceType === "even" && block.variantActive !== false) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true }); return block; }
  addFooter(text, config = {}) { const block = new DocumentHeaderFooterBlock(this, "footer", text, config); this.footers.push(block); if (!config._restore && block.referenceType === "even" && block.variantActive !== false) this.settings = normalizeDocxSettings({ ...this.settings, evenAndOddHeaders: true }); return block; }
  addBookmark(target, name, config = {}) {
    const targetEndpoint = documentBookmarkEndpoint(target ?? config.target ?? config.targetId);
    const targetId = targetEndpoint?.id;
    const bookmarkName = String(name || config.name || "");
    const existing = this.bookmarks.find((bookmark) => bookmark.name === bookmarkName);
    if (existing && existing.targetId === targetId) {
      if (config.nativeId !== undefined) existing.nativeId = Number(config.nativeId);
      const configuredEnd = config.endTarget ?? config.end ?? config.endTargetId;
      if (configuredEnd) {
        existing.endTarget = documentBookmarkEndpoint(configuredEnd);
        existing.endTargetId = existing.endTarget?.id;
      }
      return existing;
    }
    const bookmark = new DocumentBookmark(this, target, name, config);
    this.bookmarks.push(bookmark);
    return bookmark;
  }
  addNote(kind, target, text, config = {}) { const note = new DocumentNote(this, kind, target, text, config); this.notes.push(note); return note; }
  addFootnote(target, text, config = {}) { return this.addNote("footnote", target, text, config); }
  addEndnote(target, text, config = {}) { return this.addNote("endnote", target, text, config); }
  addComment(target, text, config = {}) { const targetId = typeof target === "string" ? target : target?.id; const comment = new DocumentComment(this, targetId, text, config); this.comments.push(comment); return comment; }
  replyToComment(parent, text, config = {}) { const comment = typeof parent === "string" ? this.comments.find((item) => item.id === parent) : parent; if (!comment || !this.comments.includes(comment)) throw new Error(`Unknown parent document comment: ${typeof parent === "string" ? parent : parent?.id}`); return this.addComment(comment.targetId, text, { ...config, parentId: comment.id }); }
  setSettings(settings = {}) { this.settings = normalizeDocxSettings({ ...this.settings, ...settings }); return this; }
  setSectionSettings(sectionIndex, settings = {}) { this.sectionSettings = normalizeDocxSectionSettings([...this.sectionSettings.filter((entry) => entry.sectionIndex !== Number(sectionIndex)), { sectionIndex: Number(sectionIndex), ...settings }], this.blocks); return this; }
  resolve(id) {
    const token = String(id || "");
    if (token.endsWith("/text")) return documentTextRange(this, token);
    const contentControl = this.contentControls.find((control) => control.id === token);
    if (contentControl) return contentControl;
    const cellMatch = /^(.+)\/cell\/(\d+)\/(\d+)$/.exec(token);
    if (cellMatch) {
      const table = this.blocks.find((block) => block.kind === "table" && block.id === cellMatch[1]);
      const row = Number(cellMatch[2]);
      const column = Number(cellMatch[3]);
      if (table && row < table.rows && column < table.columns) return table.getCell(row, column);
      return undefined;
    }
    return token === `${this.id}/settings` ? this.settings : token === `${this.id}/theme` ? this.theme : this.id === token ? this : this.blocks.find((block) => block.id === token) || this.headers.find((block) => block.id === token) || this.footers.find((block) => block.id === token) || this.bookmarks.find((bookmark) => bookmark.id === token || bookmark.name === token) || this.notes.find((note) => note.id === token) || this.comments.find((comment) => comment.id === token) || this.bibliographySources.find((source) => source.id === token || source.tag === token) || this.styles.get(token);
  }

  toProto() { return { id: this.id, name: this.name, designPreset: this.designPreset, theme: this.theme, defaultRunStyle: this.defaultRunStyle, settings: this.settings, bibliography: this.bibliography, bibliographySources: this.bibliographySources.map((source) => source.toProto()), sectionSettings: this.sectionSettings, styles: Object.fromEntries(this.styles.values().map((style) => [style.id, style])), blocks: this.blocks.map((block) => block.toProto()), headers: this.headers.map((block) => block.toProto()), footers: this.footers.map((block) => block.toProto()), bookmarks: this.bookmarks.map((bookmark) => bookmark.toProto()), notes: this.notes.map((note) => note.toProto()), comments: this.comments.map((comment) => comment.toProto()) }; }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["paragraph", "contentControl", "table", "listItem", "hyperlink", "field", "citation", "bibliographySource", "image", "section", "change", "bookmark", "footnote", "endnote", "comment", "header", "footer"]);
    const records = [];
    if (kinds.has("document")) records.push({ kind: "document", id: this.id, name: this.name, blocks: this.blocks.length, contentControls: this.contentControls.length, sections: this.blocks.filter((block) => block.kind === "section").length + 1, bookmarks: this.bookmarks.length, notes: this.notes.length, footnotes: this.notes.filter((note) => note.kind === "footnote").length, endnotes: this.notes.filter((note) => note.kind === "endnote").length, bibliographySources: this.bibliographySources.length, designPreset: this.designPreset, defaultRunStyle: this.defaultRunStyle, settings: this.settings, sectionSettings: this.sectionSettings });
    if (kinds.has("theme")) records.push({ kind: "theme", id: `${this.id}/theme`, ...this.theme });
    if (kinds.has("settings")) records.push({ kind: "settings", id: `${this.id}/settings`, ...this.settings });
    if (kinds.has("layout")) records.push(...documentLayoutRecords(this, options));
    this.blocks.forEach((block, index) => { if (kinds.has(block.kind)) records.push(documentInspectRecord(this, block, index)); });
    if (kinds.has("contentControl")) records.push(...this.contentControls.map((control) => control.inspectRecord()));
    if (kinds.has("tableCell")) for (const table of this.blocks.filter((block) => block.kind === "table")) for (let row = 0; row < table.rows; row += 1) for (let column = 0; column < table.columns; column += 1) records.push(table.getCell(row, column).inspectRecord());
    if (kinds.has("header")) records.push(...this.headers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("footer")) records.push(...this.footers.map((block, index) => documentInspectRecord(this, block, index)));
    if (kinds.has("bookmark")) records.push(...this.bookmarks.map((bookmark) => bookmark.inspectRecord()));
    if (kinds.has("note") || kinds.has("footnote") || kinds.has("endnote")) records.push(...this.notes.filter((note) => kinds.has("note") || kinds.has(note.kind)).map((note, index) => note.inspectRecord(index)));
    if (kinds.has("bibliographySource")) records.push(...this.bibliographySources.map((source, index) => source.inspectRecord(index)));
    if (kinds.has("comment")) records.push(...this.comments.map((comment) => comment.inspectRecord()));
    if (kinds.has("textRange")) records.push(...documentTextRangeRecords(this));
    if (kinds.has("style")) records.push(...this.styles.values().map((style) => ({ kind: "style", ...style })));
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  verify(options = {}) {
    const issues = [];
    const knownStyleIds = new Set(this.styles.values().map((style) => style.id));
    const checkStyle = (block) => {
      if (block.styleId && !knownStyleIds.has(block.styleId)) {
        issues.push(verificationIssue("document", "unknownStyle", `${block.kind} ${block.id} references missing style ${block.styleId}.`, { severity: "warning", id: block.id, styleId: block.styleId }));
      }
    };
    if (this.blocks.length === 0) issues.push(verificationIssue("document", "emptyDocument", "Document has no body blocks."));
    for (const block of [...this.blocks, ...this.headers, ...this.footers]) checkStyle(block);
    const contentControlIds = new Set();
    const nativeContentControlIds = new Set();
    for (const control of this.contentControls) {
      if (!control.id || contentControlIds.has(control.id)) issues.push(verificationIssue("document", "invalidContentControlId", `Content control ${control.tag || "(untagged)"} requires a unique non-empty ID.`, { id: control.id, targetId: control.targetId }));
      else contentControlIds.add(control.id);
      if (!control.tag || control.tag.length > 64 || /[\u0000-\u001f\u007f]/.test(control.tag)) issues.push(verificationIssue("document", "invalidContentControlTag", `Content control ${control.id} tag must contain 1 to 64 characters without controls.`, { id: control.id, tag: control.tag }));
      if (control.alias.length > 255 || /[\u0000-\u001f\u007f]/.test(control.alias)) issues.push(verificationIssue("document", "invalidContentControlAlias", `Content control ${control.id} alias must contain at most 255 characters without controls.`, { id: control.id, alias: control.alias }));
      if (control.controlType !== "text" && control.controlType !== "checkbox" && control.controlType !== "dropdown") issues.push(verificationIssue("document", "invalidContentControlType", `Content control ${control.id} type must be text, checkbox, or dropdown.`, { id: control.id, controlType: control.controlType }));
      if (control.controlType === "checkbox" && (typeof control.checked !== "boolean" || control.text !== documentCheckboxGlyph(control.checked))) issues.push(verificationIssue("document", "invalidCheckboxContentControl", `Checkbox content control ${control.id} must have boolean checked state and its canonical visible glyph.`, { id: control.id, checked: control.checked, visibleText: control.text }));
      if (control.controlType === "dropdown") {
        try {
          const choices = normalizeDocumentDropdownChoices(control.choices);
          const selected = choices.find((choice) => choice.value === control.selectedValue);
          if (!selected || control.text !== selected.displayText) throw new TypeError("selected value and visible text do not match one declared choice");
        } catch (error) {
          issues.push(verificationIssue("document", "invalidDropdownContentControl", `Drop-down content control ${control.id} has invalid choices or selected state: ${error.message}.`, { id: control.id, selectedValue: control.selectedValue, visibleText: control.text }));
        }
      }
      if (control.nativeId !== undefined && (!Number.isInteger(control.nativeId) || control.nativeId < 1 || control.nativeId > 0x7fffffff || nativeContentControlIds.has(control.nativeId))) issues.push(verificationIssue("document", "invalidContentControlNativeId", `Content control ${control.id} has an invalid or duplicate nativeId.`, { id: control.id, nativeId: control.nativeId }));
      if (control.nativeId !== undefined) nativeContentControlIds.add(control.nativeId);
    }
    const bookmarkByName = new Map();
    const bookmarkNativeIds = new Set();
    const registerBookmarkIdentity = (name, id, nativeId) => {
      const key = String(name || "").toLowerCase();
      if (!/^[A-Za-z][A-Za-z0-9_]{0,39}$/.test(String(name || ""))) issues.push(verificationIssue("document", "invalidBookmarkName", `Bookmark ${id} name must start with an ASCII letter and contain only letters, digits, or underscores (maximum 40 characters).`, { id, name }));
      if (bookmarkByName.has(key)) issues.push(verificationIssue("document", "duplicateBookmarkName", `Bookmark ${id} duplicates name ${name}.`, { id, name }));
      else bookmarkByName.set(key, id);
      if (nativeId !== undefined) {
        if (!Number.isInteger(nativeId) || nativeId < 0 || nativeId > 0xffffffff) issues.push(verificationIssue("document", "invalidBookmarkNativeId", `Bookmark ${id} has invalid native ID ${nativeId}.`, { id, nativeId }));
        else if (bookmarkNativeIds.has(nativeId)) issues.push(verificationIssue("document", "duplicateBookmarkNativeId", `Bookmark ${id} duplicates native ID ${nativeId}.`, { id, nativeId }));
        bookmarkNativeIds.add(nativeId);
      }
    };
    for (const bookmark of this.bookmarks) {
      registerBookmarkIdentity(bookmark.name, bookmark.id, bookmark.nativeId);
    }
    const bibliographyByTag = new Map();
    for (const source of this.bibliographySources) {
      try { normalizeDocxBibliographySource(source); }
      catch (error) { issues.push(verificationIssue("document", "invalidBibliographySource", error.message, { id: source.id, tag: source.tag })); }
      if (bibliographyByTag.has(source.tag)) issues.push(verificationIssue("document", "duplicateBibliographyTag", `Bibliography source ${source.id} duplicates tag ${source.tag}.`, { id: source.id, tag: source.tag }));
      else bibliographyByTag.set(source.tag, source);
    }
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        for (const [runIndex, run] of (block.runs || []).entries()) {
          if (run.style?.runStyleId && !knownStyleIds.has(run.style.runStyleId)) issues.push(verificationIssue("document", "unknownRunStyle", `Paragraph ${block.id} references missing character style ${run.style.runStyleId}.`, { severity: "warning", id: block.id, runStyleId: run.style.runStyleId }));
          if (run.inlineField?.bookmarkName || run.inlineField?.bookmarkNativeId !== undefined) {
            const id = `${block.id}/run/${runIndex}/inline-field`;
            registerBookmarkIdentity(run.inlineField.bookmarkName, id, run.inlineField.bookmarkNativeId);
            if (!String(run.inlineField.instruction || "").startsWith("SEQ ")) issues.push(verificationIssue("document", "invalidInlineFieldBookmark", `Inline field ${id} may bookmark only a canonical SEQ cached result.`, { id, instruction: run.inlineField.instruction }));
          }
        }
      }
      if (block.kind === "paragraph" && /^\s*([-*•]|\d+[.)])\s+/.test(block.text)) {
        issues.push(verificationIssue("document", "fakeList", `Paragraph ${block.id} looks like a fake list item; use addListItem instead.`, { id: block.id }));
      }
      if (block.kind === "hyperlink" && block.anchor) {
        if (block.url) issues.push(verificationIssue("document", "ambiguousHyperlink", `Hyperlink ${block.id} cannot combine an external URL with an internal anchor.`, { id: block.id, url: block.url, anchor: block.anchor }));
        if (!bookmarkByName.has(String(block.anchor || "").toLowerCase())) issues.push(verificationIssue("document", "missingHyperlinkAnchor", `Hyperlink ${block.id} targets missing bookmark ${block.anchor}.`, { id: block.id, anchor: block.anchor }));
        if (block.anchor.length > 255) issues.push(verificationIssue("document", "invalidHyperlinkAnchor", `Hyperlink ${block.id} anchor exceeds 255 characters.`, { id: block.id, anchor: block.anchor }));
      } else if (block.kind === "hyperlink" && !/^https?:\/\//.test(block.url)) {
        issues.push(verificationIssue("document", "invalidHyperlink", `Hyperlink ${block.id} is missing an absolute http(s) URL or internal bookmark anchor.`, { id: block.id, url: block.url }));
      }
      if (block.kind === "hyperlink" && block.tooltip && String(block.tooltip).length > 260) issues.push(verificationIssue("document", "invalidHyperlinkTooltip", `Hyperlink ${block.id} tooltip exceeds 260 characters.`, { id: block.id }));
      if (block.kind === "field" && !block.instruction.trim()) {
        issues.push(verificationIssue("document", "emptyField", `Field ${block.id} is missing an instruction.`, { id: block.id }));
      }
      if (block.kind === "change") {
        if (!new Set(["insert", "delete"]).has(block.changeType)) issues.push(verificationIssue("document", "invalidChangeType", `Tracked change ${block.id} must be insert or delete.`, { id: block.id, changeType: block.changeType }));
        if (!String(block.author || "").trim() || String(block.author).length > 255 || /[\u0000-\u001f\u007f]/.test(String(block.author))) issues.push(verificationIssue("document", "invalidChangeAuthor", `Tracked change ${block.id} needs an author of at most 255 characters without controls.`, { id: block.id, author: block.author }));
        if (block.date !== undefined && Number.isNaN(Date.parse(block.date))) issues.push(verificationIssue("document", "invalidChangeDate", `Tracked change ${block.id} has an invalid date.`, { id: block.id, date: block.date }));
        if (String(block.text ?? "").length > 1_000_000) issues.push(verificationIssue("document", "changeTextTooLong", `Tracked change ${block.id} exceeds 1,000,000 characters.`, { id: block.id, textChars: String(block.text ?? "").length }));
      }
      if (block.kind === "citation" && block.metadata?.url && !/^https?:\/\//.test(String(block.metadata.url))) {
        issues.push(verificationIssue("document", "invalidCitationUrl", `Citation ${block.id} has a non-http(s) URL.`, { severity: "warning", id: block.id, url: block.metadata.url }));
      }
      if (block.kind === "citation" && !bibliographyByTag.has(block.metadata?.tag)) issues.push(verificationIssue("document", "missingCitationSource", `Citation ${block.id} references missing bibliography source ${block.metadata?.tag || "(none)"}.`, { id: block.id, tag: block.metadata?.tag }));
      if (block.kind === "image") {
        if (!block.dataUrl && !block.uri && !block.prompt) issues.push(verificationIssue("document", "emptyImage", `Image ${block.id} has no dataUrl, uri, or prompt.`, { id: block.id }));
        if (block.dataUrl && !imageDataFromDataUrl(block.dataUrl)) issues.push(verificationIssue("document", "invalidImageDataUrl", `Image ${block.id} has an unsupported data URL.`, { id: block.id }));
        if (!Number.isFinite(block.widthPx) || !Number.isFinite(block.heightPx) || block.widthPx <= 0 || block.heightPx <= 0) issues.push(verificationIssue("document", "invalidImageDimensions", `Image ${block.id} has invalid dimensions.`, { id: block.id, widthPx: block.widthPx, heightPx: block.heightPx }));
      }
      if (block.kind === "section") {
        if (!["portrait", "landscape"].includes(block.orientation)) issues.push(verificationIssue("document", "invalidSectionOrientation", `Section ${block.id} has invalid orientation ${block.orientation}.`, { id: block.id, orientation: block.orientation }));
        if (!["nextPage", "continuous", "evenPage", "oddPage", ""].includes(block.breakType)) issues.push(verificationIssue("document", "invalidSectionBreak", `Section ${block.id} has invalid break type ${block.breakType}.`, { id: block.id, breakType: block.breakType }));
        for (const [side, value] of Object.entries(block.margins || {})) if (!Number.isFinite(value) || value < 0) issues.push(verificationIssue("document", "invalidSectionMargin", `Section ${block.id} has an invalid ${side} margin.`, { id: block.id, side, value }));
        for (const [dimension, value] of Object.entries(block.pageSize || {})) if (!Number.isFinite(value) || value <= 0) issues.push(verificationIssue("document", "invalidSectionPageSize", `Section ${block.id} has invalid ${dimension}.`, { id: block.id, dimension, value }));
        const horizontalMargins = Number(block.margins?.left || 0) + Number(block.margins?.right || 0);
        if (Number.isFinite(block.pageSize?.widthTwips) && horizontalMargins >= block.pageSize.widthTwips) issues.push(verificationIssue("document", "sectionMarginsExceedPage", `Section ${block.id} horizontal margins exceed page width.`, { id: block.id, margins: block.margins, pageSize: block.pageSize }));
      }
      if (block.kind === "listItem") {
        if (!Number.isInteger(block.level) || block.level < 0 || block.level > 8) issues.push(verificationIssue("document", "invalidListLevel", `List item ${block.id} level must be an integer from 0 through 8.`, { id: block.id, level: block.level }));
        if (!Number.isInteger(block.start) || block.start < 1) issues.push(verificationIssue("document", "invalidListStart", `List item ${block.id} start must be a positive integer.`, { id: block.id, start: block.start }));
        if (!String(block.numberFormat || "").trim()) issues.push(verificationIssue("document", "invalidListNumberFormat", `List item ${block.id} is missing numberFormat.`, { id: block.id, numberFormat: block.numberFormat }));
        try {
          const pictureBullet = normalizeDocumentPictureBullet(block.pictureBullet);
          if (pictureBullet && (block.listType !== "bullet" || block.numberFormat !== "bullet")) issues.push(verificationIssue("document", "invalidPictureBulletListType", `List item ${block.id} picture bullet requires bullet list semantics.`, { id: block.id, listType: block.listType, numberFormat: block.numberFormat }));
        } catch (error) {
          issues.push(verificationIssue("document", "invalidPictureBullet", `List item ${block.id} has an invalid picture bullet: ${error.message}`, { id: block.id }));
        }
      }
      if (block.kind === "table") {
        if (!block.rows || !block.columns) issues.push(verificationIssue("document", "emptyTable", `Table ${block.id} has no rows or columns.`, { id: block.id, rows: block.rows, columns: block.columns }));
        if (block.columns > 12) issues.push(verificationIssue("document", "wideTable", `Table ${block.id} has ${block.columns} columns and may not fit the page.`, { severity: "warning", id: block.id, columns: block.columns }));
        if (!Number.isFinite(block.widthDxa) || block.widthDxa <= 0) issues.push(verificationIssue("document", "invalidTableWidth", `Table ${block.id} has an invalid width.`, { id: block.id, widthDxa: block.widthDxa }));
        if (!Number.isFinite(block.indentDxa) || block.indentDxa < 0) issues.push(verificationIssue("document", "invalidTableIndent", `Table ${block.id} has an invalid indent.`, { id: block.id, indentDxa: block.indentDxa }));
        const formattingColumns = block.cells?.length ? block.gridColumns : block.columns;
        if (!Array.isArray(block.columnWidthsDxa) || block.columnWidthsDxa.length !== formattingColumns) issues.push(verificationIssue("document", "invalidTableColumnWidths", `Table ${block.id} needs one column width per logical grid column.`, { id: block.id, columns: formattingColumns, columnWidthsDxa: block.columnWidthsDxa }));
        else {
          if (block.columnWidthsDxa.some((value) => !Number.isFinite(value) || value <= 0)) issues.push(verificationIssue("document", "invalidTableColumnWidth", `Table ${block.id} contains an invalid column width.`, { id: block.id, columnWidthsDxa: block.columnWidthsDxa }));
          const widthSum = block.columnWidthsDxa.reduce((sum, value) => sum + value, 0);
          if (Number.isFinite(block.widthDxa) && widthSum !== block.widthDxa) issues.push(verificationIssue("document", "tableColumnWidthMismatch", `Table ${block.id} column widths do not equal the table width.`, { id: block.id, widthDxa: block.widthDxa, columnWidthsDxa: block.columnWidthsDxa, widthSum }));
        }
        for (const [side, value] of Object.entries(block.cellMarginsDxa || {})) if (!Number.isFinite(value) || value < 0) issues.push(verificationIssue("document", "invalidTableCellMargin", `Table ${block.id} has an invalid ${side} cell margin.`, { id: block.id, side, value }));
        if (!Number.isFinite(block.borderSize) || block.borderSize < 0) issues.push(verificationIssue("document", "invalidTableBorderSize", `Table ${block.id} has an invalid border size.`, { id: block.id, borderSize: block.borderSize }));
        if (!/^[A-F0-9]{6}$/.test(block.borderColor)) issues.push(verificationIssue("document", "invalidTableBorderColor", `Table ${block.id} has an invalid border color.`, { id: block.id, borderColor: block.borderColor }));
        if (!/^[A-F0-9]{6}$/.test(block.headerFill)) issues.push(verificationIssue("document", "invalidTableHeaderFill", `Table ${block.id} has an invalid header fill.`, { id: block.id, headerFill: block.headerFill }));
        block.values.forEach((row, rowIndex) => {
          if (row.length !== block.columns) issues.push(verificationIssue("document", "raggedTableRows", `Table ${block.id} row ${rowIndex} has ${row.length} cells; expected ${block.columns}.`, { id: block.id, row: rowIndex, cells: row.length, expected: block.columns }));
          for (const cell of row) {
            if (String(cell ?? "").length > 240) issues.push(verificationIssue("document", "tableCellTooLong", `Table ${block.id} contains paragraph-like cell content.`, { severity: "warning", id: block.id }));
          }
        });
      }
    }
    const blockIds = new Set(this.blocks.map((block) => block.id));
    const noteTargets = new Set();
    const noteNativeIds = new Set();
    for (const note of this.notes) {
      const target = this.blocks.find((block) => block.id === note.targetId);
      if (!target) issues.push(verificationIssue("document", "danglingNote", `${note.kind} ${note.id} points at missing block ${note.targetId}.`, { id: note.id, kind: note.kind, targetId: note.targetId }));
      else if (!new Set(["paragraph", "listItem"]).has(target.kind)) issues.push(verificationIssue("document", "unsupportedNoteTarget", `${note.kind} ${note.id} must target a paragraph or list item.`, { id: note.id, kind: note.kind, targetId: note.targetId, targetKind: target.kind }));
      if (noteTargets.has(note.targetId)) issues.push(verificationIssue("document", "duplicateNoteTarget", `${note.kind} ${note.id} shares target ${note.targetId}; the bounded profile permits one note per block.`, { id: note.id, kind: note.kind, targetId: note.targetId }));
      noteTargets.add(note.targetId);
      if (!String(note.text || "").length || String(note.text).length > 1_000_000 || /[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f]/.test(String(note.text))) issues.push(verificationIssue("document", "invalidNoteText", `${note.kind} ${note.id} text must contain 1 through 1,000,000 XML-safe characters.`, { id: note.id, kind: note.kind, textChars: String(note.text || "").length }));
      if (note.nativeId !== undefined) {
        const key = `${note.kind}:${note.nativeId}`;
        if (!Number.isInteger(note.nativeId) || note.nativeId < 1 || note.nativeId > 2_147_483_647) issues.push(verificationIssue("document", "invalidNoteNativeId", `${note.kind} ${note.id} native ID must be a positive 32-bit integer.`, { id: note.id, kind: note.kind, nativeId: note.nativeId }));
        else if (noteNativeIds.has(key)) issues.push(verificationIssue("document", "duplicateNoteNativeId", `${note.kind} ${note.id} duplicates native ID ${note.nativeId}.`, { id: note.id, kind: note.kind, nativeId: note.nativeId }));
        noteNativeIds.add(key);
      }
    }
    const blockIndexes = new Map(this.blocks.map((block, index) => [block.id, index]));
    for (const bookmark of this.bookmarks) {
      const validateEndpoint = (endpoint, end = false) => {
        const blockId = documentBookmarkEndpointBlockId(endpoint);
        const block = this.blocks.find((item) => item.id === blockId);
        if (!block) {
          issues.push(verificationIssue("document", end ? "missingBookmarkEndTarget" : "missingBookmarkTarget", `Bookmark ${bookmark.id} points at missing ${end ? "end" : "start"} block ${blockId}.`, { id: bookmark.id, targetId: endpoint?.id }));
          return false;
        }
        if (endpoint?.type === "tableCell") {
          if (block.kind !== "table" || !Number.isInteger(endpoint.row) || !Number.isInteger(endpoint.column) || endpoint.row < 0 || endpoint.column < 0 || endpoint.row >= block.rows || endpoint.column >= block.columns) {
            issues.push(verificationIssue("document", "invalidBookmarkTableCell", `Bookmark ${bookmark.id} points at invalid table cell ${endpoint.id}.`, { id: bookmark.id, targetId: endpoint.id, tableId: endpoint.tableId, row: endpoint.row, column: endpoint.column }));
            return false;
          }
        } else if (block.kind === "table") {
          issues.push(verificationIssue("document", "unsupportedBookmarkTarget", `Bookmark ${bookmark.id} must target a specific table cell, not the table block.`, { id: bookmark.id, targetId: endpoint?.id }));
          return false;
        }
        return true;
      };
      const startValid = validateEndpoint(bookmark.target, false);
      const endValid = validateEndpoint(bookmark.endTarget, true);
      if (startValid && endValid && compareDocumentBookmarkEndpoints(bookmark.target, bookmark.endTarget, blockIndexes) > 0) issues.push(verificationIssue("document", "reversedBookmarkRange", `Bookmark ${bookmark.id} ends before it starts.`, { id: bookmark.id, targetId: bookmark.targetId, endTargetId: bookmark.endTargetId }));
      const targetBlock = this.blocks.find((block) => block.id === bookmark.targetId);
      if (targetBlock?.kind === "paragraph" && targetBlock.runs.some((run) => run.inlineField?.bookmarkName)) issues.push(verificationIssue("document", "nestedBoundedBookmark", `Bookmark ${bookmark.id} cannot wrap paragraph ${targetBlock.id}, which already contains an inline-field bookmark.`, { id: bookmark.id, targetId: targetBlock.id }));
    }
    const finalSectionIndex = this.blocks.filter((block) => block.kind === "section").length;
    for (const block of [...this.headers, ...this.footers]) {
      if (block.sectionIndex !== undefined && (!Number.isInteger(block.sectionIndex) || block.sectionIndex < 0 || block.sectionIndex > finalSectionIndex)) {
        issues.push(verificationIssue("document", "invalidHeaderFooterSection", `${block.kind} ${block.id} targets invalid section index ${block.sectionIndex}; expected 0 through ${finalSectionIndex}.`, { id: block.id, kind: block.kind, sectionIndex: block.sectionIndex, finalSectionIndex }));
      }
    }
    const commentParaIds = new Set();
    const commentDurableIds = new Set();
    for (const comment of this.comments) {
      if (!blockIds.has(comment.targetId)) issues.push(verificationIssue("document", "danglingComment", `Comment ${comment.id} points at a missing block.`, { id: comment.id, targetId: comment.targetId }));
      if (comment.date && Number.isNaN(Date.parse(comment.date))) issues.push(verificationIssue("document", "invalidCommentDate", `Comment ${comment.id} has an invalid date.`, { id: comment.id, date: comment.date }));
      const parent = comment.parentId ? this.comments.find((item) => item.id === comment.parentId) : undefined;
      if (comment.parentId && !parent) issues.push(verificationIssue("document", "missingCommentParent", `Comment ${comment.id} points at missing parent comment ${comment.parentId}.`, { id: comment.id, parentId: comment.parentId }));
      if (parent && parent.targetId !== comment.targetId) issues.push(verificationIssue("document", "commentParentTargetMismatch", `Comment ${comment.id} and parent ${parent.id} target different blocks.`, { id: comment.id, parentId: parent.id, targetId: comment.targetId, parentTargetId: parent.targetId }));
      const commentParaNumber = /^[0-9A-Fa-f]{8}$/.test(comment.paraId || "") ? Number.parseInt(comment.paraId, 16) : undefined;
      if (comment.paraId && (!commentParaNumber || commentParaNumber >= 0x80000000)) issues.push(verificationIssue("document", "invalidCommentParaId", `Comment ${comment.id} has invalid paraId ${comment.paraId}; expected 00000001 through 7FFFFFFF.`, { id: comment.id, paraId: comment.paraId }));
      else if (comment.paraId && commentParaIds.has(comment.paraId.toUpperCase())) issues.push(verificationIssue("document", "duplicateCommentParaId", `Comment ${comment.id} duplicates paraId ${comment.paraId}.`, { id: comment.id, paraId: comment.paraId }));
      if (comment.paraId) commentParaIds.add(comment.paraId.toUpperCase());
      const durableNumber = /^[0-9A-Fa-f]{8}$/.test(comment.durableId || "") ? Number.parseInt(comment.durableId, 16) : undefined;
      if (comment.durableId && (!durableNumber || durableNumber >= 0x7FFFFFFF)) issues.push(verificationIssue("document", "invalidCommentDurableId", `Comment ${comment.id} has invalid durableId ${comment.durableId}.`, { id: comment.id, durableId: comment.durableId }));
      else if (comment.durableId && commentDurableIds.has(comment.durableId.toUpperCase())) issues.push(verificationIssue("document", "duplicateCommentDurableId", `Comment ${comment.id} duplicates durableId ${comment.durableId}.`, { id: comment.id, durableId: comment.durableId }));
      if (comment.durableId) commentDurableIds.add(comment.durableId.toUpperCase());
      if (comment.dateUtc && Number.isNaN(Date.parse(comment.dateUtc))) issues.push(verificationIssue("document", "invalidCommentDateUtc", `Comment ${comment.id} has an invalid UTC date.`, { id: comment.id, dateUtc: comment.dateUtc }));
      if (comment.parentId && comment.intelligentPlaceholder) issues.push(verificationIssue("document", "invalidCommentReplyPlaceholder", `Reply comment ${comment.id} must not be an intelligent placeholder.`, { id: comment.id, parentId: comment.parentId }));
      if (comment.person) {
        const providerId = String(comment.person.providerId || "");
        const userId = String(comment.person.userId || "");
        if (!providerId || providerId.length > 100 || !userId || userId.length > 300) issues.push(verificationIssue("document", "invalidCommentPerson", `Comment ${comment.id} has invalid person presence metadata.`, { id: comment.id, providerId, userId }));
      }
    }
    if (options.visualQa || options.renderQa) {
      const layout = this.layoutJson(options);
      for (const element of layout.elements) {
        const pageInfo = layout.pages.find((page) => page.page === element.page);
        if (pageInfo && element.bbox[3] > pageInfo.height - pageInfo.margin * 2) issues.push(verificationIssue("document", "layoutElementTooTall", `Block ${element.id} is taller than the usable page area.`, { severity: "warning", id: element.id, page: element.page, bbox: element.bbox }));
        if (pageInfo && element.bbox[1] + element.bbox[3] > pageInfo.height - pageInfo.margin + 1) issues.push(verificationIssue("document", "layoutElementOverflow", `Block ${element.id} may overflow page ${element.page}.`, { severity: "warning", id: element.id, page: element.page, bbox: element.bbox }));
      }
    }
    return verificationResult("document", issues, options);
  }

  help(query = "*", options = {}) {
    return documentHelp(query, options);
  }

  async render(options = {}) {
    if (options.format === "layout") return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "document", format: "layout", target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    if (documentRenderUsesDocxSource(options)) return renderDocumentFromDocx(this, options);
    const width = 612;
    const margin = 72;
    let y = 72;
    const parts = [`<rect width="100%" height="100%" fill="white"/>`];
    const firstPageHeaderFooter = resolveDocxPageHeaderFooter(planDocxHeaderFooterSections(this), 0, 1);
    const firstPageHeaderIds = new Set(firstPageHeaderFooter.headers);
    const firstPageFooterIds = new Set(firstPageHeaderFooter.footers);
    for (const header of this.headers.filter((block) => firstPageHeaderIds.has(block.id))) {
      parts.push(`<text x="${margin}" y="${Math.max(28, y - 36)}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(header.text)}</text>`);
    }
    let listCounters = new Map();
    for (const block of this.blocks) {
      if (block.kind === "paragraph") {
        const style = this.styles.effective(block.styleId) || this.styles.get("Normal");
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        const runs = block.runs?.length ? block.runs : [{ text: block.text, style: {} }];
        const tspans = runs.map((run, index) => {
          const runStyle = documentEffectiveRunStyle(this, block, run);
          return `<tspan${index ? "" : ` x=\"${margin}\"`} font-family="${xmlEscape(runStyle.effectiveFontFamily || runStyle.fontFamily || "Arial")}" font-size="${Math.max(5, (runStyle.effectiveFontSize || style.fontSize || 22) / 2)}" font-style="${runStyle.effectiveItalic ? "italic" : "normal"}" font-weight="${runStyle.effectiveBold ? "700" : "400"}" fill="${xmlEscape(runStyle.effectiveColor || "#111827")}">${xmlEscape(run.text)}</tspan>`;
        }).join("");
        parts.push(`<text x="${margin}" y="${y}">${tspans}</text>`);
        y += fontSize * 1.6;
      } else if (block.kind === "hyperlink") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#2563eb" text-decoration="underline">${xmlEscape(block.text)}</text>`);
        y += 20;
      } else if (block.kind === "field") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#334155">${xmlEscape(block.display)} (${xmlEscape(block.instruction)})</text>`);
        y += 20;
      } else if (block.kind === "citation") {
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="11" fill="#475569">${xmlEscape(block.text)}</text>`);
        y += 20;
      } else if (block.kind === "image") {
        const imageWidth = Math.max(16, Math.min(width - margin * 2, Number(block.widthPx) || 240));
        const imageHeight = Math.max(16, Math.min(360, Number(block.heightPx) || 160));
        if (block.dataUrl) {
          parts.push(`<image href="${attrEscape(block.dataUrl)}" x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" preserveAspectRatio="xMidYMid meet"/>`);
          parts.push(`<text x="${margin}" y="${y + imageHeight + 14}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(block.alt || block.name || "image")}</text>`);
        } else {
          parts.push(`<rect x="${margin}" y="${y}" width="${imageWidth}" height="${imageHeight}" fill="#fef3c7" stroke="#f59e0b"/>`);
          parts.push(`<text x="${margin + 8}" y="${y + 18}" font-family="Arial" font-size="11" fill="#92400e">${xmlEscape(block.alt || block.prompt || block.uri || block.name || "image")}</text>`);
        }
        y += imageHeight + (block.dataUrl ? 36 : 20);
      } else if (block.kind === "section") {
        parts.push(`<line x1="${margin}" x2="${width - margin}" y1="${y}" y2="${y}" stroke="#94a3b8" stroke-dasharray="4 4"/>`);
        parts.push(`<text x="${margin}" y="${y + 16}" font-family="Arial" font-size="10" fill="#64748b">section break: ${xmlEscape(block.breakType)} ${xmlEscape(block.orientation)} ${block.pageSize.widthTwips}x${block.pageSize.heightTwips}</text>`);
        y += 34;
      } else if (block.kind === "change") {
        const marker = block.changeType === "delete" ? "-" : "+";
        const fill = block.changeType === "delete" ? "#dc2626" : "#047857";
        const decoration = block.changeType === "delete" ? " text-decoration=\"line-through\"" : "";
        parts.push(`<text x="${margin}" y="${y}" font-family="Arial" font-size="10" fill="#64748b">tracked ${xmlEscape(block.changeType)} by ${xmlEscape(block.author)}</text>`);
        parts.push(`<text x="${margin + 92}" y="${y}" font-family="Arial" font-size="11" fill="${fill}"${decoration}>${xmlEscape(marker)} ${xmlEscape(block.text)}</text>`);
        y += 22;
      } else if (block.kind === "listItem") {
        const style = this.styles.effective(block.styleId) || this.styles.get("Normal") || {};
        const fontSize = Math.max(10, (style.fontSize || 22) / 2);
        const key = `${block.listType}:${block.level}`;
        const next = (listCounters.get(key) || 0) + 1;
        listCounters.set(key, next);
        const marker = block.listType === "number" ? `${next}.` : "•";
        const x = margin + block.level * 24;
        if (block.pictureBullet?.dataUrl) {
          const markerWidth = Math.max(4, Number(block.pictureBullet.widthPt) || fontSize);
          const markerHeight = Math.max(4, Number(block.pictureBullet.heightPt) || fontSize);
          parts.push(`<image data-picture-bullet="embedded" href="${attrEscape(block.pictureBullet.dataUrl)}" x="${x}" y="${y - markerHeight}" width="${markerWidth}" height="${markerHeight}" preserveAspectRatio="xMidYMid meet"/>`);
        } else if (block.pictureBullet?.uri) {
          const markerSize = Math.max(4, Number(block.pictureBullet.widthPt) || fontSize);
          parts.push(`<rect data-picture-bullet="external" data-uri="${attrEscape(block.pictureBullet.uri)}" x="${x}" y="${y - markerSize}" width="${markerSize}" height="${markerSize}" rx="2" fill="#dbeafe" stroke="#2563eb"/>`);
        } else {
          parts.push(`<text x="${x}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${style.italic ? "italic" : "normal"}" font-weight="${style.bold ? "700" : "400"}" fill="${xmlEscape(style.color || "#111827")}">${xmlEscape(marker)}</text>`);
        }
        parts.push(`<text x="${x + 22}" y="${y}" font-family="${xmlEscape(style.fontFamily || "Arial")}" font-size="${fontSize}" font-style="${style.italic ? "italic" : "normal"}" font-weight="${style.bold ? "700" : "400"}" fill="${xmlEscape(style.color || "#111827")}">${xmlEscape(block.text)}</text>`);
        y += Math.max(20, fontSize * 1.5);
      } else if (block.kind === "table") {
        const cellW = (width - margin * 2) / Math.max(1, block.columns);
        const cellH = 24;
        for (let r = 0; r < block.rows; r++) {
          for (let c = 0; c < block.columns; c++) {
            const x = margin + c * cellW;
            const fill = r === 0 ? "#f1f5f9" : "#ffffff";
            parts.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${fill}" stroke="#cbd5e1"/>`);
            parts.push(`<text x="${x + 5}" y="${y + 16}" font-family="Arial" font-size="11" fill="#111827">${xmlEscape(block.values[r]?.[c] ?? "")}</text>`);
          }
          y += cellH;
        }
        y += 16;
      }
    }
    const height = Math.max(792, y + 72);
    for (const footer of this.footers.filter((block) => firstPageFooterIds.has(block.id))) {
      parts.push(`<text x="${margin}" y="${height - 36}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(footer.text)}</text>`);
    }
    return new FileBlob(`<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}">${parts.join("")}</svg>`, { type: "image/svg+xml" });
  }
}

export class DocumentFile {
  static async inspectDocx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, DOCX_PACKAGE_CONFIG);
  }

  static async patchDocx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, DOCX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: DOCX_MIME, metadata: { artifactKind: "document", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportDocx(document, options = {}) {
    const { exportDocxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return exportDocxWithOpenChestnut(document, options);
  }

  static async importDocx(blobOrBuffer, options = {}) {
    const { importDocxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return importDocxWithOpenChestnut(blobOrBuffer, options);
  }

  static async addTrackedReplacement(blobOrBuffer, options = {}) {
    const { addDocxTrackedReplacementWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return addDocxTrackedReplacementWithOpenChestnut(blobOrBuffer, options);
  }

  static async finalizeRevisions(blobOrBuffer, options = {}) {
    const { finalizeDocxRevisionsWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return finalizeDocxRevisionsWithOpenChestnut(blobOrBuffer, options);
  }
}
