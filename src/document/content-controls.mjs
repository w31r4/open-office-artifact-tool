import { aid } from "../shared/ids.mjs";
import { isXmlSafeText } from "../shared/xml.mjs";

const CHECKBOX_GLYPHS = Object.freeze({ checked: "☒", unchecked: "☐" });
const MAX_CHOICES = 256;
const MAX_CHOICE_TEXT_LENGTH = 255;
const DATE_VALUE_PATTERN = /^(\d{4})-(\d{2})-(\d{2})$/;

export function documentCheckboxGlyph(checked) {
  return checked ? CHECKBOX_GLYPHS.checked : CHECKBOX_GLYPHS.unchecked;
}

export function normalizeDocumentContentControlChoices(value, controlType = "drop-down") {
  if (!Array.isArray(value) || value.length < 1 || value.length > MAX_CHOICES) {
    throw new TypeError(`Document ${controlType} content control requires 1 through ${MAX_CHOICES} choices.`);
  }
  const values = new Set();
  const displayTexts = new Set();
  return value.map((choice, index) => {
    const source = typeof choice === "string" ? { displayText: choice, value: choice } : choice;
    if (!source || typeof source !== "object") throw new TypeError(`Document ${controlType} choice ${index + 1} must be a string or object.`);
    const displayText = source.displayText ?? source.text ?? source.label ?? "";
    const itemValue = source.value ?? source.id ?? "";
    if (typeof displayText !== "string" || typeof itemValue !== "string") {
      throw new TypeError(`Document ${controlType} choice ${index + 1} displayText and value must be strings.`);
    }
    if (!displayText || !itemValue || displayText.length > MAX_CHOICE_TEXT_LENGTH || itemValue.length > MAX_CHOICE_TEXT_LENGTH || !isXmlSafeText(displayText) || !isXmlSafeText(itemValue) || /[\u0000-\u001f\u007f]/.test(displayText + itemValue)) {
      throw new TypeError(`Document ${controlType} choice ${index + 1} requires XML-safe displayText and value strings of 1 through ${MAX_CHOICE_TEXT_LENGTH} characters.`);
    }
    if (values.has(itemValue) || displayTexts.has(displayText)) {
      throw new TypeError(`Document ${controlType} content-control choice values and displayText strings must be unique.`);
    }
    values.add(itemValue);
    displayTexts.add(displayText);
    return { displayText, value: itemValue };
  });
}

export function documentContentControlChoice(control, value) {
  return control?.choices?.find((choice) => choice.value === value);
}

export function normalizeDocumentComboBoxValue(value) {
  if (typeof value !== "string") throw new TypeError("Document combo-box content-control value must be a string.");
  if (!value || value.length > MAX_CHOICE_TEXT_LENGTH || !isXmlSafeText(value) || /[\u0000-\u001f\u007f]/.test(value)) {
    throw new TypeError(`Document combo-box content-control value must be XML-safe and contain 1 through ${MAX_CHOICE_TEXT_LENGTH} characters.`);
  }
  return value;
}

export function documentComboBoxVisibleText(control, value = control?.value) {
  return documentContentControlChoice(control, value)?.displayText ?? value;
}

export function normalizeDocumentDateValue(value) {
  if (typeof value !== "string") throw new TypeError("Document date content-control dateValue must be a string in YYYY-MM-DD form.");
  const match = DATE_VALUE_PATTERN.exec(value);
  if (!match) throw new TypeError("Document date content-control dateValue must use canonical YYYY-MM-DD form.");
  const year = Number(match[1]);
  const month = Number(match[2]);
  const day = Number(match[3]);
  const leap = year % 4 === 0 && (year % 100 !== 0 || year % 400 === 0);
  const daysInMonth = [31, leap ? 29 : 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];
  if (year < 1 || month < 1 || month > 12 || day < 1 || day > daysInMonth[month - 1]) {
    throw new TypeError("Document date content-control dateValue must be a real Gregorian date from 0001-01-01 through 9999-12-31.");
  }
  return value;
}

export function normalizeDocumentContentControl(value) {
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
    : source.dateValue !== undefined ? "date"
      : source.choices !== undefined || source.items !== undefined || source.options !== undefined || source.selectedValue !== undefined ? "dropdown" : "text";
  const rawType = String(source.controlType ?? source.type ?? inferredType).trim().toLowerCase();
  const controlType = rawType === "text" || rawType === "plain-text" || rawType === "plaintext"
    ? "text"
    : rawType === "checkbox" || rawType === "check-box" ? "checkbox"
      : rawType === "dropdown" || rawType === "drop-down" || rawType === "drop_down" ? "dropdown"
        : rawType === "combobox" || rawType === "combo-box" || rawType === "combo_box" ? "comboBox"
          : rawType === "date" || rawType === "datepicker" || rawType === "date-picker" || rawType === "date_picker" ? "date" : undefined;
  if (!controlType) throw new TypeError("Document content control type must be text, checkbox, dropdown, comboBox, or date.");
  if (controlType === "checkbox" && source.checked !== undefined && typeof source.checked !== "boolean") {
    throw new TypeError("Document checkbox content control checked state must be boolean.");
  }
  const choices = controlType === "dropdown" || controlType === "comboBox"
    ? normalizeDocumentContentControlChoices(source.choices ?? source.items ?? source.options, controlType === "comboBox" ? "combo-box" : "drop-down")
    : undefined;
  const requestedSelectedValue = source.selectedValue ?? source.value;
  if (controlType === "dropdown" && requestedSelectedValue !== undefined && typeof requestedSelectedValue !== "string") {
    throw new TypeError("Document drop-down content-control selectedValue must be a string.");
  }
  const selectedValue = controlType === "dropdown" ? requestedSelectedValue ?? choices[0].value : undefined;
  if (controlType === "dropdown" && !choices.some((choice) => choice.value === selectedValue)) {
    throw new TypeError(`Document drop-down content-control selectedValue ${selectedValue} does not match a choice value.`);
  }
  const comboBoxValue = controlType === "comboBox"
    ? normalizeDocumentComboBoxValue(source.value ?? source.selectedValue ?? choices[0].value)
    : undefined;
  const dateValue = controlType === "date" ? normalizeDocumentDateValue(source.dateValue ?? source.value) : undefined;
  return {
    id: String(source.id || aid("dcc")),
    tag,
    alias: String(source.alias ?? source.title ?? tag),
    nativeId,
    controlType,
    ...(controlType === "checkbox" ? { checked: source.checked === true } : {}),
    ...(controlType === "dropdown" ? { choices, selectedValue } : {}),
    ...(controlType === "comboBox" ? { choices, value: comboBoxValue } : {}),
    ...(controlType === "date" ? { dateValue } : {}),
  };
}

function paragraphBinding(block, runIndex) {
  const placement = runIndex === undefined ? "block" : "inline";
  const run = placement === "block" ? block.runs[0] : block.runs[runIndex];
  return {
    document: block.document,
    block,
    placement,
    runIndex,
    control: placement === "block" ? block.blockContentControl : run?.contentControl,
    targetId: block.id,
    getText: () => placement === "block" ? String(block.text ?? "") : String(run?.text ?? ""),
    setText: (value) => {
      run.text = String(value ?? "");
      block._syncText();
    },
  };
}

function tableCellBinding(table, row, column) {
  const record = table.cells?.find((cell) => cell.row === row && cell.column === column);
  return {
    document: table.document,
    block: table,
    placement: "tableCell",
    row,
    column,
    control: record?.contentControl,
    targetId: `${table.id}/cell/${row}/${column}`,
    getText: () => String(table.values[row]?.[column] ?? ""),
    setText: (value) => {
      table.ensureCell(row, column);
      table.textPatches = table.textPatches.filter((patch) => patch.row !== row || patch.column !== column);
      table.values[row][column] = String(value ?? "");
    },
  };
}

class DocumentContentControlHandle {
  constructor(binding) {
    this.document = binding.document;
    this.block = binding.block;
    this.placement = binding.placement;
    this.runIndex = binding.runIndex;
    this.row = binding.row;
    this.column = binding.column;
    this.kind = "contentControl";
    this.binding = binding;
  }
  get run() { return this.placement === "inline" ? this.block.runs[this.runIndex] : this.placement === "block" ? this.block.runs[0] : undefined; }
  get control() { return this.binding.control; }
  get id() { return this.control?.id; }
  get targetId() { return this.binding.targetId; }
  get tag() { return this.control?.tag || ""; }
  set tag(value) { this.control.tag = String(value ?? "").trim(); }
  get alias() { return this.control?.alias || ""; }
  set alias(value) {
    const next = String(value ?? "");
    if ((this.placement === "block" || this.placement === "tableCell") && !next.length) throw new TypeError("Document block and table-cell text content controls require a non-empty alias.");
    this.control.alias = next;
  }
  get nativeId() { return this.control?.nativeId; }
  get controlType() { return this.control?.controlType || "text"; }
  get text() { return this.binding.getText(); }
  set text(value) {
    if (this.controlType === "checkbox") throw new TypeError("Checkbox content-control text is codec-owned; set checked instead.");
    if (this.controlType === "dropdown") throw new TypeError("Drop-down content-control text is codec-owned; set selectedValue instead.");
    if (this.controlType === "comboBox") throw new TypeError("Combo-box content-control text is codec-owned; set value instead.");
    if (this.controlType === "date") throw new TypeError("Date content-control text is codec-owned; set dateValue instead.");
    if (this.controlType !== "text") throw new TypeError(`Unsupported ${this.controlType} content-control text mutation.`);
    this.binding.setText(value);
  }
  get checked() { return this.controlType === "checkbox" ? this.control?.checked === true : undefined; }
  set checked(value) {
    if (this.controlType !== "checkbox") throw new TypeError("Only checkbox content controls have checked state.");
    if (typeof value !== "boolean") throw new TypeError("Document checkbox content control checked state must be boolean.");
    this.control.checked = value;
    this.binding.setText(documentCheckboxGlyph(value));
  }
  get choices() { return this.controlType === "dropdown" || this.controlType === "comboBox" ? this.control.choices.map((choice) => ({ ...choice })) : undefined; }
  get selectedValue() { return this.controlType === "dropdown" ? this.control?.selectedValue : undefined; }
  set selectedValue(value) {
    if (this.controlType !== "dropdown") throw new TypeError("Only drop-down content controls have selectedValue state.");
    if (typeof value !== "string") throw new TypeError("Document drop-down content-control selectedValue must be a string.");
    const choice = documentContentControlChoice(this.control, value);
    if (!choice) throw new TypeError(`Document drop-down content-control selectedValue ${value} does not match a choice value.`);
    this.control.selectedValue = value;
    this.binding.setText(choice.displayText);
  }
  get value() { return this.controlType === "comboBox" ? this.control?.value : undefined; }
  set value(value) {
    if (this.controlType !== "comboBox") throw new TypeError("Only combo-box content controls have editable value state.");
    const next = normalizeDocumentComboBoxValue(value);
    this.control.value = next;
    this.binding.setText(documentComboBoxVisibleText(this.control, next));
  }
  get dateValue() { return this.controlType === "date" ? this.control?.dateValue : undefined; }
  set dateValue(value) {
    if (this.controlType !== "date") throw new TypeError("Only date content controls have dateValue state.");
    const next = normalizeDocumentDateValue(value);
    this.control.dateValue = next;
    this.binding.setText(next);
  }
  inspectRecord() {
    return {
      kind: this.kind,
      id: this.id,
      targetId: this.targetId,
      placement: this.placement,
      runIndex: this.runIndex,
      row: this.row,
      column: this.column,
      tag: this.tag,
      alias: this.alias,
      nativeId: this.nativeId,
      controlType: this.controlType,
      ...(this.controlType === "checkbox"
        ? { checked: this.checked, visibleText: this.text }
        : this.controlType === "dropdown"
          ? { choices: this.choices, selectedValue: this.selectedValue, visibleText: this.text }
          : this.controlType === "comboBox"
            ? { choices: this.choices, value: this.value, visibleText: this.text }
            : this.controlType === "date"
              ? { dateValue: this.dateValue, visibleText: this.text }
              : { text: this.text, textChars: this.text.length }),
    };
  }
}

export function documentTableCellContentControl(table, row, column) {
  const binding = tableCellBinding(table, row, column);
  return binding.control ? new DocumentContentControlHandle(binding) : undefined;
}

export function documentContentControls(document) {
  return document.blocks.flatMap((block) => {
    if (block.kind === "paragraph") {
      return [
        ...(block.blockContentControl ? [new DocumentContentControlHandle(paragraphBinding(block))] : []),
        ...block.runs.flatMap((run, runIndex) => run.contentControl ? [new DocumentContentControlHandle(paragraphBinding(block, runIndex))] : []),
      ];
    }
    if (block.kind === "table") {
      return (block.cells || []).flatMap((cell) => cell.contentControl
        ? [new DocumentContentControlHandle(tableCellBinding(block, cell.row, cell.column))]
        : []);
    }
    return [];
  });
}
