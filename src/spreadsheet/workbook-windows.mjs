function resolveWorksheet(workbook, value, label) {
  const worksheet = typeof value === "number"
    ? workbook.worksheets.getItemAt(value)
    : typeof value === "string"
      ? workbook.worksheets.getItem(value) || workbook.worksheets.items.find((sheet) => sheet.id === value)
      : value;
  if (!worksheet || worksheet.workbook !== workbook || !workbook.worksheets.items.includes(worksheet))
    throw new Error(`Workbook ${label} ${String(value)} was not found.`);
  return worksheet;
}

function orderedSelectedWorksheets(workbook, activeWorksheetId, selectedWorksheetIds, label) {
  const active = workbook.worksheets.items.find((sheet) => sheet.id === activeWorksheetId);
  if (!active) throw new Error(`Workbook ${label} active worksheet ${activeWorksheetId} no longer exists.`);
  if (active.visibility !== "visible") throw new Error(`Workbook ${label} active worksheet ${active.name} is not visible.`);
  if (!selectedWorksheetIds) return [active];
  const selectedIds = new Set(selectedWorksheetIds);
  const selected = workbook.worksheets.items.filter((sheet) => selectedIds.has(sheet.id));
  if (selected.length !== selectedIds.size) throw new Error(`Workbook ${label} selected worksheets contain an identity that no longer exists.`);
  if (selected.some((sheet) => sheet.visibility !== "visible")) throw new Error(`Workbook ${label} selected worksheets must all be visible.`);
  if (!selected.includes(active)) throw new Error(`Workbook ${label} selected worksheets must include the active worksheet.`);
  return selected;
}

class WorkbookWindow {
  constructor(collection, ordinal) {
    this.collection = collection;
    this.workbook = collection.workbook;
    this.ordinal = ordinal;
  }

  get id() {
    return `workbook-window/${this.ordinal + 1}`;
  }

  get index() {
    return this.ordinal;
  }

  _state() {
    if (this.ordinal === 0) {
      return {
        activeWorksheetId: this.workbook._activeWorksheetId,
        selectedWorksheetIds: this.workbook._selectedWorksheetIds,
      };
    }
    const state = this.workbook._additionalWorkbookWindows[this.ordinal - 1];
    if (!state) throw new Error(`Workbook window ${this.ordinal} no longer exists.`);
    return state;
  }

  _replace(state) {
    if (this.ordinal === 0) {
      this.workbook._activeWorksheetId = state.activeWorksheetId;
      this.workbook._selectedWorksheetIds = state.selectedWorksheetIds;
      return;
    }
    if (!this.workbook._additionalWorkbookWindows[this.ordinal - 1])
      throw new Error(`Workbook window ${this.ordinal} no longer exists.`);
    this.workbook._additionalWorkbookWindows[this.ordinal - 1] = state;
  }

  getActiveWorksheet() {
    const state = this._state();
    if (state.activeWorksheetId) {
      const active = this.workbook.worksheets.items.find((sheet) => sheet.id === state.activeWorksheetId);
      if (!active) throw new Error(`Workbook window ${this.ordinal} active worksheet ${state.activeWorksheetId} no longer exists.`);
      if (active.visibility !== "visible") throw new Error(`Workbook window ${this.ordinal} active worksheet ${active.name} is not visible.`);
      return active;
    }
    const active = this.workbook.worksheets.items.find((sheet) => sheet.visibility === "visible");
    if (!active) throw new Error("Workbook has no visible worksheets; at least one visible worksheet is required. Add or show a worksheet first.");
    return active;
  }

  setActiveWorksheet(value) {
    const worksheet = resolveWorksheet(this.workbook, value, `window ${this.ordinal} active worksheet`);
    if (worksheet.visibility !== "visible") throw new Error(`Workbook window ${this.ordinal} active worksheet ${worksheet.name} must be visible.`);
    this._replace({ activeWorksheetId: worksheet.id, selectedWorksheetIds: [worksheet.id] });
    return worksheet;
  }

  getSelectedWorksheets() {
    const state = this._state();
    return orderedSelectedWorksheets(this.workbook, this.getActiveWorksheet().id, state.selectedWorksheetIds, `window ${this.ordinal}`);
  }

  setSelectedWorksheets(values) {
    const requested = Array.isArray(values) ? values : [values];
    if (requested.length === 0) throw new Error(`Workbook window ${this.ordinal} selected worksheets must contain at least one visible worksheet.`);
    const resolved = requested.map((value) => resolveWorksheet(this.workbook, value, `window ${this.ordinal} selected worksheet`));
    if (resolved.some((worksheet) => worksheet.visibility !== "visible"))
      throw new Error(`Workbook window ${this.ordinal} selected worksheets must all be visible.`);
    const ids = new Set(resolved.map((worksheet) => worksheet.id));
    if (ids.size !== resolved.length) throw new Error(`Workbook window ${this.ordinal} selected worksheets cannot contain duplicates.`);
    const current = this._state();
    const activeWorksheetId = current.activeWorksheetId && ids.has(current.activeWorksheetId)
      ? current.activeWorksheetId
      : resolved[0].id;
    const selectedWorksheetIds = this.workbook.worksheets.items
      .filter((worksheet) => ids.has(worksheet.id))
      .map((worksheet) => worksheet.id);
    this._replace({ activeWorksheetId, selectedWorksheetIds });
    return this.getSelectedWorksheets();
  }

  toJSON() {
    return {
      id: this.id,
      index: this.ordinal,
      activeWorksheet: this.getActiveWorksheet().name,
      selectedWorksheets: this.getSelectedWorksheets().map((sheet) => sheet.name),
    };
  }
}

export class WorkbookWindowCollection {
  constructor(workbook) {
    this.workbook = workbook;
    workbook._additionalWorkbookWindows = [];
    this._items = [new WorkbookWindow(this, 0)];
  }

  get items() {
    return [...this._items];
  }

  get count() {
    return this._items.length;
  }

  getItemAt(index) {
    return this._items[index];
  }

  add(options = {}) {
    if (!this.workbook.worksheets.items.length) throw new Error("Workbook windows require at least one worksheet.");
    const active = resolveWorksheet(
      this.workbook,
      options.activeWorksheet ?? options.activeSheet ?? this.getItemAt(0).getActiveWorksheet(),
      "new window active worksheet",
    );
    if (active.visibility !== "visible") throw new Error(`Workbook new window active worksheet ${active.name} must be visible.`);
    const ordinal = this._items.length;
    this.workbook._additionalWorkbookWindows.push({ activeWorksheetId: active.id, selectedWorksheetIds: [active.id] });
    const window = new WorkbookWindow(this, ordinal);
    this._items.push(window);
    if (options.selectedWorksheets !== undefined) window.setSelectedWorksheets(options.selectedWorksheets);
    return window;
  }

  _clearAdditional() {
    this.workbook._additionalWorkbookWindows = [];
    this._items = [this._items[0]];
  }

  toJSON() {
    return this._items.map((window) => window.toJSON());
  }

  [Symbol.iterator]() {
    return this._items[Symbol.iterator]();
  }
}

export function createWorkbookWindowCollection(workbook) {
  return new WorkbookWindowCollection(workbook);
}

export function workbookWindowSnapshots(workbook) {
  return workbook.windows.items.map((window) => ({
    activeWorksheetId: window.getActiveWorksheet().id,
    selectedWorksheetIds: window.getSelectedWorksheets().map((sheet) => sheet.id),
  }));
}

export function worksheetWindowMemberships(workbook, worksheetId) {
  return workbook.windows.items.flatMap((window) => {
    const active = window.index === 0
      ? workbook._activeWorksheetId === worksheetId
      : window.getActiveWorksheet().id === worksheetId;
    const selected = window.index === 0
      ? workbook._selectedWorksheetIds?.includes(worksheetId) === true
      : window.getSelectedWorksheets().some((sheet) => sheet.id === worksheetId);
    return active || selected ? [{ windowIndex: window.index, active, selected }] : [];
  });
}
