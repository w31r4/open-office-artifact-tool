const MIN_NATIVE_MASTER_ID = 2_147_483_648;
const MAX_NATIVE_ID = 4_294_967_295;

function uniqueItems(items, label) {
  const ids = new Set();
  for (const item of items) {
    const id = String(item?.id || "").trim();
    if (!id) throw new TypeError(`Presentation ${label} ID must be a non-empty string.`);
    if (ids.has(id)) throw new Error(`Duplicate presentation ${label} ID ${id}.`);
    ids.add(id);
  }
  return ids;
}

function preferredNativeId(id, prefix) {
  const value = Number(new RegExp(`^${prefix}-(\\d+)$`).exec(String(id || ""))?.[1]);
  return Number.isInteger(value) && value >= MIN_NATIVE_MASTER_ID && value <= MAX_NATIVE_ID ? value : undefined;
}

function allocateNativeId(preferred, used, next) {
  if (preferred !== undefined && !used.has(preferred)) {
    used.add(preferred);
    return preferred;
  }
  while (used.has(next.value) && next.value <= MAX_NATIVE_ID) next.value += 1;
  if (next.value > MAX_NATIVE_ID) throw new RangeError("Presentation native master/layout IDs are exhausted.");
  const value = next.value++;
  used.add(value);
  return value;
}

export function planPresentationMasterGraph(masters = [], layouts = []) {
  if (!Array.isArray(masters) || !Array.isArray(layouts)) throw new TypeError("Presentation masters and layouts must be arrays.");
  if (masters.length > 64) throw new RangeError("Presentation masters exceed 64 entries.");
  if (layouts.length > 1024) throw new RangeError("Presentation layouts exceed 1024 entries.");
  const masterIds = uniqueItems(masters, "master");
  uniqueItems(layouts, "layout");
  for (const layout of layouts) {
    if (!masterIds.has(String(layout.masterId || ""))) throw new Error(`Presentation layout ${layout.id} references missing master ${layout.masterId}.`);
  }

  const usedNativeIds = new Set();
  const nextMasterId = { value: MIN_NATIVE_MASTER_ID };
  const nextLayoutId = { value: MIN_NATIVE_MASTER_ID + masters.length };
  const masterParts = [];
  const layoutParts = [];
  for (const master of masters) {
    const ownedLayouts = layouts.filter((layout) => layout.masterId === master.id);
    if (!ownedLayouts.length) continue;
    const masterPart = {
      master,
      masterPartId: masterParts.length + 1,
      nativeMasterId: allocateNativeId(preferredNativeId(master.id, "pptx-master"), usedNativeIds, nextMasterId),
      layoutParts: [],
    };
    for (const layout of ownedLayouts) {
      const part = {
        layout,
        layoutPartId: layoutParts.length + 1,
        masterPartId: masterPart.masterPartId,
        masterRelId: `rId${masterPart.layoutParts.length + 1}`,
        nativeLayoutId: allocateNativeId(preferredNativeId(layout.id, "pptx-layout"), usedNativeIds, nextLayoutId),
      };
      masterPart.layoutParts.push(part);
      layoutParts.push(part);
    }
    masterParts.push(masterPart);
  }
  return { masterParts, layoutParts };
}
