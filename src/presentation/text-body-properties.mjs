const EMU_PER_PIXEL = 9525;
const MAX_COORDINATE_EMU = 2_147_483_647;
const BODY_PROPERTY_KEYS = new Set(["insets", "anchor", "wrap", "autoFit", "rotation", "verticalText", "verticalOverflow", "horizontalOverflow", "columns", "upright"]);
const INSET_KEYS = new Set(["left", "top", "right", "bottom"]);
const COLUMN_KEYS = new Set(["count", "spacing", "rightToLeft"]);
const ANCHORS = new Set(["top", "center", "bottom"]);
const WRAPS = new Set(["square", "none"]);
const AUTO_FIT_MODES = new Set(["none", "shrinkText", "resizeShape"]);
const VERTICAL_TEXT_MODES = new Set(["horizontal", "vertical", "vertical270"]);
const VERTICAL_OVERFLOW_MODES = new Set(["overflow", "ellipsis", "clip"]);
const HORIZONTAL_OVERFLOW_MODES = new Set(["overflow", "clip"]);
const MAX_ROTATION_DEGREES = 360;

export const DEFAULT_PRESENTATION_TEXT_BODY_PROPERTIES = Object.freeze({
  insets: Object.freeze({ left: 0, top: 0, right: 0, bottom: 0 }),
  anchor: "top",
  wrap: "square",
});

export function normalizePresentationTextBodyProperties(value, { defaults = false } = {}) {
  if (value == null) return defaults ? cloneDefaults() : {};
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation text body properties must be an object.");
  const unknown = Object.keys(value).filter((key) => !BODY_PROPERTY_KEYS.has(key));
  if (unknown.length) throw new TypeError(`Unsupported Presentation text body properties: ${unknown.join(", ")}.`);
  const result = defaults ? cloneDefaults() : {};
  if (value.insets != null) {
    if (typeof value.insets !== "object" || Array.isArray(value.insets)) throw new TypeError("Presentation text body insets must be an object.");
    const unknownInsets = Object.keys(value.insets).filter((key) => !INSET_KEYS.has(key));
    if (unknownInsets.length) throw new TypeError(`Unsupported Presentation text body insets: ${unknownInsets.join(", ")}.`);
    const insets = defaults ? { ...result.insets } : {};
    for (const key of INSET_KEYS) {
      if (value.insets[key] == null) continue;
      const pixels = Number(value.insets[key]);
      const emu = Math.round(pixels * EMU_PER_PIXEL);
      if (!Number.isFinite(pixels) || pixels < 0 || emu > MAX_COORDINATE_EMU) throw new RangeError(`Presentation text body ${key} inset is outside the supported DrawingML coordinate range.`);
      insets[key] = pixels;
    }
    if (Object.keys(insets).length) result.insets = insets;
  }
  if (value.anchor != null) {
    if (!ANCHORS.has(value.anchor)) throw new RangeError(`Unsupported Presentation text body anchor ${value.anchor}.`);
    result.anchor = value.anchor;
  }
  if (value.wrap != null) {
    if (!WRAPS.has(value.wrap)) throw new RangeError(`Unsupported Presentation text body wrap mode ${value.wrap}.`);
    result.wrap = value.wrap;
  }
  if (value.autoFit != null) {
    if (!AUTO_FIT_MODES.has(value.autoFit)) throw new RangeError(`Unsupported Presentation text body AutoFit mode ${value.autoFit}.`);
    result.autoFit = value.autoFit;
  }
  if (value.rotation != null) {
    const rotation = Number(value.rotation);
    if (!Number.isFinite(rotation) || rotation < -MAX_ROTATION_DEGREES || rotation > MAX_ROTATION_DEGREES) throw new RangeError("Presentation text body rotation must be between -360 and 360 degrees.");
    result.rotation = rotation;
  }
  if (value.verticalText != null) {
    if (!VERTICAL_TEXT_MODES.has(value.verticalText)) throw new RangeError(`Unsupported Presentation vertical text mode ${value.verticalText}.`);
    result.verticalText = value.verticalText;
  }
  if (value.verticalOverflow != null) {
    if (!VERTICAL_OVERFLOW_MODES.has(value.verticalOverflow)) throw new RangeError(`Unsupported Presentation vertical overflow mode ${value.verticalOverflow}.`);
    result.verticalOverflow = value.verticalOverflow;
  }
  if (value.horizontalOverflow != null) {
    if (!HORIZONTAL_OVERFLOW_MODES.has(value.horizontalOverflow)) throw new RangeError(`Unsupported Presentation horizontal overflow mode ${value.horizontalOverflow}.`);
    result.horizontalOverflow = value.horizontalOverflow;
  }
  if (value.columns != null) {
    if (typeof value.columns !== "object" || Array.isArray(value.columns)) throw new TypeError("Presentation text body columns must be an object.");
    const unknownColumns = Object.keys(value.columns).filter((key) => !COLUMN_KEYS.has(key));
    if (unknownColumns.length) throw new TypeError(`Unsupported Presentation text body column properties: ${unknownColumns.join(", ")}.`);
    const columns = {};
    if (value.columns.count != null) {
      const count = Number(value.columns.count);
      if (!Number.isInteger(count) || count < 1 || count > 16) throw new RangeError("Presentation text body column count must be an integer from 1 through 16.");
      columns.count = count;
    }
    if (value.columns.spacing != null) {
      const spacing = Number(value.columns.spacing);
      const emu = Math.round(spacing * EMU_PER_PIXEL);
      if (!Number.isFinite(spacing) || spacing < 0 || emu > MAX_COORDINATE_EMU) throw new RangeError("Presentation text body column spacing is outside the supported DrawingML coordinate range.");
      columns.spacing = spacing;
    }
    if (value.columns.rightToLeft != null) {
      if (typeof value.columns.rightToLeft !== "boolean") throw new TypeError("Presentation text body rightToLeft columns must be boolean.");
      columns.rightToLeft = value.columns.rightToLeft;
    }
    if (Object.keys(columns).length) result.columns = columns;
  }
  if (value.upright != null) {
    if (typeof value.upright !== "boolean") throw new TypeError("Presentation text body upright must be boolean.");
    result.upright = value.upright;
  }
  return result;
}


function cloneDefaults() {
  return { ...DEFAULT_PRESENTATION_TEXT_BODY_PROPERTIES, insets: { ...DEFAULT_PRESENTATION_TEXT_BODY_PROPERTIES.insets } };
}
