const EMU_PER_PIXEL = 9525;
const MAX_COORDINATE_EMU = 2_147_483_647;
const BODY_PROPERTY_KEYS = new Set(["insets", "anchor", "wrap", "autoFit"]);
const INSET_KEYS = new Set(["left", "top", "right", "bottom"]);
const ANCHORS = new Set(["top", "center", "bottom"]);
const WRAPS = new Set(["square", "none"]);
const AUTO_FIT_MODES = new Set(["none", "shrinkText", "resizeShape"]);

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
  return result;
}

export function presentationTextBodyPropertiesXml(value, options = {}) {
  const properties = normalizePresentationTextBodyProperties(value, options);
  const attributes = [];
  if (properties.wrap != null) attributes.push(`wrap="${properties.wrap}"`);
  for (const [key, attribute] of [["left", "lIns"], ["top", "tIns"], ["right", "rIns"], ["bottom", "bIns"]]) {
    if (properties.insets?.[key] != null) attributes.push(`${attribute}="${Math.round(properties.insets[key] * EMU_PER_PIXEL)}"`);
  }
  if (properties.anchor != null) attributes.push(`anchor="${properties.anchor === "center" ? "ctr" : properties.anchor === "bottom" ? "b" : "t"}"`);
  const autoFit = properties.autoFit === "none" ? "<a:noAutofit/>" : properties.autoFit === "shrinkText" ? "<a:normAutofit/>" : properties.autoFit === "resizeShape" ? "<a:spAutoFit/>" : "";
  return autoFit ? `<a:bodyPr${attributes.length ? ` ${attributes.join(" ")}` : ""}>${autoFit}</a:bodyPr>` : `<a:bodyPr${attributes.length ? ` ${attributes.join(" ")}` : ""}/>`;
}

export function parsePresentationTextBodyPropertiesXml(xml) {
  const match = /<(?:[A-Za-z_][\w.-]*:)?bodyPr\b([^>]*?)(?:\/>|>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?bodyPr>)/.exec(String(xml || ""));
  if (!match) return {};
  const attributes = parseAttributes(match[1]);
  const properties = {};
  const insets = {};
  for (const [key, attribute] of [["left", "lIns"], ["top", "tIns"], ["right", "rIns"], ["bottom", "bIns"]]) {
    const emu = Number(attributes[attribute]);
    if (Number.isInteger(emu) && emu >= 0 && emu <= MAX_COORDINATE_EMU) insets[key] = emu / EMU_PER_PIXEL;
  }
  if (Object.keys(insets).length) properties.insets = insets;
  if (attributes.anchor === "t") properties.anchor = "top";
  else if (attributes.anchor === "ctr") properties.anchor = "center";
  else if (attributes.anchor === "b") properties.anchor = "bottom";
  if (WRAPS.has(attributes.wrap)) properties.wrap = attributes.wrap;
  const children = match[2] || "";
  if (/<(?:[A-Za-z_][\w.-]*:)?noAutofit\b[^>]*\/>/.test(children)) properties.autoFit = "none";
  else if (/<(?:[A-Za-z_][\w.-]*:)?normAutofit\b\s*\/>/.test(children)) properties.autoFit = "shrinkText";
  else if (/<(?:[A-Za-z_][\w.-]*:)?spAutoFit\b[^>]*\/>/.test(children)) properties.autoFit = "resizeShape";
  return properties;
}

function cloneDefaults() {
  return { ...DEFAULT_PRESENTATION_TEXT_BODY_PROPERTIES, insets: { ...DEFAULT_PRESENTATION_TEXT_BODY_PROPERTIES.insets } };
}

function parseAttributes(source) {
  return Object.fromEntries([...String(source || "").matchAll(/([A-Za-z_][\w.-]*)\s*=\s*"([^"]*)"/g)].map((match) => [match[1], match[2]]));
}
