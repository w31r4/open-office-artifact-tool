import { attrEscape } from "../ooxml/source-reference-xml.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { normalizePresentationRunLink } from "./ooxml-hyperlinks.mjs";

const EMU_PER_PIXEL = 9525;
const MAX_PARAGRAPHS = 4096;
const MAX_RUNS = 16384;
const MAX_TAB_POSITION = 2147483647 / EMU_PER_PIXEL;
const FIELD_ID_PATTERN = /^\{?([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})\}?$/i;
const TAB_ALIGNMENTS = new Set(["left", "center", "right", "decimal"]);
let nextPresentationFieldId = 1;
const AUTO_NUMBER_TYPES = new Set([
  "alphaLcParenBoth", "alphaLcParenR", "alphaLcPeriod", "alphaUcParenBoth", "alphaUcParenR", "alphaUcPeriod",
  "arabic1Minus", "arabic2Minus", "arabicDbPeriod", "arabicDbPlain", "arabicParenBoth", "arabicParenR", "arabicPeriod", "arabicPlain",
  "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain", "ea1ChsPeriod", "ea1ChsPlain", "ea1ChtPeriod", "ea1ChtPlain",
  "ea1JpnChsDbPeriod", "ea1JpnKorPeriod", "ea1JpnKorPlain", "hebrew2Minus", "hindiAlpha1Period", "hindiAlphaPeriod",
  "hindiNumParenR", "hindiNumPeriod", "romanLcParenBoth", "romanLcParenR", "romanLcPeriod", "romanUcParenBoth", "romanUcParenR", "romanUcPeriod",
  "thaiAlphaParenBoth", "thaiAlphaParenR", "thaiAlphaPeriod", "thaiNumParenBoth", "thaiNumParenR", "thaiNumPeriod",
]);

export function isPresentationAutoNumberType(value) {
  return AUTO_NUMBER_TYPES.has(String(value));
}

function xmlEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

const PRESENTATION_SCHEME_COLORS = new Set(["dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2", "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink"]);

function normalizePresentationColor(value, label) {
  const raw = String(value ?? "").trim();
  if (!raw) throw new TypeError(`${label} must be a non-empty color string.`);
  if (PRESENTATION_SCHEME_COLORS.has(raw)) return raw;
  const resolved = String(resolveColorToken(raw, raw) || "").trim();
  const short = /^#?([0-9a-f]{3})$/i.exec(resolved)?.[1];
  if (short) return `#${[...short].map((character) => character.repeat(2)).join("").toLowerCase()}`;
  const full = /^#?([0-9a-f]{6})$/i.exec(resolved)?.[1];
  if (!full) throw new TypeError(`${label} must be a DrawingML scheme color, six-digit RGB color, or supported color token.`);
  return `#${full.toLowerCase()}`;
}

function finiteNumber(value, fallback) {
  const number = Number(value);
  return Number.isFinite(number) ? number : fallback;
}

function normalizeLevel(value = 0) {
  const level = Number(value);
  if (!Number.isInteger(level) || level < 0 || level > 8) throw new RangeError("Presentation paragraph level must be an integer from 0 through 8.");
  return level;
}

function normalizeRunStyle(style = {}) {
  if (!style || typeof style !== "object" || Array.isArray(style)) throw new TypeError("Presentation run style must be an object.");
  const rawFontSize = style.fontSize == null ? undefined : String(style.fontSize).trim();
  const fontSize = rawFontSize == null ? undefined : finiteNumber(rawFontSize.replace(/(?:px|pt)$/i, "")) * (/pt$/i.test(rawFontSize) ? 4 / 3 : 1);
  if (fontSize != null && !(fontSize > 0 && fontSize <= 1024)) throw new RangeError("Presentation run fontSize must be between 0 and 1024 pixels.");
  return {
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(style.underline == null ? {} : { underline: String(style.underline) }),
    ...(fontSize == null ? {} : { fontSize }),
    ...(style.fontFamily || style.typeface ? { fontFamily: String(style.fontFamily || style.typeface) } : {}),
    ...(style.color || style.fill ? { color: normalizePresentationColor(style.color || style.fill, "Presentation run color") } : {}),
  };
}

function generatedFieldId() {
  const suffix = (nextPresentationFieldId++).toString(16).padStart(12, "0");
  return `{00000000-0000-4000-8000-${suffix}}`;
}

function normalizeFieldId(value) {
  const raw = value == null ? generatedFieldId() : String(value).trim();
  const match = FIELD_ID_PATTERN.exec(raw);
  if (!match) throw new TypeError("Presentation field id must be a UUID, optionally wrapped in braces.");
  return `{${match[1].toUpperCase()}}`;
}

function normalizeTextField(value, run = {}) {
  const input = typeof value === "string" ? { type: value } : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Presentation field must be a field type string or object.");
  const type = String(input.type ?? input.fieldType ?? "").trim();
  if (!type || type.length > 255 || /[\u0000-\u001f\u007f]/.test(type)) throw new TypeError("Presentation field type must contain 1 through 255 printable characters.");
  return {
    id: normalizeFieldId(input.id ?? input.fieldId),
    type,
    text: String(input.text ?? input.value ?? run.text ?? run.run ?? run.value ?? ""),
  };
}

function normalizeTabStops(value) {
  if (value == null) return [];
  if (!Array.isArray(value)) throw new TypeError("Presentation paragraph tabStops must be an array.");
  let previous = -1;
  return value.map((item, index) => {
    const input = typeof item === "number" ? { position: item } : item;
    if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError(`Presentation tab stop ${index + 1} must be a number or object.`);
    const position = Number(input.position ?? input.pos ?? input.offset);
    if (!Number.isFinite(position) || position < 0 || position > MAX_TAB_POSITION) throw new RangeError(`Presentation tab stop ${index + 1} position must fit the DrawingML signed 32-bit EMU range.`);
    if (position <= previous) throw new RangeError("Presentation tab stops must use strictly increasing positions.");
    previous = position;
    const alignment = String(input.alignment ?? input.align ?? "left");
    if (!TAB_ALIGNMENTS.has(alignment)) throw new RangeError(`Presentation tab stop ${index + 1} alignment must be left, center, right, or decimal.`);
    return { position, alignment };
  });
}

function normalizeRun(value, options = {}) {
  if (value == null) return { text: "", style: {} };
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return { text: String(value), style: {} };
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation paragraph runs must be strings, numbers, booleans, or run objects.");
  const isBreak = value.break === true || value.lineBreak === true || value.kind === "break";
  const hasField = value.field != null || value.kind === "field";
  if (isBreak && hasField) throw new TypeError("Presentation inline cannot be both a line break and a field.");
  const link = normalizePresentationRunLink(value.link ?? value.hyperlink, { allowTargetPart: options.allowTargetPart === true });
  const common = { style: normalizeRunStyle(value.textStyle || value.style || {}), ...(link ? { link } : {}) };
  if (isBreak) {
    if (value.run != null || value.text != null || value.value != null) throw new TypeError("Presentation line break cannot also carry text.");
    return { break: true, ...common };
  }
  if (hasField) return { field: normalizeTextField(value.field ?? { type: value.fieldType, id: value.fieldId }, value), ...common };
  return { text: String(value.run ?? value.text ?? value.value ?? ""), ...common };
}

function paragraphInput(value) {
  if (value && typeof value === "object" && !Array.isArray(value)) return value;
  return { runs: Array.isArray(value) ? value : [value] };
}

function normalizeAutoNumber(value) {
  if (!value) return undefined;
  const input = typeof value === "string" ? { type: value } : value;
  const type = String(input.type || input.scheme || "arabicPeriod");
  if (!AUTO_NUMBER_TYPES.has(type)) throw new RangeError(`Unsupported Presentation auto-number type ${type}.`);
  const rawStart = input.startAt ?? input.start;
  if (rawStart == null) return { type };
  const startAt = Number(rawStart);
  if (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767) throw new RangeError("Presentation auto-number startAt must be an integer from 1 through 32767.");
  return { type, startAt };
}

function normalizeBulletImage(value) {
  if (value == null || value === false) return undefined;
  const input = typeof value === "string" ? { src: value } : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Presentation bulletImage must be a data URL, URI, or object.");
  const src = input.src == null ? undefined : String(input.src);
  const dataUrl = input.dataUrl ?? (src?.startsWith("data:") ? src : undefined);
  const uri = input.uri ?? input.url ?? (src && !src.startsWith("data:") ? src : undefined);
  const relationshipId = input.relationshipId ?? input.relId;
  const sources = [dataUrl, uri, relationshipId].filter((item) => item != null && String(item).length > 0);
  if (sources.length !== 1) throw new Error("Presentation bulletImage requires exactly one of dataUrl, uri, or relationshipId.");
  if (dataUrl != null && !/^data:image\/(?:png|jpe?g|gif|svg\+xml);base64,[A-Za-z0-9+/=\s]+$/i.test(String(dataUrl))) throw new TypeError("Presentation bulletImage dataUrl must be a base64 PNG, JPEG, GIF, or SVG image.");
  if (uri != null && !String(uri).trim()) throw new TypeError("Presentation bulletImage uri must not be empty.");
  if (relationshipId != null && !String(relationshipId).trim()) throw new TypeError("Presentation bulletImage relationshipId must not be empty.");
  const relationshipMode = input.relationshipMode || input.mode || (uri != null ? "link" : "embed");
  if (!new Set(["embed", "link"]).has(relationshipMode)) throw new TypeError("Presentation bulletImage relationshipMode must be embed or link.");
  return {
    ...(dataUrl == null ? {} : { dataUrl: String(dataUrl) }),
    ...(uri == null ? {} : { uri: String(uri) }),
    ...(relationshipId == null ? {} : { relationshipId: String(relationshipId) }),
    relationshipMode,
    ...(input.alt == null ? {} : { alt: String(input.alt) }),
  };
}

function normalizeParagraph(value, defaults = {}, options = {}) {
  const input = paragraphInput(value);
  const runsInput = input.runs ?? input.children ?? (input.text == null ? [] : [input.text]);
  const runs = (Array.isArray(runsInput) ? runsInput : [runsInput]).map((run) => normalizeRun(run, options));
  const level = normalizeLevel(input.level ?? input.depth ?? defaults.level ?? 0);
  const alignment = input.alignment || input.paragraphStyle?.alignment || defaults.alignment;
  if (alignment && !new Set(["left", "center", "right", "justify"]).has(alignment)) throw new RangeError("Presentation paragraph alignment must be left, center, right, or justify.");
  const bulletCharacter = input.bulletCharacter ?? input.bullet?.character ?? defaults.bulletCharacter;
  if (bulletCharacter != null && [...String(bulletCharacter)].length !== 1) throw new RangeError("Presentation bulletCharacter must contain exactly one Unicode character.");
  const autoNumber = normalizeAutoNumber(input.autoNumber || input.numbering || defaults.autoNumber);
  const bulletNone = Boolean(input.bulletNone ?? (input.bullet === false ? true : defaults.bulletNone));
  const bulletImage = normalizeBulletImage(input.bulletImage ?? input.pictureBullet ?? input.bullet?.image ?? defaults.bulletImage);
  if ([bulletCharacter != null, Boolean(autoNumber), bulletNone, Boolean(bulletImage)].filter(Boolean).length > 1) throw new Error("Presentation paragraph can use exactly one of bulletCharacter, autoNumber, bulletImage, or bulletNone.");
  const bulletFont = input.bulletFont ?? input.bullet?.fontFamily ?? defaults.bulletFont;
  const bulletColorInput = input.bulletColor ?? input.bullet?.color ?? defaults.bulletColor;
  const bulletColor = bulletColorInput == null ? undefined : normalizePresentationColor(bulletColorInput, "Presentation bulletColor");
  const bulletFontFollowText = Boolean(input.bulletFontFollowText ?? input.bullet?.fontFollowText ?? defaults.bulletFontFollowText);
  const bulletColorFollowText = Boolean(input.bulletColorFollowText ?? input.bullet?.colorFollowText ?? defaults.bulletColorFollowText);
  const bulletSizeFollowText = Boolean(input.bulletSizeFollowText ?? input.bullet?.sizeFollowText ?? defaults.bulletSizeFollowText);
  if (bulletFont != null && !String(bulletFont).trim()) throw new TypeError("Presentation bulletFont must not be empty.");
  if (bulletFont != null && String(bulletFont).length > 255) throw new RangeError("Presentation bulletFont exceeds 255 characters.");
  if (bulletFont != null && bulletFontFollowText) throw new Error("Presentation paragraph cannot combine bulletFont with bulletFontFollowText.");
  if (bulletColor != null && bulletColorFollowText) throw new Error("Presentation paragraph cannot combine bulletColor with bulletColorFollowText.");
  const numberFields = ["marginLeft", "indent", "spaceBefore", "spaceAfter", "spaceBeforePercent", "spaceAfterPercent", "lineSpacing", "bulletSize", "bulletSizePercent"];
  const numeric = Object.fromEntries(numberFields.flatMap((field) => {
    const bulletAlias = field === "bulletSize" ? input.bullet?.size : field === "bulletSizePercent" ? input.bullet?.sizePercent : undefined;
    const raw = input[field] ?? bulletAlias ?? input.paragraphStyle?.[field] ?? defaults[field];
    if (raw == null) return [];
    const number = finiteNumber(raw);
    if (!Number.isFinite(number)) throw new TypeError(`Presentation paragraph ${field} must be finite.`);
    if (["marginLeft", "spaceBefore", "spaceAfter", "spaceBeforePercent", "spaceAfterPercent"].includes(field) && number < 0) throw new RangeError(`Presentation paragraph ${field} must be non-negative.`);
    if (field === "lineSpacing" && number <= 0) throw new RangeError("Presentation paragraph lineSpacing must be positive.");
    if (field === "bulletSize" && !(number >= 4 / 3 && number <= 1024)) throw new RangeError("Presentation paragraph bulletSize must be between 1.333 and 1024 pixels.");
    if (field === "bulletSizePercent" && !(number >= 0.25 && number <= 4)) throw new RangeError("Presentation paragraph bulletSizePercent must be between 0.25 and 4.");
    if (Math.abs(number) > 1_000_000) throw new RangeError(`Presentation paragraph ${field} exceeds the supported coordinate range.`);
    return [[field, number]];
  }));
  if (numeric.spaceBefore != null && numeric.spaceBeforePercent != null) throw new Error("Presentation paragraph must use either spaceBefore or spaceBeforePercent, not both.");
  if (numeric.spaceAfter != null && numeric.spaceAfterPercent != null) throw new Error("Presentation paragraph must use either spaceAfter or spaceAfterPercent, not both.");
  if ([numeric.bulletSize != null, numeric.bulletSizePercent != null, bulletSizeFollowText].filter(Boolean).length > 1) throw new Error("Presentation paragraph must use exactly one of bulletSize, bulletSizePercent, or bulletSizeFollowText.");
  const tabStops = normalizeTabStops(input.tabStops ?? input.tabs ?? defaults.tabStops);
  return {
    runs,
    level,
    ...(alignment ? { alignment } : {}),
    ...(bulletCharacter != null ? { bulletCharacter: String(bulletCharacter) } : {}),
    ...(autoNumber ? { autoNumber } : {}),
    ...(bulletImage ? { bulletImage } : {}),
    ...(bulletNone ? { bulletNone: true } : {}),
    ...(bulletFont != null ? { bulletFont: String(bulletFont).trim() } : {}),
    ...(bulletColor != null ? { bulletColor } : {}),
    ...(bulletFontFollowText ? { bulletFontFollowText: true } : {}),
    ...(bulletColorFollowText ? { bulletColorFollowText: true } : {}),
    ...(bulletSizeFollowText ? { bulletSizeFollowText: true } : {}),
    ...(tabStops.length ? { tabStops } : {}),
    ...numeric,
    style: normalizeRunStyle(input.style || input.textStyle || input.paragraphStyle?.textStyle || {}),
  };
}

function isStructuredParagraph(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value) && ("runs" in value || "text" in value || "bullet" in value || "bulletCharacter" in value || "bulletImage" in value || "pictureBullet" in value || "autoNumber" in value || "numbering" in value || "paragraphStyle" in value || "level" in value || "bulletFont" in value || "bulletColor" in value || "bulletSize" in value || "bulletSizePercent" in value || "bulletFontFollowText" in value || "bulletColorFollowText" in value || "bulletSizeFollowText" in value || "tabStops" in value || "tabs" in value));
}

export function normalizePresentationParagraphs(value, options = {}) {
  let inputs;
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean" || value == null) inputs = String(value ?? "").split(/\r?\n/);
  else if (isStructuredParagraph(value)) inputs = [value];
  else if (Array.isArray(value)) {
    if (value.every((item) => typeof item === "string" || typeof item === "number" || typeof item === "boolean")) inputs = value;
    else if (value.some((item) => Array.isArray(item) || isStructuredParagraph(item))) inputs = value;
    else inputs = [{ runs: value }];
  } else throw new TypeError("Presentation text must be a string, paragraph array, structured paragraph, or run array.");
  if (inputs.length > MAX_PARAGRAPHS) throw new RangeError(`Presentation text exceeds ${MAX_PARAGRAPHS} paragraphs.`);
  const paragraphs = inputs.map((input) => normalizeParagraph(input, options.defaults || {}));
  const runCount = paragraphs.reduce((count, paragraph) => count + paragraph.runs.length, 0);
  if (runCount > MAX_RUNS) throw new RangeError(`Presentation text exceeds ${MAX_RUNS} runs.`);
  return paragraphs;
}

export function normalizePresentationParagraphStyles(value = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation paragraph styles must be an object keyed by levels 0 through 8.");
  return Object.fromEntries(Object.entries(value).map(([rawLevel, style]) => {
    const level = normalizeLevel(rawLevel);
    if (!style || typeof style !== "object" || Array.isArray(style)) throw new TypeError(`Presentation paragraph style level ${level} must be an object.`);
    const { runs: _runs, ...normalized } = normalizeParagraph({ ...style, level, runs: [] });
    return [level, normalized];
  }));
}

export function presentationParagraphText(paragraph) {
  return (paragraph?.runs || []).map((run) => run.break ? "\n" : run.field?.text ?? run.text ?? "").join("");
}

export function presentationParagraphsText(paragraphs = []) {
  return paragraphs.map(presentationParagraphText).join("\n");
}

export function presentationParagraphsNeedSerialization(paragraphs = []) {
  return paragraphs.length > 1 || paragraphs.some((paragraph) => paragraph.level || paragraph.alignment || paragraph.bulletCharacter != null || paragraph.bulletImage || paragraph.autoNumber || paragraph.bulletNone || paragraph.bulletFont != null || paragraph.bulletColor != null || paragraph.bulletSize != null || paragraph.bulletSizePercent != null || paragraph.bulletFontFollowText || paragraph.bulletColorFollowText || paragraph.bulletSizeFollowText || paragraph.tabStops?.length || paragraph.marginLeft != null || paragraph.indent != null || paragraph.spaceBefore != null || paragraph.spaceAfter != null || paragraph.spaceBeforePercent != null || paragraph.spaceAfterPercent != null || paragraph.lineSpacing != null || Object.keys(paragraph.style || {}).length || paragraph.runs.length !== 1 || paragraph.runs.some((run) => run.break || run.field || run.link || Object.keys(run.style || {}).length));
}

export function inheritPresentationParagraphs(paragraphs = [], inheritedByLevel = {}) {
  return paragraphs.map((paragraph) => {
    const inherited = inheritedByLevel[paragraph.level] || {};
    const merged = { ...inherited, ...paragraph, style: { ...(inherited.style || {}), ...(paragraph.style || {}) }, runs: paragraph.runs };
    if (paragraph.bulletCharacter != null || paragraph.bulletImage || paragraph.autoNumber || paragraph.bulletNone) {
      delete merged.bulletCharacter;
      delete merged.bulletImage;
      delete merged.autoNumber;
      delete merged.bulletNone;
      if (paragraph.bulletCharacter != null) merged.bulletCharacter = paragraph.bulletCharacter;
      else if (paragraph.bulletImage) merged.bulletImage = paragraph.bulletImage;
      else if (paragraph.autoNumber) merged.autoNumber = paragraph.autoNumber;
      else merged.bulletNone = true;
    }
    for (const [fixed, follow] of [["bulletFont", "bulletFontFollowText"], ["bulletColor", "bulletColorFollowText"]]) {
      if (paragraph[fixed] == null && !paragraph[follow]) continue;
      delete merged[fixed];
      delete merged[follow];
      if (paragraph[fixed] != null) merged[fixed] = paragraph[fixed];
      else merged[follow] = true;
    }
    if (paragraph.bulletSize != null || paragraph.bulletSizePercent != null || paragraph.bulletSizeFollowText) {
      delete merged.bulletSize;
      delete merged.bulletSizePercent;
      delete merged.bulletSizeFollowText;
      if (paragraph.bulletSize != null) merged.bulletSize = paragraph.bulletSize;
      else if (paragraph.bulletSizePercent != null) merged.bulletSizePercent = paragraph.bulletSizePercent;
      else merged.bulletSizeFollowText = true;
    }
    for (const [points, percent] of [["spaceBefore", "spaceBeforePercent"], ["spaceAfter", "spaceAfterPercent"]]) {
      if (paragraph[points] == null && paragraph[percent] == null) continue;
      delete merged[points];
      delete merged[percent];
      if (paragraph[points] != null) merged[points] = paragraph[points];
      else merged[percent] = paragraph[percent];
    }
    return normalizeParagraph(merged);
  });
}

export function replacePresentationParagraphText(paragraphs, search, replacement) {
  const needle = String(search);
  if (!needle) return paragraphs;
  for (const paragraph of paragraphs) for (const run of paragraph.runs) {
    const text = run.field?.text ?? run.text;
    if (text == null || !text.includes(needle)) continue;
    if (run.field) run.field.text = text.replace(search, replacement);
    else run.text = text.replace(search, replacement);
    return paragraphs;
  }
  return normalizePresentationParagraphs(presentationParagraphsText(paragraphs).replace(search, replacement));
}

function autoNumberLabel(autoNumber, index) {
  const number = (autoNumber.startAt || 1) + index;
  if (/roman/i.test(autoNumber.type)) {
    const table = [[1000, "M"], [900, "CM"], [500, "D"], [400, "CD"], [100, "C"], [90, "XC"], [50, "L"], [40, "XL"], [10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
    let value = number;
    const roman = table.map(([amount, symbol]) => { const count = Math.floor(value / amount); value %= amount; return symbol.repeat(count); }).join("");
    return /Lc/.test(autoNumber.type) ? roman.toLowerCase() : roman;
  }
  if (/alpha/i.test(autoNumber.type)) {
    const character = String.fromCharCode((number - 1) % 26 + (/alphaLc/.test(autoNumber.type) ? 97 : 65));
    return character;
  }
  return String(number);
}

export function presentationParagraphsSvg(paragraphs, frame, defaultStyle = {}, options = {}) {
  const escape = options.escape || xmlEscape;
  let y = frame.top + 8;
  const counters = new Map();
  return paragraphs.map((paragraph) => {
    const paragraphStyle = { ...defaultStyle, ...(paragraph.style || {}) };
    const fontSize = paragraphStyle.fontSize || 24;
    const spacing = paragraph.lineSpacing || paragraphStyle.lineSpacing || 1.2;
    const lineHeight = spacing > 10 ? spacing : fontSize * spacing;
    y += paragraph.spaceBefore ?? fontSize * (paragraph.spaceBeforePercent || 0);
    const marginLeft = paragraph.marginLeft ?? paragraph.level * 24;
    const textX = frame.left + 12 + marginLeft;
    let marker = paragraph.bulletCharacter;
    if (paragraph.autoNumber) {
      const key = `${paragraph.level}:${paragraph.autoNumber.type}:${paragraph.autoNumber.startAt || 1}`;
      const index = counters.get(key) || 0;
      counters.set(key, index + 1);
      const label = autoNumberLabel(paragraph.autoNumber, index);
      marker = /ParenBoth$/.test(paragraph.autoNumber.type) ? `(${label})` : /ParenR$/.test(paragraph.autoNumber.type) ? `${label})` : /Period$/.test(paragraph.autoNumber.type) ? `${label}.` : label;
    }
    const markerFontSize = paragraph.bulletSize ?? fontSize * (paragraph.bulletSizePercent ?? 1);
    const markerFontFamily = paragraph.bulletFontFollowText ? paragraphStyle.fontFamily : paragraph.bulletFont || paragraphStyle.fontFamily;
    const markerColor = paragraph.bulletColorFollowText ? paragraphStyle.color : paragraph.bulletColor || paragraphStyle.color;
    const markerX = textX + (paragraph.indent ?? -12);
    const markerXml = paragraph.bulletImage
      ? paragraph.bulletImage.dataUrl
        ? `<image x="${markerX}" y="${y + fontSize - markerFontSize}" width="${markerFontSize}" height="${markerFontSize}" preserveAspectRatio="xMidYMid meet" href="${attrEscape(paragraph.bulletImage.dataUrl)}"/>`
        : `<rect x="${markerX}" y="${y + fontSize - markerFontSize}" width="${markerFontSize}" height="${markerFontSize}" rx="2" fill="#cbd5e1" data-picture-bullet="external"/>`
      : marker ? `<text x="${markerX}" y="${y + fontSize}" font-family="${attrEscape(markerFontFamily || "Arial")}" font-size="${markerFontSize}" fill="${attrEscape(markerColor || "#0f172a")}">${escape(marker)}</text>` : "";
    let x = textX;
    let lineIndex = 0;
    const tabX = (text, style) => {
      const relative = Math.max(0, x - textX);
      const stop = (paragraph.tabStops || []).find((candidate) => candidate.position > relative);
      const fallback = { position: (Math.floor(relative / 48) + 1) * 48, alignment: "left" };
      const target = stop || fallback;
      const size = style.fontSize || fontSize;
      const width = text.length * size * 0.55;
      if (target.alignment === "center") return textX + target.position - width / 2;
      if (target.alignment === "right") return textX + target.position - width;
      if (target.alignment === "decimal") return textX + target.position - Math.max(0, text.split(/[.,]/, 1)[0].length * size * 0.55);
      return textX + target.position;
    };
    const runsXml = paragraph.runs.map((run) => {
      if (run.break) {
        lineIndex += 1;
        x = textX;
        return "";
      }
      const style = { ...paragraphStyle, ...(run.style || {}) };
      const text = run.field?.text ?? run.text ?? "";
      const hyperlink = run.link?.uri || run.link?.slideId || (run.link?.action ? `action:${run.link.action}` : run.link?.customShow ? `custom-show:${run.link.customShow}` : undefined);
      return text.split("\t").map((segment, segmentIndex) => {
        if (segmentIndex) x = tabX(segment, style);
        const width = segment.length * (style.fontSize || fontSize) * 0.55;
        const result = segment ? `<text x="${x}" y="${y + fontSize + lineIndex * lineHeight}" font-family="${attrEscape(style.fontFamily || "Arial")}" font-size="${style.fontSize || fontSize}" font-weight="${style.bold ? 700 : 400}" font-style="${style.italic ? "italic" : "normal"}"${style.underline || run.link ? ' text-decoration="underline"' : ""} fill="${attrEscape(style.color || (run.link ? "#2563eb" : paragraphStyle.color || "#0f172a"))}"${hyperlink ? ` data-hyperlink="${attrEscape(hyperlink)}"` : ""}${run.field ? ` data-field-type="${attrEscape(run.field.type)}" data-field-id="${attrEscape(run.field.id)}"` : ""}>${escape(segment)}</text>` : "";
        x += width;
        return result;
      }).join("");
    }).join("");
    y += lineHeight * (lineIndex + 1) + (paragraph.spaceAfter ?? fontSize * (paragraph.spaceAfterPercent || 0));
    return markerXml + runsXml;
  }).join("");
}
