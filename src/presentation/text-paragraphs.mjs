import { attributes, attrEscape, decodeXml } from "../ooxml/source-reference-xml.mjs";
import { presentationColorXml } from "./ooxml-masters.mjs";

const EMU_PER_PIXEL = 9525;
const HUNDREDTH_POINTS_PER_PIXEL = 75;
const MAX_PARAGRAPHS = 4096;
const MAX_RUNS = 16384;
const AUTO_NUMBER_TYPES = new Set([
  "alphaLcParenBoth", "alphaLcParenR", "alphaLcPeriod", "alphaUcParenBoth", "alphaUcParenR", "alphaUcPeriod",
  "arabic1Minus", "arabic2Minus", "arabicDbPeriod", "arabicDbPlain", "arabicParenBoth", "arabicParenR", "arabicPeriod", "arabicPlain",
  "circleNumDbPlain", "circleNumWdBlackPlain", "circleNumWdWhitePlain", "ea1ChsPeriod", "ea1ChsPlain", "ea1ChtPeriod", "ea1ChtPlain",
  "ea1JpnChsDbPeriod", "ea1JpnKorPeriod", "ea1JpnKorPlain", "hebrew2Minus", "hindiAlpha1Period", "hindiAlphaPeriod",
  "hindiNumParenR", "hindiNumPeriod", "romanLcParenBoth", "romanLcParenR", "romanLcPeriod", "romanUcParenBoth", "romanUcParenR", "romanUcPeriod",
  "thaiAlphaParenBoth", "thaiAlphaParenR", "thaiAlphaPeriod", "thaiNumParenBoth", "thaiNumParenR", "thaiNumPeriod",
]);

function xmlEscape(value) {
  return String(value ?? "").replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function localElementBlock(xml = "", localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*>[\\s\\S]*?<\\/(?:[A-Za-z_][\\w.-]*:)?${localName}>`).exec(String(xml))?.[0] || "";
}

function localElementOpening(xml = "", localName) {
  return new RegExp(`<(?:[A-Za-z_][\\w.-]*:)?${localName}\\b[^>]*\\/?>`).exec(String(xml))?.[0] || "";
}

function localElement(xml = "", localName) {
  return localElementBlock(xml, localName) || localElementOpening(xml, localName);
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
    ...(style.color || style.fill ? { color: style.color || style.fill } : {}),
  };
}

function normalizeRun(value) {
  if (value == null) return { text: "", style: {} };
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") return { text: String(value), style: {} };
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation paragraph runs must be strings, numbers, booleans, or run objects.");
  const text = String(value.run ?? value.text ?? value.value ?? "");
  if (value.link) throw new Error("Presentation structured-run links are not supported yet; use an unlinked run instead of silently losing the relationship.");
  return { text, style: normalizeRunStyle(value.textStyle || value.style || {}) };
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

function normalizeParagraph(value, defaults = {}) {
  const input = paragraphInput(value);
  const runsInput = input.runs ?? input.children ?? (input.text == null ? [] : [input.text]);
  const runs = (Array.isArray(runsInput) ? runsInput : [runsInput]).map(normalizeRun);
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
  const bulletColor = input.bulletColor ?? input.bullet?.color ?? defaults.bulletColor;
  const bulletFontFollowText = Boolean(input.bulletFontFollowText ?? input.bullet?.fontFollowText ?? defaults.bulletFontFollowText);
  const bulletColorFollowText = Boolean(input.bulletColorFollowText ?? input.bullet?.colorFollowText ?? defaults.bulletColorFollowText);
  const bulletSizeFollowText = Boolean(input.bulletSizeFollowText ?? input.bullet?.sizeFollowText ?? defaults.bulletSizeFollowText);
  if (bulletFont != null && !String(bulletFont).trim()) throw new TypeError("Presentation bulletFont must not be empty.");
  if (bulletFont != null && String(bulletFont).length > 255) throw new RangeError("Presentation bulletFont exceeds 255 characters.");
  if (bulletColor != null && (typeof bulletColor !== "string" || !bulletColor.trim())) throw new TypeError("Presentation bulletColor must be a non-empty color string.");
  if (bulletColor != null) presentationColorXml(bulletColor);
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
    ...numeric,
    style: normalizeRunStyle(input.style || input.textStyle || input.paragraphStyle?.textStyle || {}),
  };
}

function isStructuredParagraph(value) {
  return Boolean(value && typeof value === "object" && !Array.isArray(value) && ("runs" in value || "text" in value || "bullet" in value || "bulletCharacter" in value || "bulletImage" in value || "pictureBullet" in value || "autoNumber" in value || "numbering" in value || "paragraphStyle" in value || "level" in value || "bulletFont" in value || "bulletColor" in value || "bulletSize" in value || "bulletSizePercent" in value || "bulletFontFollowText" in value || "bulletColorFollowText" in value || "bulletSizeFollowText" in value));
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
  return (paragraph?.runs || []).map((run) => run.text).join("");
}

export function presentationParagraphsText(paragraphs = []) {
  return paragraphs.map(presentationParagraphText).join("\n");
}

export function presentationParagraphsNeedSerialization(paragraphs = []) {
  return paragraphs.length > 1 || paragraphs.some((paragraph) => paragraph.level || paragraph.alignment || paragraph.bulletCharacter != null || paragraph.bulletImage || paragraph.autoNumber || paragraph.bulletNone || paragraph.bulletFont != null || paragraph.bulletColor != null || paragraph.bulletSize != null || paragraph.bulletSizePercent != null || paragraph.bulletFontFollowText || paragraph.bulletColorFollowText || paragraph.bulletSizeFollowText || paragraph.marginLeft != null || paragraph.indent != null || paragraph.spaceBefore != null || paragraph.spaceAfter != null || paragraph.spaceBeforePercent != null || paragraph.spaceAfterPercent != null || paragraph.lineSpacing != null || Object.keys(paragraph.style || {}).length || paragraph.runs.length !== 1 || paragraph.runs.some((run) => Object.keys(run.style || {}).length));
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
    return normalizeParagraph(merged);
  });
}

export function replacePresentationParagraphText(paragraphs, search, replacement) {
  const needle = String(search);
  if (!needle) return paragraphs;
  for (const paragraph of paragraphs) for (const run of paragraph.runs) {
    if (!run.text.includes(needle)) continue;
    run.text = run.text.replace(search, replacement);
    return paragraphs;
  }
  return normalizePresentationParagraphs(presentationParagraphsText(paragraphs).replace(search, replacement));
}

function parseColor(xml = "") {
  const scheme = attributes(localElementOpening(xml, "schemeClr")).val;
  if (scheme) return scheme;
  const rgb = attributes(localElementOpening(xml, "srgbClr")).val;
  if (rgb) return `#${String(rgb).toLowerCase()}`;
  const system = attributes(localElementOpening(xml, "sysClr")).lastClr;
  return system ? `#${String(system).toLowerCase()}` : undefined;
}

function parseRunStyle(xml = "") {
  const opening = localElementOpening(xml, "rPr") || localElementOpening(xml, "defRPr") || localElementOpening(xml, "endParaRPr");
  const attrs = attributes(opening);
  const fontFamily = attributes(localElementOpening(xml, "latin")).typeface;
  const color = parseColor(xml);
  return {
    ...(attrs.sz && Number.isFinite(Number(attrs.sz)) ? { fontSize: Number(attrs.sz) / HUNDREDTH_POINTS_PER_PIXEL } : {}),
    ...(attrs.b != null ? { bold: ["1", "true", "on"].includes(String(attrs.b).toLowerCase()) } : {}),
    ...(attrs.i != null ? { italic: ["1", "true", "on"].includes(String(attrs.i).toLowerCase()) } : {}),
    ...(attrs.u && attrs.u !== "none" ? { underline: attrs.u } : {}),
    ...(fontFamily ? { fontFamily } : {}),
    ...(color ? { color } : {}),
  };
}

function parseSpacing(xml, localName) {
  const block = localElementBlock(xml, localName);
  const points = attributes(localElementOpening(block, "spcPts")).val;
  const percent = attributes(localElementOpening(block, "spcPct")).val;
  return points != null ? { points: Number(points) / HUNDREDTH_POINTS_PER_PIXEL } : percent != null ? { percent: Number(percent) / 100000 } : undefined;
}

function parseParagraphProperties(xml = "", inherited = {}) {
  const opening = localElementOpening(xml, "pPr") || localElementOpening(xml, "defPPr") || /<(?:[A-Za-z_][\w.-]*:)?lvl[1-9]pPr\b[^>]*\/?>/.exec(String(xml))?.[0] || "";
  const attrs = attributes(opening);
  const bullet = attributes(localElementOpening(xml, "buChar")).char;
  const autoNumberAttrs = attributes(localElementOpening(xml, "buAutoNum"));
  const pictureBulletBlipAttrs = attributes(localElementOpening(localElementBlock(xml, "buBlip"), "blip"));
  const embeddedPictureBulletId = Object.entries(pictureBulletBlipAttrs).find(([name]) => /:embed$/.test(name))?.[1];
  const linkedPictureBulletId = Object.entries(pictureBulletBlipAttrs).find(([name]) => /:link$/.test(name))?.[1];
  const bulletImage = embeddedPictureBulletId || linkedPictureBulletId ? { relationshipId: embeddedPictureBulletId || linkedPictureBulletId, relationshipMode: linkedPictureBulletId ? "link" : "embed" } : undefined;
  const hasBulletNone = Boolean(localElementOpening(xml, "buNone"));
  const bulletFont = decodeXml(attributes(localElementOpening(xml, "buFont")).typeface || "") || undefined;
  const bulletColor = parseColor(localElementBlock(xml, "buClr"));
  const bulletSizePoints = attributes(localElementOpening(xml, "buSzPts")).val;
  const bulletSizePercent = attributes(localElementOpening(xml, "buSzPct")).val;
  const bulletFontFollowText = Boolean(localElementOpening(xml, "buFontTx"));
  const bulletColorFollowText = Boolean(localElementOpening(xml, "buClrTx"));
  const bulletSizeFollowText = Boolean(localElementOpening(xml, "buSzTx"));
  const level = normalizeLevel(attrs.lvl ?? inherited.level ?? 0);
  const alignment = { l: "left", ctr: "center", r: "right", just: "justify" }[attrs.algn] || inherited.alignment;
  return {
    ...inherited,
    level,
    ...(alignment ? { alignment } : {}),
    ...(attrs.marL != null ? { marginLeft: Number(attrs.marL) / EMU_PER_PIXEL } : {}),
    ...(attrs.indent != null ? { indent: Number(attrs.indent) / EMU_PER_PIXEL } : {}),
    ...(bullet != null ? { bulletCharacter: decodeXml(bullet), bulletImage: undefined, autoNumber: undefined, bulletNone: undefined } : {}),
    ...(bulletImage ? { bulletImage, bulletCharacter: undefined, autoNumber: undefined, bulletNone: undefined } : {}),
    ...(autoNumberAttrs.type ? { autoNumber: { type: autoNumberAttrs.type, ...(autoNumberAttrs.startAt == null ? {} : { startAt: Number(autoNumberAttrs.startAt) }) }, bulletCharacter: undefined, bulletImage: undefined, bulletNone: undefined } : {}),
    ...(hasBulletNone ? { bulletNone: true, bulletCharacter: undefined, bulletImage: undefined, autoNumber: undefined } : {}),
    ...(bulletFont ? { bulletFont, bulletFontFollowText: undefined } : {}),
    ...(bulletFontFollowText ? { bulletFontFollowText: true, bulletFont: undefined } : {}),
    ...(bulletColor ? { bulletColor, bulletColorFollowText: undefined } : {}),
    ...(bulletColorFollowText ? { bulletColorFollowText: true, bulletColor: undefined } : {}),
    ...(bulletSizePoints != null ? { bulletSize: Number(bulletSizePoints) / HUNDREDTH_POINTS_PER_PIXEL, bulletSizePercent: undefined, bulletSizeFollowText: undefined } : {}),
    ...(bulletSizePercent != null ? { bulletSizePercent: Number(bulletSizePercent) / 100000, bulletSize: undefined, bulletSizeFollowText: undefined } : {}),
    ...(bulletSizeFollowText ? { bulletSizeFollowText: true, bulletSize: undefined, bulletSizePercent: undefined } : {}),
    ...(parseSpacing(xml, "spcBef")?.points != null ? { spaceBefore: parseSpacing(xml, "spcBef").points } : {}),
    ...(parseSpacing(xml, "spcBef")?.percent != null ? { spaceBeforePercent: parseSpacing(xml, "spcBef").percent } : {}),
    ...(parseSpacing(xml, "spcAft")?.points != null ? { spaceAfter: parseSpacing(xml, "spcAft").points } : {}),
    ...(parseSpacing(xml, "spcAft")?.percent != null ? { spaceAfterPercent: parseSpacing(xml, "spcAft").percent } : {}),
    ...(parseSpacing(xml, "lnSpc")?.points != null ? { lineSpacing: parseSpacing(xml, "lnSpc").points } : {}),
    ...(parseSpacing(xml, "lnSpc")?.percent != null ? { lineSpacing: parseSpacing(xml, "lnSpc").percent } : {}),
    style: { ...(inherited.style || {}), ...parseRunStyle(localElementBlock(xml, "defRPr") || xml) },
  };
}

function parseLevelStyles(xml = "") {
  const styles = {};
  for (let level = 0; level < 9; level += 1) {
    const block = localElement(xml, `lvl${level + 1}pPr`);
    if (block) styles[level] = parseParagraphProperties(block, { level });
  }
  return styles;
}

export function parsePresentationListStyleXml(xml = "") {
  const listStyle = localElementBlock(xml, "lstStyle");
  return listStyle ? parseLevelStyles(listStyle) : {};
}

export function parsePresentationMasterListStylesXml(xml = "") {
  return Object.fromEntries(Object.entries({ title: "titleStyle", body: "bodyStyle", other: "otherStyle" }).map(([kind, localName]) => {
    const block = localElementBlock(xml, localName);
    return [kind, block ? parseLevelStyles(block) : {}];
  }));
}

function directParagraphBlocks(xml = "") {
  const body = localElementBlock(xml, "txBody") || String(xml);
  return [...body.matchAll(/<(?:[A-Za-z_][\w.-]*:)?p\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?p>/g)].map((match) => match[0]);
}

export function parsePresentationParagraphsXml(xml = "", options = {}) {
  const inheritedByLevel = options.inheritedByLevel || {};
  return directParagraphBlocks(xml).map((block) => {
    const pPr = localElement(block, "pPr");
    const pPrAttrs = attributes(localElementOpening(pPr, "pPr"));
    const level = normalizeLevel(pPrAttrs.lvl || 0);
    const properties = parseParagraphProperties(pPr, inheritedByLevel[level] || { level });
    const runs = [...block.matchAll(/<(?:[A-Za-z_][\w.-]*:)?(?:r|fld)\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?(?:r|fld)>/g)].map((match) => {
      const runXml = match[0];
      const text = decodeXml([...runXml.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((item) => item[1]).join(""));
      return { text, style: parseRunStyle(runXml) };
    });
    if (!runs.length) {
      const text = decodeXml([...block.matchAll(/<(?:[A-Za-z_][\w.-]*:)?t\b[^>]*>([\s\S]*?)<\/(?:[A-Za-z_][\w.-]*:)?t>/g)].map((item) => item[1]).join(""));
      if (text) runs.push({ text, style: {} });
    }
    return normalizeParagraph({ ...properties, runs });
  });
}

function colorXml(style) {
  return style.color ? presentationColorXml(style.color) : "";
}

function runPropertiesXml(style = {}, tag = "a:rPr") {
  const attrs = `${style.fontSize ? ` sz="${Math.round(style.fontSize * HUNDREDTH_POINTS_PER_PIXEL)}"` : ""}${style.bold != null ? ` b="${style.bold ? 1 : 0}"` : ""}${style.italic != null ? ` i="${style.italic ? 1 : 0}"` : ""}${style.underline ? ` u="${attrEscape(style.underline)}"` : ""}`;
  const typeface = style.fontFamily ? attrEscape(style.fontFamily) : "";
  const scriptFonts = typeface ? `<a:latin typeface="${typeface}"/>${typeface.endsWith("-lt") ? `<a:ea typeface="${typeface.slice(0, -2)}ea"/><a:cs typeface="${typeface.slice(0, -2)}cs"/>` : ""}` : "";
  return `<${tag} lang="en-US"${attrs}>${colorXml(style)}${scriptFonts}</${tag}>`;
}

function spacingXml(localName, value, percent = false) {
  if (value == null) return "";
  return percent ? `<a:${localName}><a:spcPct val="${Math.round(value * 100000)}"/></a:${localName}>` : `<a:${localName}><a:spcPts val="${Math.round(value * HUNDREDTH_POINTS_PER_PIXEL)}"/></a:${localName}>`;
}

function bulletColorXml(value) {
  return presentationColorXml(value).replace(/^<a:solidFill>/, "").replace(/<\/a:solidFill>$/, "");
}

function paragraphPropertiesXml(paragraph, options = {}) {
  const alignment = { left: "l", center: "ctr", right: "r", justify: "just" }[paragraph.alignment];
  const attrs = `${paragraph.level ? ` lvl="${paragraph.level}"` : ""}${alignment ? ` algn="${alignment}"` : ""}${paragraph.marginLeft != null ? ` marL="${Math.round(paragraph.marginLeft * EMU_PER_PIXEL)}"` : ""}${paragraph.indent != null ? ` indent="${Math.round(paragraph.indent * EMU_PER_PIXEL)}"` : ""}`;
  let bullet = "";
  if (paragraph.bulletCharacter != null) bullet = `<a:buChar char="${attrEscape(paragraph.bulletCharacter)}"/>`;
  else if (paragraph.bulletImage) {
    const relationshipId = options.pictureBulletRelationshipId?.(paragraph.bulletImage);
    if (!relationshipId) throw new Error("Presentation picture bullet has no relationship in its owning OOXML part.");
    const relationshipAttribute = paragraph.bulletImage.relationshipMode === "link" ? "link" : "embed";
    bullet = `<a:buBlip><a:blip r:${relationshipAttribute}="${attrEscape(relationshipId)}"/></a:buBlip>`;
  } else if (paragraph.autoNumber) bullet = `<a:buAutoNum type="${attrEscape(paragraph.autoNumber.type)}"${paragraph.autoNumber.startAt == null ? "" : ` startAt="${paragraph.autoNumber.startAt}"`}/>`;
  else if (paragraph.bulletNone) bullet = "<a:buNone/>";
  const bulletColor = paragraph.bulletColorFollowText ? "<a:buClrTx/>" : paragraph.bulletColor != null ? `<a:buClr>${bulletColorXml(paragraph.bulletColor)}</a:buClr>` : "";
  const bulletSize = paragraph.bulletSizeFollowText ? "<a:buSzTx/>" : paragraph.bulletSizePercent != null ? `<a:buSzPct val="${Math.round(paragraph.bulletSizePercent * 100000)}"/>` : paragraph.bulletSize != null ? `<a:buSzPts val="${Math.round(paragraph.bulletSize * HUNDREDTH_POINTS_PER_PIXEL)}"/>` : "";
  const bulletFont = paragraph.bulletFontFollowText ? "<a:buFontTx/>" : paragraph.bulletFont != null ? `<a:buFont typeface="${attrEscape(paragraph.bulletFont)}"/>` : "";
  const spacing = `${spacingXml("lnSpc", paragraph.lineSpacing, paragraph.lineSpacing == null || paragraph.lineSpacing <= 10)}${spacingXml("spcBef", paragraph.spaceBefore)}${spacingXml("spcBef", paragraph.spaceBeforePercent, true)}${spacingXml("spcAft", paragraph.spaceAfter)}${spacingXml("spcAft", paragraph.spaceAfterPercent, true)}`;
  const defaultRun = Object.keys(paragraph.style || {}).length ? runPropertiesXml(paragraph.style, "a:defRPr") : "";
  return `<a:pPr${attrs}>${spacing}${bulletColor}${bulletSize}${bulletFont}${bullet}${defaultRun}</a:pPr>`;
}

export function presentationParagraphsXml(paragraphs = [], defaultStyle = {}, options = {}) {
  return paragraphs.map((paragraph) => {
    const runs = paragraph.runs.map((run) => {
      const style = { ...defaultStyle, ...(paragraph.style || {}), ...(run.style || {}) };
      const preserve = /^\s|\s$/.test(run.text) ? ' xml:space="preserve"' : "";
      return `<a:r>${runPropertiesXml(style)}<a:t${preserve}>${xmlEscape(run.text)}</a:t></a:r>`;
    }).join("");
    const endStyle = { ...defaultStyle, ...(paragraph.style || {}) };
    return `<a:p>${paragraphPropertiesXml(paragraph, options)}${runs}<a:endParaRPr lang="en-US"${endStyle.fontSize ? ` sz="${Math.round(endStyle.fontSize * HUNDREDTH_POINTS_PER_PIXEL)}"` : ""}/></a:p>`;
  }).join("");
}

export function presentationListStyleXml(styles = {}) {
  const levels = Object.entries(styles).sort(([left], [right]) => Number(left) - Number(right)).map(([level, style]) => paragraphPropertiesXml({ ...style, level: 0 }).replace(/^<a:pPr/, `<a:lvl${Number(level) + 1}pPr`).replace(/<\/a:pPr>$/, `</a:lvl${Number(level) + 1}pPr>`)).join("");
  return `<a:lstStyle>${levels}</a:lstStyle>`;
}

export function presentationMasterListStylesXml(stylesByKind = {}, fallbackTextStyles = {}) {
  const localNames = { title: "titleStyle", body: "bodyStyle", other: "otherStyle" };
  return Object.entries(localNames).map(([kind, localName]) => {
    const fallback = fallbackTextStyles[kind] || {};
    const levels = Array.from({ length: 9 }, (_, level) => {
      const explicit = stylesByKind[kind]?.[level] || {};
      const paragraph = {
        alignment: explicit.alignment ?? fallback.alignment,
        marginLeft: explicit.marginLeft ?? level * 48,
        indent: explicit.indent ?? (level ? -24 : 0),
        ...explicit,
        style: { fontSize: fallback.fontSize == null ? undefined : fallback.fontSize * 4 / 3, bold: fallback.bold, italic: fallback.italic, color: fallback.color, fontFamily: fallback.fontFamily, ...(explicit.style || {}) },
      };
      return paragraphPropertiesXml(paragraph).replace(/^<a:pPr/, `<a:lvl${level + 1}pPr`).replace(/<\/a:pPr>$/, `</a:lvl${level + 1}pPr>`).replace(/<a:defRPr lang="en-US" sz="([^"]+)"/, '<a:defRPr sz="$1"');
    }).join("");
    return `<p:${localName}>${levels}</p:${localName}>`;
  }).join("");
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
    const runsXml = paragraph.runs.map((run) => {
      const style = { ...paragraphStyle, ...(run.style || {}) };
      const width = run.text.length * (style.fontSize || fontSize) * 0.55;
      const result = `<text x="${x}" y="${y + fontSize}" font-family="${attrEscape(style.fontFamily || "Arial")}" font-size="${style.fontSize || fontSize}" font-weight="${style.bold ? 700 : 400}" font-style="${style.italic ? "italic" : "normal"}"${style.underline ? ' text-decoration="underline"' : ""} fill="${attrEscape(style.color || paragraphStyle.color || "#0f172a")}">${escape(run.text)}</text>`;
      x += width;
      return result;
    }).join("");
    y += lineHeight + (paragraph.spaceAfter ?? fontSize * (paragraph.spaceAfterPercent || 0));
    return markerXml + runsXml;
  }).join("");
}
