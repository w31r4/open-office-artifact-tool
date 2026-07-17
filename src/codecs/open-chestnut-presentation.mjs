import { ChartElement, ImageElement, Presentation, Shape, TableElement } from "../index.mjs";
import {
  ArtifactFamily,
  SpreadsheetChartDataLabelPosition,
  SpreadsheetChartLineDashStyle,
  SpreadsheetChartMarkerSymbol,
  SpreadsheetChartType,
} from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizePresentationRunLink } from "../presentation/ooxml-hyperlinks.mjs";
import { normalizePresentationThemeConfig } from "../presentation/ooxml-theme.mjs";
import { normalizePresentationTextBodyProperties } from "../presentation/text-body-properties.mjs";
import { isPresentationAutoNumberType, normalizePresentationParagraphs, normalizePresentationParagraphStyles } from "../presentation/text-paragraphs.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { createPresentationAssetCatalog, validatePictureBulletUri } from "./open-chestnut-assets.mjs";
import { OpenChestnutCodecError } from "./open-chestnut-error.mjs";
import { materializePresentationNativeGraphs, presentationNativeGraphSnapshot } from "./open-chestnut-presentation-native.mjs";

const EMU_PER_PIXEL = 9525;
const EMU_PER_POINT = 12700;
const POINTS_PER_PIXEL = 0.75;
const MAX_FONT_SIZE_PIXELS = 1024;
const MAX_PARAGRAPH_COORDINATE_EMU = 51_206_400;
const MAX_TEXT_BODY_INSET_EMU = 2_147_483_647;
const ROTATION_UNITS_PER_DEGREE = 60_000;
const MAX_PARAGRAPH_SPACING_POINTS = 1584;
const MAX_PARAGRAPH_SPACING_MULTIPLIER = 132;
const MAX_PRESENTATION_CHART_POINTS = 1_048_576;
const PRESENTATION_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-presentation-state");
const PRESENTATION_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const DEFAULT_PRESENTATION_THEME = JSON.stringify(normalizePresentationThemeConfig({}));
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "color"]);
const TEXT_FRAME_PARAGRAPH_KEYS = new Set(["alignment", "tabStops", "marginLeft", "indent", "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent"]);
const PARAGRAPH_KEYS = new Set([
  "runs", "level", "alignment", "style", "bulletCharacter", "autoNumber", "bulletImage", "bulletNone",
  "bulletFont", "bulletFontFollowText", "bulletColor", "bulletColorFollowText",
  "bulletSize", "bulletSizePercent", "bulletSizeFollowText", "tabStops", "marginLeft", "indent",
  "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent",
]);
const PRESENTATION_CHART_TYPES_TO_WIRE = new Map([
  ["bar", SpreadsheetChartType.BAR],
  ["line", SpreadsheetChartType.LINE],
  ["pie", SpreadsheetChartType.PIE],
]);
const PRESENTATION_CHART_TYPES_FROM_WIRE = new Map([...PRESENTATION_CHART_TYPES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_LINE_STYLES_TO_WIRE = new Map([
  ["solid", SpreadsheetChartLineDashStyle.SOLID],
  ["dashed", SpreadsheetChartLineDashStyle.DASHED],
  ["dotted", SpreadsheetChartLineDashStyle.DOTTED],
  ["dash-dot", SpreadsheetChartLineDashStyle.DASH_DOT],
  ["dash-dot-dot", SpreadsheetChartLineDashStyle.DASH_DOT_DOT],
]);
const PRESENTATION_CHART_LINE_STYLES_FROM_WIRE = new Map([...PRESENTATION_CHART_LINE_STYLES_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE = new Map([
  ["bestFit", SpreadsheetChartDataLabelPosition.BEST_FIT],
  ["bottom", SpreadsheetChartDataLabelPosition.BOTTOM],
  ["center", SpreadsheetChartDataLabelPosition.CENTER],
  ["insideBase", SpreadsheetChartDataLabelPosition.INSIDE_BASE],
  ["insideEnd", SpreadsheetChartDataLabelPosition.INSIDE_END],
  ["left", SpreadsheetChartDataLabelPosition.LEFT],
  ["outsideEnd", SpreadsheetChartDataLabelPosition.OUTSIDE_END],
  ["right", SpreadsheetChartDataLabelPosition.RIGHT],
  ["top", SpreadsheetChartDataLabelPosition.TOP],
]);
const PRESENTATION_CHART_LABEL_POSITIONS_FROM_WIRE = new Map([...PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE].map(([name, value]) => [value, name]));
const PRESENTATION_CHART_MARKERS_TO_WIRE = new Map([
  ["none", SpreadsheetChartMarkerSymbol.NONE],
  ["dot", SpreadsheetChartMarkerSymbol.DOT],
  ["circle", SpreadsheetChartMarkerSymbol.CIRCLE],
  ["square", SpreadsheetChartMarkerSymbol.SQUARE],
  ["diamond", SpreadsheetChartMarkerSymbol.DIAMOND],
  ["triangle", SpreadsheetChartMarkerSymbol.TRIANGLE],
  ["x", SpreadsheetChartMarkerSymbol.X],
  ["star", SpreadsheetChartMarkerSymbol.STAR],
  ["plus", SpreadsheetChartMarkerSymbol.PLUS],
  ["dash", SpreadsheetChartMarkerSymbol.DASH],
]);
const PRESENTATION_CHART_MARKERS_FROM_WIRE = new Map([...PRESENTATION_CHART_MARKERS_TO_WIRE].map(([name, value]) => [value, name]));

function assertTrustedPresentationState(state) {
  if (!state) return;
  const sourceHash = String(state.source?.packageSha256 || "").toLowerCase();
  const snapshot = state.opaqueOpc?.sourcePackage;
  const snapshotHash = String(snapshot?.sha256 || "").toLowerCase();
  if (!sourceHash || !snapshotHash || sourceHash !== snapshotHash || !snapshot?.data?.length) {
    throw new OpenChestnutCodecError("PPTX source-bound export requires its validated source package snapshot.", [], { code: "missing_source_package" });
  }
}

function emuFromPixels(value, name) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0) throw new OpenChestnutCodecError(`${name} must be a non-negative finite number.`, [], { code: "invalid_presentation_frame" });
  return BigInt(Math.round(number * EMU_PER_PIXEL));
}

function paragraphEmuFromPixels(value, name, { allowNegative = false } = {}) {
  const number = Number(value);
  const emu = Math.round(number * EMU_PER_PIXEL);
  if (!Number.isFinite(number) || (!allowNegative && number < 0) || emu < (allowNegative ? -MAX_PARAGRAPH_COORDINATE_EMU : 0) || emu > MAX_PARAGRAPH_COORDINATE_EMU) {
    throw new OpenChestnutCodecError(`${name} is outside the supported DrawingML coordinate range.`, [], { code: "invalid_presentation_text" });
  }
  return BigInt(emu);
}

function presentationRgb(value, name) {
  const source = typeof value === "string" ? value : value?.color || value?.fill;
  const raw = resolveColorToken(source, source);
  if (raw == null || raw === "transparent" || raw === "none") return "";
  const match = /^#([0-9a-f]{6})$/i.exec(String(raw));
  if (!match) throw new OpenChestnutCodecError(`${name} must be transparent or a six-digit RGB color.`, [], { code: "unsupported_presentation_features" });
  return match[1].toUpperCase();
}

function unsupportedStyleFields(style = {}) {
  return Object.keys(style).filter((key) => !RUN_STYLE_KEYS.has(key));
}

function wireTextStyle(style = {}, shapeId) {
  const unsupported = unsupportedStyleFields(style);
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a paragraph font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = style.fontFamily == null ? undefined : String(style.fontFamily);
  if (fontFamily !== undefined && (!fontFamily.trim() || fontFamily.length > 255)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid paragraph font family.`, [], { code: "invalid_presentation_text" });
  }
  let color;
  if (style.color != null) {
    const token = String(style.color).trim();
    color = PRESENTATION_SCHEME_COLORS.has(token)
      ? { case: "colorScheme", value: token }
      : { case: "colorRgb", value: presentationRgb(style.color, `${shapeId}.text.paragraphStyle.color`) };
    if (!color.value) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent paragraph color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  return {
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(color ? { color } : {}),
  };
}

function wireDefaultRunStyle(paragraph, original, shapeId) {
  if (Object.keys(paragraph.style || {}).length) {
    return { case: "defaultRunProperties", value: wireTextStyle(paragraph.style, shapeId) };
  }
  if (new Set(["defaultRunProperties", "noDefaultRunProperties"]).has(original?.defaultRunStyle?.case)) {
    return { case: "noDefaultRunProperties", value: true };
  }
  return undefined;
}

function wireHyperlink(value, original, shapeId) {
  if (value == null) {
    if (new Set(["runHyperlink", "noHyperlink"]).has(original?.hyperlink?.case)) return { case: "noHyperlink", value: true };
    return undefined;
  }
  let link;
  try {
    link = normalizePresentationRunLink(value);
  } catch (error) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid run hyperlink: ${error.message}`, [], { code: "invalid_presentation_hyperlink" });
  }
  if (link.customShow) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a custom-show hyperlink outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  const target = link.uri
    ? { case: "uri", value: link.uri }
    : link.slideId
      ? { case: "slideId", value: link.slideId }
      : link.action
        ? { case: "action", value: link.action }
        : undefined;
  if (!target) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unsupported run hyperlink target.`, [], { code: "unsupported_presentation_features" });
  return {
    case: "runHyperlink",
    value: {
      target,
      ...(link.tooltip == null ? {} : { tooltip: link.tooltip }),
      ...(link.targetFrame == null ? {} : { targetFrame: link.targetFrame }),
      ...(link.history == null ? {} : { history: link.history }),
      ...(link.highlightClick == null ? {} : { highlightClick: link.highlightClick }),
    },
  };
}

function wireRun(run, inheritedStyle, shapeId, original) {
  const unsupported = unsupportedStyleFields(run.style);
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported run style fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const style = { ...inheritedStyle, ...(run.style || {}) };
  const fontSize = style.fontSize == null ? undefined : Number(style.fontSize);
  if (fontSize !== undefined && (!Number.isFinite(fontSize) || fontSize <= 0 || fontSize > MAX_FONT_SIZE_PIXELS)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a font size outside the supported 0-${MAX_FONT_SIZE_PIXELS} pixel range.`, [], { code: "invalid_presentation_text" });
  }
  const fontFamily = style.fontFamily == null ? undefined : String(style.fontFamily);
  if (fontFamily !== undefined && (!fontFamily.trim() || fontFamily.length > 255)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid font family.`, [], { code: "invalid_presentation_text" });
  }
  const colorRgb = style.color == null ? undefined : presentationRgb(style.color, `${shapeId}.text.color`);
  if (colorRgb === "") {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent run color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
  }
  const hyperlink = wireHyperlink(run.link, original, shapeId);
  return {
    content: run.break
      ? { case: "lineBreak", value: true }
      : run.field
        ? { case: "field", value: { id: run.field.id, type: run.field.type, text: run.field.text } }
        : { case: "text", value: String(run.text ?? "") },
    ...(style.bold == null ? {} : { bold: Boolean(style.bold) }),
    ...(style.italic == null ? {} : { italic: Boolean(style.italic) }),
    ...(fontSize === undefined ? {} : { fontSizePoints: fontSize * POINTS_PER_PIXEL }),
    ...(fontFamily === undefined ? {} : { fontFamily }),
    ...(colorRgb === undefined ? {} : { colorRgb }),
    ...(hyperlink ? { hyperlink } : {}),
  };
}

function wireBullet(paragraph, original, shapeId, assetCatalog) {
  const choices = [paragraph.bulletCharacter != null, Boolean(paragraph.autoNumber), Boolean(paragraph.bulletImage), paragraph.bulletNone === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects more than one list marker.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletCharacter != null) {
    const character = String(paragraph.bulletCharacter);
    if ([...character].length !== 1) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} bullet character must contain one Unicode scalar value.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletCharacter", value: character };
  }
  if (paragraph.autoNumber) {
    const scheme = String(paragraph.autoNumber.type || paragraph.autoNumber.scheme || "");
    if (!isPresentationAutoNumberType(scheme)) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported auto-number scheme ${scheme || "(missing)"}.`, [], { code: "invalid_presentation_text" });
    const rawStart = paragraph.autoNumber.startAt ?? paragraph.autoNumber.start;
    const startAt = rawStart == null ? undefined : Number(rawStart);
    if (startAt !== undefined && (!Number.isInteger(startAt) || startAt < 1 || startAt > 32767)) {
      throw new OpenChestnutCodecError(`Presentation shape ${shapeId} auto-number start must be from 1 through 32767.`, [], { code: "invalid_presentation_text" });
    }
    return { case: "autoNumber", value: { scheme, ...(startAt === undefined ? {} : { startAt }) } };
  }
  if (paragraph.bulletImage) {
    if (paragraph.bulletImage.relationshipId) {
      throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unresolved picture-bullet relationship ID.`, [], { code: "invalid_presentation_asset" });
    }
    const source = paragraph.bulletImage.dataUrl
      ? { case: "assetId", value: assetCatalog.addDataUrl(paragraph.bulletImage.dataUrl) }
      : { case: "uri", value: validatePictureBulletUri(paragraph.bulletImage.uri) };
    return { case: "pictureBullet", value: { source } };
  }
  if (paragraph.bulletNone === true || new Set(["noBullet", "bulletCharacter", "autoNumber", "pictureBullet"]).has(original?.bullet?.case)) {
    return { case: "noBullet", value: true };
  }
  return undefined;
}

function wireBulletFont(paragraph, original, shapeId) {
  if (paragraph.bulletFont != null && paragraph.bulletFontFollowText === true) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet font and follow-text font.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletFont != null) {
    const family = String(paragraph.bulletFont).trim();
    if (!family || family.length > 255) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid bullet font family.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletFontFamily", value: family };
  }
  if (paragraph.bulletFontFollowText === true || new Set(["bulletFontFamily", "bulletFontFollowText"]).has(original?.bulletFont?.case)) {
    return { case: "bulletFontFollowText", value: true };
  }
  return undefined;
}

function wireBulletColor(paragraph, original, shapeId) {
  if (paragraph.bulletColor != null && paragraph.bulletColorFollowText === true) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects both a bullet color and follow-text color.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletColor != null) {
    const scheme = String(paragraph.bulletColor).trim();
    if (PRESENTATION_SCHEME_COLORS.has(scheme)) return { case: "bulletColorScheme", value: scheme };
    const rgb = presentationRgb(paragraph.bulletColor, `${shapeId}.text.bulletColor`);
    if (!rgb) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a transparent bullet color outside the PPTX WebAssembly text slice.`, [], { code: "unsupported_presentation_features" });
    return { case: "bulletColorRgb", value: rgb };
  }
  if (paragraph.bulletColorFollowText === true || new Set(["bulletColorRgb", "bulletColorScheme", "bulletColorFollowText"]).has(original?.bulletColor?.case)) {
    return { case: "bulletColorFollowText", value: true };
  }
  return undefined;
}

function wireBulletSize(paragraph, original, shapeId) {
  const choices = [paragraph.bulletSize != null, paragraph.bulletSizePercent != null, paragraph.bulletSizeFollowText === true];
  if (choices.filter(Boolean).length > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph selects more than one bullet size.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.bulletSize != null) {
    const pixels = Number(paragraph.bulletSize);
    if (!Number.isFinite(pixels) || pixels < 4 / 3 || pixels > MAX_FONT_SIZE_PIXELS) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid fixed bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePoints", value: pixels * POINTS_PER_PIXEL };
  }
  if (paragraph.bulletSizePercent != null) {
    const percent = Number(paragraph.bulletSizePercent);
    if (!Number.isFinite(percent) || percent < 0.25 || percent > 4) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an invalid percentage bullet size.`, [], { code: "invalid_presentation_text" });
    return { case: "bulletSizePercent", value: percent };
  }
  if (paragraph.bulletSizeFollowText === true || new Set(["bulletSizePoints", "bulletSizePercent", "bulletSizeFollowText"]).has(original?.bulletSize?.case)) {
    return { case: "bulletSizeFollowText", value: true };
  }
  return undefined;
}

function wireTabStops(paragraph, original, shapeId) {
  if (paragraph.tabStops?.length) {
    return { tabStops: paragraph.tabStops.map((tab) => ({ positionEmu: emuFromPixels(tab.position, `${shapeId}.text.tabStops.position`), alignment: tab.alignment })) };
  }
  if (original?.tabStops?.length || original?.noTabStops === true) return { noTabStops: true };
  return {};
}

function wireParagraphLayout(paragraph, original, shapeId) {
  const leftMargin = paragraph.marginLeft != null
    ? { case: "marginLeftEmu", value: paragraphEmuFromPixels(paragraph.marginLeft, `${shapeId}.text.marginLeft`) }
    : new Set(["marginLeftEmu", "noMarginLeft"]).has(original?.leftMargin?.case)
      ? { case: "noMarginLeft", value: true }
      : undefined;
  const indentation = paragraph.indent != null
    ? { case: "indentEmu", value: paragraphEmuFromPixels(paragraph.indent, `${shapeId}.text.indent`, { allowNegative: true }) }
    : new Set(["indentEmu", "noIndent"]).has(original?.indentation?.case)
      ? { case: "noIndent", value: true }
      : undefined;
  return {
    ...(leftMargin ? { leftMargin } : {}),
    ...(indentation ? { indentation } : {}),
  };
}

function paragraphSpacingValue(value, name, { allowZero = true, maximum }) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < (allowZero ? 0 : Number.EPSILON) || number > maximum) {
    throw new OpenChestnutCodecError(`${name} is outside the supported DrawingML spacing range.`, [], { code: "invalid_presentation_text" });
  }
  return number;
}

function paragraphSpacingPointsFromPixels(value, name, { allowZero = true } = {}) {
  const pixels = paragraphSpacingValue(value, name, { allowZero, maximum: MAX_PARAGRAPH_SPACING_POINTS / POINTS_PER_PIXEL });
  return pixels * POINTS_PER_PIXEL;
}

function wireParagraphSpacing(paragraph, original, shapeId) {
  let lineSpacing;
  if (paragraph.lineSpacing != null) {
    const value = Number(paragraph.lineSpacing);
    lineSpacing = value <= 10
      ? { case: "lineSpacingMultiplier", value: paragraphSpacingValue(value, `${shapeId}.text.lineSpacing`, { allowZero: false, maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) }
      : { case: "lineSpacingPoints", value: paragraphSpacingPointsFromPixels(value, `${shapeId}.text.lineSpacing`, { allowZero: false }) };
  } else if (new Set(["lineSpacingPoints", "lineSpacingMultiplier", "noLineSpacing"]).has(original?.lineSpacing?.case)) {
    lineSpacing = { case: "noLineSpacing", value: true };
  }

  if (paragraph.spaceBefore != null && paragraph.spaceBeforePercent != null) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-before, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceBefore;
  if (paragraph.spaceBefore != null) {
    spaceBefore = { case: "spaceBeforePoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceBefore, `${shapeId}.text.spaceBefore`) };
  } else if (paragraph.spaceBeforePercent != null) {
    spaceBefore = { case: "spaceBeforeMultiplier", value: paragraphSpacingValue(paragraph.spaceBeforePercent, `${shapeId}.text.spaceBeforePercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceBeforePoints", "spaceBeforeMultiplier", "noSpaceBefore"]).has(original?.spaceBefore?.case)) {
    spaceBefore = { case: "noSpaceBefore", value: true };
  }

  if (paragraph.spaceAfter != null && paragraph.spaceAfterPercent != null) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} paragraph must use either point or percentage space-after, not both.`, [], { code: "invalid_presentation_text" });
  }
  let spaceAfter;
  if (paragraph.spaceAfter != null) {
    spaceAfter = { case: "spaceAfterPoints", value: paragraphSpacingPointsFromPixels(paragraph.spaceAfter, `${shapeId}.text.spaceAfter`) };
  } else if (paragraph.spaceAfterPercent != null) {
    spaceAfter = { case: "spaceAfterMultiplier", value: paragraphSpacingValue(paragraph.spaceAfterPercent, `${shapeId}.text.spaceAfterPercent`, { maximum: MAX_PARAGRAPH_SPACING_MULTIPLIER }) };
  } else if (new Set(["spaceAfterPoints", "spaceAfterMultiplier", "noSpaceAfter"]).has(original?.spaceAfter?.case)) {
    spaceAfter = { case: "noSpaceAfter", value: true };
  }

  return {
    ...(lineSpacing ? { lineSpacing } : {}),
    ...(spaceBefore ? { spaceBefore } : {}),
    ...(spaceAfter ? { spaceAfter } : {}),
  };
}

function wireParagraph(paragraph, textStyle, original, shapeId, assetCatalog, { forceLevel = false } = {}) {
  const unsupported = Object.keys(paragraph).filter((key) => !PARAGRAPH_KEYS.has(key));
  if (unsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph fields: ${unsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const paragraphStyleUnsupported = unsupportedStyleFields(paragraph.style);
  if (paragraphStyleUnsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph text style fields: ${paragraphStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const level = Number(paragraph.level || 0);
  if (!Number.isInteger(level) || level < 0 || level > 8) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses a paragraph level outside the supported 0-8 range.`, [], { code: "invalid_presentation_text" });
  }
  if (paragraph.alignment && !new Set(["left", "center", "right", "justify"]).has(paragraph.alignment)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses unsupported paragraph alignment ${paragraph.alignment}.`, [], { code: "invalid_presentation_text" });
  }
  const originalLevel = original?.level;
  const includeLevel = forceLevel || level !== 0 || originalLevel !== undefined;
  const bullet = wireBullet(paragraph, original, shapeId, assetCatalog);
  const bulletFont = wireBulletFont(paragraph, original, shapeId);
  const bulletColor = wireBulletColor(paragraph, original, shapeId);
  const bulletSize = wireBulletSize(paragraph, original, shapeId);
  const tabs = wireTabStops(paragraph, original, shapeId);
  const layout = wireParagraphLayout(paragraph, original, shapeId);
  const spacing = wireParagraphSpacing(paragraph, original, shapeId);
  const defaultRunStyle = wireDefaultRunStyle(paragraph, original, shapeId);
  const directInheritedStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => !Object.hasOwn(paragraph.style || {}, key)));
  return {
    ...(includeLevel ? { level } : {}),
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    runs: (paragraph.runs || []).map((run, index) => wireRun(run, directInheritedStyle, shapeId, original?.runs?.[index])),
    ...(bullet ? { bullet } : {}),
    ...(bulletFont ? { bulletFont } : {}),
    ...(bulletColor ? { bulletColor } : {}),
    ...(bulletSize ? { bulletSize } : {}),
    ...tabs,
    ...layout,
    ...spacing,
    ...(defaultRunStyle ? { defaultRunStyle } : {}),
  };
}

function modelRunCase(run) {
  if (run.break) return "lineBreak";
  if (run.field) return "field";
  return "text";
}

function wireTextBodyProperties(value, original, shapeId) {
  let properties;
  try {
    properties = normalizePresentationTextBodyProperties(value);
  } catch (error) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses invalid text body properties: ${error.message}`, [], { code: "invalid_presentation_text" });
  }
  const originalProperties = original?.bodyProperties;
  const insetChoice = (key, wireName, noWireName) => {
    if (properties.insets?.[key] != null) {
      const emu = Math.round(properties.insets[key] * EMU_PER_PIXEL);
      if (emu < 0 || emu > MAX_TEXT_BODY_INSET_EMU) throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an out-of-range ${key} text inset.`, [], { code: "invalid_presentation_text" });
      return { case: wireName, value: BigInt(emu) };
    }
    const originalCase = originalProperties?.[`${key}Inset`]?.case;
    return new Set([wireName, noWireName]).has(originalCase) ? { case: noWireName, value: true } : undefined;
  };
  const leftInset = insetChoice("left", "leftInsetEmu", "noLeftInset");
  const topInset = insetChoice("top", "topInsetEmu", "noTopInset");
  const rightInset = insetChoice("right", "rightInsetEmu", "noRightInset");
  const bottomInset = insetChoice("bottom", "bottomInsetEmu", "noBottomInset");
  const anchor = properties.anchor != null
    ? { case: "verticalAnchor", value: properties.anchor }
    : new Set(["verticalAnchor", "noVerticalAnchor"]).has(originalProperties?.anchor?.case)
      ? { case: "noVerticalAnchor", value: true }
      : undefined;
  const wrapping = properties.wrap != null
    ? { case: "wrap", value: properties.wrap }
    : new Set(["wrap", "noWrap"]).has(originalProperties?.wrapping?.case)
      ? { case: "noWrap", value: true }
      : undefined;
  const autoFit = properties.autoFit != null
    ? { case: "autoFitMode", value: properties.autoFit }
    : new Set(["autoFitMode", "noAutoFitMode"]).has(originalProperties?.autoFit?.case)
      ? { case: "noAutoFitMode", value: true }
      : undefined;
  const rotation = properties.rotation != null
    ? { case: "rotationAngle60000", value: Math.round(properties.rotation * ROTATION_UNITS_PER_DEGREE) }
    : new Set(["rotationAngle60000", "noRotation"]).has(originalProperties?.rotation?.case)
      ? { case: "noRotation", value: true }
      : undefined;
  const verticalText = properties.verticalText != null
    ? { case: "verticalTextMode", value: properties.verticalText }
    : new Set(["verticalTextMode", "noVerticalTextMode"]).has(originalProperties?.verticalText?.case)
      ? { case: "noVerticalTextMode", value: true }
      : undefined;
  const verticalOverflow = properties.verticalOverflow != null
    ? { case: "verticalOverflowMode", value: properties.verticalOverflow }
    : new Set(["verticalOverflowMode", "noVerticalOverflowMode"]).has(originalProperties?.verticalOverflow?.case)
      ? { case: "noVerticalOverflowMode", value: true }
      : undefined;
  const horizontalOverflow = properties.horizontalOverflow != null
    ? { case: "horizontalOverflowMode", value: properties.horizontalOverflow }
    : new Set(["horizontalOverflowMode", "noHorizontalOverflowMode"]).has(originalProperties?.horizontalOverflow?.case)
      ? { case: "noHorizontalOverflowMode", value: true }
      : undefined;
  const columnCount = properties.columns?.count != null
    ? { case: "columns", value: properties.columns.count }
    : new Set(["columns", "noColumns"]).has(originalProperties?.columnCount?.case)
      ? { case: "noColumns", value: true }
      : undefined;
  const columnSpacing = properties.columns?.spacing != null
    ? { case: "columnSpacingEmu", value: BigInt(Math.round(properties.columns.spacing * EMU_PER_PIXEL)) }
    : new Set(["columnSpacingEmu", "noColumnSpacing"]).has(originalProperties?.columnSpacing?.case)
      ? { case: "noColumnSpacing", value: true }
      : undefined;
  const columnDirection = properties.columns?.rightToLeft != null
    ? { case: "rightToLeftColumns", value: properties.columns.rightToLeft }
    : new Set(["rightToLeftColumns", "noColumnDirection"]).has(originalProperties?.columnDirection?.case)
      ? { case: "noColumnDirection", value: true }
      : undefined;
  const uprightText = properties.upright != null
    ? { case: "upright", value: properties.upright }
    : new Set(["upright", "noUpright"]).has(originalProperties?.uprightText?.case)
      ? { case: "noUpright", value: true }
      : undefined;
  if (![leftInset, topInset, rightInset, bottomInset, anchor, wrapping, autoFit, rotation, verticalText, verticalOverflow, horizontalOverflow, columnCount, columnSpacing, columnDirection, uprightText].some(Boolean)) return undefined;
  return {
    ...(leftInset ? { leftInset } : {}),
    ...(topInset ? { topInset } : {}),
    ...(rightInset ? { rightInset } : {}),
    ...(bottomInset ? { bottomInset } : {}),
    ...(anchor ? { anchor } : {}),
    ...(wrapping ? { wrapping } : {}),
    ...(autoFit ? { autoFit } : {}),
    ...(rotation ? { rotation } : {}),
    ...(verticalText ? { verticalText } : {}),
    ...(verticalOverflow ? { verticalOverflow } : {}),
    ...(horizontalOverflow ? { horizontalOverflow } : {}),
    ...(columnCount ? { columnCount } : {}),
    ...(columnSpacing ? { columnSpacing } : {}),
    ...(columnDirection ? { columnDirection } : {}),
    ...(uprightText ? { uprightText } : {}),
  };
}

function presentationTextBody(shape, original, assetCatalog) {
  const textStyle = shape.text?.style || {};
  const textStyleUnsupported = Object.keys(textStyle).filter((key) => !RUN_STYLE_KEYS.has(key) && !TEXT_FRAME_PARAGRAPH_KEYS.has(key));
  if (textStyleUnsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
  const inheritedRunStyle = Object.fromEntries(Object.entries(textStyle).filter(([key]) => RUN_STYLE_KEYS.has(key)));
  const inheritedParagraph = Object.fromEntries(Object.entries(textStyle).filter(([key]) => TEXT_FRAME_PARAGRAPH_KEYS.has(key)));
  const paragraphs = shape.text?.paragraphs || [];
  if (original?.textBody && (original.textBody.paragraphs.length !== paragraphs.length || original.textBody.paragraphs.some((paragraph, index) => paragraph.runs.length !== (paragraphs[index]?.runs || []).length || paragraph.runs.some((run, runIndex) => run.content?.case !== modelRunCase(paragraphs[index].runs[runIndex]))))) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} changed its source-bound paragraph/inline topology.`, [], { code: "presentation_text_topology_changed" });
  }
  const originalListStyles = new Map((original?.textBody?.listStyles || []).map((style) => [Number(style.level), style]));
  const inheritedParagraphStyles = Object.entries(shape.text?.inheritedParagraphStyles || {}).sort(([left], [right]) => Number(left) - Number(right));
  const listStyles = inheritedParagraphStyles.map(([level, style]) => wireParagraph(
    { ...style, level: Number(level), runs: [] },
    {},
    originalListStyles.get(Number(level)),
    shape.id,
    assetCatalog,
    { forceLevel: true },
  ));
  const noListStyles = listStyles.length === 0 && (originalListStyles.size > 0 || original?.textBody?.noListStyles === true);
  const bodyProperties = wireTextBodyProperties(shape.text?.bodyProperties, original?.textBody, shape.id);
  return {
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph({ ...inheritedParagraph, ...paragraph }, inheritedRunStyle, original?.textBody?.paragraphs?.[index], shape.id, assetCatalog)),
    ...(listStyles.length ? { listStyles } : {}),
    ...(noListStyles ? { noListStyles: true } : {}),
    ...(bodyProperties ? { bodyProperties } : {}),
  };
}

const MASTER_STYLE_KINDS = [
  ["title", "titleLevels", "deletedTitleLevels"],
  ["body", "bodyLevels", "deletedBodyLevels"],
  ["other", "otherLevels", "deletedOtherLevels"],
];

function wireMasterTextStyles(master, original, assetCatalog) {
  const result = {};
  for (const [kind, levelsField, deletedField] of MASTER_STYLE_KINDS) {
    const originalLevels = new Map((original?.textStyles?.[levelsField] || []).map((style) => [Number(style.level), style]));
    const current = master.textParagraphStyles?.[kind] || {};
    const levels = Object.entries(current)
      .sort(([left], [right]) => Number(left) - Number(right))
      .map(([level, style]) => wireParagraph(
        { ...style, level: Number(level), runs: [] },
        {},
        originalLevels.get(Number(level)),
        `master ${master.id} ${kind} level ${Number(level) + 1}`,
        assetCatalog,
        { forceLevel: true },
      ));
    const currentLevels = new Set(Object.keys(current).map(Number));
    const deleted = [...originalLevels.keys()].filter((level) => !currentLevels.has(level)).sort((left, right) => left - right);
    result[levelsField] = levels;
    if (deleted.length) result[deletedField] = deleted;
  }
  return result;
}

function wireBackground(background, ownerId) {
  if (!background) return undefined;
  const fill = String(background.fill || "").trim();
  const color = PRESENTATION_SCHEME_COLORS.has(fill)
    ? { case: "colorScheme", value: fill }
    : { case: "colorRgb", value: presentationRgb(fill, `${ownerId}.background.fill`) };
  if (!color.value) throw new OpenChestnutCodecError(`Presentation ${ownerId} uses an unsupported transparent background.`, [], { code: "unsupported_presentation_features" });
  if (background.mode === "reference") {
    const index = Number(background.index);
    if (!Number.isInteger(index) || index < 0 || index > 4_294_967_295) {
      throw new OpenChestnutCodecError(`Presentation ${ownerId} background reference index must be an unsigned 32-bit integer.`, [], { code: "invalid_presentation_background" });
    }
    return { color, kind: { case: "styleReferenceIndex", value: index } };
  }
  if (background.mode !== "solid") throw new OpenChestnutCodecError(`Presentation ${ownerId} background mode must be solid or reference.`, [], { code: "invalid_presentation_background" });
  return { color, kind: { case: "solid", value: true } };
}

function modelBackground(background) {
  if (!background?.color?.case || !background?.kind?.case) return undefined;
  const fill = background.color.case === "colorScheme" ? background.color.value : `#${String(background.color.value).toLowerCase()}`;
  return background.kind.case === "styleReferenceIndex"
    ? { fill, mode: "reference", index: Number(background.kind.value) }
    : { fill, mode: "solid" };
}

function wirePresentationTransform(transform, ownerLabel) {
  if (transform == null) return {};
  if (typeof transform !== "object" || Array.isArray(transform)) {
    throw new OpenChestnutCodecError(`Presentation ${ownerLabel} transform must be an object.`, [], { code: "invalid_presentation_transform" });
  }
  const output = {};
  if (Object.hasOwn(transform, "rotationDegrees") && transform.rotationDegrees != null) {
    const degrees = Number(transform.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) {
      throw new OpenChestnutCodecError(`Presentation ${ownerLabel} rotation must be between -360 and 360 degrees.`, [], { code: "invalid_presentation_transform" });
    }
    output.rotationAngle60000 = Math.round(degrees * ROTATION_UNITS_PER_DEGREE);
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(transform, key) || transform[key] == null) continue;
    if (typeof transform[key] !== "boolean") {
      throw new OpenChestnutCodecError(`Presentation ${ownerLabel} ${key} must be a boolean.`, [], { code: "invalid_presentation_transform" });
    }
    output[key] = transform[key];
  }
  if (Object.keys(output).length === 0) {
    throw new OpenChestnutCodecError(`Presentation ${ownerLabel} transform must define rotationDegrees, flipHorizontal, or flipVertical.`, [], { code: "invalid_presentation_transform" });
  }
  return output;
}

function masterReadOnlySnapshot(master) {
  return JSON.stringify(master.toJSON());
}

function layoutReadOnlySnapshot(layout) {
  return JSON.stringify(layout.toJSON());
}

function presentationMasters(presentation, state, assetCatalog) {
  if (state) {
    if (presentation.masters.items.length !== state.masters.length || state.masters.some((entry, index) => presentation.masters.items[index] !== entry.model)) {
      throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.masters.length}-master topology.`, [], { code: "presentation_master_topology_changed" });
    }
    return state.masters.map((entry) => {
      if (masterReadOnlySnapshot(entry.model) !== entry.snapshot) {
        throw new OpenChestnutCodecError(`Presentation master ${entry.model.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_presentation_edit" });
      }
      return entry.wire;
    });
  }
  const master = presentation.master;
  return master ? [{ id: master.id, name: master.name, textStyles: wireMasterTextStyles(master, undefined, assetCatalog), background: wireBackground(master.background, `master ${master.id}`) }] : [];
}

function presentationLayouts(presentation, state, assetCatalog) {
  if (!state) return [];
  if (presentation.layouts.items.length !== state.layouts.length || state.layouts.some((entry, index) => presentation.layouts.items[index] !== entry.model)) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.layouts.length}-layout topology.`, [], { code: "presentation_layout_topology_changed" });
  }
  return state.layouts.map((entry) => {
    if (layoutReadOnlySnapshot(entry.model) !== entry.snapshot) {
      throw new OpenChestnutCodecError(`Presentation layout ${entry.model.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_presentation_edit" });
    }
    return entry.wire;
  });
}

function presentationShadow(shadow, shapeId) {
  if (shadow == null || shadow === false || shadow === "shadow-none") return undefined;
  const presets = {
    "shadow-sm": { color: "#000000", blurRadius: 4, distance: 2, direction: 45, opacity: 0.15 },
    shadow: { color: "#000000", blurRadius: 6, distance: 3, direction: 45, opacity: 0.18 },
    "shadow-md": { color: "#000000", blurRadius: 10, distance: 4, direction: 45, opacity: 0.2 },
    "shadow-lg": { color: "#000000", blurRadius: 15, distance: 6, direction: 45, opacity: 0.22 },
    "shadow-xl": { color: "#000000", blurRadius: 22, distance: 9, direction: 45, opacity: 0.24 },
    "shadow-2xl": { color: "#000000", blurRadius: 32, distance: 14, direction: 45, opacity: 0.25 },
  };
  let source = typeof shadow === "string" ? presets[shadow] : shadow;
  if (!source && typeof shadow === "string") {
    const match = /^(-?\d+(?:\.\d+)?)px\s+(-?\d+(?:\.\d+)?)px\s+(\d+(?:\.\d+)?)px\s+(#[0-9a-f]{6})(?:\/(\d+(?:\.\d+)?))?$/i.exec(shadow.trim());
    if (match) {
      const offsetX = Number(match[1]);
      const offsetY = Number(match[2]);
      source = {
        color: match[4],
        blurRadius: Number(match[3]),
        distance: Math.hypot(offsetX, offsetY),
        direction: (Math.atan2(offsetY, offsetX) * 180 / Math.PI + 360) % 360,
        opacity: match[5] == null ? 1 : Number(match[5]) / 100,
      };
    }
  }
  if (!source || typeof source !== "object") {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} uses an unsupported shadow.`, [], { code: "unsupported_presentation_features" });
  }
  const blurRadius = Number(source.blurRadius ?? source.blur ?? 0);
  const distance = Number(source.distance ?? 0);
  const direction = Number(source.direction ?? source.angle ?? 0);
  const opacity = Number(source.opacity ?? 0.2);
  if (![blurRadius, distance, direction, opacity].every(Number.isFinite) || blurRadius < 0 || distance < 0 || opacity < 0 || opacity > 1) {
    throw new OpenChestnutCodecError(`Presentation shape ${shapeId} has an invalid shadow.`, [], { code: "invalid_presentation_shadow" });
  }
  const normalizedDirection = ((direction % 360) + 360) % 360;
  return {
    colorRgb: presentationRgb(source.color || source.fill || "#000000", `${shapeId}.shadow.color`),
    blurRadiusEmu: emuFromPixels(blurRadius, `${shapeId}.shadow.blurRadius`),
    distanceEmu: emuFromPixels(distance, `${shapeId}.shadow.distance`),
    directionAngle60000: Math.round(normalizedDirection * ROTATION_UNITS_PER_DEGREE),
    opacityThousandthPercent: Math.round(opacity * 100_000),
  };
}

function modelPresentationShadow(shadow) {
  if (!shadow) return undefined;
  return {
    color: shadow.colorRgb ? `#${shadow.colorRgb}` : "#000000",
    blurRadius: Number(shadow.blurRadiusEmu) / EMU_PER_PIXEL,
    distance: Number(shadow.distanceEmu) / EMU_PER_PIXEL,
    direction: Number(shadow.directionAngle60000) / ROTATION_UNITS_PER_DEGREE,
    opacity: Number(shadow.opacityThousandthPercent) / 100_000,
  };
}

function presentationConnector(connector, original) {
  const type = String(connector.connectorType || "straight");
  if (!new Set(["straight", "elbow"]).has(type)) {
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses unsupported type ${type}.`, [], { code: "unsupported_presentation_features" });
  }
  const line = connector.line || {};
  if (line.style != null && !new Set(["solid", "none"]).has(String(line.style))) {
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses an unsupported line style.`, [], { code: "unsupported_presentation_features" });
  }
  const width = Number(line.width ?? 2);
  if (!Number.isFinite(width) || width < 0) throw new OpenChestnutCodecError(`Presentation connector ${connector.id} has an invalid line width.`, [], { code: "invalid_presentation_connector" });
  const arrow = (value, name) => {
    if (value == null || value === false || value === "none") return "";
    if (value === true || value === "triangle") return "triangle";
    throw new OpenChestnutCodecError(`Presentation connector ${connector.id} uses unsupported ${name} ${value}.`, [], { code: "unsupported_presentation_features" });
  };
  return {
    id: original?.id || connector.id,
    name: connector.name || original?.name || "",
    source: original?.source,
    content: {
      case: "connector",
      value: {
        connectorType: type,
        startXEmu: emuFromPixels(connector.start?.x, `${connector.id}.start.x`),
        startYEmu: emuFromPixels(connector.start?.y, `${connector.id}.start.y`),
        endXEmu: emuFromPixels(connector.end?.x, `${connector.id}.end.x`),
        endYEmu: emuFromPixels(connector.end?.y, `${connector.id}.end.y`),
        lineRgb: presentationRgb(line.fill || line.color || (width > 0 ? "#334155" : "transparent"), `${connector.id}.line.fill`),
        lineWidthEmu: BigInt(Math.round(width * EMU_PER_POINT)),
        startArrow: arrow(line.startArrow ?? connector.startArrow, "start arrow"),
        endArrow: arrow(line.endArrow ?? connector.endArrow, "end arrow"),
        startTargetId: connector.startTargetId || "",
        endTargetId: connector.endTargetId || "",
      },
    },
  };
}

function presentationChartColor(value, chart, field) {
  if (value == null || value === "") return undefined;
  const rgb = presentationRgb(value, `${chart.id}.${field}`);
  if (!rgb) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field} cannot be transparent.`, [], { code: "unsupported_presentation_features" });
  return { source: { case: "rgb", value: rgb } };
}

function presentationChartLine(line, chart, field) {
  if (!line) return undefined;
  const styleName = String(line.style || "solid");
  const aliases = { dot: "dotted", dash: "dashed", dashDot: "dash-dot", longDashDotDot: "dash-dot-dot" };
  const dashStyle = PRESENTATION_CHART_LINE_STYLES_TO_WIRE.get(aliases[styleName] || styleName);
  if (dashStyle == null) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field} uses unsupported line style ${styleName}.`, [], { code: "unsupported_presentation_features" });
  const width = Number(line.width ?? line.weight ?? 1);
  if (!Number.isFinite(width) || width < 0 || width > 1584) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field} has an invalid line width.`, [], { code: "invalid_presentation_chart" });
  return {
    ...(line.fill || line.color ? { color: presentationChartColor(line.fill || line.color, chart, `${field}.color`) } : {}),
    dashStyle,
    widthPoints: width,
  };
}

function presentationChartMarker(marker, chart, field) {
  if (!marker) return undefined;
  const symbol = PRESENTATION_CHART_MARKERS_TO_WIRE.get(String(marker.symbol || "circle"));
  if (symbol == null) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field} uses unsupported marker ${marker.symbol}.`, [], { code: "unsupported_presentation_features" });
  const size = marker.size == null ? undefined : Number(marker.size);
  if (size !== undefined && (!Number.isInteger(size) || size < 2 || size > 72)) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field} has an invalid marker size.`, [], { code: "invalid_presentation_chart" });
  return {
    symbol,
    ...(size === undefined ? {} : { size }),
    ...(marker.fill ? { fill: presentationChartColor(marker.fill, chart, `${field}.fill`) } : {}),
    ...(marker.line || marker.stroke ? { line: presentationChartLine(marker.line || marker.stroke, chart, `${field}.line`) } : {}),
  };
}

function presentationChartAxis(axis, chart, field, original) {
  const title = typeof axis?.title === "object" ? axis.title.text : axis?.title;
  const result = {
    title: String(title || ""),
    numberFormatCode: String(axis?.numberFormatCode || axis?.numberFormat || ""),
    ...(axis?.tickLabelInterval == null ? {} : { tickLabelInterval: Number(axis.tickLabelInterval) }),
    ...(axis?.min == null ? {} : { minimum: Number(axis.min) }),
    ...(axis?.max == null ? {} : { maximum: Number(axis.max) }),
    ...(axis?.majorUnit == null ? {} : { majorUnit: Number(axis.majorUnit) }),
  };
  const hasSemantics = result.title || result.numberFormatCode || result.tickLabelInterval !== undefined || result.minimum !== undefined || result.maximum !== undefined || result.majorUnit !== undefined;
  if (!hasSemantics && !original) return undefined;
  for (const [name, value] of Object.entries(result)) {
    if (typeof value === "number" && !Number.isFinite(value)) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} ${field}.${name} must be finite.`, [], { code: "invalid_presentation_chart" });
  }
  return result;
}

function presentationChartDataLabels(labels, chart, original) {
  const source = labels || {};
  const positionName = ({ b: "bottom", ctr: "center", inBase: "insideBase", inEnd: "insideEnd", l: "left", outEnd: "outsideEnd", r: "right", t: "top" })[source.position] || source.position || "bestFit";
  const position = PRESENTATION_CHART_LABEL_POSITIONS_TO_WIRE.get(positionName);
  if (position == null) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} uses unsupported data-label position ${source.position}.`, [], { code: "unsupported_presentation_features" });
  const hasSemantics = Boolean(source.showValue || source.showCategoryName || source.showSeriesName) || (source.position && source.position !== "bestFit");
  if (!hasSemantics && !original) return undefined;
  return {
    showValue: Boolean(source.showValue),
    showCategoryName: Boolean(source.showCategoryName),
    ...(source.showSeriesName == null ? {} : { showSeriesName: Boolean(source.showSeriesName) }),
    ...(source.position == null && !original?.position ? {} : { position }),
  };
}

function presentationChart(chart, original) {
  const originalChart = original?.content?.case === "chart" ? original.content.value : undefined;
  const type = PRESENTATION_CHART_TYPES_TO_WIRE.get(String(chart.chartType));
  if (type == null) throw new OpenChestnutCodecError(`Presentation chart ${chart.id} supports only bar, line, or pie.`, [], { code: "unsupported_presentation_features" });
  if (chart.externalData || chart.series.some((series) => series.categoryFormula || series.valueFormula || series.categoriesFormula || series.valuesFormula)) {
    throw new OpenChestnutCodecError(`Presentation chart ${chart.id} must use literal categories and values.`, [], { code: "unsupported_presentation_features" });
  }
  if (!Array.isArray(chart.categories) || chart.categories.length > MAX_PRESENTATION_CHART_POINTS || chart.series.length < 1 || chart.series.length > 256) {
    throw new OpenChestnutCodecError(`Presentation chart ${chart.id} exceeds the bounded category or series budget.`, [], { code: "invalid_presentation_chart" });
  }
  if (originalChart && (originalChart.type !== type || originalChart.series.length !== chart.series.length || originalChart.categories.length !== chart.categories.length)) {
    throw new OpenChestnutCodecError(`Presentation chart ${chart.id} cannot change its imported type, series count, or point topology.`, [], { code: "presentation_chart_topology_changed" });
  }
  const position = chart.position || {};
  const series = chart.series.map((item, index) => {
    if (!Array.isArray(item.values) || item.values.length !== chart.categories.length || item.values.some((value) => !Number.isFinite(Number(value)))) {
      throw new OpenChestnutCodecError(`Presentation chart ${chart.id} series ${index + 1} must contain one finite value per category.`, [], { code: "invalid_presentation_chart" });
    }
    if (item.axisGroup === "secondary" || item.chartType || item.points?.length || item.trendlines?.length || item.errorBars || item.dataLabels || item.smooth != null) {
      throw new OpenChestnutCodecError(`Presentation chart ${chart.id} series ${index + 1} uses semantics outside the bounded bar/line/pie slice.`, [], { code: "unsupported_presentation_features" });
    }
    return {
      name: String(item.name || `Series ${index + 1}`),
      values: item.values.map(Number),
      ...(item.fill || item.color ? { fill: presentationChartColor(item.fill || item.color, chart, `series[${index}].fill`) } : {}),
      ...(item.line || item.stroke ? { line: presentationChartLine(item.line || item.stroke, chart, `series[${index}].line`) } : {}),
      ...(item.marker ? { marker: presentationChartMarker(item.marker, chart, `series[${index}].marker`) } : {}),
    };
  });
  const xAxis = type === SpreadsheetChartType.PIE ? undefined : presentationChartAxis(chart.axes?.category, chart, "xAxis", originalChart?.xAxis);
  const yAxis = type === SpreadsheetChartType.PIE ? undefined : presentationChartAxis(chart.axes?.value, chart, "yAxis", originalChart?.yAxis);
  const dataLabels = presentationChartDataLabels(chart.dataLabels, chart, originalChart?.dataLabels);
  return {
    id: original?.id || chart.id,
    name: chart.name || original?.name || chart.id,
    source: original?.source,
    content: {
      case: "chart",
      value: {
        leftEmu: emuFromPixels(position.left, `${chart.id}.position.left`),
        topEmu: emuFromPixels(position.top, `${chart.id}.position.top`),
        widthEmu: emuFromPixels(position.width, `${chart.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${chart.id}.position.height`),
        type,
        title: String(chart.title || ""),
        hasLegend: Boolean(chart.legend?.visible ?? chart.hasLegend),
        categories: chart.categories.map((value) => String(value ?? "")),
        series,
        ...(xAxis ? { xAxis } : {}),
        ...(yAxis ? { yAxis } : {}),
        ...(dataLabels ? { dataLabels } : {}),
      },
    },
  };
}

function presentationShape(shape, original, assetCatalog) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse", "roundRect", "textbox", "custom"]).has(shape.geometry)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  if (shape.geometry !== "custom" && shape.customPaths?.length) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} has custom paths without custom geometry.`, [], { code: "invalid_presentation_geometry" });
  }
  const customPaths = (shape.customPaths || []).map((path) => ({
    width: BigInt(path.width),
    height: BigInt(path.height),
    commands: path.commands.map((command) => {
      if (command.moveTo) return { command: { case: "moveTo", value: { x: BigInt(command.moveTo.x), y: BigInt(command.moveTo.y) } } };
      if (command.lineTo) return { command: { case: "lineTo", value: { x: BigInt(command.lineTo.x), y: BigInt(command.lineTo.y) } } };
      if (command.cubicBezTo) return {
        command: {
          case: "cubicBezierTo",
          value: {
            control1: { x: BigInt(command.cubicBezTo.x1), y: BigInt(command.cubicBezTo.y1) },
            control2: { x: BigInt(command.cubicBezTo.x2), y: BigInt(command.cubicBezTo.y2) },
            end: { x: BigInt(command.cubicBezTo.x), y: BigInt(command.cubicBezTo.y) },
          },
        },
      };
      return { command: { case: "close", value: true } };
    }),
  }));
  if (shape.geometry === "custom" && customPaths.length === 0) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} requires custom paths.`, [], { code: "invalid_presentation_geometry" });
  }
  const position = shape.position || {};
  const lineWidth = Number(shape.line?.width ?? 1);
  if (!Number.isFinite(lineWidth) || lineWidth < 0) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} has an invalid line width.`, [], { code: "invalid_presentation_frame" });
  const textBody = presentationTextBody(shape, originalShape, assetCatalog);
  const shadow = presentationShadow(shape.shadow, shape.id);
  return {
    id: original?.id || shape.id,
    name: shape.name || original?.name || "",
    source: original?.source,
    content: {
      case: "shape",
      value: {
        geometry: shape.geometry,
        leftEmu: emuFromPixels(position.left, `${shape.id}.position.left`),
        topEmu: emuFromPixels(position.top, `${shape.id}.position.top`),
        widthEmu: emuFromPixels(position.width, `${shape.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${shape.id}.position.height`),
        text: shape.text?.value || "",
        textBody,
        fillRgb: presentationRgb(shape.fill, `${shape.id}.fill`),
        lineRgb: presentationRgb(shape.line?.fill || shape.line?.color || (lineWidth > 0 ? "#334155" : "transparent"), `${shape.id}.line.fill`),
        lineWidthEmu: BigInt(Math.round(lineWidth * EMU_PER_POINT)),
        ...(shape.transform == null ? {} : { transform: wirePresentationTransform(shape.transform, `shape ${shape.id}`) }),
        ...(shadow ? { shadow } : {}),
        ...(customPaths.length ? { customPaths } : {}),
      },
    },
  };
}

function presentationImage(image, original, assetCatalog) {
  const position = image.position || {};
  if (!image.dataUrl) {
    throw new OpenChestnutCodecError(`Presentation image ${image.id} requires an embedded dataUrl.`, [], { code: "invalid_presentation_image" });
  }
  if (image.uri || image.geometry !== "rect" || image.borderRadius != null || !new Set(["contain", "stretch"]).has(image.fit)) {
    throw new OpenChestnutCodecError(`Presentation image ${image.id} uses crop, external, geometry, or fit semantics outside the bounded PPTX image slice.`, [], { code: "unsupported_presentation_features" });
  }
  return {
    id: original?.id || image.id,
    name: image.name || original?.name || "",
    source: original?.source,
    content: {
      case: "image",
      value: {
        assetId: assetCatalog.addDataUrl(image.dataUrl),
        altText: image.alt || image.prompt || "",
        leftEmu: emuFromPixels(position.left, `${image.id}.position.left`),
        topEmu: emuFromPixels(position.top, `${image.id}.position.top`),
        widthEmu: emuFromPixels(position.width, `${image.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${image.id}.position.height`),
        ...(image.transform == null ? {} : { transform: wirePresentationTransform(image.transform, `image ${image.id}`) }),
      },
    },
  };
}

function presentationImageReadOnlySnapshot(image) {
  return JSON.stringify({
    uri: image.uri,
    contentType: image.contentType,
    fit: image.fit,
    geometry: image.geometry,
    borderRadius: image.borderRadius,
  });
}

function distributePresentationTableSize(total, count, ownerLabel) {
  const slots = Number(count);
  if (!Number.isInteger(slots) || slots < 1) {
    throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} must contain at least one row and column.`, [], { code: "invalid_presentation_table" });
  }
  const base = total / BigInt(slots);
  const remainder = Number(total % BigInt(slots));
  if (base < 1n) throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} is too small for its grid.`, [], { code: "invalid_presentation_table" });
  return Array.from({ length: slots }, (_, index) => base + (index < remainder ? 1n : 0n));
}

function scalePresentationTableSize(values, total, ownerLabel) {
  const source = values.map((value) => BigInt(value));
  const sourceTotal = source.reduce((sum, value) => sum + value, 0n);
  if (!source.length || sourceTotal < 1n) return distributePresentationTableSize(total, source.length, ownerLabel);
  const scaled = source.map((value) => ({ value: (value * total) / sourceTotal, remainder: (value * total) % sourceTotal }));
  let missing = total - scaled.reduce((sum, item) => sum + item.value, 0n);
  for (const index of scaled.map((item, index) => ({ index, remainder: item.remainder })).sort((left, right) => left.remainder === right.remainder ? left.index - right.index : left.remainder > right.remainder ? -1 : 1).map((item) => item.index)) {
    if (missing <= 0n) break;
    scaled[index].value += 1n;
    missing -= 1n;
  }
  if (scaled.some((item) => item.value < 1n)) {
    throw new OpenChestnutCodecError(`Presentation table ${ownerLabel} is too small for its imported grid.`, [], { code: "invalid_presentation_table" });
  }
  return scaled.map((item) => item.value);
}

function presentationTable(table, original) {
  const originalTable = original?.content?.case === "table" ? original.content.value : undefined;
  const rows = Number(table.rows);
  const columns = Number(table.columns);
  if (!Number.isInteger(rows) || rows < 1 || rows > 2048 || !Number.isInteger(columns) || columns < 1 || columns > 256 ||
      table.values.length !== rows || table.values.some((row) => !Array.isArray(row) || row.length !== columns)) {
    throw new OpenChestnutCodecError(`Presentation table ${table.id} requires a rectangular 1-2048 by 1-256 value matrix.`, [], { code: "invalid_presentation_table" });
  }
  const position = table.position || {};
  const leftEmu = emuFromPixels(position.left, `${table.id}.position.left`);
  const topEmu = emuFromPixels(position.top, `${table.id}.position.top`);
  const widthEmu = emuFromPixels(position.width, `${table.id}.position.width`);
  const heightEmu = emuFromPixels(position.height, `${table.id}.position.height`);
  if (widthEmu < 1n || heightEmu < 1n) throw new OpenChestnutCodecError(`Presentation table ${table.id} requires positive width and height.`, [], { code: "invalid_presentation_table" });
  if (originalTable && (originalTable.rows.length !== rows || originalTable.columnWidthsEmu.length !== columns)) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires presentation table ${table.id}'s original fixed topology.`, [], { code: "presentation_table_topology_changed" });
  }
  const columnWidthsEmu = originalTable
    ? scalePresentationTableSize(originalTable.columnWidthsEmu, widthEmu, `${table.id} columns`)
    : distributePresentationTableSize(widthEmu, columns, `${table.id} columns`);
  const rowHeightsEmu = originalTable
    ? scalePresentationTableSize(originalTable.rows.map((row) => row.heightEmu), heightEmu, `${table.id} rows`)
    : distributePresentationTableSize(heightEmu, rows, `${table.id} rows`);
  return {
    id: original?.id || table.id,
    name: String(table.name || original?.name || ""),
    source: original?.source,
    content: {
      case: "table",
      value: {
        leftEmu,
        topEmu,
        widthEmu,
        heightEmu,
        columnWidthsEmu,
        rows: table.values.map((row, rowIndex) => ({
          heightEmu: rowHeightsEmu[rowIndex],
          cells: row.map((value) => ({ text: String(value ?? "") })),
        })),
        ...(originalTable?.firstRow === undefined ? { firstRow: Boolean(table.styleOptions?.headerRow) } : { firstRow: originalTable.firstRow }),
        ...(originalTable?.bandedRows === undefined ? { bandedRows: Boolean(table.styleOptions?.bandedRows) } : { bandedRows: originalTable.bandedRows }),
      },
    },
  };
}

function presentationTableReadOnlySnapshot(table) {
  return JSON.stringify({
    id: table.id,
    nativeId: table.nativeId,
    creationId: table.creationId,
    rows: table.rows,
    columns: table.columns,
    style: table.style,
    styleOptions: table.styleOptions,
    border: table.border,
    mergeRange: table.mergeRange,
  });
}

function directSlideElements(slide) {
  return [
    ...slide.shapes.items,
    ...slide.tables.items,
    ...slide.charts.items,
    ...slide.images.items,
    ...slide.connectors.items,
    ...slide.groups.items,
    ...slide.nativeObjects.items,
  ];
}

function presentationThemeSnapshot(theme) {
  return JSON.stringify({
    name: theme.name,
    colors: theme.colors,
    fonts: theme.fonts,
    textStyles: theme.textStyles,
    colorMap: theme.colorMap,
  });
}

function presentationAdvancedSnapshot(presentation) {
  return JSON.stringify({
    theme: JSON.parse(presentationThemeSnapshot(presentation.theme)),
    commentFormat: presentation.commentFormat,
    customShows: presentation.customShows.items.map((show) => show.toJSON()),
    slides: presentation.slides.items.map((slide) => ({
      background: slide.background,
      comments: slide.comments.items.map((comment) => comment.toJSON()),
    })),
  });
}

function unsupportedPresentationFeatures(presentation) {
  const unsupported = [];
  if (presentationThemeSnapshot(presentation.theme) !== DEFAULT_PRESENTATION_THEME) unsupported.push("presentation theme customization");
  if (presentation.layouts?.items?.length) unsupported.push("slide layouts");
  if (presentation.masters?.items?.length !== 1) unsupported.push("multiple slide masters");
  const master = presentation.master;
  if (master?.configured) unsupported.push("slide master authoring");
  if (master?.theme) unsupported.push("master theme override");
  if (master?.placeholders?.length) unsupported.push("master placeholders");
  if (presentation.customShows?.items?.length) unsupported.push("custom shows");
  if (presentation.commentFormat !== "legacy") unsupported.push("modern comments");
  for (const slide of presentation.slides?.items || []) {
    const prefix = `slide ${slide.index + 1}`;
    if (slide.layoutId) unsupported.push(`${prefix} layout binding`);
    if (slide.background?.fill) unsupported.push(`${prefix} background`);
    if (slide.comments?.items?.length) unsupported.push(`${prefix} comments`);
    if (slide.groups?.items?.length) unsupported.push(`${prefix} groups`);
    if (slide.nativeObjects?.items?.length) unsupported.push(`${prefix} native objects`);
    if (slide.shapes?.items?.some((shape) => shape.placeholder)) unsupported.push(`${prefix} source-free placeholder authoring`);
  }
  return unsupported;
}

function opaquePresentationSnapshot(object) {
  const oleWorkbook = object.oleWorkbook ? {
    partPath: object.oleWorkbook.partPath,
    contentType: object.oleWorkbook.contentType,
    sourceSha256: object.oleWorkbook.sourceSha256,
    relationshipId: object.oleWorkbook.relationshipId,
  } : undefined;
  return JSON.stringify({
    id: object.id,
    name: object.name,
    position: object.position,
    nativeKind: object.nativeKind,
    rawXml: object.rawXml,
    oleWorkbook,
    ...presentationNativeGraphSnapshot(object),
  });
}

export function presentationEnvelope(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("exportPptxWithOpenChestnut expects a Presentation instance.");
  if (!presentation.slides?.items?.length) throw new OpenChestnutCodecError("Presentation must contain at least one slide.", [], { code: "missing_slides" });
  const state = presentation[PRESENTATION_STATE];
  assertTrustedPresentationState(state);
  if (!state) {
    const unsupported = unsupportedPresentationFeatures(presentation);
    if (unsupported.length) {
      throw new OpenChestnutCodecError(`OpenChestnut cannot author these source-free PPTX features: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Export fails closed; use supported features or import a trustworthy source package for opaque preservation.`, [], { code: "unsupported_presentation_features" });
    }
  } else {
    if (presentationAdvancedSnapshot(presentation) !== state.advancedSnapshot) {
      throw new OpenChestnutCodecError("Imported presentation theme, comments, slide backgrounds, and custom shows are source-bound and read-only in OpenChestnut 0.2.", [], { code: "unsupported_presentation_edit" });
    }
    if (state.slides.length !== presentation.slides.items.length) throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.slides.length}-slide topology.`, [], { code: "presentation_topology_changed" });
    if (Number(state.slideWidthEmu) !== Math.round(Number(presentation.slideSize.width) * EMU_PER_PIXEL) || Number(state.slideHeightEmu) !== Math.round(Number(presentation.slideSize.height) * EMU_PER_PIXEL)) {
      throw new OpenChestnutCodecError("Source-preserving PPTX export does not yet support changing slide dimensions.", [], { code: "unsupported_presentation_edit" });
    }
  }

  const assetCatalog = createPresentationAssetCatalog();
  const masters = presentationMasters(presentation, state, assetCatalog);
  const layouts = presentationLayouts(presentation, state, assetCatalog);
  const slides = presentation.slides.items.map((slide, slideIndex) => {
    const sourceState = state?.slides[slideIndex];
    if (sourceState) {
      if (slide.name !== sourceState.name) throw new OpenChestnutCodecError(`Source-preserving PPTX export does not yet support renaming slide ${slideIndex + 1}.`, [], { code: "unsupported_presentation_edit" });
      if ((slide.layoutId || "") !== (sourceState.wire.layoutId || "")) throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot change slide ${slideIndex + 1}'s layout binding.`, [], { code: "presentation_slide_layout_binding_changed" });
      const current = directSlideElements(slide);
      if (current.length !== sourceState.entries.length || sourceState.entries.some((entry) => !current.includes(entry.model))) {
        throw new OpenChestnutCodecError(`Source-preserving PPTX export requires slide ${slideIndex + 1}'s original ${sourceState.entries.length}-element topology.`, [], { code: "presentation_element_topology_changed" });
      }
      if (!sourceState.wire.speakerNotes && slide.speakerNotes?.text) {
        throw new OpenChestnutCodecError(`Source-preserving PPTX export cannot add speaker notes to slide ${slideIndex + 1} because the source slide has no notes part.`, [], { code: "unsupported_presentation_edit" });
      }
    }
    return {
      id: sourceState?.wire.id || slide.id,
      name: slide.name,
      source: sourceState?.wire.source,
      ...(slide.layoutId ? { layoutId: slide.layoutId } : {}),
      ...(sourceState?.wire.speakerNotes
        ? { speakerNotes: { text: slide.speakerNotes?.text || "", source: sourceState.wire.speakerNotes.source } }
        : slide.speakerNotes?.text
          ? { speakerNotes: { text: slide.speakerNotes.text } }
          : {}),
      elements: sourceState
          ? sourceState.entries.map((entry) => {
            if (entry.wire.content.case === "shape") {
              if (entry.wire.content.value.placeholder) {
                if (slidePlaceholderSnapshot(entry.model) !== entry.placeholderSnapshot) {
                  throw new OpenChestnutCodecError(`Presentation slide placeholder ${entry.model.id} is a read-only inherited projection in this codec slice.`, [], { code: "unsupported_presentation_edit" });
                }
                return entry.wire;
              }
              return presentationShape(entry.model, entry.wire, assetCatalog);
            }
            if (entry.wire.content.case === "image") {
              if (presentationImageReadOnlySnapshot(entry.model) !== entry.snapshot) {
                throw new OpenChestnutCodecError(`Presentation image ${entry.model.id} changed outside its embedded rectangular image boundary.`, [], { code: "unsupported_presentation_edit" });
              }
              return presentationImage(entry.model, entry.wire, assetCatalog);
            }
            if (entry.wire.content.case === "table") {
              if (presentationTableReadOnlySnapshot(entry.model) !== entry.snapshot) {
                throw new OpenChestnutCodecError(`Presentation table ${entry.model.id} changed outside its name/frame/plain-text boundary.`, [], { code: "unsupported_presentation_edit" });
              }
              return presentationTable(entry.model, entry.wire);
            }
            if (entry.wire.content.case === "connector") return presentationConnector(entry.model, entry.wire);
            if (entry.wire.content.case === "chart") return presentationChart(entry.model, entry.wire);
            if (opaquePresentationSnapshot(entry.model) !== entry.snapshot) throw new OpenChestnutCodecError(`Presentation native element ${entry.model.id} is source-bound and read-only in OpenChestnut 0.2.`, [], { code: "unsupported_presentation_edit" });
            return entry.wire;
          })
        : directSlideElements(slide)
          .filter((element) => element instanceof Shape || element instanceof ImageElement || element instanceof TableElement || element instanceof ChartElement || slide.connectors.items.includes(element))
          .map((element) => element instanceof ImageElement
            ? presentationImage(element, undefined, assetCatalog)
            : element instanceof TableElement
              ? presentationTable(element, undefined)
              : element instanceof ChartElement
                ? presentationChart(element, undefined)
                : slide.connectors.items.includes(element)
                  ? presentationConnector(element, undefined)
                  : presentationShape(element, undefined, assetCatalog)),
    };
  });
  return {
    protocolVersion,
    family: ArtifactFamily.PRESENTATION,
    source: state?.source,
    assets: assetCatalog.assets(),
    opaqueOpc: state?.opaqueOpc,
    diagnostics: state?.diagnostics || [],
    payload: {
      case: "presentation",
      value: {
        id: presentation.id,
        name: state?.name || "",
        slideWidthEmu: emuFromPixels(presentation.slideSize.width, "slideSize.width"),
        slideHeightEmu: emuFromPixels(presentation.slideSize.height, "slideSize.height"),
        slides,
        masters,
        layouts,
      },
    },
  };
}

function presentationNativeKind(elementName) {
  return ({ pic: "picture", graphicFrame: "graphicFrame", grpSp: "group", cxnSp: "connector", contentPart: "contentPart" })[elementName] || elementName || "nativeObject";
}

function modelRun(run) {
  const hyperlink = run.hyperlink?.case === "runHyperlink" ? modelHyperlink(run.hyperlink.value) : undefined;
  const content = run.content?.case === "lineBreak"
    ? { break: true }
    : run.content?.case === "field"
      ? { field: { id: run.content.value.id, type: run.content.value.type, text: run.content.value.text } }
      : { text: run.content?.case === "text" ? run.content.value : "" };
  return {
    ...content,
    style: {
      ...(run.bold === undefined ? {} : { bold: run.bold }),
      ...(run.italic === undefined ? {} : { italic: run.italic }),
      ...(run.fontSizePoints === undefined ? {} : { fontSize: run.fontSizePoints / POINTS_PER_PIXEL }),
      ...(run.fontFamily === undefined ? {} : { fontFamily: run.fontFamily }),
      ...(run.colorRgb === undefined ? {} : { color: `#${run.colorRgb}` }),
    },
    ...(hyperlink ? { link: hyperlink } : {}),
  };
}

function modelHyperlink(link) {
  const target = link.target?.case === "uri"
    ? { uri: link.target.value }
    : link.target?.case === "slideId"
      ? { slideId: link.target.value }
      : link.target?.case === "action"
        ? { action: link.target.value }
        : {};
  return {
    ...target,
    ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }),
    ...(link.targetFrame === undefined ? {} : { targetFrame: link.targetFrame }),
    ...(link.history === undefined ? {} : { history: link.history }),
    ...(link.highlightClick === undefined ? {} : { highlightClick: link.highlightClick }),
  };
}

function modelBullet(bullet, assetCatalog) {
  if (bullet?.case === "noBullet") return { bulletNone: true };
  if (bullet?.case === "bulletCharacter") return { bulletCharacter: bullet.value };
  if (bullet?.case === "autoNumber") return { autoNumber: { type: bullet.value.scheme, ...(bullet.value.startAt === undefined ? {} : { startAt: bullet.value.startAt }) } };
  if (bullet?.case === "pictureBullet") {
    if (bullet.value.source?.case === "assetId") return { bulletImage: { dataUrl: assetCatalog.dataUrl(bullet.value.source.value), relationshipMode: "embed" } };
    if (bullet.value.source?.case === "uri") return { bulletImage: { uri: validatePictureBulletUri(bullet.value.source.value), relationshipMode: "link" } };
    throw new OpenChestnutCodecError("Presentation picture bullet has no source.", [], { code: "invalid_presentation_asset" });
  }
  return {};
}

function modelBulletStyle(paragraph) {
  return {
    ...(paragraph.bulletFont?.case === "bulletFontFamily" ? { bulletFont: paragraph.bulletFont.value } : {}),
    ...(paragraph.bulletFont?.case === "bulletFontFollowText" ? { bulletFontFollowText: true } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorRgb" ? { bulletColor: `#${paragraph.bulletColor.value}` } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorScheme" ? { bulletColor: paragraph.bulletColor.value } : {}),
    ...(paragraph.bulletColor?.case === "bulletColorFollowText" ? { bulletColorFollowText: true } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePoints" ? { bulletSize: paragraph.bulletSize.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizePercent" ? { bulletSizePercent: paragraph.bulletSize.value } : {}),
    ...(paragraph.bulletSize?.case === "bulletSizeFollowText" ? { bulletSizeFollowText: true } : {}),
  };
}

function modelParagraphLayout(paragraph) {
  return {
    ...(paragraph.leftMargin?.case === "marginLeftEmu" ? { marginLeft: Number(paragraph.leftMargin.value) / EMU_PER_PIXEL } : {}),
    ...(paragraph.indentation?.case === "indentEmu" ? { indent: Number(paragraph.indentation.value) / EMU_PER_PIXEL } : {}),
  };
}

function modelParagraphSpacing(paragraph) {
  return {
    ...(paragraph.lineSpacing?.case === "lineSpacingPoints" ? { lineSpacing: paragraph.lineSpacing.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.lineSpacing?.case === "lineSpacingMultiplier" ? { lineSpacing: paragraph.lineSpacing.value } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforePoints" ? { spaceBefore: paragraph.spaceBefore.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceBefore?.case === "spaceBeforeMultiplier" ? { spaceBeforePercent: paragraph.spaceBefore.value } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterPoints" ? { spaceAfter: paragraph.spaceAfter.value / POINTS_PER_PIXEL } : {}),
    ...(paragraph.spaceAfter?.case === "spaceAfterMultiplier" ? { spaceAfterPercent: paragraph.spaceAfter.value } : {}),
  };
}

function modelDefaultRunStyle(paragraph) {
  if (paragraph.defaultRunStyle?.case !== "defaultRunProperties") return {};
  const style = paragraph.defaultRunStyle.value;
  return {
    ...(style.bold === undefined ? {} : { bold: style.bold }),
    ...(style.italic === undefined ? {} : { italic: style.italic }),
    ...(style.fontSizePoints === undefined ? {} : { fontSize: style.fontSizePoints / POINTS_PER_PIXEL }),
    ...(style.fontFamily === undefined ? {} : { fontFamily: style.fontFamily }),
    ...(style.color?.case === "colorRgb" ? { color: `#${style.color.value}` } : {}),
    ...(style.color?.case === "colorScheme" ? { color: style.color.value } : {}),
  };
}

function modelParagraph(paragraph, assetCatalog, { includeRuns = true } = {}) {
  return {
    ...(includeRuns ? { runs: paragraph.runs.map(modelRun) } : {}),
    level: paragraph.level ?? 0,
    ...(paragraph.alignment ? { alignment: paragraph.alignment } : {}),
    ...modelBullet(paragraph.bullet, assetCatalog),
    ...modelBulletStyle(paragraph),
    ...modelParagraphLayout(paragraph),
    ...modelParagraphSpacing(paragraph),
    ...(paragraph.tabStops?.length ? { tabStops: paragraph.tabStops.map((tab) => ({ position: Number(tab.positionEmu) / EMU_PER_PIXEL, alignment: tab.alignment })) } : {}),
    style: modelDefaultRunStyle(paragraph),
  };
}

function modelText(shape, assetCatalog) {
  if (!shape.textBody) return shape.text;
  return shape.textBody.paragraphs.map((paragraph) => modelParagraph(paragraph, assetCatalog));
}

function modelListStyles(shape, assetCatalog) {
  if (!shape.textBody) return {};
  return Object.fromEntries(shape.textBody.listStyles.map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })]));
}

function modelTextBodyProperties(shape) {
  const source = shape.textBody?.bodyProperties;
  if (!source) return {};
  const properties = {};
  const insets = {};
  for (const [key, choice] of [["left", source.leftInset], ["top", source.topInset], ["right", source.rightInset], ["bottom", source.bottomInset]]) {
    if (choice?.case?.endsWith("InsetEmu")) insets[key] = Number(choice.value) / EMU_PER_PIXEL;
  }
  if (Object.keys(insets).length) properties.insets = insets;
  if (source.anchor?.case === "verticalAnchor") properties.anchor = source.anchor.value;
  if (source.wrapping?.case === "wrap") properties.wrap = source.wrapping.value;
  if (source.autoFit?.case === "autoFitMode") properties.autoFit = source.autoFit.value;
  if (source.rotation?.case === "rotationAngle60000") properties.rotation = source.rotation.value / ROTATION_UNITS_PER_DEGREE;
  if (source.verticalText?.case === "verticalTextMode") properties.verticalText = source.verticalText.value;
  if (source.verticalOverflow?.case === "verticalOverflowMode") properties.verticalOverflow = source.verticalOverflow.value;
  if (source.horizontalOverflow?.case === "horizontalOverflowMode") properties.horizontalOverflow = source.horizontalOverflow.value;
  const columns = {};
  if (source.columnCount?.case === "columns") columns.count = source.columnCount.value;
  if (source.columnSpacing?.case === "columnSpacingEmu") columns.spacing = Number(source.columnSpacing.value) / EMU_PER_PIXEL;
  if (source.columnDirection?.case === "rightToLeftColumns") columns.rightToLeft = source.columnDirection.value;
  if (Object.keys(columns).length) properties.columns = columns;
  if (source.uprightText?.case === "upright") properties.upright = source.uprightText.value;
  return properties;
}

function modelMasterTextStyles(source, assetCatalog) {
  return Object.fromEntries(MASTER_STYLE_KINDS.map(([kind, levelsField]) => [
    kind,
    Object.fromEntries((source?.textStyles?.[levelsField] || []).map((style) => [style.level, modelParagraph(style, assetCatalog, { includeRuns: false })])),
  ]));
}

function modelPlaceholder(source, assetCatalog) {
  const shape = { textBody: source.textBody };
  const transform = modelPlaceholderTransform(source.directFrame);
  return {
    id: source.id,
    name: source.name,
    type: source.type,
    idx: source.index,
    ...(source.directFrame ? { position: modelPlaceholderFrame(source.directFrame) } : {}),
    ...(Object.keys(transform).length ? { transform } : {}),
    text: modelText(shape, assetCatalog),
    paragraphStyles: modelListStyles(shape, assetCatalog),
    textBodyProperties: modelTextBodyProperties(shape),
  };
}

function modelPlaceholderFrame(frame) {
  return {
    left: Number(frame.leftEmu) / EMU_PER_PIXEL,
    top: Number(frame.topEmu) / EMU_PER_PIXEL,
    width: Number(frame.widthEmu) / EMU_PER_PIXEL,
    height: Number(frame.heightEmu) / EMU_PER_PIXEL,
  };
}

function modelPresentationTransform(frame) {
  const transform = {};
  if (frame?.rotationAngle60000 != null) transform.rotationDegrees = frame.rotationAngle60000 / ROTATION_UNITS_PER_DEGREE;
  if (frame?.flipHorizontal != null) transform.flipHorizontal = Boolean(frame.flipHorizontal);
  if (frame?.flipVertical != null) transform.flipVertical = Boolean(frame.flipVertical);
  return transform;
}

function modelCustomGeometryPaths(shape) {
  return (shape.customPaths || []).map((path) => ({
    width: Number(path.width),
    height: Number(path.height),
    commands: path.commands.map((command) => {
      if (command.command.case === "moveTo") return { moveTo: { x: Number(command.command.value.x), y: Number(command.command.value.y) } };
      if (command.command.case === "lineTo") return { lineTo: { x: Number(command.command.value.x), y: Number(command.command.value.y) } };
      if (command.command.case === "cubicBezierTo") return {
        cubicBezTo: {
          x1: Number(command.command.value.control1.x),
          y1: Number(command.command.value.control1.y),
          x2: Number(command.command.value.control2.x),
          y2: Number(command.command.value.control2.y),
          x: Number(command.command.value.end.x),
          y: Number(command.command.value.end.y),
        },
      };
      return { close: {} };
    }),
  }));
}

function modelPlaceholderTransform(frame) {
  return modelPresentationTransform(frame);
}

function slidePlaceholderSnapshot(shape) {
  return JSON.stringify(shape.layoutJson());
}

function modelPresentationChartColor(color) {
  if (!color) return undefined;
  if (color.source?.case !== "rgb") throw new OpenChestnutCodecError("Presentation chart contains a non-RGB color outside the bounded chart slice.", [], { code: "invalid_presentation_chart" });
  return `#${color.source.value}`;
}

function modelPresentationChartLine(line) {
  if (!line) return undefined;
  const style = PRESENTATION_CHART_LINE_STYLES_FROM_WIRE.get(line.dashStyle);
  if (!style) throw new OpenChestnutCodecError("Presentation chart contains an unsupported line style.", [], { code: "invalid_presentation_chart" });
  const presentationStyle = { dashed: "dash", dotted: "dot", "dash-dot": "dashDot", "dash-dot-dot": "longDashDotDot" }[style] || style;
  return {
    ...(line.color ? { fill: modelPresentationChartColor(line.color) } : {}),
    style: presentationStyle,
    ...(line.widthPoints === undefined ? {} : { width: line.widthPoints }),
  };
}

function modelPresentationChartMarker(marker) {
  if (!marker) return undefined;
  const symbol = PRESENTATION_CHART_MARKERS_FROM_WIRE.get(marker.symbol);
  if (!symbol) throw new OpenChestnutCodecError("Presentation chart contains an unsupported series marker.", [], { code: "invalid_presentation_chart" });
  return {
    symbol,
    ...(marker.size === undefined ? {} : { size: marker.size }),
    ...(marker.fill ? { fill: modelPresentationChartColor(marker.fill) } : {}),
    ...(marker.line ? { line: modelPresentationChartLine(marker.line) } : {}),
  };
}

function modelPresentationChartAxis(axis, category) {
  if (!axis) return undefined;
  return {
    title: axis.title || "",
    ...(axis.numberFormatCode ? { numberFormatCode: axis.numberFormatCode } : {}),
    ...(category && axis.tickLabelInterval !== undefined ? { tickLabelInterval: axis.tickLabelInterval } : {}),
    ...(!category && axis.minimum !== undefined ? { min: axis.minimum } : {}),
    ...(!category && axis.maximum !== undefined ? { max: axis.maximum } : {}),
    ...(!category && axis.majorUnit !== undefined ? { majorUnit: axis.majorUnit } : {}),
  };
}

function modelPresentationChartDataLabels(labels) {
  if (!labels) return undefined;
  const position = labels.position === undefined ? undefined : PRESENTATION_CHART_LABEL_POSITIONS_FROM_WIRE.get(labels.position);
  if (labels.position !== undefined && !position) throw new OpenChestnutCodecError("Presentation chart contains an unsupported data-label position.", [], { code: "invalid_presentation_chart" });
  return {
    showValue: Boolean(labels.showValue),
    showCategoryName: Boolean(labels.showCategoryName),
    ...(labels.showSeriesName === undefined ? {} : { showSeriesName: Boolean(labels.showSeriesName) }),
    ...(position ? { position } : {}),
  };
}

function modelPresentationChart(source) {
  const chartType = PRESENTATION_CHART_TYPES_FROM_WIRE.get(source.type);
  if (!chartType) throw new OpenChestnutCodecError("Presentation chart contains an unsupported chart type.", [], { code: "invalid_presentation_chart" });
  const axes = chartType === "pie" ? undefined : {
    category: modelPresentationChartAxis(source.xAxis, true) || { title: "" },
    value: modelPresentationChartAxis(source.yAxis, false) || { title: "" },
  };
  return {
    chartType,
    position: {
      left: Number(source.leftEmu) / EMU_PER_PIXEL,
      top: Number(source.topEmu) / EMU_PER_PIXEL,
      width: Number(source.widthEmu) / EMU_PER_PIXEL,
      height: Number(source.heightEmu) / EMU_PER_PIXEL,
    },
    title: source.title,
    categories: [...source.categories],
    series: source.series.map((series) => ({
      name: series.name,
      values: [...series.values],
      ...(series.fill ? { fill: modelPresentationChartColor(series.fill) } : {}),
      ...(series.line ? { line: modelPresentationChartLine(series.line) } : {}),
      ...(series.marker ? { marker: modelPresentationChartMarker(series.marker) } : {}),
    })),
    hasLegend: source.hasLegend,
    ...(axes ? { axes } : {}),
    ...(source.dataLabels ? { dataLabels: modelPresentationChartDataLabels(source.dataLabels) } : {}),
  };
}

export async function presentationFromEnvelope(envelope) {
  if (envelope.family !== ArtifactFamily.PRESENTATION || envelope.payload.case !== "presentation") {
    throw new OpenChestnutCodecError("OpenChestnut response does not contain a presentation artifact.", [], { code: "invalid_presentation_artifact" });
  }
  const source = envelope.payload.value;
  const nativeGraph = await materializePresentationNativeGraphs(envelope);
  const assetCatalog = createPresentationAssetCatalog(envelope.assets || []);
  const presentation = Presentation.create({
    slideSize: { width: Number(source.slideWidthEmu) / EMU_PER_PIXEL, height: Number(source.slideHeightEmu) / EMU_PER_PIXEL },
  });
  presentation.id = source.id || presentation.id;
  const masterStates = [];
  if (source.masters?.length) {
    presentation.masters.items.length = 0;
    for (const sourceMaster of source.masters) {
      const model = presentation.masters.add({
        id: sourceMaster.id,
        name: sourceMaster.name,
        ...(sourceMaster.background ? { background: modelBackground(sourceMaster.background) } : {}),
        placeholders: (sourceMaster.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog)),
        textParagraphStyles: modelMasterTextStyles(sourceMaster, assetCatalog),
      });
      if (!sourceMaster.background) model.background = undefined;
      for (let index = 0; index < sourceMaster.placeholders.length; index += 1) {
        if (!sourceMaster.placeholders[index].directFrame) model.placeholders[index].position = undefined;
      }
      masterStates.push({
        wire: sourceMaster,
        model,
        snapshot: masterReadOnlySnapshot(model),
      });
    }
  }
  const layoutStates = [];
  for (const sourceLayout of source.layouts || []) {
    const model = presentation.layouts.add({
      id: sourceLayout.id,
      name: sourceLayout.name,
      type: sourceLayout.type,
      masterId: sourceLayout.masterId,
      ...(sourceLayout.background ? { background: modelBackground(sourceLayout.background) } : {}),
      placeholders: (sourceLayout.placeholders || []).map((placeholder) => modelPlaceholder(placeholder, assetCatalog)),
    });
    layoutStates.push({
      wire: sourceLayout,
      model,
      snapshot: layoutReadOnlySnapshot(model),
    });
  }
  const slideStates = [];
  for (const sourceSlide of source.slides) {
    const slide = presentation.slides.add({ name: sourceSlide.name });
    slide.id = sourceSlide.id || slide.id;
    slide.layoutId = sourceSlide.layoutId || undefined;
    slide.addNotes(sourceSlide.speakerNotes?.text || "");
    const entries = [];
    for (const element of sourceSlide.elements) {
      let model;
      if (element.content.case === "shape") {
        const shape = element.content.value;
        const placeholderIdentity = shape.placeholder;
        const layout = placeholderIdentity ? presentation.layouts.getItem(slide.layoutId) : undefined;
        // PowerPoint resolves slide placeholders against their linked layout by
        // idx. Type remains descriptive and deliberately does not participate
        // in this lookup.
        const inheritedPlaceholder = placeholderIdentity?.inheritsGeometry
          ? layout?.effectivePlaceholders().find((candidate) => candidate.idx === Number(placeholderIdentity.index))
          : undefined;
        const directFrame = shape.directFrame ? modelPlaceholderFrame(shape.directFrame) : undefined;
        const directTransform = shape.directFrame ? modelPlaceholderTransform(shape.directFrame) : undefined;
        const effectiveFrame = directFrame || inheritedPlaceholder?.position || {
          left: Number(shape.leftEmu) / EMU_PER_PIXEL,
          top: Number(shape.topEmu) / EMU_PER_PIXEL,
          width: Number(shape.widthEmu) / EMU_PER_PIXEL,
          height: Number(shape.heightEmu) / EMU_PER_PIXEL,
        };
        const geometrySource = directFrame
          ? "slide"
          : placeholderIdentity?.inheritsGeometry
            ? (inheritedPlaceholder?.geometrySource || "unresolved")
            : placeholderIdentity
              ? "slide-unrecognized"
              : undefined;
        const effectiveTransform = directFrame
          ? directTransform
          : placeholderIdentity
            ? inheritedPlaceholder?.transform
            : modelPresentationTransform(shape.transform);
        model = slide.shapes.add({
          id: element.id,
          name: element.name || inheritedPlaceholder?.name,
          geometry: shape.geometry || "rect",
          ...(shape.customPaths?.length ? { customPaths: modelCustomGeometryPaths(shape) } : {}),
          position: { ...effectiveFrame },
          ...(effectiveTransform && Object.keys(effectiveTransform).length ? { transform: effectiveTransform } : {}),
          ...(placeholderIdentity ? { placeholder: {
            layoutId: slide.layoutId,
            type: placeholderIdentity.type,
            idx: Number(placeholderIdentity.index),
            geometrySource,
          } } : {}),
          fill: shape.fillRgb ? `#${shape.fillRgb}` : "transparent",
          line: { fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent", width: Number(shape.lineWidthEmu) / EMU_PER_POINT },
          ...(shape.shadow ? { shadow: modelPresentationShadow(shape.shadow) } : {}),
          text: modelText(shape, assetCatalog),
          textBodyProperties: modelTextBodyProperties(shape),
        });
        model.text.inheritedParagraphStyles = modelListStyles(shape, assetCatalog);
      } else if (element.content.case === "image") {
        const image = element.content.value;
        model = slide.images.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(image.leftEmu) / EMU_PER_PIXEL,
            top: Number(image.topEmu) / EMU_PER_PIXEL,
            width: Number(image.widthEmu) / EMU_PER_PIXEL,
            height: Number(image.heightEmu) / EMU_PER_PIXEL,
          },
          alt: image.altText,
          dataUrl: assetCatalog.dataUrl(image.assetId),
          fit: "stretch",
          geometry: "rect",
          ...(image.transform ? { transform: modelPresentationTransform(image.transform) } : {}),
        });
      } else if (element.content.case === "table") {
        const table = element.content.value;
        model = slide.tables.add({
          id: element.id,
          name: element.name,
          position: {
            left: Number(table.leftEmu) / EMU_PER_PIXEL,
            top: Number(table.topEmu) / EMU_PER_PIXEL,
            width: Number(table.widthEmu) / EMU_PER_PIXEL,
            height: Number(table.heightEmu) / EMU_PER_PIXEL,
          },
          values: table.rows.map((row) => row.cells.map((cell) => cell.text)),
          rows: table.rows.length,
          columns: table.columnWidthsEmu.length,
          styleOptions: {
            headerRow: table.firstRow === true,
            bandedRows: table.bandedRows === true,
          },
        });
      } else if (element.content.case === "connector") {
        const connector = element.content.value;
        model = slide.connectors.add({
          id: element.id,
          name: element.name,
          connectorType: connector.connectorType || "straight",
          start: { x: Number(connector.startXEmu) / EMU_PER_PIXEL, y: Number(connector.startYEmu) / EMU_PER_PIXEL },
          end: { x: Number(connector.endXEmu) / EMU_PER_PIXEL, y: Number(connector.endYEmu) / EMU_PER_PIXEL },
          startTargetId: connector.startTargetId || undefined,
          endTargetId: connector.endTargetId || undefined,
          line: {
            fill: connector.lineRgb ? `#${connector.lineRgb}` : "transparent",
            width: Number(connector.lineWidthEmu) / EMU_PER_POINT,
            ...(connector.startArrow ? { startArrow: connector.startArrow } : {}),
            ...(connector.endArrow ? { endArrow: connector.endArrow } : {}),
          },
        });
      } else if (element.content.case === "chart") {
        model = slide.charts.add(element.content.value.type === SpreadsheetChartType.BAR ? "bar" : element.content.value.type === SpreadsheetChartType.LINE ? "line" : "pie", {
          id: element.id,
          name: element.name,
          ...modelPresentationChart(element.content.value),
        });
      } else if (element.content.case === "opaque") {
        const opaque = element.content.value;
        const sourcePart = sourceSlide.source?.partPath;
        model = slide.nativeObjects.add({
          id: element.id,
          name: element.name,
          nativeKind: opaque.nativeKind || presentationNativeKind(opaque.elementName),
          position: {
            left: Number(opaque.leftEmu) / EMU_PER_PIXEL,
            top: Number(opaque.topEmu) / EMU_PER_PIXEL,
            width: Number(opaque.widthEmu) / EMU_PER_PIXEL,
            height: Number(opaque.heightEmu) / EMU_PER_PIXEL,
          },
          rawXml: opaque.rawXml,
          sourcePart,
          editable: false,
          ...(opaque.oleWorkbook ? { oleWorkbook: {
            partPath: opaque.oleWorkbook.partPath,
            contentType: opaque.oleWorkbook.contentType,
            sourceSha256: opaque.oleWorkbook.sourceSha256,
            relationshipId: opaque.oleWorkbook.relationshipId,
          } } : {}),
          ...nativeGraph(opaque, sourcePart),
        });
      } else {
        throw new OpenChestnutCodecError(`Presentation element ${element.id} has no supported wire content.`, [], { code: "invalid_presentation_artifact" });
      }
      entries.push({
        wire: element,
        model,
        placeholderSnapshot: element.content.case === "shape" && element.content.value.placeholder
          ? slidePlaceholderSnapshot(model)
          : undefined,
        snapshot: element.content.case === "opaque"
          ? opaquePresentationSnapshot(model)
          : element.content.case === "image"
            ? presentationImageReadOnlySnapshot(model)
            : element.content.case === "table"
              ? presentationTableReadOnlySnapshot(model)
            : undefined,
      });
    }
    slideStates.push({ wire: sourceSlide, name: slide.name, entries });
  }
  Object.defineProperty(presentation, PRESENTATION_STATE, {
    configurable: true,
    value: {
      source: envelope.source,
      opaqueOpc: envelope.opaqueOpc,
      diagnostics: envelope.diagnostics,
      name: source.name,
      slideWidthEmu: source.slideWidthEmu,
      slideHeightEmu: source.slideHeightEmu,
      advancedSnapshot: presentationAdvancedSnapshot(presentation),
      masters: masterStates,
      layouts: layoutStates,
      slides: slideStates,
    },
    writable: true,
  });
  return presentation;
}
