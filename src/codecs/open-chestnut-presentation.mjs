import { createHash } from "node:crypto";
import { Presentation } from "../index.mjs";
import { ArtifactFamily } from "../generated/open_office/artifact/v1/office_artifact_pb.js";
import { normalizePresentationRunLink } from "../presentation/ooxml-hyperlinks.mjs";
import { normalizePresentationTextBodyProperties } from "../presentation/text-body-properties.mjs";
import { isPresentationAutoNumberType, normalizePresentationParagraphs, normalizePresentationParagraphStyles } from "../presentation/text-paragraphs.mjs";
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
const PRESENTATION_STATE = Symbol.for("open-office-artifact-tool.open-chestnut-presentation-state");
const PRESENTATION_SCHEME_COLORS = new Set([
  "dk1", "lt1", "dk2", "lt2", "tx1", "bg1", "tx2", "bg2",
  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6", "hlink", "folHlink",
]);
const RUN_STYLE_KEYS = new Set(["bold", "italic", "fontSize", "fontFamily", "color"]);
const PARAGRAPH_KEYS = new Set([
  "runs", "level", "alignment", "style", "bulletCharacter", "autoNumber", "bulletImage", "bulletNone",
  "bulletFont", "bulletFontFollowText", "bulletColor", "bulletColorFollowText",
  "bulletSize", "bulletSizePercent", "bulletSizeFollowText", "tabStops", "marginLeft", "indent",
  "lineSpacing", "spaceBefore", "spaceBeforePercent", "spaceAfter", "spaceAfterPercent",
]);

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
  const raw = typeof value === "string" ? value : value?.color || value?.fill;
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
  const textStyleUnsupported = unsupportedStyleFields(shape.text?.style);
  if (textStyleUnsupported.length) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported text-frame style fields: ${textStyleUnsupported.join(", ")}.`, [], { code: "unsupported_presentation_features" });
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
    paragraphs: paragraphs.map((paragraph, index) => wireParagraph(paragraph, shape.text?.style || {}, original?.textBody?.paragraphs?.[index], shape.id, assetCatalog)),
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

function sourceBackground(model, entry, ownerId) {
  const current = model.background;
  if (!current) {
    if (model._backgroundClearRequested && !entry.wire.background && entry.wire.source?.backgroundEditable === false) {
      throw new OpenChestnutCodecError(`Presentation ${ownerId} background is preserved but not safely removable by this codec slice.`, [], { code: "unsupported_presentation_edit" });
    }
    return undefined;
  }
  if (!entry.wire.background && JSON.stringify(current) === entry.backgroundSnapshot) return undefined;
  return wireBackground(current, ownerId);
}

function placeholderShape(placeholder) {
  return {
    id: placeholder.id,
    text: {
      paragraphs: normalizePresentationParagraphs(placeholder.text ?? ""),
      style: { ...(placeholder.style || {}) },
      inheritedParagraphStyles: normalizePresentationParagraphStyles(placeholder.paragraphStyles || {}),
      bodyProperties: normalizePresentationTextBodyProperties(placeholder.textBodyProperties || {}),
    },
  };
}

function wirePlaceholder(placeholder, original, assetCatalog) {
  const shape = placeholderShape(placeholder);
  return {
    id: original.id,
    name: original.name,
    type: original.type,
    index: original.index,
    textBody: presentationTextBody(shape, { textBody: original.textBody }, assetCatalog),
    source: original.source,
    ...(original.directFrame ? {
      directFrame: {
        leftEmu: emuFromPixels(placeholder.position?.left, `${placeholder.id}.position.left`),
        topEmu: emuFromPixels(placeholder.position?.top, `${placeholder.id}.position.top`),
        widthEmu: emuFromPixels(placeholder.position?.width, `${placeholder.id}.position.width`),
        heightEmu: emuFromPixels(placeholder.position?.height, `${placeholder.id}.position.height`),
      },
    } : {}),
  };
}

function placeholderReadOnlySnapshot(placeholder, { directFrameEditable = false } = {}) {
  const { text: _text, paragraphStyles: _paragraphStyles, textBodyProperties: _textBodyProperties, ...rest } = placeholder;
  const { position: _position, ...withoutPosition } = rest;
  const readOnly = directFrameEditable ? withoutPosition : rest;
  return JSON.stringify(readOnly);
}

function sourcePlaceholders(placeholders, entries, ownerId, assetCatalog) {
  if (placeholders.length !== entries.length || entries.some((entry, index) => placeholders[index] !== entry.model)) {
    throw new OpenChestnutCodecError(`Source-preserving PPTX export requires ${ownerId}'s original ${entries.length}-placeholder topology.`, [], { code: "presentation_placeholder_topology_changed" });
  }
  return entries.map((entry) => {
    if (placeholderReadOnlySnapshot(entry.model, { directFrameEditable: entry.directFrameEditable }) !== entry.snapshot) {
      throw new OpenChestnutCodecError(`Presentation placeholder ${entry.model.id} can edit only local text/paragraph/body properties and an already-present recognized direct frame in this codec slice.`, [], { code: "unsupported_presentation_edit" });
    }
    return wirePlaceholder(entry.model, entry.wire, assetCatalog);
  });
}

function masterReadOnlySnapshot(master) {
  return JSON.stringify({
    id: master.id,
    name: master.name,
    theme: master.theme?.toJSON(),
  });
}

function layoutReadOnlySnapshot(layout) {
  const { background: _background, placeholders: _placeholders, ...readOnly } = layout.toJSON();
  return JSON.stringify(readOnly);
}

function presentationMasters(presentation, state, assetCatalog) {
  if (state) {
    if (presentation.masters.items.length !== state.masters.length || state.masters.some((entry, index) => presentation.masters.items[index] !== entry.model)) {
      throw new OpenChestnutCodecError(`Source-preserving PPTX export requires the original ${state.masters.length}-master topology.`, [], { code: "presentation_master_topology_changed" });
    }
    return state.masters.map((entry) => {
      if (masterReadOnlySnapshot(entry.model) !== entry.snapshot) {
        throw new OpenChestnutCodecError(`Presentation master ${entry.model.id} changed outside the bounded textParagraphStyles slice.`, [], { code: "unsupported_presentation_edit" });
      }
      return {
        id: entry.wire.id,
        name: entry.wire.name,
        source: entry.wire.source,
        textStyles: wireMasterTextStyles(entry.model, entry.wire, assetCatalog),
        background: sourceBackground(entry.model, entry, `master ${entry.model.id}`),
        placeholders: sourcePlaceholders(entry.model.placeholders, entry.placeholders, `master ${entry.model.id}`, assetCatalog),
      };
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
      throw new OpenChestnutCodecError(`Presentation layout ${entry.model.id} is preserved but not editable by this codec slice.`, [], { code: "unsupported_presentation_edit" });
    }
    return {
      id: entry.wire.id,
      name: entry.wire.name,
      masterId: entry.wire.masterId,
      type: entry.wire.type,
      source: entry.wire.source,
      background: sourceBackground(entry.model, entry, `layout ${entry.model.id}`),
      placeholders: sourcePlaceholders(entry.model.placeholders, entry.placeholders, `layout ${entry.model.id}`, assetCatalog),
    };
  });
}

function presentationShape(shape, original, assetCatalog) {
  const originalShape = original?.content?.case === "shape" ? original.content.value : original;
  if (!new Set(["rect", "ellipse"]).has(shape.geometry)) {
    throw new OpenChestnutCodecError(`Presentation shape ${shape.id} uses unsupported geometry ${shape.geometry}.`, [], { code: "unsupported_presentation_features" });
  }
  const position = shape.position || {};
  const lineWidth = Number(shape.line?.width ?? 1);
  if (!Number.isFinite(lineWidth) || lineWidth < 0) throw new OpenChestnutCodecError(`Presentation shape ${shape.id} has an invalid line width.`, [], { code: "invalid_presentation_frame" });
  const textBody = presentationTextBody(shape, originalShape, assetCatalog);
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
      },
    },
  };
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

function unsupportedPresentationFeatures(presentation) {
  const unsupported = [];
  if (presentation.layouts?.items?.length) unsupported.push("slide layouts");
  if (presentation.masters?.items?.length !== 1) unsupported.push("multiple slide masters");
  const master = presentation.master;
  if (master?.theme) unsupported.push("master theme override");
  if (master?.placeholders?.length) unsupported.push("master placeholders");
  if (presentation.customShows?.items?.length) unsupported.push("custom shows");
  if (presentation.commentFormat !== "legacy") unsupported.push("modern comments");
  for (const slide of presentation.slides?.items || []) {
    const prefix = `slide ${slide.index + 1}`;
    if (slide.layoutId) unsupported.push(`${prefix} layout binding`);
    if (slide.speakerNotes?.text) unsupported.push(`${prefix} speaker notes`);
    if (slide.background?.fill) unsupported.push(`${prefix} background`);
    if (slide.comments?.items?.length) unsupported.push(`${prefix} comments`);
    if (slide.tables?.items?.length) unsupported.push(`${prefix} tables`);
    if (slide.charts?.items?.length) unsupported.push(`${prefix} charts`);
    if (slide.images?.items?.length) unsupported.push(`${prefix} images`);
    if (slide.connectors?.items?.length) unsupported.push(`${prefix} connectors`);
    if (slide.groups?.items?.length) unsupported.push(`${prefix} groups`);
    if (slide.nativeObjects?.items?.length) unsupported.push(`${prefix} native objects`);
  }
  return unsupported;
}

function opaquePresentationSnapshot(object, { includePlacement = true } = {}) {
  const oleWorkbook = object.oleWorkbook ? {
    partPath: object.oleWorkbook.partPath,
    contentType: object.oleWorkbook.contentType,
    sourceSha256: object.oleWorkbook.sourceSha256,
    relationshipId: object.oleWorkbook.relationshipId,
  } : undefined;
  return JSON.stringify({
    id: object.id,
    ...(includePlacement ? { name: object.name, position: object.position } : {}),
    nativeKind: object.nativeKind,
    rawXml: object.rawXml,
    oleWorkbook,
    ...presentationNativeGraphSnapshot(object, { ignoredPartPaths: oleWorkbook ? [oleWorkbook.partPath] : [] }),
  });
}

function presentationOpaqueElement(object, original, assetCatalog) {
  const name = String(object.name ?? "");
  if (name.length > 1_024) throw new OpenChestnutCodecError(`Presentation native object ${object.id} name exceeds 1024 characters.`, [], { code: "invalid_presentation_native_object" });
  const position = object.position || {};
  const sourceOleWorkbook = original.content.value.oleWorkbook;
  let oleWorkbook = sourceOleWorkbook;
  if (sourceOleWorkbook) {
    const metadata = object.oleWorkbook;
    if (!metadata || metadata.partPath !== sourceOleWorkbook.partPath || metadata.contentType !== sourceOleWorkbook.contentType ||
        metadata.sourceSha256 !== sourceOleWorkbook.sourceSha256 || metadata.relationshipId !== sourceOleWorkbook.relationshipId) {
      throw new OpenChestnutCodecError(`Presentation OLE workbook ${object.id} changed its source binding.`, [], { code: "presentation_ole_workbook_binding_mismatch" });
    }
    const matches = (object.parts || []).filter((part) => part.path === metadata.partPath);
    if (matches.length !== 1 || matches[0].contentType !== metadata.contentType) {
      throw new OpenChestnutCodecError(`Presentation OLE workbook ${object.id} no longer resolves to one source-bound XLSX part.`, [], { code: "presentation_ole_workbook_binding_mismatch" });
    }
    const digest = createHash("sha256").update(matches[0].bytes).digest("hex");
    oleWorkbook = {
      ...sourceOleWorkbook,
      replacementAssetId: digest === metadata.sourceSha256 ? "" : assetCatalog.addOleWorkbook(matches[0].bytes),
    };
  }
  return {
    id: original.id,
    name,
    source: original.source,
    content: {
      case: "opaque",
      value: {
        ...original.content.value,
        ...(oleWorkbook ? { oleWorkbook } : {}),
        leftEmu: emuFromPixels(position.left, `${object.id}.position.left`),
        topEmu: emuFromPixels(position.top, `${object.id}.position.top`),
        widthEmu: emuFromPixels(position.width, `${object.id}.position.width`),
        heightEmu: emuFromPixels(position.height, `${object.id}.position.height`),
      },
    },
  };
}

export function presentationEnvelope(presentation, protocolVersion) {
  if (!(presentation instanceof Presentation)) throw new TypeError("exportPptxWithOpenChestnut expects a Presentation instance.");
  if (!presentation.slides?.items?.length) throw new OpenChestnutCodecError("Presentation must contain at least one slide.", [], { code: "missing_slides" });
  const state = presentation[PRESENTATION_STATE];
  if (!state) {
    const unsupported = unsupportedPresentationFeatures(presentation);
    if (unsupported.length) {
      throw new OpenChestnutCodecError(`The PPTX WebAssembly vertical slice cannot author: ${unsupported.slice(0, 8).join(", ")}${unsupported.length > 8 ? `, and ${unsupported.length - 8} more` : ""}. Use PresentationFile.exportPptx until parity reaches these features.`, [], { code: "unsupported_presentation_features" });
    }
  } else {
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
    }
    return {
      id: sourceState?.wire.id || slide.id,
      name: slide.name,
      source: sourceState?.wire.source,
      ...(slide.layoutId ? { layoutId: slide.layoutId } : {}),
      elements: sourceState
          ? sourceState.entries.map((entry) => {
            if (entry.wire.content.case === "shape") return presentationShape(entry.model, entry.wire, assetCatalog);
            const placementEditable = entry.wire.source?.editable === true;
            if (opaquePresentationSnapshot(entry.model, { includePlacement: !placementEditable }) !== entry.snapshot) throw new OpenChestnutCodecError(`Presentation element ${entry.model.id} changed outside its ${placementEditable ? "name/frame" : "read-only"} native-object boundary.`, [], { code: "unsupported_presentation_edit" });
            if (placementEditable) return presentationOpaqueElement(entry.model, entry.wire, assetCatalog);
            return entry.wire;
          })
        : slide.shapes.items.map((shape) => presentationShape(shape, undefined, assetCatalog)),
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
  return {
    id: source.id,
    name: source.name,
    type: source.type,
    idx: source.index,
    ...(source.directFrame ? {
      position: {
        left: Number(source.directFrame.leftEmu) / EMU_PER_PIXEL,
        top: Number(source.directFrame.topEmu) / EMU_PER_PIXEL,
        width: Number(source.directFrame.widthEmu) / EMU_PER_PIXEL,
        height: Number(source.directFrame.heightEmu) / EMU_PER_PIXEL,
      },
    } : {}),
    text: modelText(shape, assetCatalog),
    paragraphStyles: modelListStyles(shape, assetCatalog),
    textBodyProperties: modelTextBodyProperties(shape),
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
        backgroundSnapshot: JSON.stringify(model.background),
        placeholders: (sourceMaster.placeholders || []).map((wire, index) => ({
          wire,
          model: model.placeholders[index],
          directFrameEditable: Boolean(wire.directFrame),
          snapshot: placeholderReadOnlySnapshot(model.placeholders[index], { directFrameEditable: Boolean(wire.directFrame) }),
        })),
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
      backgroundSnapshot: JSON.stringify(model.background),
      placeholders: (sourceLayout.placeholders || []).map((wire, index) => ({
        wire,
        model: model.placeholders[index],
        directFrameEditable: Boolean(wire.directFrame),
        snapshot: placeholderReadOnlySnapshot(model.placeholders[index], { directFrameEditable: Boolean(wire.directFrame) }),
      })),
    });
  }
  const slideStates = [];
  for (const sourceSlide of source.slides) {
    const slide = presentation.slides.add({ name: sourceSlide.name });
    slide.id = sourceSlide.id || slide.id;
    slide.layoutId = sourceSlide.layoutId || undefined;
    const entries = [];
    for (const element of sourceSlide.elements) {
      let model;
      if (element.content.case === "shape") {
        const shape = element.content.value;
        model = slide.shapes.add({
          id: element.id,
          name: element.name,
          geometry: shape.geometry || "rect",
          position: {
            left: Number(shape.leftEmu) / EMU_PER_PIXEL,
            top: Number(shape.topEmu) / EMU_PER_PIXEL,
            width: Number(shape.widthEmu) / EMU_PER_PIXEL,
            height: Number(shape.heightEmu) / EMU_PER_PIXEL,
          },
          fill: shape.fillRgb ? `#${shape.fillRgb}` : "transparent",
          line: { fill: shape.lineRgb ? `#${shape.lineRgb}` : "transparent", width: Number(shape.lineWidthEmu) / EMU_PER_POINT },
          text: modelText(shape, assetCatalog),
          textBodyProperties: modelTextBodyProperties(shape),
        });
        model.text.inheritedParagraphStyles = modelListStyles(shape, assetCatalog);
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
          editable: element.source?.editable === true,
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
        snapshot: element.content.case === "opaque"
          ? opaquePresentationSnapshot(model, { includePlacement: element.source?.editable !== true })
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
      masters: masterStates,
      layouts: layoutStates,
      slides: slideStates,
    },
    writable: true,
  });
  return presentation;
}
