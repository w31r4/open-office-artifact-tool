import { inspectOoxmlPackage, ooxmlResolveRelationshipTarget, ooxmlSafePartPath, patchOoxmlPackage } from "../ooxml/package.mjs";
import { validatePptxPackageSemantics } from "../ooxml/pptx-package-semantics.mjs";
import { queryHelpRecords } from "../help/index.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { resolveColorToken } from "../shared/colors.mjs";
import { aid } from "../shared/ids.mjs";
import { imageDataFromDataUrl } from "../shared/images.mjs";
import { filterInspectRecords, inspectRecordMatchesTarget, inspectTargetTokens, ndjson, normalizeKinds, verificationIssue, verificationResult } from "../shared/inspection.mjs";
import { LAYOUT_MIME } from "../shared/render-output.mjs";
import { attrEscape, xmlEscape } from "../shared/xml.mjs";
import { createTextRange, textRangeRecord } from "../shared/text-range.mjs";
import { materializeComposeNode } from "./compose.mjs";
import { normalizePresentationThemeConfig } from "./ooxml-theme.mjs";
import { mergePresentationPlaceholders, normalizePresentationBackground, resolvePresentationBackgroundColor } from "./ooxml-masters.mjs";
import { createPresentationGroupShapeClass } from "./group-shapes.mjs";
import { createNativePresentationObjectClass } from "./native-objects.mjs";
import { normalizePresentationChartAxisGroup, normalizePresentationChartDataLabels, normalizePresentationChartErrorBars, normalizePresentationChartSeriesStyle, normalizePresentationChartStyle, normalizePresentationChartTrendlines } from "./ooxml-charts.mjs";
import { normalizePresentationChartExternalData, presentationChartUsesFormulaReferences } from "./ooxml-chart-data.mjs";
import { presentationChartLineSvgAttributes, presentationChartTrendlinesSvg } from "./chart-trendline-svg.mjs";
import { planPresentationCustomShows, PresentationCustomShowCollection } from "./ooxml-custom-shows.mjs";
import { inheritPresentationParagraphs, normalizePresentationParagraphs, normalizePresentationParagraphStyles, presentationParagraphsNeedSerialization, presentationParagraphsSvg, presentationParagraphsText, replacePresentationParagraphText } from "./text-paragraphs.mjs";
import { normalizePresentationTextBodyProperties } from "./text-body-properties.mjs";
import { normalizePresentationCustomPaths, presentationCustomPathsSvg } from "./custom-geometry.mjs";
import { normalizePresentationImageCrop, normalizePresentationImageFit, presentationImageCropViewport } from "./image-crop.mjs";
import { planPresentationModernComments } from "./ooxml-modern-comments.mjs";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";

const PPTX_PACKAGE_CONFIG = {
  family: "PPTX",
  packageKind: "pptxPackage",
  partKind: "pptxPart",
  counts: { slides: /^ppt\/slides\/slide\d+\.xml$/ },
  semanticIssues: validatePptxPackageSemantics,
};

class SlideCollection {
  constructor(presentation) {
    this.presentation = presentation;
    this.items = [];
  }

  add(options = {}) {
    const slide = new Slide(this.presentation, options);
    this.items.push(slide);
    return slide;
  }

  getItem(index) { return this.items[index]; }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class PresentationTheme {
  constructor(presentation, config = {}, base = {}) {
    const normalized = normalizePresentationThemeConfig(config, base);
    this.presentation = presentation;
    this.id = config.id || "theme/default";
    this.name = normalized.name;
    this.colors = normalized.colors;
    this.fonts = normalized.fonts;
    this.textStyles = normalized.textStyles;
    this.colorMap = normalized.colorMap;
  }

  update(config = {}) {
    const normalized = normalizePresentationThemeConfig(config, this);
    Object.assign(this, normalized);
    return this;
  }

  setColors(colors = {}) { return this.update({ colors }); }
  setFonts(fonts = {}) { return this.update({ fonts }); }
  setTextStyles(textStyles = {}) { return this.update({ textStyles }); }
  setColorMap(colorMap = {}) { return this.update({ colorMap }); }
  inspectRecord() { return { kind: "theme", id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
  toJSON() { return { id: this.id, name: this.name, colors: this.colors, fonts: this.fonts, textStyles: this.textStyles, colorMap: this.colorMap }; }
}

function presentationThemeSemantics(theme) {
  const normalized = normalizePresentationThemeConfig(theme);
  return JSON.stringify({ name: normalized.name, colors: normalized.colors, fonts: normalized.fonts, textStyles: normalized.textStyles, colorMap: normalized.colorMap });
}

function normalizePresentationPlaceholderTransform(value, name = "Presentation placeholder transform") {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError(`${name} must be an object.`);
  const output = {};
  if (Object.hasOwn(value, "rotationDegrees") && value.rotationDegrees != null) {
    const degrees = Number(value.rotationDegrees);
    if (!Number.isFinite(degrees) || degrees < -360 || degrees > 360) throw new RangeError(`${name}.rotationDegrees must be between -360 and 360 degrees.`);
    output.rotationDegrees = degrees;
  }
  for (const key of ["flipHorizontal", "flipVertical"]) {
    if (!Object.hasOwn(value, key) || value[key] == null) continue;
    if (typeof value[key] !== "boolean") throw new TypeError(`${name}.${key} must be a boolean.`);
    output[key] = value[key];
  }
  if (Object.keys(output).length === 0) throw new TypeError(`${name} must define rotationDegrees, flipHorizontal, or flipVertical.`);
  return output;
}

function normalizePresentationPlaceholders(value = [], idPrefix = "placeholder", options = {}) {
  if (!Array.isArray(value)) throw new TypeError("Presentation placeholders must be an array.");
  if (value.length > 128) throw new RangeError("Presentation placeholders exceed 128 entries.");
  const placeholders = value.map((placeholder, index) => {
    const position = options.allowMissingPosition && !placeholder.position && !placeholder.frame && !["left", "top", "width", "height"].some((key) => placeholder[key] != null)
      ? undefined
      : normalizeFrame(placeholder, { left: 80, top: 80 + index * 80, width: 640, height: 64 });
    const transform = normalizePresentationPlaceholderTransform(placeholder.transform, `Presentation placeholder ${placeholder.name || index + 1} transform`);
    if (transform && !position) throw new TypeError(`Presentation placeholder ${placeholder.name || index + 1} cannot define a transform without a direct position.`);
    return {
      id: placeholder.id || `${idPrefix}/${index + 1}`,
      type: placeholder.type || "body",
      idx: Number(placeholder.idx ?? index + 1),
      name: placeholder.name || `${placeholder.type || "body"} placeholder`,
      position,
      transform,
      text: placeholder.text ?? "",
      required: Boolean(placeholder.required),
      style: { ...(placeholder.style || {}) },
      paragraphStyles: normalizePresentationParagraphStyles(placeholder.paragraphStyles || placeholder.listStyles || {}),
      textBodyProperties: normalizePresentationTextBodyProperties(placeholder.textBodyProperties || placeholder.bodyProperties || {}),
    };
  });
  if (placeholders.some((placeholder) => !Number.isInteger(placeholder.idx) || placeholder.idx < 0 || placeholder.idx > 4_294_967_295)) throw new RangeError("Presentation placeholder idx must be an unsigned 32-bit integer.");
  if (new Set(placeholders.map((placeholder) => `${placeholder.type}:${placeholder.idx}`)).size !== placeholders.length) throw new Error("Presentation placeholder type/idx pairs must be unique.");
  return placeholders;
}

function clonePresentationParagraphStyles(styles = {}) {
  return Object.fromEntries(Object.entries(styles).map(([level, style]) => [Number(level), { ...style, style: { ...(style.style || {}) } }]));
}

function mergePresentationParagraphStyles(base = {}, overrides = {}) {
  const result = clonePresentationParagraphStyles(base);
  for (const [level, style] of Object.entries(overrides || {})) {
    const inherited = { ...(result[Number(level)] || {}) };
    if (["bulletCharacter", "bulletImage", "autoNumber", "bulletNone"].some((field) => Object.hasOwn(style, field))) {
      delete inherited.bulletCharacter;
      delete inherited.bulletImage;
      delete inherited.autoNumber;
      delete inherited.bulletNone;
    }
    for (const fields of [["bulletFont", "bulletFontFollowText"], ["bulletColor", "bulletColorFollowText"], ["bulletSize", "bulletSizePercent", "bulletSizeFollowText"]]) {
      if (!fields.some((field) => Object.hasOwn(style, field))) continue;
      for (const field of fields) delete inherited[field];
    }
    result[Number(level)] = { ...inherited, ...style, style: { ...(inherited.style || {}), ...(style.style || {}) } };
  }
  return result;
}

function normalizePresentationMasterParagraphStyles(value = {}) {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation master textParagraphStyles must be an object.");
  return Object.fromEntries(["title", "body", "other"].map((kind) => [kind, normalizePresentationParagraphStyles(value[kind] || {})]));
}

function presentationPlaceholderTextStyleKind(type = "body") {
  if (["title", "ctrTitle"].includes(type)) return "title";
  if (["body", "subTitle", "obj", "chart", "tbl", "clipArt", "dgm", "media", "pic"].includes(type)) return "body";
  return "other";
}

class PresentationSlideMaster {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.configured = Object.keys(config).length > 0;
    this.id = config.id || "master/default";
    this.name = config.name || "Default Master";
    this.theme = config.theme ? new PresentationTheme(presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, presentation.theme) : undefined;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = Object.hasOwn(config, "background")
      ? normalizePresentationBackground(config.background)
      : normalizePresentationBackground(presentation.theme.colors.bg1);
    this.placeholders = normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`);
    this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles || {});
  }

  update(config = {}) {
    if (Object.keys(config).length > 0) this.configured = true;
    const previousId = this.id;
    if (config.id) this.id = String(config.id);
    if (this.theme?.id === `${previousId}/theme`) this.theme.id = `${this.id}/theme`;
    if (config.name) this.name = String(config.name);
    if (Object.hasOwn(config, "theme")) this.theme = config.theme ? new PresentationTheme(this.presentation, { ...config.theme, id: config.theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined;
    if (Object.hasOwn(config, "background")) {
      this.background = config.background == null ? undefined : normalizePresentationBackground(config.background, this.background);
      this._backgroundClearRequested = false;
    }
    if (config.placeholders) this.placeholders = normalizePresentationPlaceholders(config.placeholders, `${this.id}/ph`);
    if (config.textParagraphStyles) this.textParagraphStyles = normalizePresentationMasterParagraphStyles(config.textParagraphStyles);
    return this;
  }

  setBackground(background) { this.configured = true; this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.configured = true; this.background = undefined; this._backgroundClearRequested = true; return this; }
  setTheme(theme) { this.configured = true; this.theme = theme ? new PresentationTheme(this.presentation, { ...theme, id: theme.id || `${this.id}/theme` }, this.presentation.theme) : undefined; return this; }
  effectiveTheme() { return this.theme || this.presentation.theme; }
  effectiveBackground() { return this.background || normalizePresentationBackground(this.effectiveTheme().colors.bg1, "#ffffff"); }
  paragraphStylesForPlaceholder(type) { return this.textParagraphStyles[presentationPlaceholderTextStyleKind(type)] || {}; }
  inspectRecord() { const theme = this.effectiveTheme(); return { kind: "slideMaster", id: this.id, name: this.name, background: this.background, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, placeholderTypes: this.placeholders.map((placeholder) => placeholder.type), textParagraphStyleLevels: Object.fromEntries(Object.entries(this.textParagraphStyles).map(([kind, styles]) => [kind, Object.keys(styles).length])), hasThemeOverride: Boolean(this.theme), themeId: theme.id, themeName: theme.name }; }
  toJSON() { return { id: this.id, name: this.name, background: this.background, theme: this.theme?.toJSON(), placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })), textParagraphStyles: normalizePresentationMasterParagraphStyles(this.textParagraphStyles) }; }
}

class PresentationSlideMasterCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) {
    if (this.items.length >= 64) throw new RangeError("Presentation masters exceed 64 entries.");
    const master = config instanceof PresentationSlideMaster ? config : new PresentationSlideMaster(this.presentation, config);
    if (this.items.some((item) => item.id === master.id)) throw new Error(`Duplicate presentation master ID ${master.id}.`);
    master.presentation = this.presentation;
    if (master.theme) master.theme.presentation = this.presentation;
    this.items.push(master);
    return master;
  }
  getItem(idOrName) { return this.items.find((master) => master.id === idOrName || master.name === idOrName); }
  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class SlideLayoutTemplate {
  constructor(presentation, config = {}) {
    this.presentation = presentation;
    this.id = config.id || aid("lo");
    this.name = config.name || "Blank";
    this.type = config.type || "blank";
    this.masterId = config.masterId || presentation.master.id;
    Object.defineProperty(this, "_backgroundClearRequested", { value: false, writable: true });
    this.background = config.background ? normalizePresentationBackground(config.background) : undefined;
    this.placeholders = normalizePresentationPlaceholders(config.placeholders || [], `${this.id}/ph`, { allowMissingPosition: true });
  }

  effectiveMaster() { return this.presentation.masters.getItem(this.masterId); }
  effectiveTheme() { return this.effectiveMaster()?.effectiveTheme() || this.presentation.theme; }
  setBackground(background) { this.background = normalizePresentationBackground(background, this.background); this._backgroundClearRequested = false; return this; }
  clearBackground() { this.background = undefined; this._backgroundClearRequested = true; return this; }
  effectivePlaceholders() {
    const master = this.effectiveMaster();
    return mergePresentationPlaceholders(master?.placeholders || [], this.placeholders).map((placeholder) => ({
      ...placeholder,
      paragraphStyles: mergePresentationParagraphStyles(master?.paragraphStylesForPlaceholder(placeholder.type), placeholder.paragraphStyles),
    }));
  }
  effectiveBackground() { return this.background || this.effectiveMaster()?.effectiveBackground() || normalizePresentationBackground(this.presentation.theme.colors.bg1, "#ffffff"); }

  apply(slide) {
    slide.layoutId = this.id;
    const placeholders = this.effectivePlaceholders();
    return placeholders.map((placeholder) => {
      const shape = slide.shapes.add({
        id: placeholder.id,
        name: placeholder.name,
        geometry: "rect",
        position: placeholder.position,
        transform: placeholder.transform,
        fill: "transparent",
        line: { fill: "transparent", width: 0 },
        text: placeholder.text,
        textBodyProperties: placeholder.textBodyProperties,
        placeholder: { layoutId: this.id, type: placeholder.type, name: placeholder.name, required: placeholder.required, idx: placeholder.idx },
      });
      shape.text.style = { ...placeholder.style };
      shape.text.inheritedParagraphStyles = Object.fromEntries(Object.entries(placeholder.paragraphStyles || {}).map(([level, style]) => [level, { ...style, style: { ...(style.style || {}) } }]));
      return shape;
    });
  }

  inspectRecord() { return { kind: "layoutTemplate", id: this.id, name: this.name, type: this.type, masterId: this.masterId, themeId: this.effectiveTheme().id, background: this.background, effectiveBackground: this.effectiveBackground(), placeholders: this.placeholders.length, effectivePlaceholders: this.effectivePlaceholders().length, placeholderTypes: this.effectivePlaceholders().map((placeholder) => placeholder.type) }; }
  toJSON() { return { id: this.id, name: this.name, type: this.type, masterId: this.masterId, background: this.background, placeholders: this.placeholders.map((placeholder) => ({ ...placeholder })) }; }
}

class SlideLayoutCollection {
  constructor(presentation) { this.presentation = presentation; this.items = []; }
  add(config = {}) { const layout = new SlideLayoutTemplate(this.presentation, config); this.items.push(layout); return layout; }
  getItem(idOrName) { return this.items.find((layout) => layout.id === idOrName || layout.name === idOrName || layout.type === idOrName); }
  inspectRecords() { return this.items.map((layout) => layout.inspectRecord()); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

function svgInner(svg = "") {
  return String(svg || "").replace(/^<svg\b[^>]*>/i, "").replace(/<\/svg>\s*$/i, "");
}

function presentationMontageSvg(presentation, options = {}) {
  const slides = presentation.slides.items.length ? presentation.slides.items : [presentation.slides.add()];
  const gap = Number(options.gap ?? 24);
  const scale = Number(options.scale ?? 0.25);
  const columns = Math.max(1, Number(options.columns ?? 1) || 1);
  const slideW = Number(presentation.slideSize.width || 1280);
  const slideH = Number(presentation.slideSize.height || 720);
  const thumbW = slideW * scale;
  const thumbH = slideH * scale;
  const labelH = 20;
  const rows = Math.ceil(slides.length / columns);
  const width = Math.max(1, columns * thumbW + (columns + 1) * gap);
  const height = Math.max(1, rows * (thumbH + labelH) + (rows + 1) * gap);
  const thumbs = slides.map((slide, index) => {
    const col = index % columns;
    const row = Math.floor(index / columns);
    const x = gap + col * (thumbW + gap);
    const y = gap + row * (thumbH + labelH + gap);
    return `<g data-slide="${index + 1}"><rect x="${x - 1}" y="${y - 1}" width="${thumbW + 2}" height="${thumbH + 2}" fill="#ffffff" stroke="#94a3b8"/><g transform="translate(${x},${y}) scale(${scale})">${svgInner(slide.toSvg())}</g><text x="${x}" y="${y + thumbH + 15}" font-family="Arial" font-size="12" fill="#475569">Slide ${index + 1}${slide.title() ? ` — ${xmlEscape(slide.title()).slice(0, 80)}` : ""}</text></g>`;
  }).join("");
  return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="#f8fafc"/>${thumbs}</svg>`;
}

export class Presentation {
  constructor(options = {}) {
    this.id = aid("pr");
    this.slideSize = options.slideSize || { width: 1280, height: 720 };
    this.commentFormat = options.commentFormat || "legacy";
    this.theme = new PresentationTheme(this, options.theme || {});
    this.masters = new PresentationSlideMasterCollection(this);
    const masterConfigs = Array.isArray(options.masters) && options.masters.length ? options.masters : [options.master || {}];
    for (const master of masterConfigs) this.masters.add(master);
    this.layouts = new SlideLayoutCollection(this);
    for (const layout of options.layouts || []) this.layouts.add(layout);
    this.slides = new SlideCollection(this);
    this.customShows = new PresentationCustomShowCollection(this);
  }

  static create(options = {}) { return new Presentation(options); }
  get master() { return this.masters.items[0]; }
  set master(value) {
    const master = value instanceof PresentationSlideMaster ? value : new PresentationSlideMaster(this, value || {});
    master.presentation = this;
    if (master.theme) master.theme.presentation = this;
    if (this.masters.items.length) this.masters.items[0] = master;
    else this.masters.items.push(master);
  }

  inspect(options = {}) {
    const kinds = normalizeKinds(options.kind, ["deck", "slide", "textbox", "shape", "nativeObject", "layout"]);
    const records = [];
    if (kinds.has("deck")) records.push({ kind: "deck", id: this.id, slides: this.slides.count, customShows: this.customShows.count });
    if (kinds.has("theme")) records.push(this.theme.inspectRecord());
    if (kinds.has("slideMaster") || kinds.has("master")) records.push(...this.masters.items.map((master) => master.inspectRecord()));
    if (kinds.has("layout") || kinds.has("layoutTemplate")) records.push(...this.layouts.inspectRecords());
    if (kinds.has("customShow")) records.push(...this.customShows.items.map((show) => show.inspectRecord()));
    for (const slide of this.slides) records.push(...slide.inspectRecords(kinds));
    return ndjson(filterInspectRecords(records, options), options.maxChars ?? Infinity);
  }

  validateLayout(options = {}) {
    const issues = this.slides.items.flatMap((slide) => slide.validateLayout(options).issues);
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  verify(options = {}) {
    const issues = [];
    if (this.slides.items.length === 0) issues.push(verificationIssue("presentation", "noSlides", "Presentation has no slides."));
    try { planPresentationCustomShows(this); }
    catch (error) { issues.push(verificationIssue("presentation", "invalidCustomShow", error.message)); }
    if (this.commentFormat === "modern" || this.slides.items.some((slide) => slide.comments.items.some((thread) => thread.nativeFormat === "modern"))) {
      try { planPresentationModernComments(this.slides.items); }
      catch (error) { issues.push(verificationIssue("presentation", "invalidModernCommentMetadata", error.message)); }
    }
    const duplicateMasterIds = this.masters.items.map((master) => master.id).filter((id, index, ids) => ids.indexOf(id) !== index);
    for (const masterId of new Set(duplicateMasterIds)) issues.push(verificationIssue("presentation", "duplicateMasterId", `Presentation contains duplicate master ID ${masterId}.`, { masterId }));
    const knownMasterIds = new Set(this.masters.items.map((master) => master.id));
    for (const layout of this.layouts.items) if (!knownMasterIds.has(layout.masterId)) issues.push(verificationIssue("presentation", "missingMaster", `Layout ${layout.name || layout.id} references missing master ${layout.masterId}.`, { id: layout.id, masterId: layout.masterId }));
    issues.push(...this.validateLayout(options).issues.map((issue) => ({ ...issue, artifactKind: "presentation" })));
    for (const slide of this.slides) {
      const slideElements = presentationSlideElements(slide);
      if (slide.layoutId && !this.layouts.getItem(slide.layoutId)) issues.push(verificationIssue("presentation", "missingLayout", `Slide ${slide.index + 1} references missing layout ${slide.layoutId}.`, { slide: slide.index + 1, layoutId: slide.layoutId }));
      for (const shape of slideElements.filter((element) => element instanceof Shape)) {
        if (shape.placeholder?.required && !shape.text.value.trim()) issues.push(verificationIssue("presentation", "placeholderMissingContent", `Required ${shape.placeholder.type || "placeholder"} placeholder ${shape.name || shape.id} on slide ${slide.index + 1} is empty.`, { slide: slide.index + 1, id: shape.id, placeholder: shape.placeholder }));
      }
      for (const table of slideElements.filter((element) => element instanceof TableElement)) {
        if (!table.rows || !table.columns || table.values.length === 0 || table.values.every((row) => row.every((cell) => String(cell ?? "").trim() === ""))) issues.push(verificationIssue("presentation", "emptyTable", `Table ${table.name || table.id} on slide ${slide.index + 1} has no visible cell data.`, { slide: slide.index + 1, id: table.id }));
        if (table.values.length !== table.rows) issues.push(verificationIssue("presentation", "tableDataMismatch", `Table ${table.name || table.id} declares ${table.rows} rows but has ${table.values.length} value rows.`, { slide: slide.index + 1, id: table.id, rows: table.rows, valueRows: table.values.length }));
        if (table.values.some((row) => row.length !== table.columns)) issues.push(verificationIssue("presentation", "raggedTableRows", `Table ${table.name || table.id} has rows that do not match its declared column count.`, { slide: slide.index + 1, id: table.id, columns: table.columns, rowLengths: table.values.map((row) => row.length) }));
      }
      for (const chart of slideElements.filter((element) => element instanceof ChartElement)) {
        if (!/^(bar|line|pie|combo)$/i.test(chart.chartType)) issues.push(verificationIssue("presentation", "unsupportedChartType", `Chart ${chart.name || chart.id} uses unsupported chart type ${chart.chartType}.`, { severity: "warning", slide: slide.index + 1, id: chart.id, chartType: chart.chartType }));
        if (!chart.series.length) issues.push(verificationIssue("presentation", "emptyChart", `Chart ${chart.name || chart.id} on slide ${slide.index + 1} has no data series.`, { slide: slide.index + 1, id: chart.id }));
        for (const series of chart.series) {
          const values = Array.isArray(series.values) ? series.values : [];
          if (chart.categories.length && values.length && chart.categories.length !== values.length) issues.push(verificationIssue("presentation", "chartDataMismatch", `Chart ${chart.name || chart.id} series ${series.name || "Series"} has ${values.length} values for ${chart.categories.length} categories.`, { slide: slide.index + 1, id: chart.id, series: series.name, values: values.length, categories: chart.categories.length }));
          if (values.some((value) => value !== "" && value != null && !Number.isFinite(Number(value)))) issues.push(verificationIssue("presentation", "chartDataNonNumeric", `Chart ${chart.name || chart.id} series ${series.name || "Series"} contains non-numeric values.`, { slide: slide.index + 1, id: chart.id, series: series.name }));
        }
      }
      for (const image of slideElements.filter((element) => element instanceof ImageElement)) {
        if (!image.dataUrl && !image.uri && !image.prompt) issues.push(verificationIssue("presentation", "emptyImage", `Image ${image.name || image.id} on slide ${slide.index + 1} has no dataUrl, uri, or prompt.`, { slide: slide.index + 1, id: image.id }));
        if (image.dataUrl && !imageDataFromDataUrl(image.dataUrl)) issues.push(verificationIssue("presentation", "invalidImageDataUrl", `Image ${image.name || image.id} on slide ${slide.index + 1} has an unsupported data URL.`, { slide: slide.index + 1, id: image.id }));
      }
      for (const object of slideElements.filter((element) => element instanceof NativePresentationObject)) {
        if (!object.rawXml) issues.push(verificationIssue("presentation", "nativeObjectMarkupMissing", `Native ${object.nativeKind} object ${object.name || object.id} on slide ${slide.index + 1} has no preserved markup.`, { slide: slide.index + 1, id: object.id, nativeKind: object.nativeKind }));
        const partPaths = new Set(object.parts.map((part) => part.path));
        const sourcePart = object.sourcePart || `ppt/slides/slide${slide.index + 1}.xml`;
        for (const relationship of object.rootRelationships) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(sourcePart, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing relationship target ${target}.`, { slide: slide.index + 1, id: object.id, relationshipId: relationship.id, target }));
        }
        for (const part of object.parts) for (const relationship of part.relationships || []) {
          if (relationship.targetMode?.toLowerCase() === "external") continue;
          const target = ooxmlSafePartPath(ooxmlResolveRelationshipTarget(part.path, relationship.target), "PPTX");
          if (!partPaths.has(target)) issues.push(verificationIssue("presentation", "nativeObjectPartMissing", `Native ${object.nativeKind} object ${object.name || object.id} is missing recursive relationship target ${target}.`, { slide: slide.index + 1, id: object.id, sourcePart: part.path, relationshipId: relationship.id, target }));
        }
      }
      for (const comment of slide.comments) {
        if (comment.targetId && !slide.resolve(comment.targetId)) issues.push(verificationIssue("presentation", "danglingComment", `Slide ${slide.index + 1} comment ${comment.id} targets missing element ${comment.targetId}.`, { slide: slide.index + 1, id: comment.id, targetId: comment.targetId }));
      }
    }
    return verificationResult("presentation", issues, options);
  }

  resolve(id) {
    if (id === this.id) return this;
    if (id === this.theme.id) return this.theme;
    const master = this.masters.getItem(id);
    if (master) return master;
    const layout = this.layouts.getItem(id);
    if (layout) return layout;
    const customShow = this.customShows.getItem(id);
    if (customShow) return customShow;
    for (const slide of this.slides) {
      if (slide.id === id) return slide;
      const found = slide.resolve(id);
      if (found) return found;
    }
    return undefined;
  }

  help(query = "*", options = {}) {
    return ndjson(queryHelpRecords("presentation", query, options), options.maxChars ?? Infinity);
  }

  async export(options = {}) {
    if (options.format === "montage" || options.montage === true) return new FileBlob(presentationMontageSvg(this, options), { type: "image/svg+xml", metadata: { format: "montage", slides: this.slides.count, artifactKind: "presentation" } });
    const slide = options.slide || this.slides.getItem(0) || this.slides.add();
    if (options.format === "layout") return slide.export({ ...options, format: "layout" });
    return slide.export(options);
  }

  toProto() {
    return { id: this.id, slideSize: this.slideSize, theme: this.theme.toJSON(), master: this.master.toJSON(), masters: this.masters.items.map((master) => master.toJSON()), layouts: this.layouts.items.map((layout) => layout.toJSON()), slides: this.slides.items.map((slide) => slide.toProto()) };
  }
}

class ShapeCollection {
  constructor(slide, owner) { this.slide = slide; this.owner = owner; this.items = []; }
  add(config = {}) { const shape = new Shape(this.slide, config); shape.parentGroup = this.owner; this.items.push(shape); this.owner?._rememberChild?.(shape); return shape; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ElementCollection {
  constructor(slide, ElementClass, owner) { this.slide = slide; this.ElementClass = ElementClass; this.owner = owner; this.items = []; }
  add(...args) { const element = new this.ElementClass(this.slide, ...args); element.parentGroup = this.owner; this.items.push(element); this.owner?._rememberChild?.(element); return element; }
  getItemAt(index) { return this.items[index]; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

function normalizeFrame(config = {}, fallback = { left: 0, top: 0, width: 240, height: 160 }) {
  const source = config.position || config.frame || config;
  return {
    left: source.left ?? fallback.left,
    top: source.top ?? fallback.top,
    width: source.width ?? fallback.width,
    height: source.height ?? fallback.height,
  };
}

function resolveAutoLayoutFrame(slide, frame) {
  if (frame === "slide") return slide.frame;
  if (frame?.position) return frame.position;
  if (frame && typeof frame.left === "number" && typeof frame.top === "number" && typeof frame.width === "number" && typeof frame.height === "number") return frame;
  return slide.frame;
}

function elementFrame(element) {
  return element.position || element.frame || element.layoutJson?.().frame;
}

function elementLabel(element) {
  return element.name || element.id;
}

function overlapArea(a, b) {
  const left = Math.max(a.left, b.left);
  const top = Math.max(a.top, b.top);
  const right = Math.min(a.left + a.width, b.left + b.width);
  const bottom = Math.min(a.top + a.height, b.top + b.height);
  return Math.max(0, right - left) * Math.max(0, bottom - top);
}

function textOverflowIssue(slide, element, frame) {
  const text = element.text?.value || "";
  if (!text) return undefined;
  const paragraphs = typeof element.text.effectiveParagraphs === "function" ? element.text.effectiveParagraphs() : normalizePresentationParagraphs(text);
  const requiredHeight = paragraphs.reduce((height, paragraph) => {
    const paragraphFontSize = Math.max(element.text.style.fontSize || 24, ...paragraph.runs.map((run) => run.style?.fontSize || 0));
    const availableWidth = Math.max(1, frame.width - 18 - Math.max(0, paragraph.marginLeft || paragraph.level * 24));
    const charsPerLine = Math.max(1, Math.floor(availableWidth / (paragraphFontSize * 0.55)));
    const requiredLines = presentationParagraphsText([paragraph]).split("\n").reduce((lines, line) => lines + Math.max(1, Math.ceil(line.length / charsPerLine)), 0);
    const spacing = paragraph.lineSpacing || element.text.style.lineSpacing || 1.2;
    const lineHeight = spacing > 10 ? spacing : paragraphFontSize * spacing;
    return height + (paragraph.spaceBefore ?? paragraphFontSize * (paragraph.spaceBeforePercent || 0)) + requiredLines * lineHeight + (paragraph.spaceAfter ?? paragraphFontSize * (paragraph.spaceAfterPercent || 0));
  }, 12);
  if (requiredHeight <= frame.height) return undefined;
  return {
    kind: "layoutIssue",
    type: "textOverflow",
    severity: "warning",
    slide: slide.index + 1,
    id: element.id,
    name: element.name || undefined,
    bbox: [frame.left, frame.top, frame.width, frame.height],
    requiredHeight: Math.round(requiredHeight),
    message: `Text may overflow ${elementLabel(element)}: estimated ${Math.round(requiredHeight)}px required for ${Math.round(frame.height)}px frame.`,
  };
}

function tableOverflowIssues(slide, tableElement, frame = tableElement.position) {
  const issues = [];
  const cellW = frame.width / Math.max(1, tableElement.columns);
  const cellH = frame.height / Math.max(1, tableElement.rows);
  const fontSize = 13;
  for (let row = 0; row < tableElement.rows; row++) {
    for (let column = 0; column < tableElement.columns; column++) {
      const value = String(tableElement.values[row]?.[column] ?? "");
      const requiredWidth = value.length * fontSize * 0.55 + 12;
      if (requiredWidth > cellW || cellH < fontSize * 1.4) {
        issues.push({
          kind: "layoutIssue",
          type: "tableTextOverflow",
          severity: "warning",
          slide: slide.index + 1,
          id: tableElement.id,
          name: tableElement.name || undefined,
          row,
          column,
          bbox: [frame.left + column * cellW, frame.top + row * cellH, cellW, cellH],
          message: `Table cell ${elementLabel(tableElement)}[${row},${column}] may overflow its cell.`,
        });
      }
    }
  }
  return issues;
}

function pointFromElement(element, fallback = { x: 0, y: 0 }) {
  const frame = elementFrame(element);
  return frame ? { x: frame.left + frame.width / 2, y: frame.top + frame.height / 2 } : fallback;
}

function connectorPoint(slide, pointOrTarget, fallback = { x: 0, y: 0 }) {
  if (!pointOrTarget) return fallback;
  if (typeof pointOrTarget === "string") return pointFromElement(slide.resolve(pointOrTarget), fallback);
  if (pointOrTarget.id) return pointFromElement(slide.resolve(pointOrTarget.id) || pointOrTarget, fallback);
  if (pointOrTarget.element) return pointFromElement(pointOrTarget.element, fallback);
  if (pointOrTarget.targetId) return pointFromElement(slide.resolve(pointOrTarget.targetId), fallback);
  if (Number.isFinite(pointOrTarget.x) && Number.isFinite(pointOrTarget.y)) return { x: Number(pointOrTarget.x), y: Number(pointOrTarget.y) };
  return fallback;
}

class SlideCommentThread {
  constructor(slide, target, text, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("pc");
    this.targetId = typeof target === "string" ? target : target?.id || config.targetId;
    this.author = config.author || "User";
    this.resolved = Boolean(config.resolved);
    this.created = config.created || new Date(0).toISOString();
    this.nativeFormat = config.nativeFormat;
    this.nativeAnchor = config.nativeAnchor;
    this.position = config.position;
    this.comments = (config.comments || [{ author: this.author, text: String(text ?? ""), created: this.created }]).map((comment) => ({ ...comment, author: comment.author || this.author, text: String(comment.text ?? ""), created: comment.created || this.created }));
  }

  addReply(text, config = {}) {
    this.comments.push({ ...config, author: config.author || this.author, text: String(text ?? ""), created: config.created || new Date(0).toISOString() });
    return this;
  }

  resolve() { this.resolved = true; return this; }
  reopen() { this.resolved = false; return this; }

  inspectRecord() {
    return { kind: "comment", id: this.id, slide: this.slide.index + 1, targetId: this.targetId, author: this.author, resolved: this.resolved, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, nativeCommentIds: this.comments.map((comment) => comment.nativeId).filter(Boolean), replies: Math.max(0, this.comments.length - 1), textPreview: this.comments.map((comment) => comment.text).join("\n").slice(0, 300) };
  }

  toJSON() { return { id: this.id, targetId: this.targetId, author: this.author, resolved: this.resolved, created: this.created, nativeFormat: this.nativeFormat, nativeAnchor: this.nativeAnchor, position: this.position, comments: this.comments.map((comment) => ({ ...comment })) }; }
}

class SlideCommentCollection {
  constructor(slide) { this.slide = slide; this.items = []; }
  addThread(target, text, config = {}) { const thread = new SlideCommentThread(this.slide, target, text, config); this.items.push(thread); return thread; }
  add(target, text, config = {}) { return this.addThread(target, text, config); }
  getItem(id) { return this.items.find((thread) => thread.id === id); }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

class ConnectorElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.kind = "connector";
    this.id = config.id || aid("cx");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.connectorType = config.connectorType || config.type || "straight";
    this.startTargetId = typeof config.from === "string" ? config.from : config.from?.id || config.startTargetId;
    this.endTargetId = typeof config.to === "string" ? config.to : config.to?.id || config.endTargetId;
    this.start = config.start || connectorPoint(slide, config.from || config.startTargetId, { x: 0, y: 0 });
    this.end = config.end || connectorPoint(slide, config.to || config.endTargetId, { x: 160, y: 0 });
    this.line = config.line || { fill: "#334155", width: 2, endArrow: config.endArrow || "triangle" };
  }

  get position() {
    const left = Math.min(this.start.x, this.end.x);
    const top = Math.min(this.start.y, this.end.y);
    return { left, top, width: Math.abs(this.end.x - this.start.x), height: Math.abs(this.end.y - this.start.y) };
  }

  inspectRecord() {
    return { kind: "connector", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line };
  }

  layoutJson() { return { kind: "connector", id: this.id, name: this.name, connectorType: this.connectorType, start: this.start, end: this.end, startTargetId: this.startTargetId, endTargetId: this.endTargetId, line: this.line, frame: this.position }; }

  toSvg() {
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const width = this.line?.width ?? 2;
    const markerId = `${this.id.replace(/[^A-Za-z0-9_-]/g, "")}-arrow`;
    const marker = this.line?.endArrow ? `<defs><marker id="${markerId}" viewBox="0 0 10 10" refX="9" refY="5" markerWidth="6" markerHeight="6" orient="auto-start-reverse"><path d="M 0 0 L 10 5 L 0 10 z" fill="${xmlEscape(stroke)}"/></marker></defs>` : "";
    return `${marker}<line x1="${this.start.x}" y1="${this.start.y}" x2="${this.end.x}" y2="${this.end.y}" stroke="${xmlEscape(stroke)}" stroke-width="${width}" marker-end="${this.line?.endArrow ? `url(#${markerId})` : ""}"/>`;
  }

}

const NativePresentationObject = createNativePresentationObjectClass({ normalizeFrame });

const GroupShape = createPresentationGroupShapeClass({
  createId: aid,
  createShapeCollection: (slide, owner) => new ShapeCollection(slide, owner),
  createConnectorCollection: (slide, owner) => new ElementCollection(slide, ConnectorElement, owner),
  createGroupCollection: (slide, owner, GroupClass) => new ElementCollection(slide, GroupClass, owner),
  createTableCollection: (slide, owner) => new ElementCollection(slide, TableElement, owner),
  createChartCollection: (slide, owner) => new ElementCollection(slide, ChartElement, owner),
  createImageCollection: (slide, owner) => new ElementCollection(slide, ImageElement, owner),
  createNativeObjectCollection: (slide, owner) => new ElementCollection(slide, NativePresentationObject, owner),
  isShape: (element) => element instanceof Shape,
  isConnector: (element) => element instanceof ConnectorElement,
  isGroup: (element) => element instanceof GroupShape,
  isTable: (element) => element instanceof TableElement,
  isChart: (element) => element instanceof ChartElement,
  isImage: (element) => element instanceof ImageElement,
  isNativeObject: (element) => element instanceof NativePresentationObject,
  elementKind: (element) => presentationElementKind(element),
  validateChildLayout: (element, frame) => element instanceof TableElement ? tableOverflowIssues(element.slide, element, frame) : [],
  createTextRange: (element, id) => createTextRange(element, id, { parentKind: "shape" }),
  textRangeRecord,
  elementLabel,
});
export { GroupShape };
function slideLayoutSlice(slide, layout, options = {}) {
  const targets = inspectTargetTokens(options);
  const search = String(options.search || options.searchTerm || "").trim().toLowerCase();
  if (!targets.length && !search) return layout;
  const before = Math.max(0, Number(options.before ?? options.contextBefore ?? options.context ?? 0) || 0);
  const after = Math.max(0, Number(options.after ?? options.contextAfter ?? options.context ?? 0) || 0);
  const targetsSlide = targets.some((target) => target === slide.id || target === slide.name || target === String(slide.index + 1) || target === "slide");
  if (targetsSlide && !search) return { ...layout, slice: { targets, before, after, matchedElements: layout.elements.length, returnedElements: layout.elements.length } };
  const matches = [];
  layout.elements.forEach((element, index) => {
    const matchesSearch = !search || JSON.stringify(element).toLowerCase().includes(search);
    const matchesTarget = !targets.length || targetsSlide || inspectRecordMatchesTarget(element, targets);
    if (matchesSearch && matchesTarget) matches.push(index);
  });
  const keep = new Set();
  for (const index of matches) {
    for (let i = Math.max(0, index - before); i <= Math.min(layout.elements.length - 1, index + after); i += 1) keep.add(i);
  }
  const elements = layout.elements.filter((_, index) => keep.has(index));
  return { ...layout, elements, slice: { targets, search: search || undefined, before, after, matchedElements: matches.length, returnedElements: elements.length } };
}

class SpeakerNotes {
  constructor(slide, text = "") {
    this.slide = slide;
    this.textFrame = new TextFrame(text);
  }

  get id() { return `${this.slide.id}/notes`; }
  get text() { return this.textFrame.value; }
  set text(value) { this.textFrame.set(value); }
  setText(value) { this.textFrame.set(value); return this; }
  append(value) { this.textFrame.set(`${this.text}${String(value ?? "")}`); return this; }
  clear() { this.textFrame.set(""); return this; }
}

export class Slide {
  constructor(presentation, options = {}) {
    this.presentation = presentation;
    this.id = aid("sl");
    this.name = options.name || "";
    this.shapes = new ShapeCollection(this);
    this.images = new ElementCollection(this, ImageElement);
    this.tables = new ElementCollection(this, TableElement);
    this.charts = new ElementCollection(this, ChartElement);
    this.connectors = new ElementCollection(this, ConnectorElement);
    this.groups = new ElementCollection(this, GroupShape);
    this.nativeObjects = new ElementCollection(this, NativePresentationObject);
    this.comments = new SlideCommentCollection(this);
    this.layoutId = options.layoutId || options.layout?.id || (typeof options.layout === "string" ? options.layout : undefined);
    this.speakerNotes = new SpeakerNotes(this, options.notes || options.speakerNotes?.text || "");
    this.background = options.background ? normalizePresentationBackground(options.background) : {};
  }

  get index() { return this.presentation.slides.items.indexOf(this); }
  get frame() { return { left: 0, top: 0, ...this.presentation.slideSize }; }

  addNotes(text) { return this.speakerNotes.setText(text); }
  addComment(target, text, config = {}) { return this.comments.addThread(target, text, config); }
  addConnector(config = {}) { return this.connectors.add(config); }
  addGroup(config = {}) { return this.groups.add(config); }
  setBackground(background) { this.background = normalizePresentationBackground(background, this.background); return this; }
  clearBackground() { this.background = {}; return this; }
  applyLayout(layoutOrName) { const layout = typeof layoutOrName === "string" ? this.presentation.layouts.getItem(layoutOrName) : layoutOrName; if (!layout) throw new Error(`Unknown slide layout: ${layoutOrName}`); return layout.apply(this); }
  effectiveBackground() { const layout = this.presentation.layouts.getItem(this.layoutId); return this.background.fill ? this.background : layout?.effectiveBackground() || this.presentation.master.effectiveBackground(); }
  effectiveTheme() { const layout = this.presentation.layouts.getItem(this.layoutId); return layout?.effectiveTheme() || this.presentation.master.effectiveTheme(); }

  inspectRecords(kinds) {
    const records = [];
    if (kinds.has("layout")) { const layout = this.presentation.layouts.getItem(this.layoutId); records.push({ kind: "layout", layoutId: this.layoutId || `${this.id}/layout`, name: layout?.name || "Blank", type: layout?.type || "blank", masterId: layout?.masterId, themeId: this.effectiveTheme().id, placeholders: layout?.placeholders.length || 0 }); }
    if (kinds.has("slide")) records.push({ kind: "slide", id: this.id, slide: this.index + 1, title: this.title(), background: this.background.fill ? this.background : undefined, effectiveBackground: this.effectiveBackground(), textShapes: this.shapes.items.filter((s) => s.text.value).length, tables: this.tables.items.length, charts: this.charts.items.length, images: this.images.items.length, connectors: this.connectors.items.length, groups: this.groups.items.length, nativeObjects: this.nativeObjects.items.length, comments: this.comments.items.length, hasNotes: Boolean(this.speakerNotes.text) });
    for (const shape of this.shapes) {
      if (kinds.has("textbox") && shape.text.value) records.push(shape.inspectRecord("textbox"));
      else if (kinds.has("shape")) records.push(shape.inspectRecord("shape"));
      if (kinds.has("textRange") && shape.text.value) records.push(textRangeRecord(shape, { parentKind: "shape", record: { slide: this.index + 1, bbox: [shape.position.left, shape.position.top, shape.position.width, shape.position.height], bboxUnit: "px" } }));
    }
    if (kinds.has("table")) records.push(...this.tables.items.map((table) => table.inspectRecord()));
    if (kinds.has("chart")) records.push(...this.charts.items.map((chart) => chart.inspectRecord()));
    if (kinds.has("image")) records.push(...this.images.items.map((image) => image.inspectRecord()));
    if (kinds.has("connector")) records.push(...this.connectors.items.map((connector) => connector.inspectRecord()));
    if (kinds.has("nativeObject") || kinds.has("native")) records.push(...this.nativeObjects.items.map((object) => object.inspectRecord()));
    for (const nativeKind of ["contentPart", "oleObject", "diagram", "graphicFrame"]) if (kinds.has(nativeKind)) records.push(...this.nativeObjects.items.filter((object) => object.nativeKind === nativeKind).map((object) => object.inspectRecord()));
    for (const group of this.groups) records.push(...group.inspectRecords(kinds));
    if (kinds.has("comment") || kinds.has("thread")) records.push(...this.comments.items.map((comment) => comment.inspectRecord()));
    if (kinds.has("notes")) records.push({ kind: "notes", id: `${this.id}/notes`, slide: this.index + 1, text: this.speakerNotes.text, textPreview: this.speakerNotes.text.slice(0, 300), textChars: this.speakerNotes.text.length });
    return records;
  }

  title() { return this.shapes.items.find((shape) => shape.text.value)?.text.value || this.charts.items[0]?.title || ""; }
  resolve(id) {
    if (id === this.speakerNotes.id) return this.speakerNotes;
    if (String(id || "").endsWith("/text")) {
      const parentId = String(id).slice(0, -5);
      const shape = this.shapes.items.find((item) => item.id === parentId);
      if (shape) return createTextRange(shape, id, { parentKind: "shape" });
    }
    const direct = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.groups.items, ...this.nativeObjects.items, ...this.comments.items].find((element) => element.id === id);
    if (direct) return direct;
    for (const group of this.groups) {
      const nested = group.resolve(id);
      if (nested) return nested;
    }
    return undefined;
  }

  validateLayout(options = {}) {
    const issues = [];
    const slideFrame = this.frame;
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.groups.items, ...this.nativeObjects.items];
    const connectors = this.connectors.items;
    const minOverlapArea = options.minOverlapArea ?? 64;
    const padding = options.boundsPadding ?? 0;
    for (const element of elements) {
      const frame = elementFrame(element);
      if (!frame) continue;
      const offCanvas = frame.left < slideFrame.left - padding || frame.top < slideFrame.top - padding || frame.left + frame.width > slideFrame.left + slideFrame.width + padding || frame.top + frame.height > slideFrame.top + slideFrame.height + padding;
      if (offCanvas) {
        issues.push({
          kind: "layoutIssue",
          type: "offCanvas",
          severity: "error",
          slide: this.index + 1,
          id: element.id,
          name: element.name || undefined,
          bbox: [frame.left, frame.top, frame.width, frame.height],
          message: `${elementLabel(element)} extends outside the slide frame.`,
        });
      }
      const textIssue = textOverflowIssue(this, element, frame);
      if (textIssue) issues.push(textIssue);
      if (element instanceof TableElement) issues.push(...tableOverflowIssues(this, element));
    }
    for (const connector of connectors) {
      const points = [connector.start, connector.end];
      if (points.some((point) => point.x < slideFrame.left - padding || point.y < slideFrame.top - padding || point.x > slideFrame.left + slideFrame.width + padding || point.y > slideFrame.top + slideFrame.height + padding)) {
        issues.push({ kind: "layoutIssue", type: "connectorOffCanvas", severity: "error", slide: this.index + 1, id: connector.id, name: connector.name || undefined, start: connector.start, end: connector.end, message: `${elementLabel(connector)} connector endpoint extends outside the slide frame.` });
      }
    }
    for (const group of this.groups) issues.push(...group.validateLayout());
    for (let leftIndex = 0; leftIndex < elements.length; leftIndex++) {
      for (let rightIndex = leftIndex + 1; rightIndex < elements.length; rightIndex++) {
        const left = elements[leftIndex];
        const right = elements[rightIndex];
        const leftFrame = elementFrame(left);
        const rightFrame = elementFrame(right);
        if (!leftFrame || !rightFrame) continue;
        const area = overlapArea(leftFrame, rightFrame);
        if (area >= minOverlapArea) {
          issues.push({
            kind: "layoutIssue",
            type: "overlap",
            severity: "error",
            slide: this.index + 1,
            ids: [left.id, right.id],
            names: [elementLabel(left), elementLabel(right)],
            overlapArea: Math.round(area),
            message: `${elementLabel(left)} overlaps ${elementLabel(right)} by about ${Math.round(area)}px².`,
          });
        }
      }
    }
    return { ok: issues.length === 0, issues, ...ndjson(issues, options.maxChars ?? Infinity) };
  }

  async export(options = {}) {
    if (options.format === "layout" || options.format === LAYOUT_MIME) return new FileBlob(JSON.stringify(this.layoutJson(options), null, 2), { type: LAYOUT_MIME, metadata: { artifactKind: "presentation", format: "layout", slide: this.index + 1, target: options.target ?? options.targetId ?? options.id ?? options.anchor, search: options.search ?? options.searchTerm } });
    return new FileBlob(this.toSvg(), { type: "image/svg+xml" });
  }

  layoutJson(options = {}) {
    const elements = [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.groups.items, ...this.nativeObjects.items].map((element) => {
      const record = element.layoutJson();
      const comments = this.comments.items.filter((comment) => comment.targetId === element.id);
      return {
        ...record,
        slide: this.index + 1,
        textRangeId: element.text?.value ? `${element.id}/text` : undefined,
        commentIds: comments.length ? comments.map((comment) => comment.id) : undefined,
        commentTextPreview: comments.length ? comments.flatMap((comment) => comment.comments.map((item) => item.text)).join("\n").slice(0, 300) : undefined,
      };
    });
    return slideLayoutSlice(this, {
      schema: "open-office-artifact.layout/v1",
      unit: "px",
      slide: { id: this.id, slide: this.index + 1, frame: this.frame, background: this.effectiveBackground(), notes: this.speakerNotes.text || undefined },
      elements,
    }, options);
  }

  toSvg() {
    const { width, height } = this.presentation.slideSize;
    const elements = [...this.connectors.items, ...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.groups.items, ...this.nativeObjects.items].map((element) => element.toSvg()).join("");
    return `<svg xmlns="http://www.w3.org/2000/svg" width="${width}" height="${height}" viewBox="0 0 ${width} ${height}"><rect width="100%" height="100%" fill="${xmlEscape(resolvePresentationBackgroundColor(this.effectiveBackground(), this.effectiveTheme()))}"/>${elements}</svg>`;
  }

  toProto() { return { id: this.id, layoutId: this.layoutId, background: this.background.fill ? this.background : undefined, notes: this.speakerNotes.text || undefined, comments: this.comments.items.map((comment) => comment.toJSON()), elements: [...this.shapes.items, ...this.tables.items, ...this.charts.items, ...this.images.items, ...this.connectors.items, ...this.nativeObjects.items].map((element) => element.layoutJson()), groups: this.groups.items.map((group) => group.toProto()) }; }

  compose(composeNode, options = {}) {
    const frame = options.frame || { left: 72, top: 64, width: this.presentation.slideSize.width - 144, height: this.presentation.slideSize.height - 128 };
    return materializeComposeNode(this, composeNode, frame);
  }

  autoLayout(shapes, options = {}) {
    const items = Array.from(shapes || []).filter(Boolean);
    if (items.length === 0) return items;
    const frame = resolveAutoLayoutFrame(this, options.frame || "slide");
    const inner = innerFrame(frame, {
      left: options.horizontalPadding ?? 0,
      right: options.horizontalPadding ?? 0,
      top: options.verticalPadding ?? 0,
      bottom: options.verticalPadding ?? 0,
    });
    const direction = options.direction || "horizontal";
    const horizontal = direction === "horizontal";
    const mainSize = horizontal ? "width" : "height";
    const crossSize = horizontal ? "height" : "width";
    const requestedGap = horizontal ? options.horizontalGap : options.verticalGap;
    const totalMain = items.reduce((sum, shape) => sum + (shape.position?.[mainSize] ?? 0), 0);
    const gap = requestedGap === "auto"
      ? items.length > 1 ? Math.max(0, (inner[mainSize] - totalMain) / (items.length - 1)) : 0
      : Number(requestedGap ?? 0);
    const usedMain = totalMain + gap * Math.max(0, items.length - 1);
    const align = options.align || "center";
    const mainStart = align.includes("Right") || align === "right" || align.includes("Bottom")
      ? inner[horizontal ? "left" : "top"] + inner[mainSize] - usedMain
      : align === "center" || align === "left" || align === "right"
        ? inner[horizontal ? "left" : "top"] + Math.max(0, (inner[mainSize] - usedMain) / 2)
        : inner[horizontal ? "left" : "top"];
    let cursor = mainStart;
    for (const shape of items) {
      const crossStart = align.includes("Bottom")
        ? inner[horizontal ? "top" : "left"] + inner[crossSize] - shape.position[crossSize]
        : align.includes("Center") || align === "center" || align === "left" || align === "right"
          ? inner[horizontal ? "top" : "left"] + Math.max(0, (inner[crossSize] - shape.position[crossSize]) / 2)
          : inner[horizontal ? "top" : "left"];
      shape.position = horizontal
        ? { ...shape.position, left: cursor, top: crossStart }
        : { ...shape.position, left: crossStart, top: cursor };
      cursor += shape.position[mainSize] + gap;
    }
    return items;
  }
}

class TextFrame {
  constructor(text = "", bodyProperties, { defaultBodyProperties = false } = {}) { this._paragraphs = normalizePresentationParagraphs(text); this.style = {}; this.inheritedParagraphStyles = {}; this.bodyProperties = normalizePresentationTextBodyProperties(bodyProperties, { defaults: defaultBodyProperties }); }
  get value() { return presentationParagraphsText(this._paragraphs); }
  set value(text) { this._paragraphs = normalizePresentationParagraphs(text); }
  get paragraphs() { return normalizePresentationParagraphs(this._paragraphs); }
  set paragraphs(value) { this._paragraphs = normalizePresentationParagraphs(value); }
  effectiveParagraphs() { return inheritPresentationParagraphs(this._paragraphs, this.inheritedParagraphStyles); }
  set(text) { this._paragraphs = normalizePresentationParagraphs(text); return this; }
  setText(text) { return this.set(text); }
  replace(search, replacement) { replacePresentationParagraphText(this._paragraphs, search, replacement); return this; }
  toString() { return this.value; }
}

export class Shape {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("sh");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.geometry = config.geometry || "rect";
    this.customPaths = normalizePresentationCustomPaths(config.customPaths, { geometry: this.geometry });
    this.name = config.name || "";
    this.position = config.position || { left: 0, top: 0, width: 160, height: 80 };
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation shape ${this.name || this.id} transform`);
    this.fill = config.fill || "transparent";
    this.line = config.line || { fill: "#334155", width: 1 };
    this.borderRadius = config.borderRadius;
    this.shadow = config.shadow ? { ...config.shadow } : undefined;
    this.placeholder = config.placeholder;
    this._text = new TextFrame(config.text ?? "", config.textBodyProperties, { defaultBodyProperties: config.textBodyProperties === undefined });
    this._text.style = { ...(config.textStyle || config.style?.text || {}) };
  }

  get text() { return this._text; }
  set text(value) { this._text.set(value); }

  inspectRecord(kind = "shape") {
    const p = this.position;
    const paragraphs = this.text.effectiveParagraphs();
    return { kind, id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, text: this.text.value || undefined, textPreview: this.text.value || undefined, textChars: this.text.value.length || undefined, textLines: this.text.value ? this.text.value.split("\n").length : undefined, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, customPathCount: this.customPaths.length || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", transform: this.transform, shadow: this.shadow, placeholder: this.placeholder || undefined };
  }

  layoutJson() { const paragraphs = this.text.effectiveParagraphs(); return { kind: this.text.value ? "textbox" : "shape", id: this.id, name: this.name, geometry: this.geometry, customPaths: this.customPaths.length ? this.customPaths : undefined, frame: this.position, transform: this.transform, text: this.text.value, paragraphs: presentationParagraphsNeedSerialization(paragraphs) ? paragraphs : undefined, bodyProperties: this.text.bodyProperties, placeholder: this.placeholder, style: { fill: this.fill, line: this.line, borderRadius: this.borderRadius, shadow: this.shadow, text: this.text.style } }; }

  toSvg() {
    const p = this.position;
    const fill = typeof this.fill === "string" ? resolveColorToken(this.fill, this.fill) : this.fill?.color || "transparent";
    const stroke = resolveColorToken(this.line?.fill || this.line?.color || "#334155", "#334155");
    const sw = this.line?.width ?? 1;
    const visual = this.geometry === "custom"
      ? `<g fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}">${presentationCustomPathsSvg(this.customPaths, p, { escape: xmlEscape })}</g>`
      : this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="${xmlEscape(fill)}" stroke="${xmlEscape(stroke)}" stroke-width="${sw}"/>`;
    const text = this.text.value ? presentationParagraphsSvg(this.text.effectiveParagraphs(), p, this.text.style, { escape: xmlEscape }) : "";
    if (!this.transform) return visual + text;
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform.rotationDegrees || 0);
    const flipHorizontal = this.transform.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform.flipVertical === true ? -1 : 1;
    return `<g transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})">${visual}${text}</g>`;
  }

}

class TableCellFacade {
  constructor(table, row, column) { this.table = table; this.row = row; this.column = column; this.text = new TextFrame(); }
  get value() { return this.table.values[this.row]?.[this.column] ?? ""; }
  set value(value) { this.table.ensureCell(this.row, this.column); this.table.values[this.row][this.column] = value; }
}

export class TableElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("tb");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.rows = Number(config.rows || config.values?.length || 1);
    this.columns = Number(config.columns || config.values?.[0]?.length || 1);
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 160 });
    this.values = Array.from({ length: this.rows }, (_, r) => Array.from({ length: this.columns }, (_, c) => config.values?.[r]?.[c] ?? ""));
    this.style = config.style;
    this.styleOptions = config.styleOptions || {};
    this.cells = { set: (row, column, value) => { this.getCell(row, column).value = value; }, block: (range) => ({ table: this, range }) };
    this.borders = { assign: (configValue) => { this.border = configValue; } };
  }

  ensureCell(row, column) {
    while (this.values.length <= row) this.values.push([]);
    while (this.values[row].length <= column) this.values[row].push("");
  }

  getCell(row, column) { return new TableCellFacade(this, row, column); }
  merge(range) { this.mergeRange = range; }

  inspectRecord() {
    const p = this.position;
    return { kind: "table", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, rows: this.rows, cols: this.columns, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", values: this.values };
  }

  layoutJson() { return { kind: "table", id: this.id, name: this.name, frame: this.position, rows: this.rows, columns: this.columns, values: this.values, style: this.style, styleOptions: this.styleOptions }; }

  toSvg() {
    const p = this.position;
    const cellW = p.width / Math.max(1, this.columns);
    const cellH = p.height / Math.max(1, this.rows);
    const parts = [`<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>`];
    for (let r = 0; r < this.rows; r++) {
      for (let c = 0; c < this.columns; c++) {
        const x = p.left + c * cellW;
        const y = p.top + r * cellH;
        const fill = this.styleOptions.headerRow && r === 0 ? "#0f172a" : r % 2 ? "#f8fafc" : "#ffffff";
        const color = this.styleOptions.headerRow && r === 0 ? "#ffffff" : "#0f172a";
        parts.push(`<rect x="${x}" y="${y}" width="${cellW}" height="${cellH}" fill="${fill}" stroke="#cbd5e1"/>`);
        parts.push(`<text x="${x + 6}" y="${y + Math.min(22, cellH - 6)}" font-family="Arial" font-size="13" fill="${color}">${xmlEscape(this.values[r]?.[c] ?? "")}</text>`);
      }
    }
    return parts.join("");
  }

}

function normalizeChartSeries(seriesItems = [], chartType = "bar") {
  return (seriesItems || []).map((series, index) => {
    const values = (series.values || series.data || []).map((value) => value);
    const style = normalizePresentationChartSeriesStyle(series, values.length);
    const seriesChartType = chartType === "combo" ? String(series.chartType || series.type || "").toLowerCase() : undefined;
    if (chartType === "combo" && !new Set(["bar", "line"]).has(seriesChartType)) throw new TypeError("Presentation combo chart series chartType must be bar or line.");
    const rawAxisGroup = series.axisGroup ?? series.axis ?? (series.secondaryAxis === true ? "secondary" : "primary");
    const axisGroup = normalizePresentationChartAxisGroup(rawAxisGroup === "y2" ? "secondary" : rawAxisGroup === "y1" ? "primary" : String(rawAxisGroup).toLowerCase(), seriesChartType || chartType);
    return {
      name: series.name || `Series ${index + 1}`,
      values,
      categories: series.categories,
      color: style.color || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4],
      ...(style.line ? { line: style.line } : {}),
      ...(style.points.length ? { points: style.points } : {}),
      ...(style.marker ? { marker: style.marker } : {}),
      ...(style.smooth == null ? {} : { smooth: style.smooth }),
      ...(series.dataLabels === undefined ? {} : { dataLabels: normalizePresentationChartDataLabels(series.dataLabels) }),
      ...((series.trendlines ?? series.trendline) == null ? {} : { trendlines: normalizePresentationChartTrendlines(series.trendlines ?? series.trendline, values.length, seriesChartType || chartType) }),
      ...(series.errorBars == null ? {} : { errorBars: normalizePresentationChartErrorBars(series.errorBars, seriesChartType || chartType, values.length) }),
      ...(seriesChartType ? { chartType: seriesChartType } : {}),
      ...(axisGroup === "secondary" ? { axisGroup } : {}),
    };
  });
}

function normalizeChartAxes(config = {}, hasSecondary = false) {
  const axes = config.axes || {};
  const axisTitles = config.axisTitles || {};
  const secondary = axes.secondary || {};
  const secondaryAxisTitles = axisTitles.secondary || config.secondaryAxisTitles || {};
  return {
    category: { ...(axes.category || axes.x || {}), title: axes.category?.title || axes.x?.title || axisTitles.category || axisTitles.x || config.categoryAxisTitle || config.xAxisTitle || "" },
    value: { ...(axes.value || axes.y || {}), title: axes.value?.title || axes.y?.title || axisTitles.value || axisTitles.y || config.valueAxisTitle || config.yAxisTitle || "" },
    ...(hasSecondary ? {
      secondary: {
        category: { ...(secondary.category || secondary.x || axes.secondaryCategory || {}), title: secondary.category?.title || secondary.x?.title || axes.secondaryCategory?.title || secondaryAxisTitles.category || secondaryAxisTitles.x || config.secondaryCategoryAxisTitle || config.secondaryXAxisTitle || "" },
        value: { ...(secondary.value || secondary.y || axes.secondaryValue || axes.y2 || {}), title: secondary.value?.title || secondary.y?.title || axes.secondaryValue?.title || axes.y2?.title || secondaryAxisTitles.value || secondaryAxisTitles.y || config.secondaryValueAxisTitle || config.secondaryYAxisTitle || "" },
      },
    } : {}),
  };
}

function normalizeChartLegend(config = {}, seriesLength = 0) {
  const raw = config.legend;
  if (raw === false || config.hasLegend === false) return { visible: false, position: "r" };
  if (typeof raw === "string") return { visible: true, position: raw };
  return { visible: raw?.visible ?? config.hasLegend ?? seriesLength > 1, position: raw?.position || config.legendPosition || "r" };
}

function normalizeChartDataLabels(config = {}) {
  const raw = config.dataLabels ?? config.labels ?? {};
  if (raw === true || raw === false) return normalizePresentationChartDataLabels(raw);
  return normalizePresentationChartDataLabels({
    ...raw,
    showValue: raw.showValue ?? config.showValues,
    showCategoryName: raw.showCategoryName ?? raw.showCategory ?? config.showCategoryLabels,
  });
}

function pieSlicePath(cx, cy, radius, startAngle, endAngle) {
  const startX = cx + radius * Math.cos(startAngle);
  const startY = cy + radius * Math.sin(startAngle);
  const endX = cx + radius * Math.cos(endAngle);
  const endY = cy + radius * Math.sin(endAngle);
  const largeArc = endAngle - startAngle > Math.PI ? 1 : 0;
  return `M ${cx} ${cy} L ${startX} ${startY} A ${radius} ${radius} 0 ${largeArc} 1 ${endX} ${endY} Z`;
}

function presentationChartMarkerSvg(marker, x, y, color) {
  if (!marker || marker.symbol === "none") return "";
  const size = Math.max(2, Number(marker.size) || 5);
  const radius = size / 2;
  const stroke = xmlEscape(color);
  if (marker.symbol === "square") return `<rect x="${x - radius}" y="${y - radius}" width="${size}" height="${size}" fill="${stroke}"/>`;
  if (marker.symbol === "diamond") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y} L ${x} ${y + radius} L ${x - radius} ${y} Z" fill="${stroke}"/>`;
  if (marker.symbol === "triangle") return `<path d="M ${x} ${y - radius} L ${x + radius} ${y + radius} L ${x - radius} ${y + radius} Z" fill="${stroke}"/>`;
  if (marker.symbol === "x") return `<path d="M ${x - radius} ${y - radius} L ${x + radius} ${y + radius} M ${x + radius} ${y - radius} L ${x - radius} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "plus") return `<path d="M ${x - radius} ${y} L ${x + radius} ${y} M ${x} ${y - radius} L ${x} ${y + radius}" fill="none" stroke="${stroke}" stroke-width="1.5"/>`;
  if (marker.symbol === "dash") return `<line x1="${x - radius}" y1="${y}" x2="${x + radius}" y2="${y}" stroke="${stroke}" stroke-width="2"/>`;
  return `<circle cx="${x}" cy="${y}" r="${marker.symbol === "dot" ? Math.max(1, radius / 2) : radius}" fill="${stroke}"/>`;
}

function presentationChartDataLabelText(dataLabels, category, value) {
  if (!dataLabels?.showValue && !dataLabels?.showCategoryName) return "";
  if (dataLabels.showValue && dataLabels.showCategoryName) return `${category}: ${value}`;
  return dataLabels.showCategoryName ? String(category ?? "") : String(value ?? "");
}

function presentationChartErrorBarsSvg(series, points, plot, max) {
  const errorBars = series.errorBars;
  if (!errorBars || !points.length) return "";
  const numericValues = (series.values || []).map(Number).filter(Number.isFinite);
  const mean = numericValues.reduce((sum, value) => sum + value, 0) / Math.max(1, numericValues.length);
  const deviation = Math.sqrt(numericValues.reduce((sum, value) => sum + (value - mean) ** 2, 0) / Math.max(1, numericValues.length));
  const magnitudeFor = (value, index, side) => errorBars.valueType === "cust" ? Number(errorBars[`${side}Values`]?.[index]) || 0
    : errorBars.valueType === "percentage" ? Math.abs(Number(value) || 0) * (errorBars.value || 0) / 100
    : errorBars.valueType === "stdDev" ? deviation * (errorBars.value || 1)
      : errorBars.valueType === "stdErr" ? deviation / Math.sqrt(Math.max(1, numericValues.length))
        : errorBars.value || 0;
  const attributes = presentationChartLineSvgAttributes(errorBars.line || { fill: series.color || "#475569", width: 1, style: "solid" });
  return points.map((point, index) => {
    const pointIndex = point.index ?? index;
    const scale = (errorBars.direction === "x" ? plot.width : plot.height) / Math.max(1, max);
    const minus = errorBars.type !== "plus" ? magnitudeFor(series.values?.[pointIndex], pointIndex, "minus") * scale : 0;
    const plus = errorBars.type !== "minus" ? magnitudeFor(series.values?.[pointIndex], pointIndex, "plus") * scale : 0;
    const x1 = errorBars.direction === "x" ? point.x - minus : point.x;
    const x2 = errorBars.direction === "x" ? point.x + plus : point.x;
    const y1 = errorBars.direction === "y" ? point.y + minus : point.y;
    const y2 = errorBars.direction === "y" ? point.y - plus : point.y;
    const caps = errorBars.noEndCap ? "" : errorBars.direction === "x"
      ? `${minus > 0 ? `<line x1="${x1}" y1="${point.y - 4}" x2="${x1}" y2="${point.y + 4}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${x2}" y1="${point.y - 4}" x2="${x2}" y2="${point.y + 4}"${attributes}/>` : ""}`
      : `${minus > 0 ? `<line x1="${point.x - 4}" y1="${y1}" x2="${point.x + 4}" y2="${y1}"${attributes}/>` : ""}${plus > 0 ? `<line x1="${point.x - 4}" y1="${y2}" x2="${point.x + 4}" y2="${y2}"${attributes}/>` : ""}`;
    return `<line x1="${x1}" y1="${y1}" x2="${x2}" y2="${y2}"${attributes}/>${caps}`;
  }).join("");
}

export class ChartElement {
  constructor(slide, chartType = "bar", config = {}) {
    this.slide = slide;
    this.id = config.id || aid("ch");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.chartType = String(chartType || config.chartType || "bar").toLowerCase();
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 360, height: 220 });
    this.title = config.title || "";
    this.categories = config.categories || [];
    this.series = normalizeChartSeries(config.series || [], this.chartType);
    this.externalData = normalizePresentationChartExternalData(config.externalData ?? config.sourceWorkbook);
    if (presentationChartUsesFormulaReferences(this) && !this.externalData) throw new TypeError("Presentation chart formula references require externalData with an embedded workbook or external workbook URI.");
    if (this.chartType === "combo" && (!this.series.some((series) => series.chartType === "bar") || !this.series.some((series) => series.chartType === "line"))) throw new TypeError("Presentation combo chart requires at least one bar series and one line series.");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const hasConfiguredSecondaryAxes = Boolean(config.axes?.secondary || config.axes?.secondaryCategory || config.axes?.secondaryValue || config.axes?.y2 || config.secondaryAxisTitles || config.secondaryCategoryAxisTitle || config.secondaryValueAxisTitle || config.secondaryXAxisTitle || config.secondaryYAxisTitle);
    if (hasConfiguredSecondaryAxes && !hasSecondary) throw new TypeError("Presentation secondary axes require at least one chart series with axisGroup secondary.");
    if (hasSecondary && !this.series.some((series) => series.axisGroup !== "secondary")) throw new TypeError("Presentation secondary-axis charts require at least one primary-axis series.");
    this.axes = normalizeChartAxes(config, hasSecondary);
    this.legend = normalizeChartLegend(config, this.series.length);
    this.hasLegend = this.legend.visible;
    this.dataLabels = normalizeChartDataLabels(config);
    Object.assign(this, normalizePresentationChartStyle(this.chartType, config));
  }

  inspectRecord() {
    const p = this.position;
    return { kind: "chart", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, chartType: this.chartType, title: this.title, categories: this.categories, series: this.series.length, seriesDetails: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px" };
  }

  layoutJson() { return { kind: "chart", id: this.id, name: this.name, chartType: this.chartType, title: this.title, frame: this.position, categories: this.categories, series: this.series, axes: this.axes, legend: this.legend, dataLabels: this.dataLabels, externalData: this.externalData ? { embedded: Boolean(this.externalData.bytes), uri: this.externalData.uri, autoUpdate: this.externalData.autoUpdate, bytes: this.externalData.bytes?.byteLength } : undefined, styleId: this.styleId, varyColors: this.varyColors, barOptions: ["bar", "combo"].includes(this.chartType) ? this.barOptions : undefined, lineOptions: ["line", "combo"].includes(this.chartType) ? this.lineOptions : undefined }; }

  toSvg() {
    const p = this.position;
    const categories = this.categories.length ? this.categories : Array.from({ length: Math.max(0, ...this.series.map((series) => series.values?.length || 0)) }, (_, index) => String(index + 1));
    const barSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "bar") : this.chartType === "bar" ? this.series : [];
    const lineSeries = this.chartType === "combo" ? this.series.filter((series) => series.chartType === "line") : this.chartType === "line" ? this.series : [];
    const stackedBars = barSeries.length > 0 && this.barOptions.grouping !== "clustered";
    const stackedLines = lineSeries.length > 0 && this.lineOptions.grouping !== "standard";
    const forAxisGroup = (series, axisGroup) => series.filter((item) => (item.axisGroup || "primary") === axisGroup);
    const stackedTotals = (series) => categories.map((_, categoryIndex) => series.reduce((sum, item) => sum + Math.max(0, Number(item.values?.[categoryIndex]) || 0), 0));
    const barByAxis = { primary: forAxisGroup(barSeries, "primary"), secondary: forAxisGroup(barSeries, "secondary") };
    const lineByAxis = { primary: forAxisGroup(lineSeries, "primary"), secondary: forAxisGroup(lineSeries, "secondary") };
    const barStackedMax = { primary: stackedTotals(barByAxis.primary), secondary: stackedTotals(barByAxis.secondary) };
    const lineStackedMax = { primary: stackedTotals(lineByAxis.primary), secondary: stackedTotals(lineByAxis.secondary) };
    const groupMax = (series, stacked, stackedValues, percentStacked) => percentStacked
      ? 1
      : Math.max(0, ...(stacked ? stackedValues : series.flatMap((item) => item.values || []).map((value) => Math.max(0, Number(value) || 0))));
    const barMax = {
      primary: groupMax(barByAxis.primary, stackedBars, barStackedMax.primary, this.barOptions?.grouping === "percentStacked"),
      secondary: groupMax(barByAxis.secondary, stackedBars, barStackedMax.secondary, this.barOptions?.grouping === "percentStacked"),
    };
    const lineMax = {
      primary: groupMax(lineByAxis.primary, stackedLines, lineStackedMax.primary, this.lineOptions?.grouping === "percentStacked"),
      secondary: groupMax(lineByAxis.secondary, stackedLines, lineStackedMax.secondary, this.lineOptions?.grouping === "percentStacked"),
    };
    const maxForAxisGroup = (axisGroup) => Math.max(
      1,
      barMax[axisGroup],
      lineMax[axisGroup],
    );
    const primaryMax = maxForAxisGroup("primary");
    const secondaryMax = maxForAxisGroup("secondary");
    const hasSecondary = this.series.some((series) => series.axisGroup === "secondary");
    const plot = { left: p.left + 42, top: p.top + 42, width: Math.max(0, p.width - 72), height: Math.max(0, p.height - 82) };
    const title = `<text x="${p.left + 12}" y="${p.top + 24}" font-family="Arial" font-size="16" font-weight="700" fill="#0f172a">${xmlEscape(this.title || this.chartType)}</text>`;
    const axes = `<line x1="${plot.left}" y1="${plot.top + plot.height}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#94a3b8"/><line x1="${plot.left}" y1="${plot.top}" x2="${plot.left}" y2="${plot.top + plot.height}" stroke="#94a3b8"/>${hasSecondary ? `<line x1="${plot.left}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top}" stroke="#64748b"/><line x1="${plot.left + plot.width}" y1="${plot.top}" x2="${plot.left + plot.width}" y2="${plot.top + plot.height}" stroke="#64748b"/>` : ""}${this.axes.category.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${p.top + p.height - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.category.title)}</text>` : ""}${this.axes.value.title ? `<text x="${p.left + 8}" y="${plot.top + 10}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.value.title)}</text>` : ""}${this.axes.secondary?.category?.title ? `<text x="${plot.left + plot.width / 2 - 24}" y="${plot.top - 4}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.category.title)}</text>` : ""}${this.axes.secondary?.value?.title ? `<text x="${plot.left + plot.width - 2}" y="${plot.top + 10}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(this.axes.secondary.value.title)}</text>` : ""}`;
    const legend = this.legend.visible ? this.series.map((series, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.color, series.color))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(series.name)}</text>`).join("") : "";
    if (/^pie$/i.test(this.chartType)) {
      const series = this.series[0] || { values: [] };
      const values = (series.values || []).map((value) => Math.max(0, Number(value) || 0));
      const total = values.reduce((sum, value) => sum + value, 0) || 1;
      const radius = Math.max(8, Math.min(plot.width, plot.height) / 2);
      const cx = plot.left + plot.width / 2;
      const cy = plot.top + plot.height / 2;
      let angle = -Math.PI / 2;
      const slices = values.map((value, index) => {
        const next = angle + (value / total) * Math.PI * 2;
        const point = series.points?.find((item) => item.idx === index);
        const color = resolveColorToken(point?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9");
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labelText = presentationChartDataLabelText(effectiveLabels, categories[index], value);
        const label = labelText ? `<text x="${cx + (radius + 8) * Math.cos((angle + next) / 2)}" y="${cy + (radius + 8) * Math.sin((angle + next) / 2)}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const path = `<path d="${pieSlicePath(cx, cy, radius, angle, next)}" fill="${xmlEscape(color)}"${presentationChartLineSvgAttributes(point?.line || series.line) || ' stroke="#ffffff"'}/>${label}`;
        angle = next;
        return path;
      }).join("");
      const categoryLegend = categories.map((category, index) => `<rect x="${p.left + p.width - 82}" y="${p.top + 18 + index * 16}" width="10" height="10" fill="${xmlEscape(resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || ["#0ea5e9", "#f97316", "#22c55e", "#a855f7"][index % 4], "#0ea5e9"))}"/><text x="${p.left + p.width - 68}" y="${p.top + 27 + index * 16}" font-family="Arial" font-size="10" fill="#334155">${xmlEscape(category)}</text>`).join("");
      return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${slices}${this.legend.visible ? categoryLegend : ""}`;
    }
    const lineBody = lineSeries.map((series, seriesIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const points = (series.values || []).map((value, index) => {
          const stackedValue = stackedLines ? lineSeries.slice(0, seriesIndex + 1).filter((item) => (item.axisGroup || "primary") === axisGroup).reduce((sum, item) => sum + Math.max(0, Number(item.values?.[index]) || 0), 0) : Number(value) || 0;
          const plottedValue = this.lineOptions.grouping === "percentStacked" ? stackedValue / (lineStackedMax[axisGroup][index] || 1) : stackedValue;
          const x = plot.left + (categories.length <= 1 ? plot.width / 2 : (index / Math.max(1, categories.length - 1)) * plot.width);
          const y = plot.top + plot.height - (plottedValue / seriesMax) * plot.height;
          return { x, y, index };
        });
        const color = resolveColorToken(series.line?.fill || series.color, series.color);
        const smooth = series.smooth ?? this.lineOptions.smooth;
        const strokeAttributes = presentationChartLineSvgAttributes(series.line) || ` stroke="${xmlEscape(color)}" stroke-width="2"`;
        const line = smooth && points.length > 2
          ? `<path d="M ${points[0].x} ${points[0].y} ${points.slice(1, -1).map((point, index) => { const next = points[index + 2]; return `Q ${point.x} ${point.y} ${(point.x + next.x) / 2} ${(point.y + next.y) / 2}`; }).join(" ")} T ${points.at(-1).x} ${points.at(-1).y}" fill="none"${strokeAttributes}/>`
          : `<polyline points="${points.map((point) => `${point.x},${point.y}`).join(" ")}" fill="none"${strokeAttributes}/>`;
        const marker = series.marker || this.lineOptions.marker;
        const effectiveLabels = series.dataLabels || this.dataLabels;
        const labels = points.map((point, index) => {
          const label = presentationChartDataLabelText(effectiveLabels, categories[index], series.values?.[index]);
          return label ? `<text x="${point.x + 4}" y="${point.y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(label)}</text>` : "";
        }).join("");
        return `${line}${presentationChartErrorBarsSvg(series, points, plot, seriesMax)}${points.map((point, index) => presentationChartMarkerSvg(marker, point.x, point.y, resolveColorToken(series.points?.find((item) => item.idx === index)?.fill || color, color))).join("")}${labels}`;
      }).join("");
    const horizontal = barSeries.length > 0 && this.barOptions.direction === "bar";
    const barBody = (() => {
      const groupExtent = categories.length ? (horizontal ? plot.height : plot.width) / categories.length : 0;
      const gapRatio = Math.max(0.12, 100 / (100 + this.barOptions.gapWidth));
      const barExtent = stackedBars ? groupExtent * gapRatio : groupExtent * gapRatio / Math.max(1, barSeries.length);
      const offsets = { primary: categories.map(() => 0), secondary: categories.map(() => 0) };
      return barSeries.flatMap((series, seriesIndex) => (series.values || []).map((rawValue, categoryIndex) => {
        const axisGroup = series.axisGroup || "primary";
        const seriesMax = axisGroup === "secondary" ? secondaryMax : primaryMax;
        const total = barStackedMax[axisGroup][categoryIndex] || 1;
        const value = Math.max(0, Number(rawValue) || 0);
        const ratio = this.barOptions.grouping === "percentStacked" ? value / total : value / seriesMax;
        const offset = offsets[axisGroup][categoryIndex];
        offsets[axisGroup][categoryIndex] += ratio;
        const point = series.points?.find((item) => item.idx === categoryIndex);
        const color = xmlEscape(resolveColorToken(point?.fill || series.color, series.color));
        const stroke = presentationChartLineSvgAttributes(point?.line || series.line);
        const labelText = presentationChartDataLabelText(series.dataLabels || this.dataLabels, categories[categoryIndex], rawValue);
        if (horizontal) {
          const width = plot.width * ratio;
          const x = plot.left + (stackedBars ? plot.width * offset : 0);
          const y = plot.top + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
          const label = labelText ? `<text x="${x + width + 3}" y="${y + barExtent - 2}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
          const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + width, y: y + Math.max(1, barExtent - 2) / 2, index: categoryIndex }], plot, seriesMax);
          return `<rect x="${x}" y="${y}" width="${width}" height="${Math.max(1, barExtent - 2)}" fill="${color}"${stroke}/>${errorBars}${label}`;
        }
        const height = plot.height * ratio;
        const x = plot.left + categoryIndex * groupExtent + (stackedBars ? (groupExtent - barExtent) / 2 : (groupExtent - barExtent * barSeries.length) / 2 + seriesIndex * barExtent);
        const y = plot.top + plot.height - height - (stackedBars ? plot.height * offset : 0);
        const label = labelText ? `<text x="${x}" y="${y - 4}" font-family="Arial" font-size="9" fill="#334155">${xmlEscape(labelText)}</text>` : "";
        const errorBars = presentationChartErrorBarsSvg(series, [{ x: x + Math.max(1, barExtent - 2) / 2, y, index: categoryIndex }], plot, seriesMax);
        return `<rect x="${x}" y="${y}" width="${Math.max(1, barExtent - 2)}" height="${height}" fill="${color}"${stroke}/>${errorBars}${label}`;
      })).join("");
    })();
    const trendlineBody = `${barSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length, { horizontal, centered: true })).join("")}${lineSeries.map((series) => presentationChartTrendlinesSvg(series, plot, series.axisGroup === "secondary" ? secondaryMax : primaryMax, categories.length)).join("")}`;
    const body = `${barBody}${lineBody}${trendlineBody}`;
    const labels = this.chartType === "bar" && horizontal
      ? categories.map((category, index) => `<text x="${plot.left - 4}" y="${plot.top + (index + 0.6) * (plot.height / Math.max(1, categories.length))}" text-anchor="end" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("")
      : categories.map((category, index) => `<text x="${plot.left + index * (plot.width / Math.max(1, categories.length))}" y="${p.top + p.height - 18}" font-family="Arial" font-size="10" fill="#475569">${xmlEscape(category)}</text>`).join("");
    return `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" fill="#ffffff" stroke="#cbd5e1"/>${title}${axes}${body}${labels}${legend}`;
  }

}

export class ImageElement {
  constructor(slide, config = {}) {
    this.slide = slide;
    this.id = config.id || aid("im");
    this.nativeId = config.nativeId;
    this.creationId = config.creationId;
    this.name = config.name || "";
    this.position = normalizeFrame(config, { left: 0, top: 0, width: 320, height: 180 });
    this.alt = config.alt || "";
    this.prompt = config.prompt;
    this.uri = config.uri;
    this.dataUrl = config.dataUrl;
    this.contentType = config.contentType;
    this.fit = config.fit || "contain";
    this.crop = config.crop;
    this.geometry = config.geometry || "rect";
    this.borderRadius = config.borderRadius;
    this.transform = config.transform == null ? undefined : normalizePresentationPlaceholderTransform(config.transform, `Presentation image ${this.name || this.id} transform`);
  }

  get frame() { return this.position; }
  set frame(value) { this.position = normalizeFrame(value, this.position); }
  get fit() { return this._fit; }
  set fit(value) { this._fit = normalizePresentationImageFit(value); }
  get crop() { return this._crop; }
  set crop(value) { this._crop = normalizePresentationImageCrop(value); }
  replace(config = {}) { Object.assign(this, config); }

  inspectRecord() {
    const p = this.position;
    return { kind: "image", id: this.id, slide: this.slide.index + 1, name: this.name || undefined, nativeId: this.nativeId, creationId: this.creationId, alt: this.alt || undefined, prompt: this.prompt || undefined, bbox: [p.left, p.top, p.width, p.height], bboxUnit: "px", fit: this.fit, crop: this.crop, transform: this.transform };
  }

  layoutJson() { return { kind: "image", id: this.id, name: this.name, frame: this.position, alt: this.alt, prompt: this.prompt, uri: this.uri, dataUrl: this.dataUrl, fit: this.fit, crop: this.crop, geometry: this.geometry, borderRadius: this.borderRadius, transform: this.transform }; }

  toSvg() {
    const p = this.position;
    const label = this.alt || this.prompt || this.uri || "image";
    const cx = p.left + p.width / 2;
    const cy = p.top + p.height / 2;
    const rotation = Number(this.transform?.rotationDegrees || 0);
    const flipHorizontal = this.transform?.flipHorizontal === true ? -1 : 1;
    const flipVertical = this.transform?.flipVertical === true ? -1 : 1;
    const transform = this.transform ? ` transform="translate(${cx} ${cy}) rotate(${rotation}) scale(${flipHorizontal} ${flipVertical}) translate(${-cx} ${-cy})"` : "";
    if (this.dataUrl) {
      const viewport = presentationImageCropViewport({ crop: this.crop, fit: this.fit, dataUrl: this.dataUrl, frame: p });
      if (viewport) {
        const cropped = `<svg x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" viewBox="${viewport.x} ${viewport.y} ${viewport.width} ${viewport.height}" preserveAspectRatio="none" overflow="hidden"><image href="${attrEscape(this.dataUrl)}" x="0" y="0" width="${viewport.imageWidth}" height="${viewport.imageHeight}" preserveAspectRatio="none"/></svg>`;
        return transform ? `<g${transform}>${cropped}</g>` : cropped;
      }
      const aspect = this.fit === "cover" ? "xMidYMid slice" : this.fit === "stretch" ? "none" : "xMidYMid meet";
      return `<image href="${attrEscape(this.dataUrl)}" x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" preserveAspectRatio="${aspect}"${transform}/>`;
    }
    const rect = this.geometry === "ellipse"
      ? `<ellipse cx="${p.left + p.width / 2}" cy="${p.top + p.height / 2}" rx="${p.width / 2}" ry="${p.height / 2}" fill="#e0f2fe" stroke="#0284c7"/>`
      : `<rect x="${p.left}" y="${p.top}" width="${p.width}" height="${p.height}" rx="${this.borderRadius ? 12 : 0}" fill="#e0f2fe" stroke="#0284c7"/>`;
    const fallback = `${rect}<text x="${p.left + 12}" y="${p.top + 28}" font-family="Arial" font-size="14" fill="#075985">${xmlEscape(label)}</text>`;
    return transform ? `<g${transform}>${fallback}</g>` : fallback;
  }

}

export class PresentationFile {
  static async inspectPptx(blobOrBuffer, options = {}) {
    return inspectOoxmlPackage(blobOrBuffer, options, PPTX_PACKAGE_CONFIG);
  }

  static async patchPptx(blobOrBuffer, patches = [], options = {}) {
    const patched = await patchOoxmlPackage(blobOrBuffer, patches, options, PPTX_PACKAGE_CONFIG);
    return new FileBlob(patched.bytes, { type: PPTX_MIME, metadata: { artifactKind: "presentation", patchedParts: patched.patchedParts, recipesApplied: patched.recipesApplied, contentTypesUpdated: patched.contentTypesUpdated, relationshipsUpdated: patched.relationshipsUpdated, sourceReferencesUpdated: patched.sourceReferencesUpdated, validated: patched.validated, validationIssues: patched.validationIssues } });
  }

  static async exportPptx(presentation, options = {}) {
    const { exportPptxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return exportPptxWithOpenChestnut(presentation, options);
  }

  static async importPptx(blobOrBuffer, options = {}) {
    const { importPptxWithOpenChestnut } = await import("../codecs/open-chestnut.mjs");
    return importPptxWithOpenChestnut(blobOrBuffer, options);
  }
}

function presentationElementKind(element) {
  if (element instanceof NativePresentationObject) return "nativeObject";
  if (element instanceof ConnectorElement) return "connector";
  if (element instanceof GroupShape) return "groupShape";
  if (element instanceof TableElement) return "table";
  if (element instanceof ChartElement) return "chart";
  if (element instanceof ImageElement) return "image";
  return "shape";
}

function presentationSlideElements(slide) {
  const direct = [...slide.connectors.items, ...slide.shapes.items, ...slide.tables.items, ...slide.charts.items, ...slide.images.items, ...slide.groups.items, ...slide.nativeObjects.items];
  return direct.flatMap((element) => element instanceof GroupShape ? element.allElements() : [element]);
}
