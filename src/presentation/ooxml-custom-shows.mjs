import { attributes, attrEscape, decodeXml } from "../ooxml/source-reference-xml.mjs";

const MAX_CUSTOM_SHOWS = 4096;
const MAX_CUSTOM_SHOW_SLIDES = 16384;
const MAX_CUSTOM_SHOW_NAME = 255;
const MAX_NATIVE_ID = 4_294_967_295;

function normalizeName(value) {
  const name = String(value || "").trim();
  if (!name || name.length > MAX_CUSTOM_SHOW_NAME) throw new RangeError(`Presentation custom show name must contain 1 through ${MAX_CUSTOM_SHOW_NAME} characters.`);
  if (/[\u0000-\u001f\u007f]/.test(name)) throw new TypeError("Presentation custom show name must not contain control characters.");
  return name;
}

function normalizeNativeId(value) {
  if (value == null) return undefined;
  const nativeId = Number(value);
  if (!Number.isInteger(nativeId) || nativeId < 0 || nativeId > MAX_NATIVE_ID) throw new RangeError("Presentation custom show nativeId must be an unsigned 32-bit integer.");
  return nativeId;
}

function normalizeSlideId(presentation, value) {
  const slide = typeof value === "number" ? presentation.slides.getItem(value) : value && typeof value === "object" ? value : presentation.slides.items.find((candidate) => candidate.id === String(value));
  if (!slide || slide.presentation !== presentation || !presentation.slides.items.includes(slide)) throw new Error(`Presentation custom show references missing slide ${value?.id || value}.`);
  return slide.id;
}

function normalizeSlideIds(presentation, value) {
  if (!Array.isArray(value) || value.length < 1) throw new RangeError("Presentation custom show requires at least one slide.");
  if (value.length > MAX_CUSTOM_SHOW_SLIDES) throw new RangeError(`Presentation custom show exceeds ${MAX_CUSTOM_SHOW_SLIDES} slides.`);
  return value.map((slide) => normalizeSlideId(presentation, slide));
}

export class PresentationCustomShow {
  constructor(presentation, config = {}, modelId) {
    this.presentation = presentation;
    this.id = String(config.id || modelId);
    this.name = normalizeName(config.name);
    this.nativeId = normalizeNativeId(config.nativeId);
    this.slideIds = normalizeSlideIds(presentation, config.slideIds || config.slides);
  }

  setSlides(slides) {
    this.slideIds = normalizeSlideIds(this.presentation, slides);
    return this;
  }

  get slides() { return this.slideIds.map((id) => this.presentation.slides.items.find((slide) => slide.id === id)); }
  inspectRecord() { return { kind: "customShow", id: this.id, name: this.name, nativeId: this.nativeId, slideIds: [...this.slideIds], slides: this.slideIds.length }; }
  toJSON() { return { id: this.id, name: this.name, nativeId: this.nativeId, slideIds: [...this.slideIds] }; }
}

export class PresentationCustomShowCollection {
  constructor(presentation) {
    this.presentation = presentation;
    this.items = [];
  }

  add(nameOrConfig, slides) {
    if (this.items.length >= MAX_CUSTOM_SHOWS) throw new RangeError(`Presentation custom shows exceed ${MAX_CUSTOM_SHOWS} entries.`);
    const config = typeof nameOrConfig === "string" ? { name: nameOrConfig, slides } : nameOrConfig;
    if (!config || typeof config !== "object" || Array.isArray(config)) throw new TypeError("Presentation custom show must be a name/config object plus slides.");
    const name = normalizeName(config.name);
    if (this.items.some((show) => show.name.toLowerCase() === name.toLowerCase())) throw new Error(`Presentation custom show name ${name} already exists.`);
    const show = new PresentationCustomShow(this.presentation, { ...config, name }, `custom-show/${this.items.length + 1}`);
    if (this.items.some((candidate) => candidate.id === show.id)) throw new Error(`Presentation custom show ID ${show.id} already exists.`);
    if (show.nativeId != null && this.items.some((candidate) => candidate.nativeId === show.nativeId)) throw new Error(`Presentation custom show nativeId ${show.nativeId} already exists.`);
    this.items.push(show);
    return show;
  }

  getItem(idOrNameOrIndex) {
    if (typeof idOrNameOrIndex === "number") return this.items[idOrNameOrIndex];
    const value = String(idOrNameOrIndex || "");
    return this.items.find((show) => show.id === value || show.name === value);
  }

  get count() { return this.items.length; }
  [Symbol.iterator]() { return this.items[Symbol.iterator](); }
}

export function planPresentationCustomShows(presentation) {
  const usedNativeIds = new Set();
  let nextNativeId = 0;
  const entries = presentation.customShows.items.map((show) => {
    show.name = normalizeName(show.name);
    show.slideIds = normalizeSlideIds(presentation, show.slideIds);
    const candidate = normalizeNativeId(show.nativeId);
    if (candidate != null && !usedNativeIds.has(candidate)) {
      show.nativeId = candidate;
      usedNativeIds.add(candidate);
    } else {
      while (usedNativeIds.has(nextNativeId)) nextNativeId += 1;
      show.nativeId = nextNativeId;
      usedNativeIds.add(nextNativeId);
    }
    return show;
  });
  const names = new Set();
  for (const show of entries) {
    const key = show.name.toLowerCase();
    if (names.has(key)) throw new Error(`Presentation custom show name ${show.name} already exists.`);
    names.add(key);
  }
  return {
    entries,
    idByName: new Map(entries.map((show) => [show.name, show.nativeId])),
    nameById: new Map(entries.map((show) => [show.nativeId, show.name])),
  };
}

export function presentationCustomShowsXml(plan, relationshipIdBySlideId = new Map()) {
  if (!plan?.entries?.length) return "";
  const shows = plan.entries.map((show) => {
    const slides = show.slideIds.map((slideId) => {
      const relationshipId = relationshipIdBySlideId.get(slideId);
      if (!relationshipId) throw new Error(`Presentation custom show ${show.name} references missing slide ${slideId}.`);
      return `<p:sld r:id="${attrEscape(relationshipId)}"/>`;
    }).join("");
    return `<p:custShow name="${attrEscape(show.name)}" id="${show.nativeId}"><p:sldLst>${slides}</p:sldLst></p:custShow>`;
  }).join("");
  return `<p:custShowLst>${shows}</p:custShowLst>`;
}

export function parsePresentationCustomShowsXml(xml = "", context = {}) {
  const list = /<(?:[A-Za-z_][\w.-]*:)?custShowLst\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?custShowLst>/.exec(String(xml))?.[0];
  if (!list) return [];
  const relationships = context.relationships || [];
  const shows = [...list.matchAll(/<(?:[A-Za-z_][\w.-]*:)?custShow\b[^>]*>[\s\S]*?<\/(?:[A-Za-z_][\w.-]*:)?custShow>/g)].map((match, index) => {
    const opening = /<(?:[A-Za-z_][\w.-]*:)?custShow\b[^>]*>/.exec(match[0])?.[0] || "";
    const attrs = attributes(opening);
    const name = normalizeName(decodeXml(attrs.name));
    const nativeId = normalizeNativeId(attrs.id);
    if (nativeId == null) throw new Error(`Presentation custom show ${name} is missing its native ID.`);
    const slideIds = [...match[0].matchAll(/<(?:[A-Za-z_][\w.-]*:)?sld\b[^>]*\/?\s*>/g)].map((slideMatch) => {
      const slideAttrs = attributes(slideMatch[0]);
      const relationshipId = Object.entries(slideAttrs).find(([key]) => key === "id" || key.endsWith(":id"))?.[1];
      if (!relationshipId) throw new Error(`Presentation custom show ${name} contains a slide without a relationship ID.`);
      const relationship = relationships.find((item) => item.id === relationshipId);
      if (!relationship) throw new Error(`Presentation custom show ${name} references missing relationship ${relationshipId}.`);
      if (!String(relationship.type || "").endsWith("/slide") || String(relationship.targetMode || "").toLowerCase() === "external") throw new Error(`Presentation custom show ${name} relationship ${relationshipId} must target an internal slide.`);
      const targetPart = context.resolveTarget?.(context.partPath || "ppt/presentation.xml", relationship.target);
      const slideId = context.slideIdByPart?.get(targetPart);
      if (!slideId) throw new Error(`Presentation custom show ${name} targets missing slide part ${targetPart || relationship.target}.`);
      return slideId;
    });
    if (!slideIds.length) throw new Error(`Presentation custom show ${name} requires at least one slide.`);
    return { id: `custom-show/${index + 1}`, name, nativeId, slideIds };
  });
  const names = new Set();
  const ids = new Set();
  for (const show of shows) {
    const name = show.name.toLowerCase();
    if (names.has(name)) throw new Error(`Presentation custom show name ${show.name} already exists.`);
    if (ids.has(show.nativeId)) throw new Error(`Presentation custom show nativeId ${show.nativeId} already exists.`);
    names.add(name);
    ids.add(show.nativeId);
  }
  return shows;
}
