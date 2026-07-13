import path from "node:path";
import { attributes, attrEscape, decodeXml } from "../ooxml/source-reference-xml.mjs";

export const PRESENTATION_HYPERLINK_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink";
export const PRESENTATION_SLIDE_RELATIONSHIP_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide";

const PRESENTATION_RUN_ACTION_URIS = new Map([
  ["nextSlide", "ppaction://hlinkshowjump?jump=nextslide"],
  ["previousSlide", "ppaction://hlinkshowjump?jump=previousslide"],
  ["firstSlide", "ppaction://hlinkshowjump?jump=firstslide"],
  ["lastSlide", "ppaction://hlinkshowjump?jump=lastslide"],
  ["endShow", "ppaction://hlinkshowjump?jump=endshow"],
]);

const PRESENTATION_RUN_ACTION_ALIASES = new Map([
  ["next", "nextSlide"],
  ["nextslide", "nextSlide"],
  ["previous", "previousSlide"],
  ["prev", "previousSlide"],
  ["previousslide", "previousSlide"],
  ["first", "firstSlide"],
  ["firstslide", "firstSlide"],
  ["last", "lastSlide"],
  ["lastslide", "lastSlide"],
  ["end", "endShow"],
  ["endshow", "endShow"],
]);

const PRESENTATION_RUN_ACTIONS_BY_URI = new Map([...PRESENTATION_RUN_ACTION_URIS].map(([action, uri]) => [uri, action]));

function optionalBoolean(value, name) {
  if (value == null) return undefined;
  if (typeof value !== "boolean") throw new TypeError(`Presentation run hyperlink ${name} must be a boolean.`);
  return value;
}

function relationshipAttribute(attributesByName, localName) {
  return Object.entries(attributesByName).find(([name]) => name === localName || name.endsWith(`:${localName}`))?.[1];
}

function normalizeExternalUri(value) {
  const uri = String(value || "").trim();
  if (!uri) throw new TypeError("Presentation run hyperlink uri must not be empty.");
  if (uri.length > 4096) throw new RangeError("Presentation run hyperlink uri exceeds 4096 characters.");
  if (/[\u0000-\u001f\u007f]/.test(uri)) throw new TypeError("Presentation run hyperlink uri must not contain control characters.");
  if (!/^[A-Za-z][A-Za-z0-9+.-]*:/.test(uri)) throw new TypeError("Presentation run hyperlink uri must be absolute.");
  if (/^(?:javascript|data):/i.test(uri)) throw new TypeError("Presentation run hyperlink uri uses a forbidden scheme.");
  return uri;
}

function normalizePresentationRunAction(value) {
  const raw = String(value || "").trim();
  const action = PRESENTATION_RUN_ACTION_URIS.has(raw) ? raw : PRESENTATION_RUN_ACTION_ALIASES.get(raw.toLowerCase());
  if (!action) throw new RangeError(`Unsupported Presentation run hyperlink action ${raw || "(empty)"}.`);
  return action;
}

export function normalizePresentationRunLink(value, options = {}) {
  if (value == null || value === false) return undefined;
  const input = typeof value === "string" ? { uri: value } : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("Presentation run hyperlink must be an absolute URI or object.");
  const uriValue = input.uri ?? input.url ?? input.href;
  const slideIdValue = input.slideId ?? input.targetSlideId ?? input.targetId;
  const actionValue = input.action ?? input.jump;
  const targetPartValue = options.allowTargetPart ? input.targetPart : undefined;
  const targets = [uriValue, slideIdValue, actionValue, targetPartValue].filter((item) => item != null && String(item).trim());
  if (targets.length !== 1) throw new Error("Presentation run hyperlink requires exactly one of uri, slideId, or action.");
  const tooltip = input.tooltip == null ? undefined : String(input.tooltip);
  if (tooltip != null && tooltip.length > 1024) throw new RangeError("Presentation run hyperlink tooltip exceeds 1024 characters.");
  const targetFrame = input.targetFrame ?? input.tgtFrame;
  if (targetFrame != null && (!String(targetFrame).trim() || String(targetFrame).length > 255)) throw new RangeError("Presentation run hyperlink targetFrame must contain 1 through 255 characters.");
  const history = optionalBoolean(input.history, "history");
  const highlightClick = optionalBoolean(input.highlightClick, "highlightClick");
  const common = {
    ...(tooltip == null ? {} : { tooltip }),
    ...(targetFrame == null ? {} : { targetFrame: String(targetFrame) }),
    ...(history == null ? {} : { history }),
    ...(highlightClick == null ? {} : { highlightClick }),
  };
  if (uriValue != null) return { uri: normalizeExternalUri(uriValue), ...common };
  if (slideIdValue != null) {
    const slideId = String(slideIdValue).trim();
    if (!slideId || slideId.length > 512) throw new RangeError("Presentation run hyperlink slideId must contain 1 through 512 characters.");
    return { slideId, ...common };
  }
  if (actionValue != null) return { action: normalizePresentationRunAction(actionValue), ...common };
  const targetPart = String(targetPartValue || "").trim();
  if (!targetPart || targetPart.length > 1024) throw new RangeError("Imported Presentation run hyperlink target part is invalid.");
  return { targetPart, ...common };
}

export function presentationRunHyperlinkKey(link, slidePartById = new Map()) {
  if (!link) return undefined;
  if (link.action) return undefined;
  if (link.uri) return `external:${link.uri}`;
  const targetPart = link.targetPart || slidePartById.get(link.slideId);
  return targetPart ? `slide:${targetPart}` : undefined;
}

export function presentationRunHyperlinkReferencesFromParagraphs(paragraphs = []) {
  return paragraphs.flatMap((paragraph) => (paragraph?.runs || []).map((run) => run?.link).filter(Boolean));
}

export function planPresentationRunHyperlinks(options = {}) {
  const slidePartById = options.slidePartById || new Map();
  const byOwner = new Map();
  for (const owner of options.owners || []) {
    let nextRelationshipIndex = Number(owner.startRelationshipIndex || 1);
    if (!Number.isInteger(nextRelationshipIndex) || nextRelationshipIndex < 1) throw new RangeError("Presentation hyperlink relationship index must be a positive integer.");
    const relationshipIds = new Map();
    const relationships = [];
    for (const link of owner.references || []) {
      if (link.action) continue;
      const key = presentationRunHyperlinkKey(link, slidePartById);
      if (!key) throw new Error(`Presentation run hyperlink references missing slide ${link.slideId || "(unknown)"}.`);
      if (relationshipIds.has(key)) continue;
      const id = `rId${nextRelationshipIndex++}`;
      relationshipIds.set(key, id);
      if (link.uri) {
        relationships.push({ id, type: PRESENTATION_HYPERLINK_RELATIONSHIP_TYPE, target: link.uri, targetMode: "External" });
      } else {
        const targetPart = link.targetPart || slidePartById.get(link.slideId);
        const target = path.posix.relative(path.posix.dirname(owner.partPath), targetPart) || path.posix.basename(targetPart);
        relationships.push({ id, type: PRESENTATION_SLIDE_RELATIONSHIP_TYPE, target });
      }
    }
    byOwner.set(owner.key, { relationshipIds, relationships, nextRelationshipIndex, slidePartById });
  }
  return { byOwner, slidePartById };
}

export function parsePresentationRunLinkXml(runXml = "", context = {}) {
  const opening = /<(?:[A-Za-z_][\w.-]*:)?hlinkClick\b[^>]*\/?\s*>/.exec(String(runXml))?.[0];
  if (!opening) return undefined;
  const linkAttributes = attributes(opening);
  const relationshipId = relationshipAttribute(linkAttributes, "id");
  const actionUri = linkAttributes.action == null ? undefined : decodeXml(linkAttributes.action);
  const action = PRESENTATION_RUN_ACTIONS_BY_URI.get(actionUri);
  if (action) {
    if (relationshipId) throw new Error(`Presentation run action ${action} must not reference relationship ${relationshipId}.`);
    return normalizePresentationRunLink({
      action,
      ...(linkAttributes.tooltip == null ? {} : { tooltip: decodeXml(linkAttributes.tooltip) }),
      ...(linkAttributes.tgtFrame == null ? {} : { targetFrame: decodeXml(linkAttributes.tgtFrame) }),
      ...(linkAttributes.history == null ? {} : { history: ["1", "true", "on"].includes(String(linkAttributes.history).toLowerCase()) }),
      ...(linkAttributes.highlightClick == null ? {} : { highlightClick: ["1", "true", "on"].includes(String(linkAttributes.highlightClick).toLowerCase()) }),
    });
  }
  if (actionUri && !relationshipId) throw new RangeError(`Unsupported Presentation run hyperlink action ${actionUri}.`);
  if (!relationshipId) throw new Error("Presentation run hyperlink is missing its relationship ID.");
  const relationship = (context.relationships || []).find((item) => item.id === relationshipId);
  if (!relationship) throw new Error(`Presentation run hyperlink references missing relationship ${relationshipId}.`);
  const relationshipType = String(relationship.type || "");
  let target;
  if (relationshipType.endsWith("/hyperlink")) {
    if (String(relationship.targetMode || "").toLowerCase() !== "external") throw new Error(`Presentation run hyperlink relationship ${relationshipId} must be external.`);
    target = { uri: relationship.target };
  } else if (relationshipType.endsWith("/slide")) {
    if (String(relationship.targetMode || "").toLowerCase() === "external") throw new Error(`Presentation run slide hyperlink relationship ${relationshipId} must be internal.`);
    const targetPart = context.resolveTarget?.(context.partPath, relationship.target);
    if (!targetPart) throw new Error(`Presentation run hyperlink relationship ${relationshipId} has an invalid slide target.`);
    const slideId = context.slideIdByPart?.get(targetPart);
    if (context.slideIdByPart && !slideId) throw new Error(`Presentation run hyperlink targets missing slide part ${targetPart}.`);
    target = slideId ? { slideId } : { targetPart };
  } else {
    throw new Error(`Presentation run hyperlink relationship ${relationshipId} has unsupported type ${relationshipType || "(missing)"}.`);
  }
  return normalizePresentationRunLink({
    ...target,
    ...(linkAttributes.tooltip == null ? {} : { tooltip: decodeXml(linkAttributes.tooltip) }),
    ...(linkAttributes.tgtFrame == null ? {} : { targetFrame: decodeXml(linkAttributes.tgtFrame) }),
    ...(linkAttributes.history == null ? {} : { history: ["1", "true", "on"].includes(String(linkAttributes.history).toLowerCase()) }),
    ...(linkAttributes.highlightClick == null ? {} : { highlightClick: ["1", "true", "on"].includes(String(linkAttributes.highlightClick).toLowerCase()) }),
  }, { allowTargetPart: true });
}

export function presentationRunHyperlinkXml(link, relationshipId) {
  if (!link) return "";
  if (link.action) {
    const actionUri = PRESENTATION_RUN_ACTION_URIS.get(link.action);
    if (!actionUri) throw new RangeError(`Unsupported Presentation run hyperlink action ${link.action}.`);
    const attributesXml = `${link.tooltip == null ? "" : ` tooltip="${attrEscape(link.tooltip)}"`}${link.targetFrame == null ? "" : ` tgtFrame="${attrEscape(link.targetFrame)}"`}${link.history == null ? "" : ` history="${link.history ? 1 : 0}"`}${link.highlightClick == null ? "" : ` highlightClick="${link.highlightClick ? 1 : 0}"`}`;
    return `<a:hlinkClick r:id="" action="${attrEscape(actionUri)}"${attributesXml}/>`;
  }
  if (!relationshipId) throw new Error("Presentation run hyperlink has no relationship in its owning OOXML part.");
  const attributesXml = `${link.slideId || link.targetPart ? ' action="ppaction://hlinksldjump"' : ""}${link.tooltip == null ? "" : ` tooltip="${attrEscape(link.tooltip)}"`}${link.targetFrame == null ? "" : ` tgtFrame="${attrEscape(link.targetFrame)}"`}${link.history == null ? "" : ` history="${link.history ? 1 : 0}"`}${link.highlightClick == null ? "" : ` highlightClick="${link.highlightClick ? 1 : 0}"`}`;
  return `<a:hlinkClick r:id="${attrEscape(relationshipId)}"${attributesXml}/>`;
}

export function resolvePresentationRunHyperlinkTargets(paragraphs = [], slideIdByPart = new Map()) {
  return paragraphs.map((paragraph) => ({
    ...paragraph,
    runs: (paragraph.runs || []).map((run) => {
      if (!run.link?.targetPart) return run;
      const slideId = slideIdByPart.get(run.link.targetPart);
      if (!slideId) throw new Error(`Presentation run hyperlink targets missing slide part ${run.link.targetPart}.`);
      const { targetPart: _targetPart, ...rest } = run.link;
      return { ...run, link: { slideId, ...rest } };
    }),
  }));
}
