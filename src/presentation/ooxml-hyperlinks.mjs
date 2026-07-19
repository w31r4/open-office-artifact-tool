const PRESENTATION_RUN_ACTIONS = new Set(["nextSlide", "previousSlide", "firstSlide", "lastSlide", "endShow"]);

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


function optionalBoolean(value, name) {
  if (value == null) return undefined;
  if (typeof value !== "boolean") throw new TypeError(`Presentation run hyperlink ${name} must be a boolean.`);
  return value;
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
  const action = PRESENTATION_RUN_ACTIONS.has(raw) ? raw : PRESENTATION_RUN_ACTION_ALIASES.get(raw.toLowerCase());
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
  const customShowValue = input.customShow ?? input.show;
  const targetPartValue = options.allowTargetPart ? input.targetPart : undefined;
  const targets = [uriValue, slideIdValue, actionValue, customShowValue, targetPartValue].filter((item) => item != null && String(item?.name ?? item).trim());
  if (targets.length !== 1) throw new Error("Presentation run hyperlink requires exactly one of uri, slideId, action, or customShow.");
  const tooltip = input.tooltip == null ? undefined : String(input.tooltip);
  if (tooltip != null && tooltip.length > 1024) throw new RangeError("Presentation run hyperlink tooltip exceeds 1024 characters.");
  const targetFrame = input.targetFrame ?? input.tgtFrame;
  if (targetFrame != null && (!String(targetFrame).trim() || String(targetFrame).length > 255)) throw new RangeError("Presentation run hyperlink targetFrame must contain 1 through 255 characters.");
  const history = optionalBoolean(input.history, "history");
  const highlightClick = optionalBoolean(input.highlightClick, "highlightClick");
  const returnToSlide = optionalBoolean(input.returnToSlide ?? input.return, "returnToSlide");
  if (returnToSlide != null && customShowValue == null) throw new Error("Presentation run hyperlink returnToSlide requires customShow.");
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
  if (customShowValue != null) {
    const customShow = String(customShowValue?.name ?? customShowValue).trim();
    if (!customShow || customShow.length > 255) throw new RangeError("Presentation run hyperlink customShow must contain 1 through 255 characters.");
    return { customShow, ...(returnToSlide == null ? {} : { returnToSlide }), ...common };
  }
  const targetPart = String(targetPartValue || "").trim();
  if (!targetPart || targetPart.length > 1024) throw new RangeError("Imported Presentation run hyperlink target part is invalid.");
  return { targetPart, ...common };
}
