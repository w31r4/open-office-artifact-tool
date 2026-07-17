const IMAGE_FITS = new Set(["contain", "cover", "stretch"]);
const MAX_CROP = 1;
const CROP_SCALE = 100_000;

function finite(value, label) {
  const number = Number(value ?? 0);
  if (!Number.isFinite(number)) throw new TypeError(`${label} must be finite.`);
  return number;
}

function validDimension(value) {
  return Number.isFinite(value) && value > 0 && value <= 10_000_000;
}

function dimensions(width, height, label) {
  if (!validDimension(width) || !validDimension(height)) throw new TypeError(`${label} does not expose bounded positive intrinsic dimensions.`);
  return { width, height };
}

function pngDimensions(bytes) {
  if (bytes.length < 24 || !bytes.subarray(0, 8).equals(Buffer.from("89504e470d0a1a0a", "hex"))) return undefined;
  return dimensions(bytes.readUInt32BE(16), bytes.readUInt32BE(20), "Presentation PNG image");
}

function gifDimensions(bytes) {
  if (bytes.length < 10 || !new Set(["GIF87a", "GIF89a"]).has(bytes.subarray(0, 6).toString("ascii"))) return undefined;
  return dimensions(bytes.readUInt16LE(6), bytes.readUInt16LE(8), "Presentation GIF image");
}

function jpegDimensions(bytes) {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return undefined;
  const sof = new Set([0xc0, 0xc1, 0xc2, 0xc3, 0xc5, 0xc6, 0xc7, 0xc9, 0xca, 0xcb, 0xcd, 0xce, 0xcf]);
  let offset = 2;
  while (offset + 3 < bytes.length) {
    while (offset < bytes.length && bytes[offset] !== 0xff) offset += 1;
    while (offset < bytes.length && bytes[offset] === 0xff) offset += 1;
    if (offset >= bytes.length) break;
    const marker = bytes[offset++];
    if (marker === 0xd8 || marker === 0xd9 || marker === 0x01 || (marker >= 0xd0 && marker <= 0xd7)) continue;
    if (offset + 2 > bytes.length) break;
    const length = bytes.readUInt16BE(offset);
    if (length < 2 || offset + length > bytes.length) break;
    if (sof.has(marker)) {
      if (length < 7) break;
      return dimensions(bytes.readUInt16BE(offset + 3), bytes.readUInt16BE(offset + 5), "Presentation JPEG image");
    }
    offset += length;
  }
  throw new TypeError("Presentation JPEG image does not expose intrinsic dimensions in a supported SOF segment.");
}

function svgDimensions(bytes) {
  const source = bytes.toString("utf8").replace(/^\uFEFF/, "");
  const root = /<svg\b([^>]*)>/i.exec(source)?.[1];
  if (root == null) return undefined;
  const numericAttribute = (name) => {
    const match = new RegExp(`\\b${name}\\s*=\\s*(["'])([+]?(?:\\d+(?:\\.\\d*)?|\\.\\d+))(?:px)?\\1`, "i").exec(root);
    return match ? Number(match[2]) : undefined;
  };
  const width = numericAttribute("width");
  const height = numericAttribute("height");
  if (validDimension(width) && validDimension(height)) return { width, height };
  const viewBox = /\bviewBox\s*=\s*(["'])([-+0-9.eE\s,]+)\1/i.exec(root)?.[2]
    ?.trim().split(/[\s,]+/).map(Number);
  if (viewBox?.length === 4) return dimensions(viewBox[2], viewBox[3], "Presentation SVG image");
  throw new TypeError("Presentation SVG image requires bounded numeric width/height or viewBox dimensions for contain/cover fitting.");
}

export function presentationImageDataUrlDimensions(value) {
  const match = /^data:(image\/(?:png|jpe?g|gif|svg\+xml));base64,([A-Za-z0-9+/=\s]+)$/i.exec(String(value || ""));
  if (!match) throw new TypeError("Presentation image contain/cover fitting requires an embedded base64 PNG, JPEG, GIF, or SVG dataUrl.");
  const bytes = Buffer.from(match[2].replace(/\s/g, ""), "base64");
  const type = match[1].toLowerCase().replace("image/jpg", "image/jpeg");
  if (type === "image/png") return pngDimensions(bytes) || dimensions(0, 0, "Presentation PNG image");
  if (type === "image/jpeg") return jpegDimensions(bytes) || dimensions(0, 0, "Presentation JPEG image");
  if (type === "image/gif") return gifDimensions(bytes) || dimensions(0, 0, "Presentation GIF image");
  return svgDimensions(bytes) || dimensions(0, 0, "Presentation SVG image");
}

export function normalizePresentationImageFit(value = "contain") {
  const fit = String(value || "contain");
  if (!IMAGE_FITS.has(fit)) throw new TypeError("Presentation image fit must be contain, cover, or stretch.");
  return fit;
}

export function normalizePresentationImageCrop(value) {
  if (value == null) return undefined;
  if (typeof value !== "object" || Array.isArray(value)) throw new TypeError("Presentation image crop must be an object.");
  const crop = {
    left: finite(value.left, "Presentation image crop.left"),
    top: finite(value.top, "Presentation image crop.top"),
    right: finite(value.right, "Presentation image crop.right"),
    bottom: finite(value.bottom, "Presentation image crop.bottom"),
  };
  if (Object.values(crop).some((edge) => edge < -MAX_CROP || edge > MAX_CROP) || crop.left + crop.right >= 1 || crop.top + crop.bottom >= 1) {
    throw new RangeError("Presentation image crop edges must be between -1 and 1 and opposing sums must remain below 1.");
  }
  return crop;
}

function normalizedFrame(frame = {}) {
  const width = Number(frame.width);
  const height = Number(frame.height);
  if (!validDimension(width) || !validDimension(height)) throw new RangeError("Presentation image contain/cover fitting requires a positive bounded frame.");
  return { width, height };
}

function roundedCrop(crop) {
  const rounded = Object.fromEntries(Object.entries(crop).map(([key, value]) => [key, Math.round(value * CROP_SCALE) / CROP_SCALE]));
  return normalizePresentationImageCrop(rounded);
}

export function effectivePresentationImageCrop({ crop, fit = "contain", dataUrl, frame } = {}) {
  const normalizedFit = normalizePresentationImageFit(fit);
  const manual = normalizePresentationImageCrop(crop);
  if (normalizedFit === "stretch") return manual;
  const image = presentationImageDataUrlDimensions(dataUrl);
  const target = normalizedFrame(frame);
  const result = { ...(manual || { left: 0, top: 0, right: 0, bottom: 0 }) };
  const sourceWidth = 1 - result.left - result.right;
  const sourceHeight = 1 - result.top - result.bottom;
  const sourceAspect = (image.width * sourceWidth) / (image.height * sourceHeight);
  const targetAspect = target.width / target.height;
  if (Math.abs(sourceAspect - targetAspect) > 1e-12) {
    if (normalizedFit === "cover" && sourceAspect > targetAspect) {
      const desired = sourceHeight * targetAspect * image.height / image.width;
      const delta = (sourceWidth - desired) / 2;
      result.left += delta;
      result.right += delta;
    } else if (normalizedFit === "cover") {
      const desired = sourceWidth * image.width / (targetAspect * image.height);
      const delta = (sourceHeight - desired) / 2;
      result.top += delta;
      result.bottom += delta;
    } else if (sourceAspect > targetAspect) {
      const desired = sourceWidth * image.width / (targetAspect * image.height);
      const delta = (desired - sourceHeight) / 2;
      result.top -= delta;
      result.bottom -= delta;
    } else {
      const desired = sourceHeight * targetAspect * image.height / image.width;
      const delta = (desired - sourceWidth) / 2;
      result.left -= delta;
      result.right -= delta;
    }
  }
  const normalized = roundedCrop(result);
  if (manual == null && Object.values(normalized).every((edge) => edge === 0)) return undefined;
  return normalized;
}

export function presentationImageCropToWire(crop) {
  const normalized = normalizePresentationImageCrop(crop);
  if (!normalized) return undefined;
  return {
    leftThousandthPercent: Math.round(normalized.left * CROP_SCALE),
    topThousandthPercent: Math.round(normalized.top * CROP_SCALE),
    rightThousandthPercent: Math.round(normalized.right * CROP_SCALE),
    bottomThousandthPercent: Math.round(normalized.bottom * CROP_SCALE),
  };
}

export function presentationImageCropFromWire(crop) {
  if (!crop) return undefined;
  return normalizePresentationImageCrop({
    left: Number(crop.leftThousandthPercent) / CROP_SCALE,
    top: Number(crop.topThousandthPercent) / CROP_SCALE,
    right: Number(crop.rightThousandthPercent) / CROP_SCALE,
    bottom: Number(crop.bottomThousandthPercent) / CROP_SCALE,
  });
}

export function presentationImageCropViewport({ crop, fit, dataUrl, frame } = {}) {
  const effective = effectivePresentationImageCrop({ crop, fit, dataUrl, frame });
  if (!effective) return undefined;
  const image = presentationImageDataUrlDimensions(dataUrl);
  return {
    x: effective.left * image.width,
    y: effective.top * image.height,
    width: (1 - effective.left - effective.right) * image.width,
    height: (1 - effective.top - effective.bottom) * image.height,
    imageWidth: image.width,
    imageHeight: image.height,
  };
}
