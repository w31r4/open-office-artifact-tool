const MAX_PICTURE_BULLET_BYTES = 8 * 1024 * 1024;

function pictureBulletData(dataUrl) {
  const match = /^data:(image\/(?:png|jpeg|gif));base64,([A-Za-z0-9+/=]+)$/i.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const contentType = match[1].toLowerCase();
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.split("/")[1];
  const bytes = Buffer.from(match[2], "base64");
  if (!bytes.length || bytes.length > MAX_PICTURE_BULLET_BYTES) throw new RangeError(`DOCX picture bullet must contain 1 through ${MAX_PICTURE_BULLET_BYTES} decoded bytes.`);
  return { contentType, extension, bytes };
}

function boundedPoints(value, fallback, name) {
  const points = Number(value ?? fallback);
  if (!Number.isFinite(points) || points < 4 || points > 72) throw new RangeError(`DOCX picture bullet ${name} must be between 4 and 72 points.`);
  return points;
}

export function normalizeDocumentPictureBullet(value) {
  if (value == null || value === false) return undefined;
  const input = typeof value === "string"
    ? (String(value).startsWith("data:") ? { dataUrl: value } : { uri: value })
    : value;
  if (!input || typeof input !== "object" || Array.isArray(input)) throw new TypeError("DOCX picture bullet must be an embedded image data URL, absolute http(s) URI, or configuration object.");
  const dataUrl = input.dataUrl || input.data || input.src;
  const uri = input.uri || input.url;
  if (Boolean(dataUrl) === Boolean(uri)) throw new TypeError("DOCX picture bullet requires exactly one of dataUrl or uri.");
  if (dataUrl && !pictureBulletData(dataUrl)) throw new TypeError("DOCX picture bullet supports embedded PNG, JPEG, or GIF base64 data URLs.");
  if (uri && !/^https?:\/\//i.test(String(uri))) throw new TypeError("DOCX external picture bullet URI must be absolute http(s).");
  const widthPt = boundedPoints(input.widthPt ?? input.sizePt ?? input.size, 12, "width");
  const heightPt = boundedPoints(input.heightPt ?? input.sizePt ?? input.size, widthPt, "height");
  const alt = String(input.alt || input.description || "Picture bullet").trim().slice(0, 255) || "Picture bullet";
  return { dataUrl: dataUrl ? String(dataUrl) : undefined, uri: uri ? String(uri) : undefined, widthPt, heightPt, alt };
}
