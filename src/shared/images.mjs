export function imageDataFromDataUrl(dataUrl) {
  const match = /^data:([^;,]+);base64,(.+)$/i.exec(String(dataUrl || ""));
  if (!match) return undefined;
  const contentType = match[1].toLowerCase();
  const extension = contentType === "image/jpeg" ? "jpg" : contentType.split("/")[1] || "bin";
  return { contentType, extension: extension === "svg+xml" ? "svg" : extension, bytes: Buffer.from(match[2], "base64") };
}

export function imageContentTypeFromExtension(extension) {
  const normalized = String(extension || "").toLowerCase().replace(/^\./, "");
  if (normalized === "jpg" || normalized === "jpeg") return "image/jpeg";
  if (normalized === "png") return "image/png";
  if (normalized === "gif") return "image/gif";
  if (normalized === "svg" || normalized === "svg+xml") return "image/svg+xml";
  return "application/octet-stream";
}
