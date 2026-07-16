import { toUint8Array } from "./binary.mjs";
import { FileBlob } from "./file-blob.mjs";

export const LAYOUT_MIME = "application/vnd.open-office-artifact.layout+json";

const RENDER_MIME_BY_FORMAT = {
  svg: "image/svg+xml",
  png: "image/png",
  webp: "image/webp",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  pdf: "application/pdf",
  layout: LAYOUT_MIME,
};

export function renderTypeForOptions(options = {}, fallbackType = "application/octet-stream") {
  const format = String(options.format || "").trim().toLowerCase();
  if (!format) return fallbackType;
  return RENDER_MIME_BY_FORMAT[format] || (format.includes("/") ? format : fallbackType);
}

export async function fileBlobFromRenderOutput(output, type, metadata = {}) {
  if (output instanceof FileBlob) {
    output.metadata = { ...(output.metadata || {}), ...metadata };
    return output;
  }
  if (output?.data !== undefined) return fileBlobFromRenderOutput(output.data, output.type || type, { ...metadata, ...(output.metadata || {}) });
  if (output?.arrayBuffer) return new FileBlob(new Uint8Array(await output.arrayBuffer()), { type: output.type || type, metadata });
  return new FileBlob(output instanceof Uint8Array || output instanceof ArrayBuffer || ArrayBuffer.isView(output) ? toUint8Array(output) : String(output ?? ""), { type, metadata });
}
