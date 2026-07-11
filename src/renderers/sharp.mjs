import { FileBlob } from "../index.mjs";

const MIME_BY_FORMAT = {
  png: "image/png",
  webp: "image/webp",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
};

function normalizeMime(type = "") {
  return String(type || "").split(";")[0].trim().toLowerCase();
}

function normalizeFormat(format, outputType) {
  const raw = String(format || "").trim().toLowerCase();
  if (raw === "image/png") return "png";
  if (raw === "image/webp") return "webp";
  if (raw === "image/jpeg") return "jpeg";
  if (raw === "jpg") return "jpeg";
  if (raw) return raw;
  const type = normalizeMime(outputType);
  return Object.entries(MIME_BY_FORMAT).find(([, mime]) => mime === type)?.[0] || "png";
}

async function readBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new Error("Sharp renderer requires a FileBlob, Blob, ArrayBuffer, or Uint8Array input.");
}

async function loadSharp(options = {}) {
  if (options.sharp) return options.sharp;
  try {
    const mod = await import("sharp");
    return mod.default || mod;
  } catch (error) {
    throw new Error(`Sharp renderer requires the optional peer dependency "sharp". Install it with "npm install -D sharp". Original error: ${error.message}`);
  }
}

export async function renderWithSharp(request = {}, defaultOptions = {}) {
  const input = request.input || request.source;
  const inputType = normalizeMime(request.inputType || input?.type || defaultOptions.inputType || "image/svg+xml");
  if (!input) throw new Error("Sharp renderer requires request.input or request.source.");
  if (!["image/svg+xml", "image/png", "image/jpeg", "image/webp"].includes(inputType)) {
    throw new Error(`Sharp renderer supports SVG, PNG, JPEG, and WebP input, not ${inputType || "unknown"}.`);
  }
  const options = { ...defaultOptions, ...(request.options?.sharp || {}), ...(request.sharp || {}) };
  const format = normalizeFormat(request.format || options.format, request.outputType);
  const outputType = request.outputType || MIME_BY_FORMAT[format];
  if (!outputType || !MIME_BY_FORMAT[format]) throw new Error(`Sharp renderer cannot produce ${request.format || request.outputType || "unknown"}; supported formats are png, webp, and jpeg.`);

  const sharp = await loadSharp(options);
  let pipeline = sharp(Buffer.from(await readBytes(input)), options.inputOptions || {});
  if (options.resize) pipeline = pipeline.resize(options.resize);
  if (options.flatten) pipeline = pipeline.flatten(typeof options.flatten === "object" ? options.flatten : { background: options.background || "white" });
  if (format === "png") pipeline = pipeline.png(options.pngOptions || {});
  else if (format === "webp") pipeline = pipeline.webp(options.webpOptions || {});
  else pipeline = pipeline.jpeg(options.jpegOptions || {});
  const bytes = await pipeline.toBuffer();
  return new FileBlob(bytes, {
    type: outputType,
    metadata: {
      renderer: "sharp",
      artifactKind: request.artifactKind,
      inputType,
      outputType,
      format,
      ...(options.metadata || {}),
    },
  });
}

export function createSharpRenderer(defaultOptions = {}) {
  return async function sharpRendererAdapter(request = {}) {
    return renderWithSharp(request, defaultOptions);
  };
}

export const sharpRenderer = createSharpRenderer();
