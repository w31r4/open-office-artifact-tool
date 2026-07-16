import { FileBlob } from "../shared/file-blob.mjs";

const MIME_BY_FORMAT = {
  png: "image/png",
  jpeg: "image/jpeg",
  jpg: "image/jpeg",
};

const SUPPORTED_INPUT_TYPES = ["image/svg+xml", "image/png", "image/jpeg", "image/webp"];

function normalizeMime(type = "") {
  return String(type || "").split(";")[0].trim().toLowerCase();
}

function normalizeFormat(format, outputType) {
  const raw = String(format || "").trim().toLowerCase();
  if (raw === "image/png") return "png";
  if (raw === "image/jpeg") return "jpeg";
  if (raw === "jpg") return "jpeg";
  if (raw === "image/webp" || raw === "webp") return "webp";
  if (raw) return raw;
  const type = normalizeMime(outputType);
  if (type === "image/webp") return "webp";
  return Object.entries(MIME_BY_FORMAT).find(([, mime]) => mime === type)?.[0] || "png";
}

async function readBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new Error("Canvas renderer requires a FileBlob, Blob, ArrayBuffer, or Uint8Array input.");
}

async function loadCanvasLib(options = {}) {
  if (options.canvas) return options.canvas;
  try {
    const mod = await import("canvas");
    return mod.default || mod;
  } catch (error) {
    throw new Error(`Canvas renderer requires the optional peer dependency "canvas". Install it with "npm install -D canvas". Original error: ${error.message}`);
  }
}

const NUMBER_WITH_UNIT = /^\s*([\d.]+)\s*(px|pt|in|cm|mm)?\s*$/i;

function svgSizeFromText(text) {
  if (!text) return null;
  const rootMatch = /<svg\b[^>]*>/i.exec(text);
  if (!rootMatch) return null;
  const tag = rootMatch[0];
  const pick = (attr) => {
    const re = new RegExp(`\\s${attr}\\s*=\\s*["']([^"']*)["']`, "i");
    const m = re.exec(tag);
    if (!m) return null;
    const parsed = NUMBER_WITH_UNIT.exec(m[1]);
    if (!parsed) return null;
    const value = Number(parsed[1]);
    return Number.isFinite(value) && value > 0 ? value : null;
  };
  const width = pick("width");
  const height = pick("height");
  let vbW = null;
  let vbH = null;
  const viewBox = /viewBox\s*=\s*["']\s*([\d.\s-]+)["']/i.exec(tag);
  if (viewBox) {
    const parts = viewBox[1].trim().split(/[\s,]+/).map(Number);
    if (parts.length >= 4 && Number.isFinite(parts[2]) && Number.isFinite(parts[3])) {
      vbW = parts[2] > 0 ? parts[2] : null;
      vbH = parts[3] > 0 ? parts[3] : null;
    }
  }
  return { width: width || vbW, height: height || vbH };
}

async function loadImageFrom(canvasLib, bytes, options) {
  if (typeof canvasLib.loadImage === "function") {
    return canvasLib.loadImage(Buffer.from(bytes));
  }
  if (typeof canvasLib.Image === "function") {
    const image = new canvasLib.Image();
    image.src = Buffer.from(bytes);
    return image;
  }
  throw new Error('Canvas renderer library must expose loadImage (or Image) and createCanvas.');
}

export async function renderWithCanvas(request = {}, defaultOptions = {}) {
  const input = request.input || request.source;
  const inputType = normalizeMime(request.inputType || input?.type || defaultOptions.inputType || "image/svg+xml");
  if (!input) throw new Error("Canvas renderer requires request.input or request.source.");
  if (!SUPPORTED_INPUT_TYPES.includes(inputType)) {
    throw new Error(`Canvas renderer supports SVG, PNG, JPEG, and WebP input, not ${inputType || "unknown"}.`);
  }
  const options = { ...defaultOptions, ...(request.options?.canvas || {}), ...(request.canvas || {}) };
  const format = normalizeFormat(request.format || options.format, request.outputType);
  const outputType = request.outputType || MIME_BY_FORMAT[format];
  if (!outputType || !MIME_BY_FORMAT[format]) {
    throw new Error(`Canvas renderer cannot produce ${request.format || request.outputType || "unknown"}; supported formats are png and jpeg. Use the sharp or Playwright renderer for webp or pdf output.`);
  }

  const canvasLib = await loadCanvasLib(options);
  if (typeof canvasLib.createCanvas !== "function") {
    throw new Error('Canvas renderer library must expose createCanvas and loadImage (or Image).');
  }
  const bytes = await readBytes(input);
  const image = await loadImageFrom(canvasLib, bytes, options);
  let width = Number(options.width) || image?.width || image?.naturalWidth;
  let height = Number(options.height) || image?.height || image?.naturalHeight;
  if ((!width || !height) && inputType === "image/svg+xml") {
    const size = svgSizeFromText(new TextDecoder().decode(bytes));
    if (size) {
      width = width || size.width;
      height = height || size.height;
    }
  }
  width = Number.isFinite(width) && width > 0 ? width : 800;
  height = Number.isFinite(height) && height > 0 ? height : 600;

  const canvas = canvasLib.createCanvas(width, height);
  const ctx = canvas.getContext("2d");
  const background = options.background || (format === "jpeg" ? "white" : null);
  if (background && typeof ctx.fillRect === "function") {
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, width, height);
  }
  if (typeof ctx.drawImage === "function") ctx.drawImage(image, 0, 0, width, height);
  const out = canvas.toBuffer(MIME_BY_FORMAT[format], options.outputOptions || {});
  const outBytes = out instanceof Uint8Array ? out : new Uint8Array(out);
  return new FileBlob(outBytes, {
    type: outputType,
    metadata: {
      renderer: "canvas",
      artifactKind: request.artifactKind,
      inputType,
      outputType,
      format,
      width,
      height,
      ...(options.metadata || {}),
    },
  });
}

export function createCanvasRenderer(defaultOptions = {}) {
  return async function canvasRendererAdapter(request = {}) {
    return renderWithCanvas(request, defaultOptions);
  };
}

export const canvasRenderer = createCanvasRenderer();
