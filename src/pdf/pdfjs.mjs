import { deflateSync } from "node:zlib";

const DEFAULT_PAGE_SIZE = { width: 612, height: 792 };
const encoder = new TextEncoder();

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (ArrayBuffer.isView(data)) return new Uint8Array(data.buffer, data.byteOffset, data.byteLength);
  throw new Error("PDF.js parser requires binary PDF data.");
}

async function readBytes(input) {
  if (input?.bytes) return toUint8Array(input.bytes);
  if (input?.input && typeof input.input.arrayBuffer === "function") return new Uint8Array(await input.input.arrayBuffer());
  if (input?.source && typeof input.source.arrayBuffer === "function") return new Uint8Array(await input.source.arrayBuffer());
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  return toUint8Array(input);
}

async function loadPdfjs(options = {}) {
  if (options.pdfjs?.getDocument) return options.pdfjs;
  try {
    return await import("pdfjs-dist/legacy/build/pdf.mjs");
  } catch (legacyError) {
    try {
      return await import("pdfjs-dist");
    } catch (error) {
      throw new Error(
        `PDF.js parser requires the optional peer dependency \"pdfjs-dist\". Install it with \"npm install -D pdfjs-dist\". Original error: ${error.message || legacyError.message}`,
      );
    }
  }
}

function normalizeTextItem(item, pageHeight, index) {
  const transform = item.transform || [1, 0, 0, 1, Number(item.x || 0), Number(item.y || 0)];
  const x = Number(transform[4] ?? item.x ?? 0);
  const rawY = Number(transform[5] ?? item.y ?? 0);
  const width = Number(item.width || 0);
  const height = Number(item.height || Math.abs(transform[3] || 0) || 0);
  const top = Math.max(0, pageHeight - rawY - height);
  return {
    id: `txt/${index + 1}`,
    text: String(item.str || item.text || ""),
    bbox: [x, top, width, height],
    fontName: item.fontName,
    dir: item.dir,
  };
}

function buildLines(textItems) {
  const rows = [];
  for (const item of [...textItems].sort((a, b) => (a.bbox[1] - b.bbox[1]) || (a.bbox[0] - b.bbox[0]))) {
    if (!item.text.trim()) continue;
    let row = rows.find((candidate) => Math.abs(candidate.top - item.bbox[1]) <= 4);
    if (!row) {
      row = { top: item.bbox[1], items: [] };
      rows.push(row);
    }
    row.items.push(item);
  }
  return rows.map((row) => {
    const items = row.items.sort((a, b) => a.bbox[0] - b.bbox[0]);
    return {
      text: items.map((item) => item.text).join(" ").replace(/\s+/g, " ").trim(),
      items,
      bbox: bboxForItems(items),
    };
  });
}

function bboxForItems(items) {
  if (!items.length) return [0, 0, 0, 0];
  const left = Math.min(...items.map((item) => item.bbox[0]));
  const top = Math.min(...items.map((item) => item.bbox[1]));
  const right = Math.max(...items.map((item) => item.bbox[0] + item.bbox[2]));
  const bottom = Math.max(...items.map((item) => item.bbox[1] + item.bbox[3]));
  return [left, top, Math.max(1, right - left), Math.max(1, bottom - top)];
}

function inferTables(lines, pageIndex) {
  const pipeRows = lines.filter((line) => line.text.includes("|")).map((line) => line.text.split("|").map((cell) => cell.trim()).filter(Boolean));
  if (pipeRows.length >= 2) {
    return [{ name: `pdfjs-pipe-table-${pageIndex + 1}`, values: pipeRows, bbox: bboxForItems(lines.filter((line) => line.text.includes("|")).flatMap((line) => line.items)) }];
  }

  const candidateRows = lines.map((line) => line.items.map((item) => item.text.trim()).filter(Boolean)).filter((row) => row.length >= 2);
  if (candidateRows.length < 2) return [];
  const commonWidth = candidateRows.filter((row) => Math.abs(row.length - candidateRows[0].length) <= 1);
  if (commonWidth.length < 2) return [];
  return [{ name: `pdfjs-position-table-${pageIndex + 1}`, values: commonWidth, bbox: bboxForItems(lines.flatMap((line) => line.items)) }];
}

function multiplyMatrix(left, right) {
  return [
    left[0] * right[0] + left[2] * right[1],
    left[1] * right[0] + left[3] * right[1],
    left[0] * right[2] + left[2] * right[3],
    left[1] * right[2] + left[3] * right[3],
    left[0] * right[4] + left[2] * right[5] + left[4],
    left[1] * right[4] + left[3] * right[5] + left[5],
  ];
}

function imageBbox(matrix, viewport, pageHeight) {
  const points = [[0, 0], [1, 0], [0, 1], [1, 1]].map(([x, y]) => [matrix[0] * x + matrix[2] * y + matrix[4], matrix[1] * x + matrix[3] * y + matrix[5]]);
  const left = Math.min(...points.map(([x]) => x));
  const right = Math.max(...points.map(([x]) => x));
  const bottom = Math.min(...points.map(([, y]) => y));
  const top = Math.max(...points.map(([, y]) => y));
  if (typeof viewport?.convertToViewportRectangle === "function") {
    const rect = viewport.convertToViewportRectangle([left, bottom, right, top]);
    return [Math.min(rect[0], rect[2]), Math.min(rect[1], rect[3]), Math.max(1, Math.abs(rect[2] - rect[0])), Math.max(1, Math.abs(rect[3] - rect[1]))];
  }
  return [left, Math.max(0, pageHeight - top), Math.max(1, right - left), Math.max(1, top - bottom)];
}

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) {
    crc ^= byte;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function pngChunk(type, data = new Uint8Array()) {
  const payload = Buffer.from(data);
  const chunk = Buffer.alloc(12 + payload.length);
  chunk.writeUInt32BE(payload.length, 0);
  Buffer.from(encoder.encode(type)).copy(chunk, 4);
  payload.copy(chunk, 8);
  chunk.writeUInt32BE(crc32(new Uint8Array(chunk.subarray(4, 8 + payload.length))), 8 + payload.length);
  return chunk;
}

function normalizeRgbColor(value) {
  const candidate = Array.isArray(value) && value.length === 1 ? value[0] : value;
  if (typeof candidate === "string") {
    const hex = candidate.match(/^#([0-9a-f]{6})$/i)?.[1];
    if (hex) return [Number.parseInt(hex.slice(0, 2), 16), Number.parseInt(hex.slice(2, 4), 16), Number.parseInt(hex.slice(4, 6), 16)];
    const rgb = candidate.match(/^rgba?\(\s*([\d.]+)[, ]+\s*([\d.]+)[, ]+\s*([\d.]+)/i);
    if (rgb) return rgb.slice(1, 4).map((part) => Math.max(0, Math.min(255, Math.round(Number(part)))));
  }
  const parts = (Array.isArray(candidate) || ArrayBuffer.isView(candidate)) ? [...candidate].slice(0, 3).map(Number) : [];
  if (parts.length === 3 && parts.every(Number.isFinite)) {
    const scale = parts.every((part) => part >= 0 && part <= 1) ? 255 : 1;
    return parts.map((part) => Math.max(0, Math.min(255, Math.round(part * scale))));
  }
  return [0, 0, 0];
}

function encodeRawImagePng(image, options = {}) {
  const width = Number(image?.width);
  const height = Number(image?.height);
  const source = image?.data && toUint8Array(image.data);
  if (!Number.isInteger(width) || !Number.isInteger(height) || width < 1 || height < 1 || !source) throw new Error("PDF.js image object is missing raw geometry or pixels");
  const pixels = width * height;
  const maxPixels = Math.max(1, Number(options.maxImagePixels ?? 20_000_000));
  if (pixels > maxPixels) throw new Error(`PDF.js image has ${pixels} pixels; maxImagePixels is ${maxPixels}`);
  let channels;
  let data = source;
  if (options.isMask) {
    const stride = Math.ceil(width / 8);
    if (source.length < stride * height) throw new Error(`PDF.js image mask buffer is truncated (${source.length} bytes for ${width}x${height})`);
    channels = 4;
    data = new Uint8Array(pixels * channels);
    const color = normalizeRgbColor(options.maskColor);
    for (let y = 0; y < height; y += 1) {
      for (let x = 0; x < width; x += 1) {
        const bit = source[y * stride + (x >> 3)] & (0x80 >> (x & 7));
        const offset = (y * width + x) * channels;
        data[offset] = color[0];
        data[offset + 1] = color[1];
        data[offset + 2] = color[2];
        data[offset + 3] = bit ? 0 : 255;
      }
    }
  } else if (source.length >= pixels * 4) { channels = 4; data = source.subarray(0, pixels * 4); }
  else if (source.length >= pixels * 3) { channels = 3; data = source.subarray(0, pixels * 3); }
  else if (source.length >= pixels) { channels = 1; data = source.subarray(0, pixels); }
  else if (source.length >= Math.ceil(width / 8) * height) {
    channels = 1;
    data = new Uint8Array(pixels);
    const stride = Math.ceil(width / 8);
    for (let y = 0; y < height; y += 1) for (let x = 0; x < width; x += 1) data[y * width + x] = (source[y * stride + (x >> 3)] & (0x80 >> (x & 7))) ? 0 : 255;
  } else throw new Error(`PDF.js image pixel buffer is truncated (${source.length} bytes for ${width}x${height})`);
  const rowBytes = width * channels;
  const raw = Buffer.alloc((rowBytes + 1) * height);
  for (let row = 0; row < height; row += 1) Buffer.from(data.subarray(row * rowBytes, (row + 1) * rowBytes)).copy(raw, row * (rowBytes + 1) + 1);
  const ihdr = Buffer.alloc(13);
  ihdr.writeUInt32BE(width, 0);
  ihdr.writeUInt32BE(height, 4);
  ihdr[8] = 8;
  ihdr[9] = channels === 4 ? 6 : channels === 3 ? 2 : 0;
  return new Uint8Array(Buffer.concat([Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]), pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(raw)), pngChunk("IEND")]));
}

async function resolvePageObject(store, id, timeoutMs = 5_000) {
  if (!store || typeof store.get !== "function" || !id) return undefined;
  return new Promise((resolve, reject) => {
    let settled = false;
    const timer = setTimeout(() => { if (!settled) reject(new Error(`Timed out resolving PDF.js image object ${id}`)); }, timeoutMs);
    const done = (value) => { if (settled) return; settled = true; clearTimeout(timer); resolve(value); };
    try { const value = store.get(id, done); if (value) done(value); } catch (error) { clearTimeout(timer); reject(error); }
  });
}

async function extractImages(page, pdfjs, pageIndex, viewport, width, height, options = {}) {
  const ops = pdfjs.OPS || {};
  const maskOps = new Set([ops.paintImageMaskXObject, ops.paintSolidColorImageMask].filter((value) => value != null));
  const imageOps = new Set([ops.paintImageXObject, ops.paintJpegXObject, ops.paintInlineImageXObject, ...maskOps].filter((value) => value != null));
  if (!imageOps.size || typeof page.getOperatorList !== "function") return [];
  try {
    const operatorList = await page.getOperatorList();
    const saveOp = ops.save;
    const restoreOp = ops.restore;
    const transformOp = ops.transform;
    let matrix = [1, 0, 0, 1, 0, 0];
    let fillColor = [0, 0, 0];
    const stack = [];
    let count = 0;
    const images = [];
    for (let index = 0; index < operatorList.fnArray.length; index += 1) {
      const fn = operatorList.fnArray[index];
      const args = operatorList.argsArray[index] || [];
      if (fn === saveOp) { stack.push({ matrix: [...matrix], fillColor: [...fillColor] }); continue; }
      if (fn === restoreOp) { const state = stack.pop(); matrix = state?.matrix || [1, 0, 0, 1, 0, 0]; fillColor = state?.fillColor || [0, 0, 0]; continue; }
      if (fn === transformOp && args.length >= 6) { matrix = multiplyMatrix(matrix, args.map(Number)); continue; }
      if (fn === ops.setFillRGBColor) { fillColor = normalizeRgbColor(args); continue; }
      if (!imageOps.has(fn)) continue;
      count += 1;
      if (count > Math.max(1, Number(options.maxImagesPerPage ?? 50))) break;
      const bbox = imageBbox(matrix, viewport, height);
      const isMask = maskOps.has(fn);
      const descriptor = args[0];
      const sourceObject = typeof descriptor === "string" ? descriptor : typeof descriptor?.data === "string" ? descriptor.data : undefined;
      const resolved = sourceObject ? await resolvePageObject(page.objs, sourceObject, options.imageObjectTimeoutMs).catch(() => undefined) : undefined;
      const image = fn === ops.paintSolidColorImageMask
        ? { width: 1, height: 1, data: new Uint8Array([0]) }
        : typeof descriptor === "object"
          ? resolved ? { ...descriptor, ...(typeof resolved === "object" ? resolved : {}), data: resolved?.data || resolved } : descriptor
          : resolved;
      try {
        const png = encodeRawImagePng(image, { ...options, isMask, maskColor: fillColor });
        images.push({ name: `pdfjs-image-${pageIndex + 1}-${count}`, alt: `PDF image ${count}`, bbox, bytes: png, contentType: "image/png", sourceObject, sourceOperator: index, pixelWidth: image.width, pixelHeight: image.height, ...(isMask ? { isMask: true, fillColor: `#${fillColor.map((part) => part.toString(16).padStart(2, "0")).join("")}` } : {}) });
      } catch (error) {
        images.push({ name: `pdfjs-image-${pageIndex + 1}-${count}`, alt: `PDF image ${count}`, bbox, prompt: `PDF.js image operator ${index}: ${error.message}` });
      }
    }
    return images;
  } catch {
    return [];
  }
}

export async function parsePdfWithPdfjs(request = {}, defaultOptions = {}) {
  const options = { ...defaultOptions, ...(request.options?.pdfjs || {}), ...(request.pdfjsOptions || {}) };
  const pdfjs = await loadPdfjs(options);
  const bytes = await readBytes(request);
  const loadingTask = pdfjs.getDocument({
    data: bytes,
    disableFontFace: true,
    isEvalSupported: false,
    useSystemFonts: true,
    ...(options.getDocumentOptions || {}),
  });
  const document = await loadingTask.promise;
  const pages = [];

  for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber++) {
    const page = await document.getPage(pageNumber);
    const viewport = page.getViewport?.({ scale: 1 }) || DEFAULT_PAGE_SIZE;
    const width = Number(viewport.width || DEFAULT_PAGE_SIZE.width);
    const height = Number(viewport.height || DEFAULT_PAGE_SIZE.height);
    const textContent = await page.getTextContent({ normalizeWhitespace: true, disableCombineTextItems: false, ...(options.textContentOptions || {}) });
    const textItems = (textContent.items || []).map((item, index) => normalizeTextItem(item, height, index)).filter((item) => item.text);
    const lines = buildLines(textItems);
    const tables = inferTables(lines, pageNumber - 1);
    const images = await extractImages(page, pdfjs, pageNumber - 1, viewport, width, height, options);
    pages.push({
      id: `pdfjs/page/${pageNumber}`,
      width,
      height,
      text: lines.map((line) => line.text).join("\n"),
      textItems,
      regions: lines.map((line, index) => ({ id: `region/${pageNumber}/${index + 1}`, kind: "textLine", label: line.text.slice(0, 80), bbox: line.bbox })),
      tables,
      images,
    });
  }

  await loadingTask.destroy?.().catch?.(() => undefined);
  return { parser: "pdfjs", metadata: { parser: "pdfjs", pages: document.numPages }, pages };
}

export function createPdfjsParser(defaultOptions = {}) {
  return async function pdfjsParserAdapter(request = {}) {
    return parsePdfWithPdfjs(request, defaultOptions);
  };
}

export const pdfjsParser = createPdfjsParser();
