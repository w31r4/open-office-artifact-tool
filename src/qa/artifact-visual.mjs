import { deflateSync } from "node:zlib";

import { decoder, encoder } from "../shared/binary.mjs";
import { FileBlob } from "../shared/file-blob.mjs";
import { ndjson, verificationIssue } from "../shared/inspection.mjs";
import { decodePngRgba, isPngBytes } from "../shared/png.mjs";
import { fileBlobFromRenderOutput, renderTypeForOptions } from "../shared/render-output.mjs";

export function createArtifactVisualQaApi({ inferArtifactKind }) {
  if (typeof inferArtifactKind !== "function") throw new TypeError("createArtifactVisualQaApi requires inferArtifactKind.");

  function attachRenderMetadata(blob, artifactKind, options = {}, format = options.format || blob.type) {
    blob.metadata = {
      ...(blob.metadata || {}),
      artifactKind,
      format,
      page: options.page,
      pageIndex: options.pageIndex,
      slide: options.slide,
      sheetName: options.sheetName,
      range: options.range,
    };
    return blob;
  }

  function stableByteHash(bytes) {
    let hash = 2166136261;
    for (const byte of bytes || []) {
      hash ^= byte;
      hash = Math.imul(hash, 16777619) >>> 0;
    }
    return hash.toString(16).padStart(8, "0");
  }

  async function bytesForVisualBaseline(input) {
    if (input == null) return undefined;
    if (input instanceof FileBlob) return input.bytes;
    if (typeof input === "string") return new TextEncoder().encode(input);
    if (input instanceof Uint8Array) return input;
    if (input instanceof ArrayBuffer) return new Uint8Array(input);
    if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
    if (typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
    return undefined;
  }

  function svgDimensions(svgText = "") {
    const open = String(svgText || "").match(/<svg\b[^>]*>/i)?.[0] || "";
    const width = Number.parseFloat(/\bwidth=["']([^"']+)/i.exec(open)?.[1] || "");
    const height = Number.parseFloat(/\bheight=["']([^"']+)/i.exec(open)?.[1] || "");
    const viewBox = /\bviewBox=["']([^"']+)["']/i.exec(open)?.[1]?.trim().split(/[\s,]+/).map(Number);
    return {
      width: Number.isFinite(width) ? width : Number.isFinite(viewBox?.[2]) ? viewBox[2] : undefined,
      height: Number.isFinite(height) ? height : Number.isFinite(viewBox?.[3]) ? viewBox[3] : undefined,
    };
  }

  function isPpmBytes(bytes) {
    return bytes?.[0] === 0x50 && (bytes?.[1] === 0x36 || bytes?.[1] === 0x33);
  }

  function ppmTokens(bytes) {
    const tokens = [];
    let i = 0;
    while (i < bytes.length) {
      while (i < bytes.length && /\s/.test(String.fromCharCode(bytes[i]))) i += 1;
      if (bytes[i] === 0x23) { while (i < bytes.length && bytes[i] !== 0x0a) i += 1; continue; }
      if (i >= bytes.length) break;
      const start = i;
      while (i < bytes.length && !/\s/.test(String.fromCharCode(bytes[i])) && bytes[i] !== 0x23) i += 1;
      tokens.push({ text: decoder.decode(bytes.slice(start, i)), end: i });
      if (tokens.length >= 4) break;
    }
    return tokens;
  }

  function decodePpmRgba(bytes) {
    if (!isPpmBytes(bytes)) throw new Error("not a PPM file");
    const tokens = ppmTokens(bytes);
    const magic = tokens[0]?.text;
    const width = Number(tokens[1]?.text);
    const height = Number(tokens[2]?.text);
    const max = Number(tokens[3]?.text);
    if (!Number.isFinite(width) || !Number.isFinite(height) || width <= 0 || height <= 0 || max <= 0) throw new Error("PPM is missing valid geometry or max value");
    const rgba = new Uint8Array(width * height * 4);
    if (magic === "P6") {
      let offset = tokens[3].end;
      while (offset < bytes.length && /\s/.test(String.fromCharCode(bytes[offset]))) offset += 1;
      const needed = width * height * 3;
      if (bytes.length - offset < needed) throw new Error("PPM image data is truncated");
      for (let p = 0; p < width * height; p += 1) {
        const input = offset + p * 3;
        const output = p * 4;
        rgba[output] = Math.round(bytes[input] * 255 / max);
        rgba[output + 1] = Math.round(bytes[input + 1] * 255 / max);
        rgba[output + 2] = Math.round(bytes[input + 2] * 255 / max);
        rgba[output + 3] = 255;
      }
    } else if (magic === "P3") {
      const text = decoder.decode(bytes.slice(tokens[3].end));
      const values = text.replace(/#[^\n\r]*/g, " ").trim().split(/\s+/).filter(Boolean).map(Number);
      if (values.length < width * height * 3) throw new Error("PPM image data is truncated");
      for (let p = 0; p < width * height; p += 1) {
        const input = p * 3;
        const output = p * 4;
        rgba[output] = Math.round(values[input] * 255 / max);
        rgba[output + 1] = Math.round(values[input + 1] * 255 / max);
        rgba[output + 2] = Math.round(values[input + 2] * 255 / max);
        rgba[output + 3] = 255;
      }
    } else {
      throw new Error(`unsupported PPM magic ${magic}`);
    }
    return { width, height, pixels: rgba };
  }

  function isJpegBytes(bytes) {
    return bytes?.[0] === 0xff && bytes?.[1] === 0xd8 && bytes?.[2] === 0xff;
  }

  function isWebpBytes(bytes) {
    return bytes?.[0] === 0x52 && bytes?.[1] === 0x49 && bytes?.[2] === 0x46 && bytes?.[3] === 0x46
      && bytes?.[8] === 0x57 && bytes?.[9] === 0x45 && bytes?.[10] === 0x42 && bytes?.[11] === 0x50;
  }

  function rasterByteFormat(bytes) {
    if (isPngBytes(bytes)) return "png";
    if (isPpmBytes(bytes)) return "ppm";
    if (isJpegBytes(bytes)) return "jpeg";
    if (isWebpBytes(bytes)) return "webp";
    return undefined;
  }

  async function decodeRasterRgba(bytes, options = {}) {
    const format = rasterByteFormat(bytes);
    if (format === "png") return decodePngRgba(bytes);
    if (format === "ppm") return decodePpmRgba(bytes);
    if (format !== "jpeg" && format !== "webp") throw new Error("unsupported raster format");
    let sharp;
    try {
      const module = await import("sharp");
      sharp = module.default || module;
    } catch (error) {
      throw new Error(`JPEG/WebP pixel diff requires the optional peer dependency \"sharp\": ${error.message}`);
    }
    const maxPixels = Math.max(1, Number(options.maxDecodedPixels ?? options.maxPixels ?? 40_000_000));
    const decoded = await sharp(Buffer.from(bytes), { failOn: "error", limitInputPixels: maxPixels })
      .ensureAlpha()
      .raw()
      .toBuffer({ resolveWithObject: true });
    if (!decoded.info?.width || !decoded.info?.height || decoded.info.channels !== 4) throw new Error(`Sharp did not return RGBA pixels for ${format}`);
    return { width: decoded.info.width, height: decoded.info.height, pixels: new Uint8Array(decoded.data) };
  }

  async function compareRasterPixels(bytes, baselineBytes, options = {}) {
    const actualFormat = rasterByteFormat(bytes);
    const baselineFormat = rasterByteFormat(baselineBytes);
    if (!actualFormat || !baselineFormat) throw new Error("pixelDiff supports PNG, JPEG, WebP, and PPM raster baselines");
    const [actual, expected] = await Promise.all([decodeRasterRgba(bytes, options), decodeRasterRgba(baselineBytes, options)]);
    const threshold = Math.max(0, Number(options.pixelThreshold ?? options.threshold ?? 0));
    const format = actualFormat === baselineFormat ? actualFormat : `${actualFormat}/${baselineFormat}`;
    const metrics = compareRgbaPixels(actual, expected, { ...options, threshold, format, actualFormat, baselineFormat });
    let diffBytes;
    if (metrics.changed && metrics.diffPixels) diffBytes = encodePngRgba(metrics.diffWidth || metrics.width, metrics.diffHeight || metrics.height, metrics.diffPixels);
    delete metrics.diffPixels;
    return { metrics, diffBytes };
  }

  function pngCrc32(bytes) {
    let crc = 0xffffffff;
    for (const byte of bytes) {
      crc ^= byte;
      for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function pngChunk(type, data = new Uint8Array()) {
    const typeBytes = encoder.encode(type);
    const payload = Buffer.from(data);
    const chunk = Buffer.alloc(12 + payload.length);
    chunk.writeUInt32BE(payload.length, 0);
    Buffer.from(typeBytes).copy(chunk, 4);
    payload.copy(chunk, 8);
    chunk.writeUInt32BE(pngCrc32(new Uint8Array(chunk.subarray(4, 8 + payload.length))), 8 + payload.length);
    return chunk;
  }

  function encodePngRgba(width, height, pixels) {
    const signature = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);
    const ihdr = Buffer.alloc(13);
    ihdr.writeUInt32BE(width, 0);
    ihdr.writeUInt32BE(height, 4);
    ihdr[8] = 8;
    ihdr[9] = 6;
    const rowBytes = width * 4;
    const raw = Buffer.alloc((rowBytes + 1) * height);
    for (let row = 0; row < height; row += 1) Buffer.from(pixels.subarray(row * rowBytes, (row + 1) * rowBytes)).copy(raw, row * (rowBytes + 1) + 1);
    return new Uint8Array(Buffer.concat([signature, pngChunk("IHDR", ihdr), pngChunk("IDAT", deflateSync(raw)), pngChunk("IEND")]));
  }

  function qaRgbColor(value, fallback) {
    if (typeof value === "string") {
      const short = /^#([0-9a-f])([0-9a-f])([0-9a-f])$/i.exec(value);
      if (short) return short.slice(1).map((part) => Number.parseInt(part + part, 16));
      const hex = /^#([0-9a-f]{6})$/i.exec(value)?.[1];
      if (hex) return [0, 2, 4].map((offset) => Number.parseInt(hex.slice(offset, offset + 2), 16));
    }
    if (Array.isArray(value) && value.length >= 3) return value.slice(0, 3).map((part) => Math.max(0, Math.min(255, Math.round(Number(part) || 0))));
    return fallback;
  }

  function rgbaAt(image, canvasX, canvasY, offsetX, offsetY) {
    const x = canvasX - offsetX;
    const y = canvasY - offsetY;
    if (x < 0 || y < 0 || x >= image.width || y >= image.height) return [0, 0, 0, 0];
    const index = (y * image.width + x) * 4;
    return image.pixels.subarray(index, index + 4);
  }

  function rgbaPixelChanged(actual, expected, canvasX, canvasY, actualOffsetX, actualOffsetY, expectedOffsetX, expectedOffsetY, threshold) {
    const actualX = canvasX - actualOffsetX;
    const actualY = canvasY - actualOffsetY;
    const expectedX = canvasX - expectedOffsetX;
    const expectedY = canvasY - expectedOffsetY;
    const actualIndex = actualX >= 0 && actualY >= 0 && actualX < actual.width && actualY < actual.height ? (actualY * actual.width + actualX) * 4 : -1;
    const expectedIndex = expectedX >= 0 && expectedY >= 0 && expectedX < expected.width && expectedY < expected.height ? (expectedY * expected.width + expectedX) * 4 : -1;
    for (let channel = 0; channel < 4; channel += 1) {
      const actualValue = actualIndex < 0 ? 0 : actual.pixels[actualIndex + channel];
      const expectedValue = expectedIndex < 0 ? 0 : expected.pixels[expectedIndex + channel];
      if (Math.abs(actualValue - expectedValue) > threshold) return true;
    }
    return false;
  }

  function countRgbaMismatches(actual, expected, geometry, threshold, expectedShiftX = 0, expectedShiftY = 0, stride = 1) {
    let differentPixels = 0;
    let sampledPixels = 0;
    for (let y = 0; y < geometry.canvasHeight; y += stride) {
      for (let x = 0; x < geometry.canvasWidth; x += stride) {
        const expectedX = x - geometry.expectedOffsetX - expectedShiftX;
        const expectedY = y - geometry.expectedOffsetY - expectedShiftY;
        if ((expectedShiftX || expectedShiftY) && (expectedX < 0 || expectedY < 0 || expectedX >= expected.width || expectedY >= expected.height)) continue;
        sampledPixels += 1;
        if (rgbaPixelChanged(actual, expected, x, y, geometry.actualOffsetX, geometry.actualOffsetY, geometry.expectedOffsetX + expectedShiftX, geometry.expectedOffsetY + expectedShiftY, threshold)) differentPixels += 1;
      }
    }
    return { differentPixels, sampledPixels };
  }

  function pixelRegistrationConfig(options = {}) {
    const request = options.pixelRegistration ?? options.registration;
    if (!request) return undefined;
    const config = typeof request === "object" ? request : {};
    const requestedOffset = request === true ? 2 : typeof request === "number" ? request : config.maxOffset ?? config.maxPixels ?? 2;
    const maxOffset = Math.max(0, Math.min(8, Math.floor(Number(requestedOffset) || 0)));
    if (!maxOffset) return undefined;
    return {
      maxOffset,
      minImprovementRatio: Math.max(0, Math.min(1, Number(config.minImprovementRatio ?? config.minImprovement ?? 0.05))),
      maxSamples: Math.max(1_000, Math.min(1_000_000, Math.floor(Number(config.maxSamples ?? config.samples ?? 100_000) || 100_000))),
    };
  }

  function findPixelRegistration(actual, expected, geometry, threshold, config) {
    const candidateCount = (config.maxOffset * 2 + 1) ** 2;
    const samplesPerCandidate = Math.max(16, Math.floor(config.maxSamples / candidateCount));
    const stride = Math.max(1, Math.ceil(Math.sqrt((geometry.canvasWidth * geometry.canvasHeight) / samplesPerCandidate)));
    const baseline = countRgbaMismatches(actual, expected, geometry, threshold, 0, 0, stride);
    let best = { x: 0, y: 0, ...baseline };
    for (let y = -config.maxOffset; y <= config.maxOffset; y += 1) {
      for (let x = -config.maxOffset; x <= config.maxOffset; x += 1) {
        if (x === 0 && y === 0) continue;
        const candidate = countRgbaMismatches(actual, expected, geometry, threshold, x, y, stride);
        const candidateRatio = candidate.sampledPixels ? candidate.differentPixels / candidate.sampledPixels : 1;
        const bestRatio = best.sampledPixels ? best.differentPixels / best.sampledPixels : 1;
        const candidateDistance = Math.abs(x) + Math.abs(y);
        const bestDistance = Math.abs(best.x) + Math.abs(best.y);
        if (candidateRatio < bestRatio || (candidateRatio === bestRatio && candidateDistance < bestDistance)) best = { x, y, ...candidate };
      }
    }
    const baselineRatio = baseline.sampledPixels ? baseline.differentPixels / baseline.sampledPixels : 0;
    const bestRatio = best.sampledPixels ? best.differentPixels / best.sampledPixels : 1;
    const improvementRatio = baselineRatio ? (baselineRatio - bestRatio) / baselineRatio : 0;
    const applied = (best.x !== 0 || best.y !== 0) && improvementRatio >= config.minImprovementRatio;
    return {
      requested: true,
      applied,
      maxOffset: config.maxOffset,
      minImprovementRatio: config.minImprovementRatio,
      maxSamples: config.maxSamples,
      candidateCount,
      samplesPerCandidate,
      estimatedComparisons: baseline.sampledPixels * candidateCount,
      sampleStride: stride,
      sampledPixels: baseline.sampledPixels,
      sampledPixelsAfter: applied ? best.sampledPixels : baseline.sampledPixels,
      sampledDifferentPixelsBefore: baseline.differentPixels,
      sampledDifferentPixelsAfter: applied ? best.differentPixels : baseline.differentPixels,
      sampledImprovementRatio: applied ? improvementRatio : 0,
      offset: applied ? { x: best.x, y: best.y } : { x: 0, y: 0 },
    };
  }

  function compareRgbaPixels(actual, expected, options = {}) {
    const threshold = Math.max(0, Number(options.threshold ?? options.pixelThreshold ?? 0));
    const requestedAlignment = String(options.diffAlignment ?? options.alignment ?? "strict").trim().toLowerCase();
    const alignment = ["strict", "top-left", "center"].includes(requestedAlignment) ? requestedAlignment : "strict";
    const dimensionMismatch = actual.width !== expected.width || actual.height !== expected.height;
    const requestedRegistrationConfig = pixelRegistrationConfig(options);
    const canvasWidth = dimensionMismatch && alignment !== "strict" ? Math.max(actual.width, expected.width) : actual.width;
    const canvasHeight = dimensionMismatch && alignment !== "strict" ? Math.max(actual.height, expected.height) : actual.height;
    const result = {
      format: options.format || "rgba",
      actualFormat: options.actualFormat,
      baselineFormat: options.baselineFormat,
      width: actual.width,
      height: actual.height,
      baselineWidth: expected.width,
      baselineHeight: expected.height,
      threshold,
      pixels: actual.width * actual.height,
      differentPixels: 0,
      mismatchRatio: 0,
      maxChannelDelta: 0,
      meanChannelDelta: 0,
      alignment,
    };
    if (dimensionMismatch && alignment === "strict") {
      result.dimensionMismatch = true;
      result.differentPixels = Math.max(result.pixels, expected.width * expected.height);
      result.mismatchRatio = 1;
      result.changed = true;
      if (requestedRegistrationConfig) result.registration = { requested: true, applied: false, skipped: "dimensionMismatch", ...requestedRegistrationConfig };
      return result;
    }
    if (dimensionMismatch) {
      result.dimensionMismatch = true;
      result.diffWidth = canvasWidth;
      result.diffHeight = canvasHeight;
      result.comparisonPixels = canvasWidth * canvasHeight;
      result.pixels = result.comparisonPixels;
    }
    const actualOffsetX = alignment === "center" ? Math.floor((canvasWidth - actual.width) / 2) : 0;
    const actualOffsetY = alignment === "center" ? Math.floor((canvasHeight - actual.height) / 2) : 0;
    const baseExpectedOffsetX = alignment === "center" ? Math.floor((canvasWidth - expected.width) / 2) : 0;
    const baseExpectedOffsetY = alignment === "center" ? Math.floor((canvasHeight - expected.height) / 2) : 0;
    const geometry = { canvasWidth, canvasHeight, actualOffsetX, actualOffsetY, expectedOffsetX: baseExpectedOffsetX, expectedOffsetY: baseExpectedOffsetY };
    const registrationConfig = dimensionMismatch ? undefined : requestedRegistrationConfig;
    const registration = registrationConfig ? findPixelRegistration(actual, expected, geometry, threshold, registrationConfig) : undefined;
    const expectedOffsetX = baseExpectedOffsetX + (registration?.offset.x || 0);
    const expectedOffsetY = baseExpectedOffsetY + (registration?.offset.y || 0);
    const diffPixels = options.diffImage === false ? undefined : new Uint8Array(canvasWidth * canvasHeight * 4);
    const palette = options.diffPalette || options.palette || {};
    const changedColor = qaRgbColor(palette.changed, [255, 24, 72]);
    const unchangedColor = palette.unchanged == null ? undefined : qaRgbColor(palette.unchanged, [64, 64, 64]);
    const changedAlpha = Math.max(0, Math.min(255, Math.round(Number(palette.changedAlpha ?? 255))));
    const unchangedAlpha = Math.max(0, Math.min(255, Math.round(Number(palette.unchangedAlpha ?? 255))));
    let changedPixels = 0;
    let channelDeltaSum = 0;
    let registrationIgnoredPixels = 0;
    for (let y = 0; y < canvasHeight; y += 1) for (let x = 0; x < canvasWidth; x += 1) {
      const actualPixel = rgbaAt(actual, x, y, actualOffsetX, actualOffsetY);
      const expectedX = x - expectedOffsetX;
      const expectedY = y - expectedOffsetY;
      const ignoredRegistrationEdge = registration?.applied && (expectedX < 0 || expectedY < 0 || expectedX >= expected.width || expectedY >= expected.height);
      const expectedPixel = ignoredRegistrationEdge ? actualPixel : rgbaAt(expected, x, y, expectedOffsetX, expectedOffsetY);
      const i = (y * canvasWidth + x) * 4;
      let pixelChanged = false;
      if (ignoredRegistrationEdge) registrationIgnoredPixels += 1;
      for (let c = 0; c < 4; c += 1) {
        const delta = Math.abs(actualPixel[c] - expectedPixel[c]);
        channelDeltaSum += delta;
        if (delta > result.maxChannelDelta) result.maxChannelDelta = delta;
        if (delta > threshold) pixelChanged = true;
      }
      if (diffPixels) {
        const brightness = Math.round((actualPixel[0] + actualPixel[1] + actualPixel[2]) / 3 * 0.35 + 32);
        const color = pixelChanged ? changedColor : unchangedColor || [brightness, brightness, brightness];
        diffPixels[i] = color[0];
        diffPixels[i + 1] = color[1];
        diffPixels[i + 2] = color[2];
        diffPixels[i + 3] = pixelChanged ? changedAlpha : unchangedAlpha;
      }
      if (pixelChanged) changedPixels += 1;
    }
    result.differentPixels = changedPixels;
    const comparedPixels = Math.max(0, result.pixels - registrationIgnoredPixels);
    result.mismatchRatio = comparedPixels ? changedPixels / comparedPixels : 0;
    result.meanChannelDelta = comparedPixels ? channelDeltaSum / (comparedPixels * 4) : 0;
    result.diffPalette = { changed: changedColor, changedAlpha, unchanged: unchangedColor || "actual-grayscale", unchangedAlpha };
    if (registration) {
      const before = registration.applied ? countRgbaMismatches(actual, expected, geometry, threshold) : { differentPixels: changedPixels, sampledPixels: result.pixels };
      result.registration = {
        ...registration,
        differentPixelsBefore: before.differentPixels,
        mismatchRatioBefore: result.pixels ? before.differentPixels / result.pixels : 0,
        differentPixelsAfter: changedPixels,
        mismatchRatioAfter: result.mismatchRatio,
        comparedPixelsAfter: comparedPixels,
        ignoredEdgePixels: registrationIgnoredPixels,
      };
    } else if (requestedRegistrationConfig && dimensionMismatch) result.registration = { requested: true, applied: false, skipped: "dimensionMismatch", ...requestedRegistrationConfig };
    result.changed = changedPixels > 0;
    if (result.changed && diffPixels) result.diffPixels = diffPixels;
    return result;
  }

  async function renderArtifact(artifact, options = {}) {
    const artifactKind = inferArtifactKind(artifact);
    if (!artifact || (typeof artifact.render !== "function" && typeof artifact.export !== "function")) {
      throw new Error("Artifact does not expose a render() or export() method.");
    }
    const renderer = typeof artifact.render === "function" ? artifact.render.bind(artifact) : artifact.export.bind(artifact);
    let blob = await renderer(options);
    if (!(blob instanceof FileBlob)) {
      blob = await fileBlobFromRenderOutput(blob, blob?.type || "application/octet-stream", { artifactKind, format: options.format || blob?.type || "unknown" });
    }
    const desiredType = renderTypeForOptions(options, blob.type);
    const wantsConversion = options.format && desiredType !== blob.type;
    if (wantsConversion) {
      const adapter = options.renderer || options.rasterRenderer || options.renderAdapter;
      if (typeof adapter !== "function") {
        throw new Error(`renderArtifact requested ${options.format} output, but no renderer adapter was provided.`);
      }
      const converted = await adapter({ input: blob, source: blob, inputType: blob.type, outputType: desiredType, format: options.format, artifactKind, options });
      blob = await fileBlobFromRenderOutput(converted, desiredType, { artifactKind, format: options.format, renderedFrom: blob.type });
      if (!blob.type || blob.type === "application/octet-stream") blob.type = desiredType;
    }
    return attachRenderMetadata(blob, artifactKind, options, options.format || blob.metadata?.format || blob.type);
  }

  async function visualQaArtifact(artifact, options = {}) {
    const blob = await renderArtifact(artifact, options);
    const artifactKind = blob.metadata?.artifactKind || inferArtifactKind(artifact);
    const bytes = blob.bytes || new Uint8Array(await blob.arrayBuffer());
    const hash = stableByteHash(bytes);
    const issues = [];
    let diffBlob;
    const summary = { kind: "visualQa", artifactKind, type: blob.type, format: blob.metadata?.format || options.format || blob.type, bytes: bytes.byteLength, hash };
    if (bytes.byteLength === 0) issues.push(verificationIssue(artifactKind, "emptyRender", "Rendered artifact is empty.", { severity: "error", type: blob.type }));
    if (options.minBytes != null && bytes.byteLength < Number(options.minBytes)) issues.push(verificationIssue(artifactKind, "renderTooSmall", `Rendered artifact has ${bytes.byteLength} bytes; expected at least ${options.minBytes}.`, { severity: "warning", bytes: bytes.byteLength, minBytes: Number(options.minBytes) }));
    if (options.maxBytes != null && bytes.byteLength > Number(options.maxBytes)) issues.push(verificationIssue(artifactKind, "renderTooLarge", `Rendered artifact has ${bytes.byteLength} bytes; expected at most ${options.maxBytes}.`, { severity: "warning", bytes: bytes.byteLength, maxBytes: Number(options.maxBytes) }));
    if (blob.type === "image/svg+xml") {
      const text = await blob.text();
      const dimensions = svgDimensions(text);
      summary.width = dimensions.width;
      summary.height = dimensions.height;
      if (!dimensions.width || !dimensions.height || dimensions.width <= 0 || dimensions.height <= 0) issues.push(verificationIssue(artifactKind, "invalidRenderGeometry", "SVG render is missing positive width/height geometry.", { severity: "error", dimensions }));
      if (!/<(text|image|rect|path|line|polyline|polygon|circle|ellipse)\b/i.test(text)) issues.push(verificationIssue(artifactKind, "blankSvgRender", "SVG render has no recognizable visible elements.", { severity: "warning" }));
    }
    const baselineBytes = await bytesForVisualBaseline(options.baseline || options.expected || options.baselineBlob);
    if (baselineBytes) {
      const baselineHash = stableByteHash(baselineBytes);
      summary.baselineHash = baselineHash;
      summary.byteChanged = baselineHash !== hash;
      summary.changed = summary.byteChanged;
      const pixelDiffEnabled = options.pixelDiff === true || typeof options.pixelDiff === "object";
      if (pixelDiffEnabled) {
        if (rasterByteFormat(bytes) && rasterByteFormat(baselineBytes)) {
          try {
            const pixelDiffOptions = typeof options.pixelDiff === "object" ? { ...options, ...options.pixelDiff } : options;
            const compared = await compareRasterPixels(bytes, baselineBytes, pixelDiffOptions);
            const pixelDiff = compared.metrics;
            summary.pixelDiff = pixelDiff;
            summary.changed = pixelDiff.changed;
            if (compared.diffBytes) {
              diffBlob = new FileBlob(compared.diffBytes, { type: "image/png", metadata: { artifactKind, format: "pixel-diff", actualFormat: pixelDiff.actualFormat, baselineFormat: pixelDiff.baselineFormat, alignment: pixelDiff.alignment, registration: pixelDiff.registration, width: pixelDiff.diffWidth || pixelDiff.width, height: pixelDiff.diffHeight || pixelDiff.height, palette: pixelDiff.diffPalette } });
              summary.diff = { type: diffBlob.type, bytes: diffBlob.bytes.length, hash: stableByteHash(diffBlob.bytes) };
            }
            if (pixelDiff.changed && options.allowChange !== true && options.allowPixelChange !== true) {
              issues.push(verificationIssue(artifactKind, "visualPixelDiff", `Rendered ${pixelDiff.format.toUpperCase()} differs from the baseline in ${pixelDiff.differentPixels} pixels.`, { severity: options.diffSeverity || "warning", ...pixelDiff }));
            }
          } catch (error) {
            summary.pixelDiff = { skipped: true, reason: error.message };
          }
        } else {
          summary.pixelDiff = { skipped: true, reason: "pixelDiff supports PNG, JPEG, WebP, and PPM raster baselines" };
        }
      }
      const pixelsEquivalent = summary.pixelDiff && !summary.pixelDiff.skipped && summary.pixelDiff.changed === false;
      if (baselineHash !== hash && !pixelsEquivalent && options.allowChange !== true) issues.push(verificationIssue(artifactKind, "visualDiff", "Rendered output differs from the supplied baseline.", { severity: options.diffSeverity || "warning", hash, baselineHash }));
    }
    const records = [summary, ...issues];
    return { artifactKind, ok: issues.length === 0, blob, diffBlob, summary, issues, ...ndjson(records, options.maxChars ?? Infinity) };
  }

  return { renderArtifact, visualQaArtifact };
}
