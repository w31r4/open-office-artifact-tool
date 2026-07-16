import { inflateSync } from "node:zlib";

import { decoder } from "./binary.mjs";

const PNG_SIGNATURE = [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a];

export function isPngBytes(bytes) {
  return PNG_SIGNATURE.every((byte, index) => bytes?.[index] === byte);
}

function paethPredictor(a, b, c) {
  const p = a + b - c;
  const pa = Math.abs(p - a);
  const pb = Math.abs(p - b);
  const pc = Math.abs(p - c);
  if (pa <= pb && pa <= pc) return a;
  if (pb <= pc) return b;
  return c;
}

export function decodePngRgba(bytes) {
  if (!isPngBytes(bytes)) throw new Error("not a PNG file");
  let offset = 8;
  let width = 0;
  let height = 0;
  let bitDepth = 0;
  let colorType = 0;
  let compression = 0;
  let filterMethod = 0;
  let interlace = 0;
  const idat = [];
  while (offset + 12 <= bytes.byteLength) {
    const length = (bytes[offset] << 24) | (bytes[offset + 1] << 16) | (bytes[offset + 2] << 8) | bytes[offset + 3];
    const type = decoder.decode(bytes.slice(offset + 4, offset + 8));
    const dataStart = offset + 8;
    const dataEnd = dataStart + length;
    if (length < 0 || dataEnd + 4 > bytes.byteLength) throw new Error("truncated PNG chunk");
    const data = bytes.slice(dataStart, dataEnd);
    if (type === "IHDR") {
      width = (data[0] << 24) | (data[1] << 16) | (data[2] << 8) | data[3];
      height = (data[4] << 24) | (data[5] << 16) | (data[6] << 8) | data[7];
      bitDepth = data[8];
      colorType = data[9];
      compression = data[10];
      filterMethod = data[11];
      interlace = data[12];
    } else if (type === "IDAT") {
      idat.push(Buffer.from(data));
    } else if (type === "IEND") {
      break;
    }
    offset = dataEnd + 4;
  }
  if (!width || !height) throw new Error("PNG is missing IHDR geometry");
  if (bitDepth !== 8 || compression !== 0 || filterMethod !== 0 || interlace !== 0) throw new Error("only 8-bit non-interlaced PNGs are supported for pixel diff");
  const channels = { 0: 1, 2: 3, 4: 2, 6: 4 }[colorType];
  if (!channels) throw new Error(`unsupported PNG color type ${colorType}`);
  const rowBytes = width * channels;
  const inflated = inflateSync(Buffer.concat(idat));
  const expected = (rowBytes + 1) * height;
  if (inflated.byteLength < expected) throw new Error("PNG image data is truncated");
  const raw = new Uint8Array(width * height * channels);
  let inOffset = 0;
  for (let y = 0; y < height; y += 1) {
    const filter = inflated[inOffset++];
    const rowStart = y * rowBytes;
    const priorStart = (y - 1) * rowBytes;
    for (let x = 0; x < rowBytes; x += 1) {
      const left = x >= channels ? raw[rowStart + x - channels] : 0;
      const up = y > 0 ? raw[priorStart + x] : 0;
      const upLeft = y > 0 && x >= channels ? raw[priorStart + x - channels] : 0;
      const value = inflated[inOffset++];
      if (filter === 0) raw[rowStart + x] = value;
      else if (filter === 1) raw[rowStart + x] = (value + left) & 0xff;
      else if (filter === 2) raw[rowStart + x] = (value + up) & 0xff;
      else if (filter === 3) raw[rowStart + x] = (value + Math.floor((left + up) / 2)) & 0xff;
      else if (filter === 4) raw[rowStart + x] = (value + paethPredictor(left, up, upLeft)) & 0xff;
      else throw new Error(`unsupported PNG row filter ${filter}`);
    }
  }
  const rgba = new Uint8Array(width * height * 4);
  for (let i = 0, p = 0; p < width * height; p += 1, i += channels) {
    const out = p * 4;
    if (colorType === 0) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i];
      rgba[out + 2] = raw[i];
      rgba[out + 3] = 255;
    } else if (colorType === 2) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i + 1];
      rgba[out + 2] = raw[i + 2];
      rgba[out + 3] = 255;
    } else if (colorType === 4) {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i];
      rgba[out + 2] = raw[i];
      rgba[out + 3] = raw[i + 1];
    } else {
      rgba[out] = raw[i];
      rgba[out + 1] = raw[i + 1];
      rgba[out + 2] = raw[i + 2];
      rgba[out + 3] = raw[i + 3];
    }
  }
  return { width, height, pixels: rgba };
}
