#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import pako from "pako";

const PNG_SIGNATURE = Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]);
const MAX_PNG_BYTES = 16 * 1024 * 1024;
const MAX_INFLATED_BYTES = 128 * 1024 * 1024;
const ADAM7 = Object.freeze([
  [0, 0, 8, 8],
  [4, 0, 8, 8],
  [0, 4, 4, 8],
  [2, 0, 4, 4],
  [0, 2, 2, 4],
  [1, 0, 2, 2],
  [0, 1, 1, 2],
]);

export const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
export const SKILL_PNG_ROOTS = Object.freeze([
  "skills/documents",
  "skills/spreadsheets",
  "skills/presentations",
  "skills/pdf",
]);

const CRC32_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let value = 0; value < table.length; value += 1) {
    let crc = value;
    for (let bit = 0; bit < 8; bit += 1) crc = (crc >>> 1) ^ ((crc & 1) ? 0xedb88320 : 0);
    table[value] = crc >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (const byte of bytes) crc = CRC32_TABLE[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function pngError(message) {
  return new Error(`Invalid PNG: ${message}`);
}

function passLength(total, start, step) {
  return total <= start ? 0 : Math.ceil((total - start) / step);
}

function expectedInflatedLength(ihdr) {
  const { width, height, bitDepth, colorType, interlace } = ihdr;
  const channels = new Map([[0, 1], [2, 3], [3, 1], [4, 2], [6, 4]]).get(colorType);
  if (!channels) throw pngError(`unsupported color type ${colorType}`);
  const permittedDepths = {
    0: [1, 2, 4, 8, 16],
    2: [8, 16],
    3: [1, 2, 4, 8],
    4: [8, 16],
    6: [8, 16],
  }[colorType];
  if (!permittedDepths.includes(bitDepth)) throw pngError(`unsupported bit depth ${bitDepth} for color type ${colorType}`);
  const bitsPerPixel = channels * bitDepth;
  const passBytes = (passWidth, passHeight) => passWidth === 0 || passHeight === 0
    ? 0
    : passHeight * (1 + Math.ceil((passWidth * bitsPerPixel) / 8));
  const expected = interlace === 0
    ? passBytes(width, height)
    : ADAM7.reduce((total, [x, y, dx, dy]) => total + passBytes(
      passLength(width, x, dx),
      passLength(height, y, dy),
    ), 0);
  if (!Number.isSafeInteger(expected) || expected > MAX_INFLATED_BYTES) {
    throw pngError(`inflated scanline budget exceeded (${expected} bytes)`);
  }
  return expected;
}

export function parsePng(input) {
  const bytes = Buffer.from(input);
  if (bytes.length > MAX_PNG_BYTES) throw pngError(`file budget exceeded (${bytes.length} bytes)`);
  if (bytes.length < PNG_SIGNATURE.length || !bytes.subarray(0, PNG_SIGNATURE.length).equals(PNG_SIGNATURE)) {
    throw pngError("signature mismatch");
  }

  const chunks = [];
  let offset = PNG_SIGNATURE.length;
  let seenIhdr = false;
  let seenIdat = false;
  let leftIdatRun = false;
  let seenIend = false;
  let ihdr;

  while (offset < bytes.length) {
    if (bytes.length - offset < 12) throw pngError(`truncated chunk header at byte ${offset}`);
    const length = bytes.readUInt32BE(offset);
    const end = offset + 12 + length;
    if (!Number.isSafeInteger(end) || end > bytes.length) throw pngError(`truncated chunk payload at byte ${offset}`);
    const typeBytes = bytes.subarray(offset + 4, offset + 8);
    const type = typeBytes.toString("ascii");
    if (!/^[A-Za-z]{4}$/.test(type)) throw pngError(`invalid chunk type at byte ${offset}`);
    const data = bytes.subarray(offset + 8, offset + 8 + length);
    const storedCrc = bytes.readUInt32BE(offset + 8 + length);
    const actualCrc = crc32(Buffer.concat([typeBytes, data]));
    if (storedCrc !== actualCrc) throw pngError(`${type} CRC mismatch at byte ${offset}`);
    const raw = bytes.subarray(offset, end);

    if (!seenIhdr) {
      if (type !== "IHDR") throw pngError("IHDR must be the first chunk");
      if (length !== 13) throw pngError("IHDR must contain 13 bytes");
      const width = data.readUInt32BE(0);
      const height = data.readUInt32BE(4);
      if (width === 0 || height === 0) throw pngError("image dimensions must be positive");
      if (data[10] !== 0 || data[11] !== 0 || ![0, 1].includes(data[12])) {
        throw pngError("unsupported compression, filter, or interlace method");
      }
      ihdr = {
        width,
        height,
        bitDepth: data[8],
        colorType: data[9],
        interlace: data[12],
      };
      seenIhdr = true;
    } else if (type === "IHDR") {
      throw pngError("multiple IHDR chunks");
    }

    if (type === "IDAT") {
      if (leftIdatRun) throw pngError("IDAT chunks must be consecutive");
      seenIdat = true;
    } else if (seenIdat && type !== "IEND") {
      leftIdatRun = true;
    }

    if (type === "IEND") {
      if (length !== 0) throw pngError("IEND must be empty");
      if (!seenIdat) throw pngError("missing IDAT chunk");
      if (end !== bytes.length) throw pngError("IEND must be the final chunk with no trailing bytes");
      seenIend = true;
    }

    chunks.push({ type, data: Buffer.from(data), raw: Buffer.from(raw) });
    offset = end;
    if (seenIend) break;
  }

  if (!seenIhdr) throw pngError("missing IHDR chunk");
  if (!seenIend) throw pngError("missing IEND chunk");
  return { bytes, chunks, ihdr, expectedInflatedBytes: expectedInflatedLength(ihdr) };
}

function inflateIdat(parsed) {
  const compressed = Buffer.concat(parsed.chunks.filter(({ type }) => type === "IDAT").map(({ data }) => data));
  let inflated;
  try {
    inflated = Buffer.from(pako.inflate(compressed));
  } catch (error) {
    throw pngError(`IDAT inflate failed: ${error.message}`);
  }
  if (inflated.length !== parsed.expectedInflatedBytes) {
    throw pngError(`inflated scanline length ${inflated.length} does not match expected ${parsed.expectedInflatedBytes}`);
  }
  return inflated;
}

function makeChunk(type, data) {
  const typeBytes = Buffer.from(type, "ascii");
  const output = Buffer.allocUnsafe(12 + data.length);
  output.writeUInt32BE(data.length, 0);
  typeBytes.copy(output, 4);
  data.copy(output, 8);
  output.writeUInt32BE(crc32(Buffer.concat([typeBytes, data])), 8 + data.length);
  return output;
}

export function pngIdentity(input) {
  const parsed = parsePng(input);
  return {
    inflated: inflateIdat(parsed),
    nonIdat: Buffer.concat(parsed.chunks.filter(({ type }) => type !== "IDAT").map(({ raw }) => raw)),
  };
}

export function optimizePngBytes(input) {
  const parsed = parsePng(input);
  const inflated = inflateIdat(parsed);
  const compressed = Buffer.from(pako.deflate(inflated, { level: 9 }));
  const originalIdatBytes = parsed.chunks
    .filter(({ type }) => type === "IDAT")
    .reduce((total, { raw }) => total + raw.length, 0);
  const replacement = makeChunk("IDAT", compressed);
  if (replacement.length >= originalIdatBytes) {
    return { bytes: Buffer.from(parsed.bytes), changed: false, savings: 0 };
  }

  const output = [PNG_SIGNATURE];
  let wroteIdat = false;
  for (const chunk of parsed.chunks) {
    if (chunk.type === "IDAT") {
      if (!wroteIdat) output.push(replacement);
      wroteIdat = true;
    } else {
      output.push(chunk.raw);
    }
  }
  const bytes = Buffer.concat(output);
  const candidate = parsePng(bytes);
  const candidateInflated = inflateIdat(candidate);
  const sourceNonIdat = Buffer.concat(parsed.chunks.filter(({ type }) => type !== "IDAT").map(({ raw }) => raw));
  const candidateNonIdat = Buffer.concat(candidate.chunks.filter(({ type }) => type !== "IDAT").map(({ raw }) => raw));
  if (!candidateInflated.equals(inflated)) throw new Error("PNG optimizer changed the inflated scanline stream");
  if (!candidateNonIdat.equals(sourceNonIdat)) throw new Error("PNG optimizer changed a non-IDAT chunk");
  const repeated = Buffer.from(pako.deflate(candidateInflated, { level: 9 }));
  if (!repeated.equals(compressed)) throw new Error("PNG optimizer output is not deterministic");
  return { bytes, changed: true, savings: parsed.bytes.length - bytes.length };
}

async function regularPngFiles(root, relative = "") {
  const directory = path.join(root, relative);
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const output = [];
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name, "en"))) {
    const child = relative ? path.join(relative, entry.name) : entry.name;
    if (entry.isDirectory()) output.push(...await regularPngFiles(root, child));
    else if (entry.isFile()) {
      if (entry.name.toLowerCase().endsWith(".png")) output.push(path.join(root, child));
    } else {
      throw new Error(`Skill asset tree contains a non-regular path: ${path.relative(REPO_ROOT, path.join(root, child))}`);
    }
  }
  return output;
}

export async function listSkillPngFiles(repoRoot = REPO_ROOT) {
  const files = [];
  for (const relativeRoot of SKILL_PNG_ROOTS) files.push(...await regularPngFiles(path.join(repoRoot, relativeRoot)));
  return files.sort((left, right) => left.localeCompare(right, "en"));
}

async function optimizeRepositoryAssets(mode) {
  const files = await listSkillPngFiles();
  const candidates = [];
  for (const filename of files) {
    const source = await fs.readFile(filename);
    const result = optimizePngBytes(source);
    if (result.changed) candidates.push({ filename, ...result });
  }

  const savings = candidates.reduce((total, candidate) => total + candidate.savings, 0);
  if (mode === "check") {
    if (candidates.length) {
      throw new Error([
        `${candidates.length} Skill PNG asset(s) can still be losslessly recompressed (${savings} bytes):`,
        ...candidates.map(({ filename, savings: fileSavings }) => `- ${path.relative(REPO_ROOT, filename)}: ${fileSavings} bytes`),
        "Run npm run assets:png:optimize and review the binary diff.",
      ].join("\n"));
    }
    console.log(`Skill PNG assets optimized: ${files.length} files, no remaining lossless IDAT savings`);
    return;
  }

  let serial = 0;
  for (const candidate of candidates) {
    const stat = await fs.stat(candidate.filename);
    const temporary = `${candidate.filename}.tmp-${process.pid}-${serial++}`;
    try {
      await fs.writeFile(temporary, candidate.bytes, { mode: stat.mode });
      await fs.rename(temporary, candidate.filename);
    } catch (error) {
      await fs.rm(temporary, { force: true });
      throw error;
    }
  }
  console.log(`Skill PNG assets optimized: ${candidates.length}/${files.length} files, ${savings} bytes saved`);
}

const entry = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entry === fileURLToPath(import.meta.url)) {
  const flag = process.argv[2] || "--check";
  if (process.argv.length > 3 || !["--check", "--write"].includes(flag)) {
    console.error("Usage: node scripts/optimize-skill-pngs.mjs [--check|--write]");
    process.exitCode = 2;
  } else {
    await optimizeRepositoryAssets(flag === "--write" ? "write" : "check").catch((error) => {
      console.error(error.message);
      process.exitCode = 1;
    });
  }
}
