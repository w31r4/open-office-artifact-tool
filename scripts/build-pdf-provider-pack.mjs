#!/usr/bin/env node
/**
 * Build one deterministic, installer-compatible PDF capability-pack archive.
 *
 * This is release tooling, not a customer installer. It turns a checked,
 * platform-native payload into the deliberately small USTAR+gzip format that
 * `safeExtractTarGz()` accepts. The installer consequently never needs a
 * platform tar implementation or an archive compatibility fallback.
 */

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { gzipSync } from "node:zlib";

const BLOCK_SIZE = 512;
const PACK_SCHEMA = "open-office-artifact-tool.pdf-provider-pack.v1";
const SUPPORTED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const SHA256 = /^[a-f0-9]{64}$/i;

function fail(message) {
  throw new Error(`PDF capability-pack build: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function safeSegment(value, label) {
  if (!nonEmptyString(value) || value.includes("/") || value.includes("\\") || value === "." || value === "..") {
    fail(`${label} must be one safe path segment.`);
  }
  return value;
}

function safeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\") || value.startsWith("/")) return false;
  const normalized = path.posix.normalize(value);
  return normalized !== "." && normalized !== ".." && !normalized.startsWith("../") && !normalized.includes("/../");
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function parseArguments(argv) {
  const values = {};
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
    values[name] = value;
    index += 1;
  }
  for (const required of ["pack", "version", "platform", "payload", "output", "source-url", "source-sha256", "license", "notices"]) {
    if (!nonEmptyString(values[required])) fail(`--${required} is required.`);
  }
  safeSegment(values.pack, "pack");
  safeSegment(values.version, "version");
  if (!SUPPORTED_PLATFORMS.has(values.platform)) fail(`platform must be one of ${[...SUPPORTED_PLATFORMS].join(", ")}.`);
  if (!SHA256.test(values["source-sha256"])) fail("--source-sha256 must be exactly 64 hexadecimal characters.");
  let sourceUrl;
  try {
    sourceUrl = new URL(values["source-url"]);
  } catch {
    fail("--source-url must be an HTTPS URL.");
  }
  if (sourceUrl.protocol !== "https:") fail("--source-url must be an HTTPS URL.");
  return {
    pack: values.pack,
    version: values.version,
    platform: values.platform,
    payload: path.resolve(values.payload),
    output: path.resolve(values.output),
    sourceUrl: sourceUrl.href,
    sourceSha256: values["source-sha256"].toLowerCase(),
    license: values.license.trim(),
    notices: path.resolve(values.notices),
  };
}

async function regularFileBytes(filePath, label) {
  const stat = await fs.lstat(filePath);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(`${label} must be a regular non-symlink file: ${filePath}.`);
  return fs.readFile(filePath);
}

async function listPayload(root) {
  const rootStat = await fs.lstat(root);
  if (!rootStat.isDirectory() || rootStat.isSymbolicLink()) fail("--payload must be a real directory.");
  const entries = [];
  async function walk(relative = "") {
    const directory = relative ? path.join(root, ...relative.split("/")) : root;
    const children = await fs.readdir(directory, { withFileTypes: true });
    children.sort((left, right) => left.name.localeCompare(right.name, "en"));
    for (const child of children) {
      const childRelative = relative ? `${relative}/${child.name}` : child.name;
      if (!safeRelativePath(childRelative)) fail(`payload contains unsafe path ${childRelative}.`);
      const absolute = path.join(directory, child.name);
      const stat = await fs.lstat(absolute);
      if (stat.isSymbolicLink()) fail(`payload contains a symlink: ${childRelative}.`);
      if (stat.isDirectory()) {
        entries.push({ path: childRelative, type: "directory", mode: 0o755, bytes: Buffer.alloc(0) });
        await walk(childRelative);
      } else if (stat.isFile()) {
        if (stat.nlink > 1) fail(`payload contains a hard-linked file: ${childRelative}.`);
        const bytes = await fs.readFile(absolute);
        entries.push({ path: childRelative, type: "file", mode: (stat.mode & 0o111) !== 0 ? 0o755 : 0o644, bytes });
      } else {
        fail(`payload contains an unsupported filesystem entry: ${childRelative}.`);
      }
    }
  }
  await walk();
  if (!entries.some((entry) => entry.type === "file")) fail("--payload must contain at least one regular file.");
  return entries;
}

function writeString(buffer, offset, length, value) {
  const bytes = Buffer.from(value, "utf8");
  if (bytes.length > length) fail(`USTAR field is too long: ${value}.`);
  bytes.copy(buffer, offset);
}

function writeOctal(buffer, offset, length, value) {
  const text = Number(value).toString(8);
  if (text.length > length - 1) fail(`USTAR numeric field is too large: ${value}.`);
  writeString(buffer, offset, length - 1, text.padStart(length - 1, "0"));
  buffer[offset + length - 1] = 0;
}

function splitUstarPath(entryPath) {
  const bytes = Buffer.byteLength(entryPath, "utf8");
  if (bytes <= 100) return { name: entryPath, prefix: "" };
  const segments = entryPath.split("/");
  for (let index = 1; index < segments.length; index += 1) {
    const prefix = segments.slice(0, index).join("/");
    const name = segments.slice(index).join("/");
    if (Buffer.byteLength(prefix, "utf8") <= 155 && Buffer.byteLength(name, "utf8") <= 100) return { name, prefix };
  }
  fail(`payload path cannot be represented by strict USTAR: ${entryPath}.`);
}

function tarHeader({ entryPath, type, mode, size }) {
  const header = Buffer.alloc(BLOCK_SIZE);
  const { name, prefix } = splitUstarPath(entryPath);
  writeString(header, 0, 100, name);
  writeOctal(header, 100, 8, mode);
  writeOctal(header, 108, 8, 0);
  writeOctal(header, 116, 8, 0);
  writeOctal(header, 124, 12, size);
  writeOctal(header, 136, 12, 0);
  header.fill(0x20, 148, 156);
  header[156] = type === "directory" ? "5".charCodeAt(0) : "0".charCodeAt(0);
  writeString(header, 257, 6, "ustar");
  writeString(header, 263, 2, "00");
  writeString(header, 345, 155, prefix);
  let checksum = 0;
  for (const byte of header) checksum += byte;
  writeString(header, 148, 6, checksum.toString(8).padStart(6, "0"));
  header[154] = 0;
  header[155] = 0x20;
  return header;
}

function tarGzip(entries) {
  const records = [];
  for (const entry of entries) {
    const bytes = entry.bytes;
    records.push(tarHeader({ entryPath: entry.path, type: entry.type, mode: entry.mode, size: bytes.length }));
    if (entry.type === "file") {
      records.push(bytes);
      const padding = (BLOCK_SIZE - (bytes.length % BLOCK_SIZE)) % BLOCK_SIZE;
      if (padding) records.push(Buffer.alloc(padding));
    }
  }
  records.push(Buffer.alloc(BLOCK_SIZE * 2));
  return gzipSync(Buffer.concat(records), { mtime: 0 });
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
}

function deterministicBomSerial(...values) {
  const digest = sha256(Buffer.from(values.join("\u0000"), "utf8"));
  // UUIDv5-shaped, deterministic identifier: the SBOM must be reproducible
  // across reviewed rebuilds while still satisfying CycloneDX's serialNumber
  // requirement for GitHub's SBOM attestation parser.
  return `urn:uuid:${digest.slice(0, 8)}-${digest.slice(8, 12)}-5${digest.slice(13, 16)}-${(Number.parseInt(digest[16], 16) & 0x3 | 0x8).toString(16)}${digest.slice(17, 20)}-${digest.slice(20, 32)}`;
}

function sbomFor({ pack, version, platform, sourceUrl, sourceSha256, license, payloadEntries }) {
  return {
    bomFormat: "CycloneDX",
    specVersion: "1.5",
    serialNumber: deterministicBomSerial(pack, version, platform, sourceUrl, sourceSha256),
    version: 1,
    metadata: {
      component: {
        type: "application",
        name: `open-office-artifact-tool-pdf-provider-${pack}`,
        version,
        properties: [
          { name: "open-office-artifact-tool:platform", value: platform },
          { name: "open-office-artifact-tool:archive-format", value: "ustar+gzip" },
        ],
      },
    },
    components: [
      {
        type: "application",
        name: pack,
        version,
        licenses: [{ license: { id: license } }],
        externalReferences: [{ type: "distribution", url: sourceUrl, hashes: [{ alg: "SHA-256", content: sourceSha256 }] }],
      },
    ],
    properties: [
      { name: "open-office-artifact-tool:payload-file-count", value: String(payloadEntries.filter((entry) => entry.type === "file").length) },
      { name: "open-office-artifact-tool:payload-unpacked-bytes", value: String(payloadEntries.filter((entry) => entry.type === "file").reduce((total, entry) => total + entry.bytes.length, 0)) },
    ],
  };
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const [payloadEntries, noticeBytes] = await Promise.all([
    listPayload(options.payload),
    regularFileBytes(options.notices, "--notices"),
  ]);
  const baseName = `${options.pack}-${options.version}-${options.platform}`;
  const sbomName = `${baseName}.sbom.cdx.json`;
  const noticesName = `${baseName}.THIRD_PARTY_NOTICES.md`;
  const sbomBytes = Buffer.from(stableJson(sbomFor({ ...options, payloadEntries })), "utf8");
  const archiveEntries = [
    ...payloadEntries,
    { path: "THIRD_PARTY_NOTICES.md", type: "file", mode: 0o644, bytes: noticeBytes },
    { path: "sbom.cdx.json", type: "file", mode: 0o644, bytes: sbomBytes },
  ].sort((left, right) => left.path.localeCompare(right.path, "en"));
  const archive = tarGzip(archiveEntries);
  const unpackedBytes = archiveEntries.filter((entry) => entry.type === "file").reduce((total, entry) => total + entry.bytes.length, 0);
  const asset = `${baseName}.tar.gz`;
  const manifest = {
    schema: PACK_SCHEMA,
    schemaVersion: 1,
    pack: options.pack,
    version: options.version,
    platform: options.platform,
    artifact: {
      asset,
      sha256: sha256(archive),
      downloadBytes: archive.length,
      unpackedBytes,
      archiveFormat: "tar.gz",
    },
    source: { url: options.sourceUrl, sha256: options.sourceSha256 },
    sbom: { asset: sbomName, sha256: sha256(sbomBytes), bytes: sbomBytes.length },
    thirdPartyNotices: { asset: noticesName, sha256: sha256(noticeBytes), bytes: noticeBytes.length },
    payload: {
      entries: archiveEntries.map((entry) => ({ path: entry.path, type: entry.type, mode: entry.mode, bytes: entry.bytes.length })),
      fileCount: archiveEntries.filter((entry) => entry.type === "file").length,
    },
  };
  await fs.mkdir(options.output, { recursive: true, mode: 0o700 });
  await Promise.all([
    fs.writeFile(path.join(options.output, asset), archive, { mode: 0o600 }),
    fs.writeFile(path.join(options.output, sbomName), sbomBytes, { mode: 0o600 }),
    fs.writeFile(path.join(options.output, noticesName), noticeBytes, { mode: 0o600 }),
    fs.writeFile(path.join(options.output, `${baseName}.manifest.json`), stableJson(manifest), { mode: 0o600 }),
  ]);
  process.stdout.write(`${stableJson(manifest)}`);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
