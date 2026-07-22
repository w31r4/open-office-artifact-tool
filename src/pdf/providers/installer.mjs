/**
 * Private, hash-pinned managed-pack installer.
 *
 * This module is intentionally not a package export. The public facade only
 * calls it after resolver policy checks. Keeping the primitive here makes the
 * archive/cache hardening directly testable without creating a second catalog
 * or a user-facing global installer.
 */

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { createGunzip } from "node:zlib";

import {
  PDF_PROVIDER_CATALOG,
  PDF_PROVIDER_CATALOG_SHA256,
  clonePdfProviderValue,
  currentPdfProviderPlatform,
  isSafePdfProviderRelativePath,
  pdfPackById,
  pdfProviderById,
} from "./catalog.mjs";

const RECEIPT_SCHEMA = "open-office-artifact-tool.pdf-provider-receipt.v1";
const RECEIPT_FILE = ".receipt.json";
const MAX_RECEIPT_BYTES = 128 * 1024;
const LOCK_TIMEOUT_MS = 20_000;
const LOCK_RETRY_MS = 25;
const MAX_DOWNLOAD_REDIRECTS = 5;

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function safePathSegment(value, label) {
  if (!nonEmptyString(value) || !isSafePdfProviderRelativePath(value) || value.includes("/")) {
    throw new Error(`${label} must be a single safe path segment.`);
  }
  return value;
}

function containedPath(root, relativePath) {
  if (!isSafePdfProviderRelativePath(relativePath)) throw new Error(`Unsafe capability-pack path: ${relativePath}.`);
  const resolvedRoot = path.resolve(root);
  const target = path.resolve(resolvedRoot, ...relativePath.split("/"));
  if (target === resolvedRoot || !target.startsWith(`${resolvedRoot}${path.sep}`)) throw new Error(`Capability-pack path escapes its cache: ${relativePath}.`);
  return target;
}

async function ensureRealDirectory(directory, { create = false } = {}) {
  if (create) {
    try {
      await fs.promises.mkdir(directory, { recursive: false, mode: 0o700 });
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
    }
  }
  const stat = await fs.promises.lstat(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`Managed provider cache path must be a real directory: ${directory}.`);
  return stat;
}

async function ensureDirectoryWithin(root, relativeDirectory) {
  let cursor = root;
  for (const segment of relativeDirectory.split("/").filter(Boolean)) {
    cursor = path.join(cursor, segment);
    try {
      const stat = await fs.promises.lstat(cursor);
      if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`Capability-pack archive uses an unsafe directory: ${relativeDirectory}.`);
    } catch (error) {
      if (error?.code !== "ENOENT") throw error;
      try {
        await fs.promises.mkdir(cursor, { mode: 0o700 });
      } catch (mkdirError) {
        if (mkdirError?.code !== "EEXIST") throw mkdirError;
        const stat = await fs.promises.lstat(cursor);
        if (!stat.isDirectory() || stat.isSymbolicLink()) throw new Error(`Capability-pack archive uses an unsafe directory: ${relativeDirectory}.`);
      }
    }
  }
}

function tarText(header, offset, length) {
  const raw = header.subarray(offset, offset + length);
  const zero = raw.indexOf(0);
  return raw.subarray(0, zero === -1 ? raw.length : zero).toString("utf8").trim();
}

function tarOctal(header, offset, length, label) {
  const raw = tarText(header, offset, length).replace(/^0+/, "") || "0";
  if (!/^[0-7]+$/.test(raw)) throw new Error(`Malformed tar ${label}.`);
  const value = Number.parseInt(raw, 8);
  if (!Number.isSafeInteger(value) || value < 0) throw new Error(`Unsafe tar ${label}.`);
  return value;
}

function tarChecksum(header) {
  let sum = 0;
  for (let index = 0; index < header.length; index += 1) sum += index >= 148 && index < 156 ? 32 : header[index];
  return sum;
}

function isZeroBlock(block) {
  for (const byte of block) if (byte !== 0) return false;
  return true;
}

async function gunzipLimited(compressed, maxBytes) {
  const gunzip = createGunzip();
  const chunks = [];
  let total = 0;
  const completed = new Promise((resolve, reject) => {
    gunzip.on("data", (chunk) => {
      total += chunk.length;
      if (total > maxBytes) {
        gunzip.destroy(new Error(`Capability-pack archive exceeds declared unpacked limit of ${maxBytes} bytes.`));
        return;
      }
      chunks.push(Buffer.from(chunk));
    });
    gunzip.once("end", resolve);
    gunzip.once("error", reject);
  });
  gunzip.end(compressed);
  await completed;
  return Buffer.concat(chunks, total);
}

/** Safely extract a strict USTAR tar.gz payload into a fresh staging directory. */
export async function safeExtractTarGz(compressed, destination, maxUnpackedBytes) {
  if (!Buffer.isBuffer(compressed)) throw new TypeError("Capability-pack archive must be bytes.");
  if (!Number.isSafeInteger(maxUnpackedBytes) || maxUnpackedBytes <= 0) throw new TypeError("Capability-pack unpacked limit must be a positive safe integer.");
  await ensureRealDirectory(destination);
  const tar = await gunzipLimited(compressed, maxUnpackedBytes + 1024 * 1024);
  const seen = new Set();
  let offset = 0;
  let unpackedBytes = 0;
  while (offset < tar.length) {
    if (offset + 512 > tar.length) throw new Error("Truncated capability-pack tar header.");
    const header = tar.subarray(offset, offset + 512);
    offset += 512;
    if (isZeroBlock(header)) {
      if (offset < tar.length && !isZeroBlock(tar.subarray(offset, Math.min(offset + 512, tar.length)))) throw new Error("Unexpected data after tar terminator.");
      break;
    }
    const expectedChecksum = tarOctal(header, 148, 8, "checksum");
    if (tarChecksum(header) !== expectedChecksum) throw new Error("Capability-pack tar header checksum mismatch.");
    const name = tarText(header, 0, 100);
    const prefix = tarText(header, 345, 155);
    const entryPath = prefix ? `${prefix}/${name}` : name;
    const type = String.fromCharCode(header[156] || 0);
    const size = tarOctal(header, 124, 12, "entry size");
    const mode = tarOctal(header, 100, 8, "entry mode") & 0o777;
    if (!isSafePdfProviderRelativePath(entryPath) || seen.has(entryPath)) throw new Error(`Unsafe or duplicate capability-pack archive entry: ${entryPath}.`);
    seen.add(entryPath);
    if (offset + size > tar.length) throw new Error(`Truncated capability-pack archive entry: ${entryPath}.`);
    const body = tar.subarray(offset, offset + size);
    const paddedSize = Math.ceil(size / 512) * 512;
    if (offset + paddedSize > tar.length) throw new Error(`Truncated capability-pack archive padding: ${entryPath}.`);
    offset += paddedSize;

    if (type === "5") {
      if (size !== 0) throw new Error(`Capability-pack directory ${entryPath} unexpectedly contains data.`);
      await ensureDirectoryWithin(destination, entryPath);
      continue;
    }
    // Reject hard links, symlinks, devices, FIFOs, GNU long names, and PAX.
    if (type !== "\0" && type !== "0") throw new Error(`Unsupported or unsafe capability-pack tar entry type for ${entryPath}.`);
    unpackedBytes += size;
    if (unpackedBytes > maxUnpackedBytes) throw new Error(`Capability-pack archive exceeds declared unpacked limit of ${maxUnpackedBytes} bytes.`);
    const parent = path.posix.dirname(entryPath);
    if (parent !== ".") await ensureDirectoryWithin(destination, parent);
    const target = containedPath(destination, entryPath);
    await fs.promises.writeFile(target, body, { flag: "wx", mode: mode || 0o600 });
  }
  return { unpackedBytes, entries: [...seen].sort() };
}

function artifactForPlatform(pack, platform) {
  const artifacts = pack.artifacts.filter((artifact) => artifact.platform === platform);
  if (artifacts.length !== 1) throw new Error(`Pack ${pack.id || "<unknown>"} does not have exactly one artifact for ${platform}.`);
  return artifacts[0];
}

function collectPackIds(requestedPackIds) {
  const ordered = [];
  const visiting = new Set();
  const visited = new Set();
  const visit = (packId) => {
    if (visited.has(packId)) return;
    if (visiting.has(packId)) throw new Error(`Capability-pack dependency cycle at ${packId}.`);
    visiting.add(packId);
    const pack = pdfPackById(packId);
    for (const dependency of pack.requiresPackIds) visit(dependency);
    visiting.delete(packId);
    visited.add(packId);
    ordered.push(packId);
  };
  for (const packId of requestedPackIds) visit(packId);
  return ordered;
}

function licenseAccepted(license, policy) {
  if (!license.requiresAcknowledgement) return true;
  const accepted = new Set(policy.acceptedLicenses);
  return Boolean((license.id && accepted.has(license.id)) || license.acceptedValues?.some((value) => accepted.has(value)));
}

function assertManagedAuthorization(provider, packIds, policy) {
  if (policy.installPolicy !== "managed") throw new Error("Managed capability installation requires installPolicy: managed.");
  if (!policy.allowedProviders.includes(provider.id)) throw new Error(`Policy does not allow provider ${provider.id || "<unknown>"}.`);
  if (!licenseAccepted(provider.license, policy)) throw new Error(`Provider ${provider.id || "<unknown>"} requires an unaccepted license acknowledgement.`);
  for (const packId of packIds) {
    const pack = pdfPackById(packId);
    if (!policy.allowedPacks.includes(packId)) throw new Error(`Policy does not allow capability pack ${packId}.`);
    if (!licenseAccepted(pack.license, policy)) throw new Error(`Capability pack ${packId} requires an unaccepted license acknowledgement.`);
  }
}

function resolveArtifactUrl(artifact, enterpriseMirror) {
  if (!enterpriseMirror) return artifact.url;
  const base = new URL(enterpriseMirror);
  const resolved = new URL(artifact.asset, base);
  if (resolved.origin !== base.origin || !resolved.pathname.startsWith(base.pathname)) {
    throw new Error("Enterprise mirror would escape its declared base URL.");
  }
  return resolved.href;
}

function isRedirectStatus(status) {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

async function fetchPinnedArtifact(initialUrl, asset, fetchImpl) {
  let url = initialUrl;
  for (let redirects = 0; redirects <= MAX_DOWNLOAD_REDIRECTS; redirects += 1) {
    let response;
    try {
      // Release assets legitimately redirect to a short-lived object URL. We
      // retain control of every hop instead of allowing the runtime's generic
      // redirect behavior, and still bind the final bytes to the catalog hash.
      response = await fetchImpl(url, { redirect: "manual" });
    } catch (error) {
      throw new Error(`Capability-pack download failed for ${asset}: ${String(error?.message || error)}.`);
    }
    if (!response) throw new Error(`Capability-pack download failed for ${asset}: no HTTP response.`);
    if (!isRedirectStatus(response.status)) return response;
    if (redirects === MAX_DOWNLOAD_REDIRECTS) throw new Error(`Capability-pack download exceeded ${MAX_DOWNLOAD_REDIRECTS} HTTPS redirects for ${asset}.`);
    const location = response.headers?.get?.("location");
    if (!nonEmptyString(location)) throw new Error(`Capability-pack download returned a redirect without a location for ${asset}.`);
    let next;
    try {
      next = new URL(location, url);
    } catch {
      throw new Error(`Capability-pack download returned an invalid redirect for ${asset}.`);
    }
    if (next.protocol !== "https:" || next.username || next.password) {
      throw new Error(`Capability-pack download redirect must use credential-free HTTPS for ${asset}.`);
    }
    url = next.href;
  }
  throw new Error(`Capability-pack download exceeded ${MAX_DOWNLOAD_REDIRECTS} HTTPS redirects for ${asset}.`);
}

async function downloadArtifact(artifact, enterpriseMirror, fetchImpl) {
  if (typeof fetchImpl !== "function") throw new Error("No fetch implementation is available for managed capability installation.");
  const url = resolveArtifactUrl(artifact, enterpriseMirror);
  const response = await fetchPinnedArtifact(url, artifact.asset, fetchImpl);
  if (!response || response.ok !== true || !response.body) throw new Error(`Capability-pack download failed for ${artifact.asset}: HTTP ${response?.status ?? "unknown"}.`);
  const chunks = [];
  let downloadedBytes = 0;
  for await (const chunk of response.body) {
    const bytes = Buffer.from(chunk);
    downloadedBytes += bytes.length;
    if (downloadedBytes > artifact.downloadBytes) throw new Error(`Capability-pack download exceeds pinned size for ${artifact.asset}.`);
    chunks.push(bytes);
  }
  if (downloadedBytes !== artifact.downloadBytes) throw new Error(`Capability-pack download size mismatch for ${artifact.asset}.`);
  const bytes = Buffer.concat(chunks, downloadedBytes);
  if (sha256(bytes) !== artifact.sha256.toLowerCase()) throw new Error(`Capability-pack SHA-256 mismatch for ${artifact.asset}.`);
  return { bytes, downloadedBytes, url };
}

function receiptFor({ packId, pack, platform, artifact, extraction, catalogSha256 = PDF_PROVIDER_CATALOG_SHA256 }) {
  return {
    schema: RECEIPT_SCHEMA,
    schemaVersion: 1,
    catalogSha256,
    packId,
    version: pack.version,
    platform,
    artifact: {
      asset: artifact.asset,
      version: artifact.version,
      sha256: artifact.sha256.toLowerCase(),
      downloadBytes: artifact.downloadBytes,
      unpackedBytes: artifact.unpackedBytes,
    },
    entrypoints: clonePdfProviderValue(pack.entrypoints),
    extractedBytes: extraction.unpackedBytes,
    installedAt: new Date().toISOString(),
  };
}

async function writeReceipt(root, receipt) {
  const target = path.join(root, RECEIPT_FILE);
  await fs.promises.writeFile(target, `${JSON.stringify(receipt, null, 2)}\n`, { encoding: "utf8", flag: "wx", mode: 0o600 });
}

async function readReceipt(root) {
  const target = path.join(root, RECEIPT_FILE);
  const stat = await fs.promises.lstat(target);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > MAX_RECEIPT_BYTES) throw new Error("Managed capability receipt is unsafe.");
  return JSON.parse(await fs.promises.readFile(target, "utf8"));
}

async function verifyEntrypoints(root, entrypoints) {
  for (const entry of entrypoints) {
    const target = containedPath(root, entry.path);
    const stat = await fs.promises.lstat(target);
    if (stat.isSymbolicLink() || (entry.kind === "file" && !stat.isFile()) || (entry.kind === "directory" && !stat.isDirectory())) {
      throw new Error(`Managed capability entrypoint is invalid: ${entry.path}.`);
    }
    if (entry.executable && process.platform !== "win32" && (stat.mode & 0o111) === 0) {
      throw new Error(`Managed capability entrypoint is not executable: ${entry.path}.`);
    }
  }
}

function locationForPack(cacheRoot, packId, version, platform) {
  return path.join(cacheRoot, safePathSegment(packId, "packId"), safePathSegment(version, "pack version"), safePathSegment(platform, "platform"));
}

async function installedPackRecord({ cacheRoot, packId, pack, platform, catalogSha256 = PDF_PROVIDER_CATALOG_SHA256 }) {
  if (pack.state !== "published") return { ready: false, reason: `${pack.state}-pack` };
  const root = locationForPack(cacheRoot, packId, pack.version, platform);
  try {
    const rootStat = await fs.promises.lstat(root);
    if (!rootStat.isDirectory() || rootStat.isSymbolicLink()) return { ready: false, reason: "unsafe-cache-root", root };
    const artifact = artifactForPlatform(pack, platform);
    const receipt = await readReceipt(root);
    if (!isPlainObject(receipt)
      || receipt.schema !== RECEIPT_SCHEMA
      || receipt.schemaVersion !== 1
      || receipt.catalogSha256 !== catalogSha256
      || receipt.packId !== packId
      || receipt.version !== pack.version
      || receipt.platform !== platform
      || receipt.artifact?.asset !== artifact.asset
      || receipt.artifact?.sha256 !== artifact.sha256.toLowerCase()
      || receipt.artifact?.downloadBytes !== artifact.downloadBytes
      || receipt.artifact?.unpackedBytes !== artifact.unpackedBytes) {
      return { ready: false, reason: "receipt-mismatch", root };
    }
    await verifyEntrypoints(root, pack.entrypoints);
    return { ready: true, root, receipt };
  } catch (error) {
    if (error?.code === "ENOENT") return { ready: false, reason: "not-installed", root };
    return { ready: false, reason: "invalid-cache", root, error: String(error?.message || error) };
  }
}

async function installedPack(cacheRoot, packId, platform) {
  return installedPackRecord({ cacheRoot, packId, pack: pdfPackById(packId), platform });
}

async function acquireLock(lockPath) {
  const deadline = Date.now() + LOCK_TIMEOUT_MS;
  while (true) {
    try {
      await fs.promises.mkdir(lockPath, { mode: 0o700 });
      return async () => {
        await fs.promises.rm(lockPath, { recursive: true, force: true });
      };
    } catch (error) {
      if (error?.code !== "EEXIST") throw error;
      const stat = await fs.promises.lstat(lockPath).catch(() => undefined);
      if (stat && (!stat.isDirectory() || stat.isSymbolicLink())) throw new Error("Capability-pack lock path is unsafe.");
      if (Date.now() >= deadline) throw new Error(`Timed out waiting for managed capability-pack lock: ${lockPath}.`);
      await new Promise((resolve) => setTimeout(resolve, LOCK_RETRY_MS));
    }
  }
}

async function installPackRecord({ cacheRoot, packId, pack, platform, enterpriseMirror, fetchImpl, catalogSha256 = PDF_PROVIDER_CATALOG_SHA256 }) {
  if (pack.state !== "published") throw new Error(`Capability pack ${packId} is ${pack.state}; no immutable release artifact is available.`);
  const artifact = artifactForPlatform(pack, platform);
  const target = locationForPack(cacheRoot, packId, pack.version, platform);
  const parent = path.dirname(target);
  await ensureDirectoryWithin(cacheRoot, `${packId}/${pack.version}`);
  const releaseLock = await acquireLock(`${target}.lock`);
  let temporary;
  try {
    const existing = await installedPackRecord({ cacheRoot, packId, pack, platform, catalogSha256 });
    if (existing.ready) return { ...existing, reused: true };
    if (existing.reason !== "not-installed") throw new Error(`Refusing to replace invalid managed capability cache for ${packId}: ${existing.reason}.`);
    temporary = await fs.promises.mkdtemp(path.join(parent, `.${packId}.tmp-`));
    await ensureRealDirectory(temporary);
    const { bytes, downloadedBytes, url } = await downloadArtifact(artifact, enterpriseMirror, fetchImpl);
    await fs.promises.writeFile(path.join(temporary, "download.tar.gz"), bytes, { flag: "wx", mode: 0o600 });
    const staging = path.join(temporary, "payload");
    await fs.promises.mkdir(staging, { mode: 0o700 });
    const extraction = await safeExtractTarGz(bytes, staging, artifact.unpackedBytes);
    await verifyEntrypoints(staging, pack.entrypoints);
    await writeReceipt(staging, receiptFor({ packId, pack, platform, artifact, extraction, catalogSha256 }));
    await fs.promises.rename(staging, target);
    return {
      ready: true,
      root: target,
      receipt: await readReceipt(target),
      reused: false,
      downloadedBytes,
      sourceUrl: url,
    };
  } finally {
    if (temporary) await fs.promises.rm(temporary, { recursive: true, force: true });
    await releaseLock();
  }
}

async function installOnePack({ cacheRoot, packId, platform, enterpriseMirror, fetchImpl }) {
  return installPackRecord({ cacheRoot, packId, pack: pdfPackById(packId), platform, enterpriseMirror, fetchImpl });
}

function managedRuntime(provider, installed, languages = []) {
  const root = installed[provider.packId]?.root;
  if (!root) throw new Error(`Managed runtime is missing provider pack ${provider.packId}.`);
  const runtime = provider.managedRuntime;
  const absolute = (reference, label) => {
    const target = typeof reference === "string" ? { packId: provider.packId, path: reference } : reference;
    if (!isPlainObject(target) || !nonEmptyString(target.packId) || !isSafePdfProviderRelativePath(target.path)) {
      throw new Error(`Managed runtime has an unsafe ${label} reference.`);
    }
    const targetRoot = installed[target.packId]?.root;
    if (!targetRoot) throw new Error(`Managed runtime is missing dependency pack ${target.packId}.`);
    return containedPath(targetRoot, target.path);
  };
  const commandPaths = Object.fromEntries(Object.entries(runtime.commandPaths || {}).map(([command, target]) => [command, absolute(target, `command ${command}`)]));
  const languagePacks = languages.map((language) => {
    const packId = PDF_PROVIDER_CATALOG.ocrLanguagePacks[language];
    const languageRoot = packId ? installed[packId]?.root : undefined;
    if (!packId || !languageRoot) throw new Error(`Managed OCR language pack is missing: ${language}.`);
    const entrypoint = pdfPackById(packId).entrypoints.find((entry) => entry.kind === "file");
    if (!entrypoint) throw new Error(`Managed OCR language pack has no data entrypoint: ${packId}.`);
    const dataPath = containedPath(languageRoot, entrypoint.path);
    return { language, packId, dataPath, dataDirectory: path.dirname(dataPath) };
  });
  const environment = Object.fromEntries(Object.entries(runtime.environment || {}).map(([name, target]) => [name, absolute(target, `environment ${name}`)]));
  if (runtime.languageDirectoryEnvironment && languagePacks.length) {
    environment[runtime.languageDirectoryEnvironment] = [...new Set(languagePacks.map((language) => language.dataDirectory))].join(path.delimiter);
  }
  return {
    root,
    pythonPath: runtime.pythonPath ? absolute(runtime.pythonPath) : undefined,
    commandPaths,
    environment,
    languagePacks,
  };
}

/** Inspect one installed managed pack without downloading or repairing it. */
export async function probeManagedPack({ cacheRoot, packId, platform = currentPdfProviderPlatform() }) {
  if (!nonEmptyString(cacheRoot)) return { ready: false, reason: "missing-private-cache" };
  try {
    await ensureRealDirectory(cacheRoot);
  } catch (error) {
    if (error?.code === "ENOENT") return { ready: false, reason: "not-installed" };
    return { ready: false, reason: "unsafe-cache-root", error: String(error?.message || error) };
  }
  return installedPack(cacheRoot, packId, platform);
}

/**
 * Install an already-authorized set of published packs. `fetchImpl` is an
 * internal dependency seam for deterministic fake-artifact tests; it cannot
 * override catalog URLs, hashes, platforms, or policy checks.
 */
export async function ensureManagedPacks({ providerId, packIds, policyContext, languages = [], fetchImpl = globalThis.fetch }) {
  if (!isPlainObject(policyContext) || !isPlainObject(policyContext.policy)) throw new TypeError("Managed installation requires a resolved file-backed policy context.");
  if (!nonEmptyString(policyContext.cacheRoot) || !nonEmptyString(policyContext.policyPath)) {
    throw new Error("Managed capability installation requires a project policy file and private cache.");
  }
  const provider = pdfProviderById(providerId);
  const orderedPackIds = collectPackIds(packIds);
  assertManagedAuthorization({ ...provider, id: providerId }, orderedPackIds, policyContext.policy);
  await ensureRealDirectory(path.dirname(policyContext.cacheRoot));
  await ensureRealDirectory(policyContext.cacheRoot, { create: true });
  const platform = currentPdfProviderPlatform();
  const installed = {};
  for (const packId of orderedPackIds) {
    installed[packId] = await installOnePack({
      cacheRoot: policyContext.cacheRoot,
      packId,
      platform,
      enterpriseMirror: policyContext.policy.enterpriseMirror,
      fetchImpl,
    });
  }
  return {
    platform,
    packIds: orderedPackIds,
    installed,
    runtime: managedRuntime({ ...provider, id: providerId }, installed, languages),
  };
}

/** Return the provider-owned paths from verified receipts, without downloading. */
export async function probeManagedProviderRuntime({ providerId, packIds, policyContext, languages = [] }) {
  if (!isPlainObject(policyContext) || !nonEmptyString(policyContext.cacheRoot)) return { ready: false, reason: "missing-private-cache" };
  const provider = pdfProviderById(providerId);
  const platform = currentPdfProviderPlatform();
  const orderedPackIds = collectPackIds(packIds);
  const installed = {};
  for (const packId of orderedPackIds) {
    const result = await probeManagedPack({ cacheRoot: policyContext.cacheRoot, packId, platform });
    if (!result.ready) return { ready: false, reason: result.reason, packId, details: result };
    installed[packId] = result;
  }
  return { ready: true, platform, packIds: orderedPackIds, installed, runtime: managedRuntime({ ...provider, id: providerId }, installed, languages) };
}

/**
 * Deterministic leaf-test seam. It is deliberately absent from package exports:
 * tests can exercise the real downloader, lock, receipt, and extractor using a
 * fake hash-pinned release record without publishing a fake production pack.
 */
export async function installManagedPackForTest({ cacheRoot, packId = "fixture-pack", pack, platform, enterpriseMirror, fetchImpl, catalogSha256 = PDF_PROVIDER_CATALOG_SHA256 }) {
  if (!nonEmptyString(cacheRoot) || !nonEmptyString(packId) || !isPlainObject(pack)) throw new TypeError("Fixture install requires cacheRoot, packId, and pack.");
  if (pack.state !== "published" || !nonEmptyString(pack.version) || !Array.isArray(pack.artifacts) || !Array.isArray(pack.entrypoints)) {
    throw new TypeError("Fixture pack must be a published immutable pack record.");
  }
  await ensureRealDirectory(cacheRoot, { create: true });
  return installPackRecord({
    cacheRoot,
    packId,
    pack,
    platform: platform || currentPdfProviderPlatform(),
    enterpriseMirror,
    fetchImpl,
    catalogSha256,
  });
}

export const PDF_PROVIDER_RECEIPT_SCHEMA = RECEIPT_SCHEMA;
