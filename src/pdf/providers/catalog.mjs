/**
 * Canonical, immutable capability-pack catalog.
 *
 * Keep this leaf free of provider imports and network operations. Importing the
 * public `./pdf/providers` subpath must be as inert as importing JSON metadata:
 * in particular it must not initialize MuPDF's WASM runtime.
 */

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const CATALOG_PATH = fileURLToPath(new URL("./provider-catalog.v1.json", import.meta.url));
const PACK_STATES = new Set(["built-in", "unpublished", "published"]);
const PACK_DELIVERIES = new Set(["npm-package", "versioned-signed-platform-release-asset"]);
const SUPPORTED_MANAGED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const SHA256 = /^[a-f0-9]{64}$/i;

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function catalogError(message) {
  return new Error(`Invalid PDF provider catalog: ${message}`);
}

function assertStringArray(value, label, { allowEmpty = false } = {}) {
  if (!Array.isArray(value) || (!allowEmpty && value.length === 0) || value.some((entry) => !nonEmptyString(entry))) {
    throw catalogError(`${label} must be ${allowEmpty ? "a" : "a non-empty"} string array.`);
  }
}

function isSafeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\0")) return false;
  const normalized = path.posix.normalize(value);
  return !path.posix.isAbsolute(value)
    && normalized !== "."
    && normalized !== ".."
    && !normalized.startsWith("../")
    && !normalized.includes("/../")
    && !value.startsWith("/")
    && !value.includes("\\");
}

function assertLicense(value, label) {
  if (!isPlainObject(value) || !nonEmptyString(value.expression) || typeof value.requiresAcknowledgement !== "boolean") {
    throw catalogError(`${label} must declare expression and requiresAcknowledgement.`);
  }
  if (value.id !== undefined && !nonEmptyString(value.id)) throw catalogError(`${label}.id must be a non-empty string.`);
  if (value.acceptedValues !== undefined) assertStringArray(value.acceptedValues, `${label}.acceptedValues`);
}

function assertEntrypoints(value, label) {
  if (!Array.isArray(value)) throw catalogError(`${label} must be an array.`);
  const seen = new Set();
  for (const entry of value) {
    if (!isPlainObject(entry) || !isSafeRelativePath(entry.path) || !["file", "directory"].includes(entry.kind) || typeof entry.executable !== "boolean") {
      throw catalogError(`${label} contains an invalid entrypoint.`);
    }
    if (seen.has(entry.path)) throw catalogError(`${label} contains duplicate entrypoint ${entry.path}.`);
    seen.add(entry.path);
  }
}

function assertHttpsAssetUrl(value, label) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw catalogError(`${label} must be an HTTPS URL.`);
  }
  if (parsed.protocol !== "https:" || /(^|[/-])latest([/-]|$)/i.test(parsed.pathname)) {
    throw catalogError(`${label} must be a versioned HTTPS release URL, never latest.`);
  }
}

function assertEvidenceAsset(value, label) {
  if (!isPlainObject(value) || !nonEmptyString(value.asset) || !SHA256.test(value.sha256 || "") || !nonEmptyString(value.url)) {
    throw catalogError(`${label} must include asset, HTTPS URL, and sha256.`);
  }
  if (!isSafeRelativePath(value.asset)) throw catalogError(`${label}.asset must be a safe relative path.`);
  assertHttpsAssetUrl(value.url, `${label}.url`);
}

function assertArtifact(value, packId, pack, index) {
  const label = `pack ${packId}.artifacts[${index}]`;
  if (!isPlainObject(value) || !nonEmptyString(value.platform) || !nonEmptyString(value.asset) || !nonEmptyString(value.version)) {
    throw catalogError(`${label} is incomplete.`);
  }
  if (!SUPPORTED_MANAGED_PLATFORMS.has(value.platform) || !pack.platforms.includes(value.platform)) {
    throw catalogError(`${label}.platform is outside the declared managed platforms.`);
  }
  if (value.version !== pack.version || !isSafeRelativePath(value.asset) || !nonEmptyString(value.url)) {
    throw catalogError(`${label} must pin the pack version, a safe asset name, and a URL.`);
  }
  assertHttpsAssetUrl(value.url, `${label}.url`);
  if (!SHA256.test(value.sha256 || "") || !Number.isSafeInteger(value.downloadBytes) || value.downloadBytes <= 0 || !Number.isSafeInteger(value.unpackedBytes) || value.unpackedBytes <= 0) {
    throw catalogError(`${label} must pin SHA-256 and positive download/unpacked sizes.`);
  }
  if (value.archiveFormat !== "tar.gz") throw catalogError(`${label}.archiveFormat must be tar.gz.`);
}

function assertManagedRuntime(provider, pack, providerId) {
  if (pack.state === "built-in") return;
  const runtime = provider.managedRuntime;
  if (!isPlainObject(runtime)) throw catalogError(`provider ${providerId} must declare managedRuntime for pack ${provider.packId}.`);
  const entrypoints = new Map(pack.entrypoints.map((entry) => [entry.path, entry]));
  const referenced = [];
  if (runtime.pythonPath !== undefined) {
    if (!isSafeRelativePath(runtime.pythonPath)) throw catalogError(`provider ${providerId}.managedRuntime.pythonPath is unsafe.`);
    referenced.push(runtime.pythonPath);
  }
  if (runtime.commandPaths !== undefined) {
    if (!isPlainObject(runtime.commandPaths)) throw catalogError(`provider ${providerId}.managedRuntime.commandPaths must be an object.`);
    for (const [command, target] of Object.entries(runtime.commandPaths)) {
      if (!nonEmptyString(command) || !isSafeRelativePath(target)) throw catalogError(`provider ${providerId}.managedRuntime.commandPaths is unsafe.`);
      referenced.push(target);
    }
  }
  if (!referenced.length) throw catalogError(`provider ${providerId}.managedRuntime must declare a runtime path.`);
  for (const target of referenced) {
    const entrypoint = entrypoints.get(target);
    if (!entrypoint || entrypoint.kind !== "file") throw catalogError(`provider ${providerId} references undeclared entrypoint ${target}.`);
  }
  if (runtime.environment !== undefined) {
    if (!isPlainObject(runtime.environment)) throw catalogError(`provider ${providerId}.managedRuntime.environment must be an object.`);
    for (const [name, target] of Object.entries(runtime.environment)) {
      if (!/^[A-Z][A-Z0-9_]*$/.test(name) || !isSafeRelativePath(target)) throw catalogError(`provider ${providerId}.managedRuntime.environment is unsafe.`);
      if (!entrypoints.has(target) && !target.startsWith("share/")) throw catalogError(`provider ${providerId}.managedRuntime.environment references undeclared runtime ${target}.`);
    }
  }
}

/** Validate a candidate catalog before it can inform a download or runtime path. */
export function validatePdfProviderCatalog(catalog) {
  if (!isPlainObject(catalog)) throw catalogError("catalog must be an object.");
  if (catalog.schema !== "open-office-artifact-tool.pdf-provider-catalog.v1" || catalog.schemaVersion !== 1) {
    throw catalogError("unsupported schema or schemaVersion.");
  }
  if (!isPlainObject(catalog.releasePolicy)
    || catalog.releasePolicy.managedArtifacts !== "versioned-signed-platform-release-assets"
    || catalog.releasePolicy.forbidLatestUrls !== true
    || catalog.releasePolicy.forbidGlobalInstallers !== true
    || catalog.releasePolicy.cacheScope !== "project-private"
    || catalog.releasePolicy.defaultInstallPolicy !== "disabled") {
    throw catalogError("releasePolicy does not enforce managed, disabled-by-default releases.");
  }
  assertStringArray(catalog.releasePolicy.managedPlatforms, "releasePolicy.managedPlatforms");
  if (catalog.releasePolicy.managedPlatforms.some((platform) => !SUPPORTED_MANAGED_PLATFORMS.has(platform))) {
    throw catalogError("releasePolicy lists an unsupported managed platform.");
  }
  if (!isPlainObject(catalog.packs) || Object.keys(catalog.packs).length === 0) throw catalogError("packs must be a non-empty object.");
  if (!isPlainObject(catalog.providers) || Object.keys(catalog.providers).length === 0) throw catalogError("providers must be a non-empty object.");
  if (!isPlainObject(catalog.tasks) || Object.keys(catalog.tasks).length === 0) throw catalogError("tasks must be a non-empty object.");

  for (const [packId, pack] of Object.entries(catalog.packs)) {
    if (!nonEmptyString(packId) || !isSafeRelativePath(packId) || !isPlainObject(pack)) throw catalogError(`pack ${packId || "<empty>"} is invalid.`);
    if (!PACK_STATES.has(pack.state) || !PACK_DELIVERIES.has(pack.delivery)) throw catalogError(`pack ${packId} has an invalid state or delivery.`);
    assertStringArray(pack.platforms, `pack ${packId}.platforms`);
    assertStringArray(pack.requiresPackIds, `pack ${packId}.requiresPackIds`, { allowEmpty: true });
    assertStringArray(pack.dependencyClosure, `pack ${packId}.dependencyClosure`);
    assertEntrypoints(pack.entrypoints, `pack ${packId}.entrypoints`);
    assertLicense(pack.license, `pack ${packId}.license`);
    if (!Array.isArray(pack.artifacts)) throw catalogError(`pack ${packId}.artifacts must be an array.`);
    if (pack.state === "built-in") {
      if (pack.delivery !== "npm-package" || !pack.platforms.includes("any") || pack.artifacts.length !== 0) throw catalogError(`built-in pack ${packId} is malformed.`);
    } else {
      if (pack.delivery !== "versioned-signed-platform-release-asset" || pack.platforms.some((platform) => !SUPPORTED_MANAGED_PLATFORMS.has(platform))) {
        throw catalogError(`managed pack ${packId} declares an unsupported delivery or platform.`);
      }
    }
    if (pack.state === "published") {
      if (!nonEmptyString(pack.version) || pack.artifacts.length === 0 || !isPlainObject(pack.releaseEvidence)) throw catalogError(`published pack ${packId} lacks version, artifacts, or release evidence.`);
      for (const [index, artifact] of pack.artifacts.entries()) assertArtifact(artifact, packId, pack, index);
      assertEvidenceAsset(pack.releaseEvidence.sbom, `pack ${packId}.releaseEvidence.sbom`);
      assertEvidenceAsset(pack.releaseEvidence.thirdPartyNotices, `pack ${packId}.releaseEvidence.thirdPartyNotices`);
      assertStringArray(pack.releaseEvidence.verifiedPlatforms, `pack ${packId}.releaseEvidence.verifiedPlatforms`);
    } else if (pack.state === "unpublished" && (pack.artifacts.length !== 0 || pack.version !== null)) {
      throw catalogError(`${pack.state} pack ${packId} must not advertise unpublished artifacts or a version.`);
    }
  }

  for (const [packId, pack] of Object.entries(catalog.packs)) {
    for (const dependency of pack.requiresPackIds) {
      if (!catalog.packs[dependency] || dependency === packId) throw catalogError(`pack ${packId} has an invalid dependency ${dependency}.`);
    }
  }

  for (const [providerId, provider] of Object.entries(catalog.providers)) {
    if (!nonEmptyString(providerId) || !isPlainObject(provider) || !nonEmptyString(provider.kind) || !nonEmptyString(provider.packId)
      || !nonEmptyString(provider.role) || !nonEmptyString(provider.integration)) {
      throw catalogError(`provider ${providerId || "<empty>"} is incomplete.`);
    }
    const pack = catalog.packs[provider.packId];
    if (!pack) throw catalogError(`provider ${providerId} references unknown pack ${provider.packId}.`);
    assertStringArray(provider.taskIds, `provider ${providerId}.taskIds`, { allowEmpty: true });
    assertLicense(provider.license, `provider ${providerId}.license`);
    if (provider.kind === "node-package" && (!nonEmptyString(provider.package) || !nonEmptyString(provider.exactVersion))) {
      throw catalogError(`node-package provider ${providerId} is incomplete.`);
    }
    if (provider.kind === "python-module" && !nonEmptyString(provider.module)) throw catalogError(`python provider ${providerId} is missing module.`);
    if (provider.kind === "command" && (!Array.isArray(provider.commands) || provider.commands.some((command) => !nonEmptyString(command)))) throw catalogError(`command provider ${providerId} is missing commands.`);
    assertManagedRuntime(provider, pack, providerId);
  }

  if (!isPlainObject(catalog.ocrLanguagePacks)) throw catalogError("ocrLanguagePacks must be an object.");
  for (const [language, packId] of Object.entries(catalog.ocrLanguagePacks)) {
    if (!/^[a-z0-9_]+$/i.test(language) || !catalog.packs[packId]) throw catalogError(`OCR language ${language} maps to an unknown pack.`);
  }

  for (const [taskId, task] of Object.entries(catalog.tasks)) {
    if (!nonEmptyString(taskId) || !isPlainObject(task)) throw catalogError(`task ${taskId || "<empty>"} is invalid.`);
    assertStringArray(task.providers, `task ${taskId}.providers`);
    assertStringArray(task.strategies, `task ${taskId}.strategies`);
    if (!nonEmptyString(task.input)) throw catalogError(`task ${taskId} must specify an input mode.`);
    if (task.defaultProvider !== undefined && !task.providers.includes(task.defaultProvider)) throw catalogError(`task ${taskId} has an invalid default provider.`);
    if (task.credentials !== undefined) assertStringArray(task.credentials, `task ${taskId}.credentials`);
    for (const providerId of task.providers) {
      const provider = catalog.providers[providerId];
      if (!provider || !provider.taskIds.includes(taskId)) throw catalogError(`task ${taskId} and provider ${providerId} disagree.`);
    }
  }

  const mupdf = catalog.providers["mupdf-js"];
  if (!mupdf || mupdf.packId !== "mupdf-js" || mupdf.exactVersion !== "1.28.0" || catalog.packs["mupdf-js"].state !== "built-in") {
    throw catalogError("mupdf-js must remain the pinned, built-in lazy default.");
  }
  return true;
}

const catalogBytes = fs.readFileSync(CATALOG_PATH);
const loadedCatalog = JSON.parse(catalogBytes.toString("utf8"));
validatePdfProviderCatalog(loadedCatalog);

function deepFreeze(value) {
  if (!value || typeof value !== "object" || Object.isFrozen(value)) return value;
  for (const child of Object.values(value)) deepFreeze(child);
  return Object.freeze(value);
}

export const PDF_PROVIDER_CATALOG = deepFreeze(loadedCatalog);
export const PDF_PROVIDER_CATALOG_SCHEMA = PDF_PROVIDER_CATALOG.schema;
export const PDF_PROVIDER_CATALOG_SHA256 = crypto.createHash("sha256").update(catalogBytes).digest("hex");

export function clonePdfProviderValue(value) {
  return structuredClone(value);
}

export function pdfProviderById(providerId) {
  const provider = PDF_PROVIDER_CATALOG.providers[providerId];
  if (!provider) throw new TypeError(`Unknown PDF provider: ${providerId}.`);
  return provider;
}

export function pdfPackById(packId) {
  const pack = PDF_PROVIDER_CATALOG.packs[packId];
  if (!pack) throw new TypeError(`Unknown PDF capability pack: ${packId}.`);
  return pack;
}

export function pdfTaskById(taskId) {
  const task = PDF_PROVIDER_CATALOG.tasks[taskId];
  if (!task) throw new TypeError(`Unknown PDF capability task: ${taskId}.`);
  return task;
}

export function currentPdfProviderPlatform() {
  return `${process.platform}-${process.arch}`;
}

export function isSafePdfProviderRelativePath(value) {
  return isSafeRelativePath(value);
}
