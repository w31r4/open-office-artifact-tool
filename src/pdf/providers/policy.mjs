/** Project-scoped policy loading for managed PDF capability packs. */

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

import { PDF_PROVIDER_CATALOG, clonePdfProviderValue } from "./catalog.mjs";

export const DEFAULT_PDF_PROVIDER_POLICY_RELATIVE_PATH = path.join(".open-office-artifact-tool", "pdf-providers.json");
export const DEFAULT_OCR_LANGUAGES = Object.freeze(["eng", "chi_sim"]);

const MAX_POLICY_BYTES = 256 * 1024;
const INSTALL_POLICIES = new Set(["disabled", "system-only", "managed"]);

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function normalizeStringArray(value, label, fallback = []) {
  if (value === undefined) return [...fallback];
  if (!Array.isArray(value) || value.some((entry) => !nonEmptyString(entry))) throw new TypeError(`${label} must be an array of non-empty strings.`);
  return [...new Set(value.map((entry) => entry.trim()))];
}

function normalizeBudget(value, label) {
  if (value === undefined) return 0;
  if (!Number.isSafeInteger(value) || value < 0) throw new TypeError(`${label} must be a non-negative safe integer.`);
  return value;
}

function normalizeEnterpriseMirror(value) {
  if (value === undefined) return undefined;
  if (!nonEmptyString(value)) throw new TypeError("policy.enterpriseMirror must be a non-empty HTTPS URL.");
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw new TypeError("policy.enterpriseMirror must be a valid HTTPS URL.");
  }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password || /(^|[/-])latest([/-]|$)/i.test(parsed.pathname)) {
    throw new TypeError("policy.enterpriseMirror must be a versioned HTTPS base URL without credentials.");
  }
  return parsed.href.endsWith("/") ? parsed.href : `${parsed.href}/`;
}

/** Normalize an inline or JSON policy to its public, serializable form. */
export function normalizePdfCapabilityPolicy(value = {}) {
  if (!isPlainObject(value)) throw new TypeError("PDF capability policy must be an object.");
  const legacyMode = value.mode === "offline" ? "disabled" : value.mode;
  if (legacyMode !== undefined && value.installPolicy !== undefined && legacyMode !== value.installPolicy) {
    throw new TypeError("policy.mode and policy.installPolicy must agree when both are supplied.");
  }
  const requestedInstallPolicy = value.installPolicy ?? legacyMode ?? PDF_PROVIDER_CATALOG.releasePolicy.defaultInstallPolicy;
  const installPolicy = requestedInstallPolicy === "offline" ? "disabled" : requestedInstallPolicy;
  if (!INSTALL_POLICIES.has(installPolicy)) throw new TypeError("policy.installPolicy must be disabled, system-only, or managed.");
  if (value.providerPython !== undefined && !nonEmptyString(value.providerPython)) throw new TypeError("policy.providerPython must be a non-empty path when supplied.");
  return {
    installPolicy,
    allowedProviders: normalizeStringArray(value.allowedProviders, "policy.allowedProviders"),
    allowedPacks: normalizeStringArray(value.allowedPacks, "policy.allowedPacks"),
    acceptedLicenses: normalizeStringArray(value.acceptedLicenses, "policy.acceptedLicenses"),
    allowedOcrLanguages: normalizeStringArray(value.allowedOcrLanguages, "policy.allowedOcrLanguages", DEFAULT_OCR_LANGUAGES),
    maxDownloadBytes: normalizeBudget(value.maxDownloadBytes, "policy.maxDownloadBytes"),
    maxUnpackedBytes: normalizeBudget(value.maxUnpackedBytes, "policy.maxUnpackedBytes"),
    enterpriseMirror: normalizeEnterpriseMirror(value.enterpriseMirror),
    providerPython: value.providerPython?.trim(),
  };
}

async function lstatRegularFile(filePath) {
  const container = await fs.promises.lstat(path.dirname(filePath));
  if (!container.isDirectory() || container.isSymbolicLink()) throw new Error("policyPath parent must be a real non-symlink directory.");
  const stat = await fs.promises.lstat(filePath);
  if (!stat.isFile() || stat.isSymbolicLink()) throw new Error("policyPath must be a regular non-symlink file.");
  if (stat.size > MAX_POLICY_BYTES) throw new Error(`policyPath exceeds ${MAX_POLICY_BYTES} bytes.`);
  return stat;
}

/** Read an explicitly requested bounded policy file. It never creates one. */
export async function loadPdfCapabilityPolicy(policyPath) {
  if (!nonEmptyString(policyPath)) throw new TypeError("policyPath must be a non-empty file path.");
  const absolute = path.resolve(policyPath);
  await lstatRegularFile(absolute);
  let parsed;
  try {
    parsed = JSON.parse(await fs.promises.readFile(absolute, "utf8"));
  } catch (error) {
    throw new Error(`policyPath must contain valid JSON: ${error.message}`);
  }
  return normalizePdfCapabilityPolicy(parsed);
}

function fingerprint(policy) {
  return crypto.createHash("sha256").update(JSON.stringify(policy)).digest("hex");
}

/**
 * Resolve either a supplied policy object, an explicit policyPath, or the
 * conventional project file. A missing conventional file intentionally means
 * disabled: resolution remains read-only and does not write boilerplate.
 */
export async function resolvePdfCapabilityPolicy(options = {}) {
  if (!isPlainObject(options)) throw new TypeError("provider options must be an object.");
  if (options.policy !== undefined && options.policyPath !== undefined) throw new TypeError("Use either policy or policyPath, not both.");
  if (options.policy !== undefined) {
    const policy = normalizePdfCapabilityPolicy(options.policy);
    return { policy, policyPath: undefined, source: "inline", fingerprint: fingerprint(policy), cacheRoot: undefined };
  }
  const selectedPath = options.policyPath === undefined
    ? path.resolve(DEFAULT_PDF_PROVIDER_POLICY_RELATIVE_PATH)
    : path.resolve(options.policyPath);
  try {
    await fs.promises.access(selectedPath, fs.constants.F_OK);
  } catch {
    if (options.policyPath !== undefined) throw new Error(`policyPath does not exist: ${selectedPath}`);
    const policy = normalizePdfCapabilityPolicy({});
    return {
      policy,
      policyPath: selectedPath,
      source: "default-missing",
      fingerprint: fingerprint(policy),
      cacheRoot: path.join(path.dirname(selectedPath), "providers"),
    };
  }
  const policy = await loadPdfCapabilityPolicy(selectedPath);
  return {
    policy,
    policyPath: selectedPath,
    source: options.policyPath === undefined ? "default-file" : "explicit-file",
    fingerprint: fingerprint(policy),
    cacheRoot: path.join(path.dirname(selectedPath), "providers"),
  };
}

export function publicPolicyContext(context) {
  return {
    policy: clonePdfProviderValue(context.policy),
    policyPath: context.policyPath,
    policySource: context.source,
    policyFingerprint: context.fingerprint,
    cacheRoot: context.cacheRoot,
  };
}
