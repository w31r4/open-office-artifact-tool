/**
 * Explicit PDF capability-pack API.
 *
 * Importing this subpath reads only catalog/policy code. It never imports
 * MuPDF, starts WASM, downloads a file, or writes the project cache. Network
 * and cache mutation can happen only through `PdfProviders.ensure()` after a
 * caller has resolved an explicit managed policy.
 */

import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { createRequire } from "node:module";
import { promisify } from "node:util";

import {
  PDF_PROVIDER_CATALOG,
  PDF_PROVIDER_CATALOG_SCHEMA,
  PDF_PROVIDER_CATALOG_SHA256,
  clonePdfProviderValue,
  currentPdfProviderPlatform,
  pdfPackById,
  pdfProviderById,
  pdfTaskById,
  validatePdfProviderCatalog,
} from "./catalog.mjs";
import {
  DEFAULT_OCR_LANGUAGES,
  DEFAULT_PDF_PROVIDER_POLICY_RELATIVE_PATH,
  loadPdfCapabilityPolicy,
  normalizePdfCapabilityPolicy,
  publicPolicyContext,
  resolvePdfCapabilityPolicy,
} from "./policy.mjs";
import { ensureManagedPacks, probeManagedProviderRuntime } from "./installer.mjs";

const execFile = promisify(execFileCallback);
const require = createRequire(import.meta.url);
const VERSION_PATTERN = /(?<!\d)(\d+)(?:\.(\d+))?(?:\.(\d+))?/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/i;

function isPlainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function normalizeStringArray(value, label) {
  if (value === undefined) return [];
  if (!Array.isArray(value) || value.some((entry) => !nonEmptyString(entry))) throw new TypeError(`${label} must be an array of non-empty strings.`);
  return [...new Set(value.map((entry) => entry.trim()))];
}

function inspectionEvidence(value) {
  if (!isPlainObject(value)) return undefined;
  const sourceSha256 = value.summary?.sourceSha256 ?? value.sourceSha256;
  if (!nonEmptyString(sourceSha256) || !SHA256_PATTERN.test(sourceSha256)) return undefined;
  return { sourceSha256: sourceSha256.toLowerCase() };
}

function reason(code, message, details = undefined) {
  return { code, message, ...(details === undefined ? {} : { details }) };
}

function versionParts(value) {
  const match = String(value || "").match(VERSION_PATTERN);
  return match ? [Number(match[1]), Number(match[2] || 0), Number(match[3] || 0)] : undefined;
}

function compareVersions(left, right) {
  const a = Array.isArray(left) ? left : versionParts(left);
  const b = Array.isArray(right) ? right : versionParts(right);
  if (!a || !b) return undefined;
  for (let index = 0; index < 3; index += 1) {
    if (a[index] !== b[index]) return a[index] < b[index] ? -1 : 1;
  }
  return 0;
}

function versionInRange(version, provider) {
  const parsed = versionParts(version);
  if (!parsed) return false;
  if (provider.exactVersion) return compareVersions(parsed, provider.exactVersion) === 0;
  if (provider.minimumVersion && compareVersions(parsed, provider.minimumVersion) < 0) return false;
  if (provider.maximumVersionExclusive && compareVersions(parsed, provider.maximumVersionExclusive) >= 0) return false;
  return true;
}

function providerLicenseAccepted(provider, policy) {
  if (!provider.license.requiresAcknowledgement) return true;
  const accepted = new Set(policy.acceptedLicenses);
  return Boolean((provider.license.id && accepted.has(provider.license.id)) || provider.license.acceptedValues?.some((entry) => accepted.has(entry)));
}

function packLicenseAccepted(pack, policy) {
  if (!pack.license.requiresAcknowledgement) return true;
  const accepted = new Set(policy.acceptedLicenses);
  return Boolean((pack.license.id && accepted.has(pack.license.id)) || pack.license.acceptedValues?.some((entry) => accepted.has(entry)));
}

function expandPackIds(packIds) {
  const result = [];
  const visiting = new Set();
  const visited = new Set();
  const visit = (packId) => {
    if (visited.has(packId)) return;
    if (visiting.has(packId)) throw new Error(`PDF capability-pack dependency cycle at ${packId}.`);
    visiting.add(packId);
    const pack = pdfPackById(packId);
    for (const dependency of pack.requiresPackIds) visit(dependency);
    visiting.delete(packId);
    visited.add(packId);
    result.push(packId);
  };
  for (const packId of packIds) visit(packId);
  return result;
}

function packIdsFor(provider, task, languages) {
  const ids = [provider.packId];
  const missingLanguages = [];
  if (task.ocrLanguages) {
    for (const language of languages) {
      const packId = PDF_PROVIDER_CATALOG.ocrLanguagePacks[language];
      if (!packId) missingLanguages.push(language);
      else ids.push(packId);
    }
  }
  return { packIds: expandPackIds(ids), missingLanguages };
}

function packArtifactForPlatform(pack, platform) {
  return pack.artifacts.find((artifact) => artifact.platform === platform);
}

function runtimeRequirement(provider, packIds, taskId = undefined) {
  return {
    providerPackId: provider.packId,
    requiredPackIds: [...packIds],
    managedRuntime: provider.managedRuntime ? clonePdfProviderValue(provider.managedRuntime) : null,
    systemOnlyEnvironment: provider.environment || null,
    minimumMajor: provider.minimumMajor ?? null,
    minimumVersion: provider.minimumVersion ?? null,
    taskMinimumVersion: taskId ? provider.taskMinimumVersions?.[taskId] ?? null : null,
  };
}

function policyAllowsProvider(providerId, provider, packIds, policy) {
  const builtin = pdfPackById(provider.packId).state === "built-in";
  if (builtin) return true;
  if (!policy.allowedProviders.includes(providerId)) return false;
  if (policy.installPolicy === "managed") return packIds.every((packId) => policy.allowedPacks.includes(packId));
  return policy.installPolicy === "system-only";
}

function buildInstallPlan(providerId, provider, task, policy, requestedLanguages, taskId = undefined) {
  const platform = currentPdfProviderPlatform();
  const { packIds, missingLanguages } = packIdsFor(provider, task, requestedLanguages);
  const packs = packIds.map((packId) => {
    const pack = pdfPackById(packId);
    const artifact = pack.state === "published" ? packArtifactForPlatform(pack, platform) : undefined;
    return {
      packId,
      state: pack.state,
      delivery: pack.delivery,
      version: pack.version,
      platform,
      dependencyClosure: [...pack.dependencyClosure],
      estimatedDownloadBytes: pack.estimatedDownloadBytes,
      estimatedUnpackedBytes: pack.estimatedUnpackedBytes,
      entrypoints: clonePdfProviderValue(pack.entrypoints),
      artifact: artifact ? clonePdfProviderValue(artifact) : null,
      license: clonePdfProviderValue(pack.license),
    };
  });
  const external = packs.filter((pack) => pack.state !== "built-in");
  const allBuiltIn = external.length === 0;
  const unpublished = external.find((pack) => pack.state === "unpublished");
  const noPlatformArtifact = external.find((pack) => pack.state === "published" && !pack.artifact);
  const downloadBytes = external.reduce((total, pack) => total + (pack.artifact?.downloadBytes || 0), 0);
  const unpackedBytes = external.reduce((total, pack) => total + (pack.artifact?.unpackedBytes || 0), 0);
  let action = "install";
  let availabilityReason = "managed-artifacts-ready";
  if (allBuiltIn) {
    action = "none";
    availabilityReason = "built-in";
  } else if (missingLanguages.length) {
    action = "unavailable";
    availabilityReason = "ocr-language-pack-unpublished";
  } else if (unpublished) {
    action = "unavailable";
    availabilityReason = "managed-artifact-unpublished";
  } else if (noPlatformArtifact) {
    action = "unavailable";
    availabilityReason = "platform-artifact-unavailable";
  } else if (downloadBytes > policy.maxDownloadBytes) {
    action = "unavailable";
    availabilityReason = "download-budget-exceeded";
  } else if (unpackedBytes > policy.maxUnpackedBytes) {
    action = "unavailable";
    availabilityReason = "unpacked-budget-exceeded";
  }
  return {
    providerId,
    platform,
    packIds,
    packs,
    action,
    reason: availabilityReason,
    performsDownload: action === "install",
    downloadBytes,
    unpackedBytes,
    missingLanguages,
    runtime: runtimeRequirement(provider, packIds, taskId),
    requiresExplicitInstallPolicy: !allBuiltIn,
  };
}

function executableFromPath(command, environmentName, environment = process.env) {
  const configured = environmentName ? String(environment[environmentName] || "").trim() : "";
  const candidates = configured ? [configured] : String(environment.PATH || "").split(path.delimiter).filter(Boolean).flatMap((directory) => {
    const extensions = process.platform === "win32" ? (environment.PATHEXT || ".EXE;.CMD;.BAT;.COM").split(";").filter(Boolean) : [""];
    return extensions.map((extension) => path.join(directory, `${command}${extension}`));
  });
  for (const candidate of candidates) {
    try {
      const stat = fs.statSync(candidate);
      if (stat.isFile() && (process.platform === "win32" || (stat.mode & 0o111) !== 0)) return path.resolve(candidate);
    } catch { /* probe failure remains explicit */ }
  }
  return undefined;
}

async function commandVersion(executable, requireVersionOutput) {
  for (const args of [["-v"], ["--version"], ["-version"]]) {
    try {
      const { stdout, stderr } = await execFile(executable, args, { timeout: 2_500, maxBuffer: 16 * 1024, windowsHide: true });
      const line = String(stdout || stderr || "").trim().split(/\r?\n/, 1)[0];
      if (line && !/couldn't open file/i.test(line)) return line.slice(0, 300);
    } catch (error) {
      const line = String(error?.stdout || error?.stderr || "").trim().split(/\r?\n/, 1)[0];
      if (line && !/couldn't open file/i.test(line)) return line.slice(0, 300);
    }
  }
  return requireVersionOutput ? undefined : executable;
}

async function probeCommand(provider, commandPaths = undefined, taskId = undefined) {
  const commands = {};
  for (const command of provider.commands) {
    const executable = commandPaths?.[command] || executableFromPath(command, provider.environment);
    const version = executable ? await commandVersion(executable, provider.requireVersionOutput) : undefined;
    commands[command] = { executable, version };
  }
  const entries = Object.values(commands);
  let available = entries.length > 0 && entries.every((entry) => entry.executable && entry.version);
  if (available && provider.minimumMajor !== undefined) available = entries.every((entry) => (versionParts(entry.version)?.[0] || 0) >= provider.minimumMajor);
  const taskMinimumVersion = taskId ? provider.taskMinimumVersions?.[taskId] : undefined;
  const versionRequirement = taskMinimumVersion ? { ...provider, minimumVersion: taskMinimumVersion } : provider;
  if (available && (versionRequirement.minimumVersion || versionRequirement.maximumVersionExclusive)) {
    available = entries.every((entry) => versionInRange(entry.version, versionRequirement));
  }
  return { available, evidence: { commands, ...(taskMinimumVersion ? { taskMinimumVersion } : {}) } };
}

function pythonExecutable(policy, explicitPath = undefined) {
  if (explicitPath) return executableFromPath(explicitPath, undefined, { ...process.env, PATH: "" });
  if (policy.providerPython) return executableFromPath(policy.providerPython, undefined, { ...process.env, PATH: "" });
  const configured = String(process.env.OPEN_OFFICE_PDF_PROVIDER_PYTHON || "").trim();
  if (configured) return executableFromPath(configured, undefined, { ...process.env, PATH: "" });
  return executableFromPath("python3");
}

async function probePythonModule(provider, policy, explicitPython = undefined) {
  const executable = pythonExecutable(policy, explicitPython);
  if (!executable) return { available: false, evidence: { python: undefined, module: provider.module } };
  const payload = {
    module: provider.module,
    distribution: provider.distribution || provider.module,
    companionModule: provider.companionModule,
    companionDistribution: provider.companionDistribution || provider.companionModule,
  };
  const program = [
    "import importlib.metadata as m, importlib.util as u, json, sys",
    "p=json.loads(sys.argv[1])",
    "def v(n):\n try: return m.version(n)\n except Exception: return None",
    "print(json.dumps({'moduleFound': u.find_spec(p['module']) is not None, 'version': v(p['distribution']), 'companionFound': (not p.get('companionModule')) or u.find_spec(p['companionModule']) is not None, 'companionVersion': v(p['companionDistribution']) if p.get('companionModule') else None}))",
  ].join("; ");
  try {
    const { stdout } = await execFile(executable, ["-c", program, JSON.stringify(payload)], { timeout: 3_000, maxBuffer: 16 * 1024, windowsHide: true });
    const probe = JSON.parse(stdout);
    let available = probe.moduleFound === true && nonEmptyString(probe.version) && versionInRange(probe.version, provider);
    if (provider.companionModule) {
      available = available && probe.companionFound === true && nonEmptyString(probe.companionVersion) && versionInRange(probe.companionVersion, {
        minimumVersion: provider.companionMinimumVersion,
        maximumVersionExclusive: provider.companionMaximumVersionExclusive,
      });
    }
    return { available, evidence: { python: executable, ...probe } };
  } catch (error) {
    return { available: false, evidence: { python: executable, module: provider.module, error: String(error?.stderr || error?.message || error).slice(0, 300) } };
  }
}

function probeNodePackage(provider) {
  try {
    let cursor = path.dirname(require.resolve(provider.package));
    let packagePath;
    while (true) {
      const candidate = path.join(cursor, "package.json");
      try {
        const metadata = JSON.parse(fs.readFileSync(candidate, "utf8"));
        if (metadata.name === provider.package) {
          packagePath = candidate;
          break;
        }
      } catch { /* locate metadata without importing the package entrypoint */ }
      const parent = path.dirname(cursor);
      if (parent === cursor) break;
      cursor = parent;
    }
    if (!packagePath) throw new Error(`could not locate ${provider.package} package metadata`);
    const metadata = JSON.parse(fs.readFileSync(packagePath, "utf8"));
    return { available: metadata.version === provider.exactVersion, evidence: { package: provider.package, packagePath, version: metadata.version, expectedVersion: provider.exactVersion } };
  } catch (error) {
    return { available: false, evidence: { package: provider.package, expectedVersion: provider.exactVersion, error: String(error?.message || error).slice(0, 300) } };
  }
}

async function probeLocalRuntime(provider, policy, runtime = undefined, taskId = undefined) {
  if (provider.kind === "core") return { available: true, evidence: { runtime: "open-office-artifact-tool" } };
  if (provider.kind === "node-package") return probeNodePackage(provider);
  if (provider.kind === "command") return probeCommand(provider, runtime?.commandPaths, taskId);
  if (provider.kind === "python-module") return probePythonModule(provider, policy, runtime?.pythonPath);
  return { available: false, evidence: { reason: `unsupported probe kind ${provider.kind}` } };
}

function providerResultBase(providerId, provider, policyContext, installPlan) {
  return {
    schema: PDF_PROVIDER_CATALOG_SCHEMA,
    schemaVersion: PDF_PROVIDER_CATALOG.schemaVersion,
    catalogSha256: PDF_PROVIDER_CATALOG_SHA256,
    provider: {
      id: providerId,
      packId: provider.packId,
      role: provider.role,
      integration: provider.integration,
      kind: provider.kind,
      license: clonePdfProviderValue(provider.license),
      matrixBoundary: provider.matrixBoundary,
    },
    ...publicPolicyContext(policyContext),
    installPlan,
    silentFallback: false,
  };
}

function normalizeProbeRequest(providerOrRequest, options) {
  if (typeof providerOrRequest === "string") return { provider: providerOrRequest, ...(options || {}) };
  if (!isPlainObject(providerOrRequest)) throw new TypeError("PDF provider probe requires a provider id or request object.");
  return providerOrRequest;
}

/**
 * Probe exactly one selected provider. It does not try a different provider,
 * download a pack, alter the cache, or import MuPDF.
 */
export async function probePdfProvider(providerOrRequest, options = undefined) {
  const request = normalizeProbeRequest(providerOrRequest, options);
  const providerId = request.provider;
  const provider = pdfProviderById(providerId);
  const policyContext = await resolvePdfCapabilityPolicy(request);
  const task = request.task ? pdfTaskById(request.task) : { ocrLanguages: false };
  const languages = normalizeStringArray(request.languages ?? request.ocrLanguages, "languages");
  const installPlan = buildInstallPlan(providerId, provider, task, policyContext.policy, languages, request.task);
  const base = providerResultBase(providerId, provider, policyContext, installPlan);
  const packIds = installPlan.packIds;
  const allBuiltin = installPlan.packs.every((pack) => pack.state === "built-in");
  if (task.ocrLanguages && !languages.length) {
    return { ...base, status: "blocked", reason: reason("ocr-language-required", "OCR probe requires one or more explicit language identifiers."), runtime: null };
  }
  if (task.ocrLanguages) {
    const disallowed = languages.filter((language) => !policyContext.policy.allowedOcrLanguages.includes(language));
    if (disallowed.length) {
      return { ...base, status: "blocked", reason: reason("ocr-language-not-allowed", "Policy does not allow every requested OCR language.", { languages: disallowed }), runtime: null };
    }
    if (installPlan.missingLanguages.length) {
      return { ...base, status: "blocked", reason: reason("ocr-language-pack-unpublished", "The catalog has no immutable language pack for every requested OCR language.", { languages: installPlan.missingLanguages }), runtime: null };
    }
  }
  if (!policyAllowsProvider(providerId, provider, packIds, policyContext.policy)) {
    return { ...base, status: "blocked", reason: reason("provider-or-pack-not-allowed", "The policy does not whitelist this provider and every required capability pack."), runtime: null };
  }
  if (!providerLicenseAccepted(provider, policyContext.policy)) {
    return { ...base, status: "blocked", reason: reason("provider-license-acknowledgement-required", "The provider's license acknowledgement is missing."), runtime: null };
  }
  const unacceptedPack = packIds.map((packId) => pdfPackById(packId)).find((pack) => !packLicenseAccepted(pack, policyContext.policy));
  if (unacceptedPack) {
    return { ...base, status: "blocked", reason: reason("pack-license-acknowledgement-required", "A required capability-pack license acknowledgement is missing.", { packId: packIds.find((packId) => pdfPackById(packId) === unacceptedPack) }), runtime: null };
  }
  if (allBuiltin) {
    const runtime = await probeLocalRuntime(provider, policyContext.policy, undefined, request.task);
    return runtime.available
      ? { ...base, status: "ready", reason: reason("built-in-ready", "The required built-in runtime is present."), runtime }
      : { ...base, status: "blocked", reason: reason("built-in-runtime-unavailable", "The required bundled runtime is not resolvable."), runtime };
  }
  if (policyContext.policy.installPolicy === "disabled") {
    return { ...base, status: "blocked", reason: reason("install-policy-disabled", "The default disabled policy does not activate external providers."), runtime: null };
  }
  if (policyContext.policy.installPolicy === "system-only") {
    const runtime = await probeLocalRuntime(provider, policyContext.policy, undefined, request.task);
    return runtime.available
      ? { ...base, status: "ready", reason: reason("system-provider-ready", "The explicitly selected system provider meets its pinned range."), runtime }
      : { ...base, status: "blocked", reason: reason("system-provider-unavailable", "The selected system provider is unavailable or outside its pinned version range."), runtime };
  }
  if (installPlan.action !== "install") {
    return { ...base, status: "blocked", reason: reason(installPlan.reason, "No immutable managed artifact is currently available within policy budgets."), runtime: null };
  }
  const managed = await probeManagedProviderRuntime({ providerId, packIds, policyContext, languages });
  if (!managed.ready) {
    return { ...base, status: "installable", reason: reason("managed-install-required", "Pinned artifacts are allowed but are not yet installed in the private project cache.", { missingPackId: managed.packId, cacheReason: managed.reason }), runtime: null };
  }
  const runtime = await probeLocalRuntime(provider, policyContext.policy, managed.runtime, request.task);
  return runtime.available
    ? { ...base, status: "ready", reason: reason("managed-provider-ready", "A verified private managed runtime meets its pinned range."), runtime: { ...runtime, managed: managed.runtime, packs: managed.packIds } }
    : { ...base, status: "blocked", reason: reason("managed-runtime-invalid", "The verified pack cache does not provide a usable provider runtime."), runtime: { ...runtime, managed: managed.runtime, packs: managed.packIds } };
}

function taskResultBase(taskId, task, policyContext, providerId, savePolicy, inspection) {
  const evidence = inspectionEvidence(inspection);
  return {
    schema: PDF_PROVIDER_CATALOG_SCHEMA,
    schemaVersion: PDF_PROVIDER_CATALOG.schemaVersion,
    catalogSha256: PDF_PROVIDER_CATALOG_SHA256,
    task: taskId,
    inputMode: task.input,
    savePolicy,
    providerId,
    candidateProviderIds: [...task.providers],
    inspectionProvided: inspection !== undefined,
    inspectionEvidence: evidence || null,
    mutation: Boolean(task.mutation),
    ...publicPolicyContext(policyContext),
    silentFallback: false,
  };
}

/**
 * Resolve an Agent's explicit intent to one provider. A catalog default is a
 * declared preference, never an automatic alternate route when it fails.
 */
export async function resolvePdfCapability(request = {}) {
  if (!isPlainObject(request)) throw new TypeError("PDF capability request must be an object.");
  const task = pdfTaskById(request.task);
  const policyContext = await resolvePdfCapabilityPolicy(request);
  const providerId = request.provider || task.defaultProvider || (task.providers.length === 1 ? task.providers[0] : undefined);
  const base = taskResultBase(request.task, task, policyContext, providerId, request.savePolicy, request.inspection);
  if (!providerId) return { ...base, status: "blocked", reason: reason("provider-selection-required", "This task has multiple providers; select one explicitly without fallback."), installPlan: null, consents: {} };
  if (!task.providers.includes(providerId)) return { ...base, status: "blocked", reason: reason("provider-not-supported-for-task", "The selected provider cannot perform this task."), installPlan: null, consents: {} };
  if (!task.strategies.includes(request.savePolicy)) return { ...base, status: "blocked", reason: reason("save-policy-required", `Choose one of: ${task.strategies.join(", ")}.`), installPlan: null, consents: {} };
  if (task.input === "existing" && request.task !== "inspect" && !base.inspectionEvidence) {
    return {
      ...base,
      status: "blocked",
      reason: reason("inspection-required", "Pass exact-source inspection or preflight evidence containing sourceSha256 before resolving this existing-PDF task."),
      installPlan: null,
      consents: {},
    };
  }

  const provider = pdfProviderById(providerId);
  const requestedLanguages = normalizeStringArray(request.languages ?? request.ocrLanguages, "ocrLanguages");
  const declaredCredentials = normalizeStringArray(request.credentials, "credentials");
  const requiredCredentials = task.credentials || [];
  const installPlan = buildInstallPlan(providerId, provider, task, policyContext.policy, requestedLanguages, request.task);
  const consents = {
    mutation: { required: Boolean(task.mutation), authorized: request.mutationAuthorized === true },
    invalidateSignatures: { required: Boolean(task.invalidateSignatures), authorized: request.invalidateSignaturesAuthorized === true },
    credentials: { required: [...requiredCredentials], declared: declaredCredentials, automaticAcquisition: false },
    ocrLanguages: { requested: requestedLanguages, allowed: [...policyContext.policy.allowedOcrLanguages], automaticAcquisition: false },
  };
  if (task.mutation && request.mutationAuthorized !== true) return { ...base, status: "blocked", reason: reason("mutation-authorization-required", "Destructive PDF mutation requires explicit caller authorization."), installPlan, consents };
  if (task.invalidateSignatures && request.invalidateSignaturesAuthorized !== true) return { ...base, status: "blocked", reason: reason("signature-invalidation-authorization-required", "This task can invalidate signatures and requires explicit authorization."), installPlan, consents };
  if (requiredCredentials.some((credential) => !declaredCredentials.includes(credential))) return { ...base, status: "blocked", reason: reason("credential-declaration-required", "Credentials must be declared by the caller and are never acquired automatically."), installPlan, consents };
  if (task.ocrLanguages) {
    if (!requestedLanguages.length) return { ...base, status: "blocked", reason: reason("ocr-language-required", "OCR requires one or more explicit language identifiers."), installPlan, consents };
    const disallowed = requestedLanguages.filter((language) => !policyContext.policy.allowedOcrLanguages.includes(language));
    if (disallowed.length) return { ...base, status: "blocked", reason: reason("ocr-language-not-allowed", "Policy does not allow every requested OCR language.", { languages: disallowed }), installPlan, consents };
    if (installPlan.missingLanguages.length) return { ...base, status: "blocked", reason: reason("ocr-language-pack-unpublished", "The catalog has no immutable language pack for every requested OCR language.", { languages: installPlan.missingLanguages }), installPlan, consents };
  }
  const probePolicy = request.policyPath !== undefined
    ? { policyPath: request.policyPath }
    : request.policy !== undefined
      ? { policy: request.policy }
      : {};
  const probe = await probePdfProvider({ provider: providerId, task: request.task, languages: requestedLanguages, ...probePolicy });
  return {
    ...base,
    status: probe.status,
    reason: probe.reason,
    provider: probe.provider,
    installPlan,
    probe,
    consents,
  };
}

/**
 * Install only an `installable` resolution under the same project policy that
 * created it. This function never chooses an alternative provider or obtains
 * credentials; it returns a fresh `ready` probe after receipt verification.
 */
export async function ensurePdfCapability({ resolution, policyPath } = {}) {
  if (!isPlainObject(resolution)) throw new TypeError("ensure requires a resolution object.");
  if (resolution.status !== "installable") throw new Error(`ensure requires an installable resolution, received ${resolution.status || "unknown"}.`);
  if (resolution.catalogSha256 !== PDF_PROVIDER_CATALOG_SHA256) throw new Error("Resolution catalog digest does not match this package.");
  if (!nonEmptyString(resolution.providerId) || !Array.isArray(resolution.installPlan?.packIds)) throw new Error("Resolution lacks a selected provider or pack plan.");
  const policyContext = await resolvePdfCapabilityPolicy({ policyPath });
  if (policyContext.fingerprint !== resolution.policyFingerprint) throw new Error("Policy changed after resolution; resolve again before installing.");
  const installation = await ensureManagedPacks({ providerId: resolution.providerId, packIds: resolution.installPlan.packIds, policyContext });
  const probe = await probePdfProvider({
    provider: resolution.providerId,
    task: resolution.task,
    languages: resolution.consents?.ocrLanguages?.requested,
    policyPath,
  });
  return { ...probe, installation };
}

export function listPdfProviders() {
  return Object.entries(PDF_PROVIDER_CATALOG.providers).map(([id, provider]) => ({
    id,
    packId: provider.packId,
    role: provider.role,
    integration: provider.integration,
    taskIds: [...provider.taskIds],
    license: clonePdfProviderValue(provider.license),
    pack: clonePdfProviderValue(pdfPackById(provider.packId)),
  }));
}

export function listPdfCapabilityTasks() {
  return Object.entries(PDF_PROVIDER_CATALOG.tasks).map(([id, task]) => ({ id, ...clonePdfProviderValue(task) }));
}

/** The compact public API promised by the capability-pack workflow. */
export const PdfProviders = Object.freeze({
  resolve: resolvePdfCapability,
  ensure: ensurePdfCapability,
  probe: probePdfProvider,
});

export {
  DEFAULT_OCR_LANGUAGES,
  DEFAULT_PDF_PROVIDER_POLICY_RELATIVE_PATH,
  loadPdfCapabilityPolicy,
  normalizePdfCapabilityPolicy,
  validatePdfProviderCatalog,
};

export { PDF_PROVIDER_CATALOG, PDF_PROVIDER_CATALOG_SCHEMA, PDF_PROVIDER_CATALOG_SHA256 } from "./catalog.mjs";
