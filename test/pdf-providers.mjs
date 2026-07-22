import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { gzipSync } from "node:zlib";

import {
  PDF_PROVIDER_CATALOG,
  PdfProviders,
  resolvePdfCapability,
  validatePdfProviderCatalog,
} from "../src/pdf/providers/index.mjs";
import {
  PDF_PROVIDER_RECEIPT_SCHEMA,
  installManagedPackForTest,
  safeExtractTarGz,
} from "../src/pdf/providers/installer.mjs";

const repoRoot = path.resolve(import.meta.dirname, "..");
const platform = `${process.platform}-${process.arch}`;
const inspectedPdf = Object.freeze({ summary: { sourceSha256: "a".repeat(64) } });

function writeString(buffer, offset, length, value) {
  Buffer.from(value, "utf8").copy(buffer, offset, 0, Math.min(Buffer.byteLength(value), length));
}

function writeOctal(buffer, offset, length, value) {
  const text = Number(value).toString(8).padStart(length - 1, "0");
  writeString(buffer, offset, length - 1, text);
  buffer[offset + length - 1] = 0;
}

function tarHeader({ name, bytes, type = "0", mode = 0o755 }) {
  const header = Buffer.alloc(512);
  writeString(header, 0, 100, name);
  writeOctal(header, 100, 8, mode);
  writeOctal(header, 108, 8, 0);
  writeOctal(header, 116, 8, 0);
  writeOctal(header, 124, 12, bytes.length);
  writeOctal(header, 136, 12, 0);
  header.fill(0x20, 148, 156);
  header[156] = type.charCodeAt(0);
  writeString(header, 257, 6, "ustar");
  writeString(header, 263, 2, "00");
  let checksum = 0;
  for (const value of header) checksum += value;
  writeString(header, 148, 6, checksum.toString(8).padStart(6, "0"));
  header[154] = 0;
  header[155] = 0x20;
  return header;
}

function tarGz(entries) {
  const records = [];
  for (const entry of entries) {
    const bytes = Buffer.from(entry.bytes || "", "utf8");
    records.push(tarHeader({ ...entry, bytes }), bytes, Buffer.alloc((512 - (bytes.length % 512)) % 512));
  }
  records.push(Buffer.alloc(1024));
  return gzipSync(Buffer.concat(records));
}

function fixturePack(archive, overrides = {}) {
  const digest = crypto.createHash("sha256").update(archive).digest("hex");
  const artifact = {
    platform,
    asset: "fixture.tar.gz",
    version: "1.2.3",
    url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/fixture.tar.gz",
    sha256: digest,
    downloadBytes: archive.length,
    unpackedBytes: 16 * 1024,
    archiveFormat: "tar.gz",
    ...(overrides.artifact || {}),
  };
  return {
    state: "published",
    version: "1.2.3",
    platforms: [platform],
    artifacts: [artifact],
    entrypoints: [{ path: "bin/tool", kind: "file", executable: true }],
    ...(overrides.pack || {}),
  };
}

function fakeFetch(bytes, calls = undefined) {
  return async (url) => {
    calls?.push(String(url));
    return new Response(bytes, { status: 200 });
  };
}

async function listTree(root) {
  const entries = [];
  async function walk(directory) {
    for (const entry of await fs.readdir(directory, { withFileTypes: true })) {
      entries.push(path.relative(root, path.join(directory, entry.name)));
      if (entry.isDirectory()) await walk(path.join(directory, entry.name));
    }
  }
  await walk(root);
  return entries.sort();
}

assert.equal(PDF_PROVIDER_CATALOG.releasePolicy.defaultInstallPolicy, "disabled");
assert.deepEqual(PDF_PROVIDER_CATALOG.releasePolicy.managedPlatforms, ["darwin-arm64", "linux-x64"]);
assert.equal(PdfProviders.resolve, resolvePdfCapability);
assert.deepEqual(Object.keys(PdfProviders).sort(), ["ensure", "probe", "resolve"]);
assert.equal(PDF_PROVIDER_CATALOG.providers.qpdf.packId, "qpdf");
assert.equal(PDF_PROVIDER_CATALOG.providers.qpdf.taskMinimumVersions.encrypt, "11.7.0");
assert.equal(PDF_PROVIDER_CATALOG.packs.qpdf.state, "published");
assert.equal(PDF_PROVIDER_CATALOG.packs.qpdf.version, "12.3.2-oat.1");
assert.deepEqual(PDF_PROVIDER_CATALOG.packs.qpdf.releaseEvidence.verifiedPlatforms, ["darwin-arm64", "linux-x64"]);
assert.equal(PDF_PROVIDER_CATALOG.packs["python-foundation"].state, "published");
assert.equal(PDF_PROVIDER_CATALOG.packs["python-foundation"].version, "3.13.14-oat.1");
assert.deepEqual(PDF_PROVIDER_CATALOG.packs["python-foundation"].releaseEvidence.verifiedPlatforms, ["darwin-arm64", "linux-x64"]);
assert.equal(PDF_PROVIDER_CATALOG.packs["python-specialists"].state, "published");
assert.equal(PDF_PROVIDER_CATALOG.packs["python-specialists"].version, "3.13.14-oat.1");
assert.deepEqual(PDF_PROVIDER_CATALOG.packs["python-specialists"].releaseEvidence.verifiedPlatforms, ["darwin-arm64", "linux-x64"]);
assert.equal(PDF_PROVIDER_CATALOG.packs.verapdf.state, "published");
assert.equal(PDF_PROVIDER_CATALOG.packs.verapdf.version, "1.30.2-oat.1");
assert.deepEqual(PDF_PROVIDER_CATALOG.packs.verapdf.releaseEvidence.verifiedPlatforms, ["darwin-arm64", "linux-x64"]);
assert.equal(PDF_PROVIDER_CATALOG.packs.verapdf.license.expression, "MPL-2.0 AND GPL-3.0-or-later AND GPL-2.0-only WITH Classpath-exception-2.0");
assert.equal(PDF_PROVIDER_CATALOG.providers.verapdf.probeTimeoutMs, 20_000);
assert.ok(!("managedPack" in PDF_PROVIDER_CATALOG.providers.qpdf), "pack metadata must have one canonical top-level home");

const invalidPlatformCatalog = structuredClone(PDF_PROVIDER_CATALOG);
invalidPlatformCatalog.packs.qpdf.state = "published";
invalidPlatformCatalog.packs.qpdf.version = "1.2.3";
invalidPlatformCatalog.packs.qpdf.artifacts = [{
  platform: "win32-x64",
  asset: "qpdf.tar.gz",
  version: "1.2.3",
  url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf.tar.gz",
  sha256: "a".repeat(64),
  downloadBytes: 1,
  unpackedBytes: 1,
  archiveFormat: "tar.gz",
}];
invalidPlatformCatalog.packs.qpdf.releaseEvidence = {
  sbom: { asset: "qpdf.cdx.json", url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf.cdx.json", sha256: "b".repeat(64) },
  thirdPartyNotices: { asset: "qpdf-notices.txt", url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf-notices.txt", sha256: "c".repeat(64) },
  verifiedPlatforms: ["win32-x64"],
};
assert.throws(() => validatePdfProviderCatalog(invalidPlatformCatalog), /unsupported managed platform|outside the declared managed platforms/);
const unsignedPublishedCatalog = structuredClone(PDF_PROVIDER_CATALOG);
unsignedPublishedCatalog.packs.qpdf.state = "published";
unsignedPublishedCatalog.packs.qpdf.version = "1.2.3";
unsignedPublishedCatalog.packs.qpdf.artifacts = ["darwin-arm64", "linux-x64"].map((candidatePlatform) => ({
  platform: candidatePlatform,
  asset: `qpdf-${candidatePlatform}.tar.gz`,
  version: "1.2.3",
  url: `https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf-${candidatePlatform}.tar.gz`,
  sha256: "a".repeat(64),
  downloadBytes: 1,
  unpackedBytes: 1,
  archiveFormat: "tar.gz",
}));
unsignedPublishedCatalog.packs.qpdf.releaseEvidence = {
  sbom: { asset: "qpdf.cdx.json", url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf.cdx.json", sha256: "b".repeat(64) },
  thirdPartyNotices: { asset: "qpdf-notices.txt", url: "https://releases.example.test/open-office-artifact-tool/v1.2.3/qpdf-notices.txt", sha256: "c".repeat(64) },
  verifiedPlatforms: ["darwin-arm64", "linux-x64"],
};
assert.throws(() => validatePdfProviderCatalog(unsignedPublishedCatalog), /artifact-attestation/);
const ocrRuntimeCatalog = structuredClone(PDF_PROVIDER_CATALOG);
assert.deepEqual(ocrRuntimeCatalog.packs["ocr-core"].requiresPackIds, ["qpdf"], "OCR core bundles its minimal pdftotext sidecar instead of silently selecting the separate Poppler QA route");
assert.equal(validatePdfProviderCatalog(ocrRuntimeCatalog), true, "OCR may reference only declared files in its qpdf dependency closure");
assert.equal(ocrRuntimeCatalog.providers.ocrmypdf.managedRuntime.languageDirectoryEnvironment, "OPEN_OFFICE_PDF_TESSDATA_DIRS");
const escapedOcrRuntimeCatalog = structuredClone(PDF_PROVIDER_CATALOG);
escapedOcrRuntimeCatalog.providers.ocrmypdf.managedRuntime.commandPaths.qpdf = { packId: "python-foundation", path: "bin/python3" };
assert.throws(() => validatePdfProviderCatalog(escapedOcrRuntimeCatalog), /outside its dependency closure/);
const invalidTaskMinimumCatalog = structuredClone(PDF_PROVIDER_CATALOG);
invalidTaskMinimumCatalog.providers.qpdf.taskMinimumVersions = { unknown: "11.7" };
assert.throws(() => validatePdfProviderCatalog(invalidTaskMinimumCatalog), /taskMinimumVersions contains an invalid task or version/);
const invalidProbeTimeoutCatalog = structuredClone(PDF_PROVIDER_CATALOG);
invalidProbeTimeoutCatalog.providers.verapdf.probeTimeoutMs = 60_001;
assert.throws(() => validatePdfProviderCatalog(invalidProbeTimeoutCatalog), /probeTimeoutMs must be an integer/);

const builtIn = await PdfProviders.resolve({ task: "inspect", savePolicy: "read-only", inspection: inspectedPdf });
assert.equal(builtIn.status, "ready");
assert.equal(builtIn.providerId, "mupdf-js");
assert.equal(builtIn.policy.installPolicy, "disabled");
assert.deepEqual(builtIn.policy.allowedOcrLanguages, ["eng", "chi_sim"]);
assert.equal(builtIn.silentFallback, false);

const missingInspection = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  savePolicy: "rewrite",
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
});
assert.equal(missingInspection.status, "blocked");
assert.equal(missingInspection.reason.code, "inspection-required");

const malformedInspection = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: { summary: { sourceSha256: "not-a-hash" } },
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
});
assert.equal(malformedInspection.status, "blocked");
assert.equal(malformedInspection.reason.code, "inspection-required");

const disabledQpdf = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
});
assert.equal(disabledQpdf.status, "blocked");
assert.equal(disabledQpdf.reason.code, "provider-or-pack-not-allowed");

const managedQpdf = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
  policy: {
    installPolicy: "managed",
    allowedProviders: ["qpdf"],
    allowedPacks: ["qpdf"],
    maxDownloadBytes: 100_000_000,
    maxUnpackedBytes: 100_000_000,
  },
});
assert.equal(managedQpdf.status, "installable");
assert.equal(managedQpdf.reason.code, "managed-install-required");
assert.equal(managedQpdf.installPlan.performsDownload, true);
await assert.rejects(() => PdfProviders.ensure({ resolution: managedQpdf }), /Policy changed after resolution/);

const managedReportlab = await PdfProviders.resolve({
  task: "create-layout",
  provider: "reportlab",
  savePolicy: "rewrite",
  policy: {
    installPolicy: "managed",
    allowedProviders: ["reportlab"],
    allowedPacks: ["python-foundation"],
    maxDownloadBytes: 100_000_000,
    maxUnpackedBytes: 200_000_000,
  },
});
assert.equal(managedReportlab.status, "installable");
assert.equal(managedReportlab.reason.code, "managed-install-required");
assert.equal(managedReportlab.installPlan.packIds[0], "python-foundation");
assert.equal(managedReportlab.installPlan.runtime.managedRuntime.pythonPath, "bin/python3");

const managedVeraPdf = await PdfProviders.resolve({
  task: "validate-conformance",
  provider: "verapdf",
  savePolicy: "read-only",
  inspection: inspectedPdf,
  policy: {
    installPolicy: "managed",
    allowedProviders: ["verapdf"],
    allowedPacks: ["verapdf"],
    maxDownloadBytes: 70_000_000,
    maxUnpackedBytes: 200_000_000,
  },
});
assert.equal(managedVeraPdf.status, "installable");
assert.equal(managedVeraPdf.reason.code, "managed-install-required");
assert.deepEqual(managedVeraPdf.installPlan.packIds, ["verapdf"]);
assert.equal(managedVeraPdf.installPlan.runtime.managedRuntime.commandPaths.verapdf, "bin/verapdf");

const unacknowledgedSpecialists = await PdfProviders.resolve({
  task: "inspect",
  provider: "pymupdf",
  savePolicy: "read-only",
  policy: {
    installPolicy: "managed",
    allowedProviders: ["pymupdf"],
    allowedPacks: ["python-specialists", "qpdf"],
    maxDownloadBytes: 128_000_000,
    maxUnpackedBytes: 300_000_000,
  },
});
assert.equal(unacknowledgedSpecialists.status, "blocked");
assert.equal(unacknowledgedSpecialists.reason.code, "provider-license-acknowledgement-required");

const managedPymupdf = await PdfProviders.resolve({
  task: "inspect",
  provider: "pymupdf",
  savePolicy: "read-only",
  policy: {
    installPolicy: "managed",
    allowedProviders: ["pymupdf"],
    allowedPacks: ["python-specialists", "qpdf"],
    acceptedLicenses: ["agpl"],
    maxDownloadBytes: 128_000_000,
    maxUnpackedBytes: 300_000_000,
  },
});
assert.equal(managedPymupdf.status, "installable");
assert.equal(managedPymupdf.reason.code, "managed-install-required");
assert.deepEqual(managedPymupdf.installPlan.packIds, ["qpdf", "python-specialists"]);
assert.equal(managedPymupdf.installPlan.runtime.managedRuntime.pythonPath, "bin/python3");

const encryptWithoutCredentialDeclaration = await PdfProviders.resolve({
  task: "encrypt",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
  policy: {
    installPolicy: "managed",
    allowedProviders: ["qpdf"],
    allowedPacks: ["qpdf"],
    maxDownloadBytes: 100_000_000,
    maxUnpackedBytes: 100_000_000,
  },
});
assert.equal(encryptWithoutCredentialDeclaration.status, "blocked");
assert.equal(encryptWithoutCredentialDeclaration.reason.code, "credential-declaration-required");
assert.deepEqual(encryptWithoutCredentialDeclaration.consents.credentials.required, ["caller-owned-user-and-owner-password-files"]);
const managedEncrypt = await PdfProviders.resolve({
  task: "encrypt",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
  credentials: ["caller-owned-user-and-owner-password-files"],
  policy: {
    installPolicy: "managed",
    allowedProviders: ["qpdf"],
    allowedPacks: ["qpdf"],
    maxDownloadBytes: 100_000_000,
    maxUnpackedBytes: 100_000_000,
  },
});
assert.equal(managedEncrypt.status, "installable");
assert.equal(managedEncrypt.reason.code, "managed-install-required");
assert.equal(managedEncrypt.installPlan.runtime.taskMinimumVersion, "11.7.0");
assert.deepEqual(managedEncrypt.consents.credentials.declared, ["caller-owned-user-and-owner-password-files"]);

const systemOnlyQpdf = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
  policy: {
    installPolicy: "system-only",
    allowedProviders: ["qpdf"],
    allowedPacks: [],
    maxDownloadBytes: 0,
    maxUnpackedBytes: 0,
  },
});
assert.ok(["ready", "blocked"].includes(systemOnlyQpdf.status));
assert.notEqual(systemOnlyQpdf.reason.code, "managed-artifact-unpublished");
assert.equal(systemOnlyQpdf.silentFallback, false);

const missingCredential = await PdfProviders.resolve({
  task: "sign",
  provider: "pyhanko",
  savePolicy: "incremental",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  policy: { installPolicy: "managed", allowedProviders: ["pyhanko"], allowedPacks: ["python-specialists", "qpdf"], acceptedLicenses: ["agpl"], maxDownloadBytes: 1, maxUnpackedBytes: 1 },
});
assert.equal(missingCredential.status, "blocked");
assert.equal(missingCredential.reason.code, "credential-declaration-required");

const ocrLanguagePolicy = await PdfProviders.resolve({
  task: "ocr",
  provider: "ocrmypdf",
  savePolicy: "rewrite",
  inspection: inspectedPdf,
  mutationAuthorized: true,
  ocrLanguages: ["fra"],
  policy: { installPolicy: "managed", allowedProviders: ["ocrmypdf"], allowedPacks: ["ocr-core", "qpdf"], allowedOcrLanguages: ["eng", "chi_sim", "fra"], maxDownloadBytes: 1, maxUnpackedBytes: 1 },
});
assert.equal(ocrLanguagePolicy.status, "blocked");
assert.equal(ocrLanguagePolicy.reason.code, "ocr-language-pack-unpublished");

const missingOcrProbeLanguage = await PdfProviders.probe({
  provider: "ocrmypdf",
  task: "ocr",
  policy: { installPolicy: "system-only", allowedProviders: ["ocrmypdf"] },
});
assert.equal(missingOcrProbeLanguage.status, "blocked");
assert.equal(missingOcrProbeLanguage.reason.code, "ocr-language-required");

const rootImport = await import("node:child_process").then(({ spawnSync }) => spawnSync(process.execPath, ["--input-type=module", "--eval", [
  "globalThis.fetch=()=>{throw new Error('network must remain unused')}",
  "await import('open-office-artifact-tool/pdf/providers')",
  "process.stdout.write('providers-import-ok')",
].join(";")], { cwd: repoRoot, encoding: "utf8" }));
assert.equal(rootImport.status, 0, rootImport.stderr);
assert.equal(rootImport.stdout, "providers-import-ok");

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-providers-"));
try {
  const oldQpdf = path.join(tempRoot, "old-qpdf.mjs");
  await fs.writeFile(oldQpdf, "#!/usr/bin/env node\nif (process.argv.includes('--version')) console.log('qpdf version 11.6.3'); else { console.error('qpdf: unrecognized argument -v'); process.exit(2); }\n", "utf8");
  await fs.chmod(oldQpdf, 0o755);
  const previousQpdf = process.env.OPEN_OFFICE_PDF_QPDF;
  process.env.OPEN_OFFICE_PDF_QPDF = oldQpdf;
  try {
    const oldRepairProbe = await PdfProviders.probe({
      provider: "qpdf",
      task: "repair",
      policy: { installPolicy: "system-only", allowedProviders: ["qpdf"] },
    });
    assert.equal(oldRepairProbe.status, "ready", "qpdf 11.6 remains valid for repair");
    const oldEncryptProbe = await PdfProviders.probe({
      provider: "qpdf",
      task: "encrypt",
      policy: { installPolicy: "system-only", allowedProviders: ["qpdf"] },
    });
    assert.equal(oldEncryptProbe.status, "blocked");
    assert.equal(oldEncryptProbe.reason.code, "system-provider-unavailable");
    assert.equal(oldEncryptProbe.runtime.evidence.taskMinimumVersion, "11.7.0");
  } finally {
    if (previousQpdf === undefined) delete process.env.OPEN_OFFICE_PDF_QPDF;
    else process.env.OPEN_OFFICE_PDF_QPDF = previousQpdf;
  }

  // A system-only Python provider may be selected through a caller-owned
  // absolute executable path. This keeps the managed-runtime direct-path
  // probe and its explicit policy equivalent: neither depends on ambient PATH.
  const explicitPython = path.join(tempRoot, "explicit-python.mjs");
  await fs.writeFile(explicitPython, [
    "#!/usr/bin/env node",
    "if (!process.argv.includes('-c')) process.exit(2);",
    "const payload = JSON.parse(process.argv.at(-1));",
    "if (payload.module !== 'reportlab' || payload.distribution !== 'reportlab') process.exit(3);",
    "process.stdout.write(JSON.stringify({ moduleFound: true, version: '4.4.9', companionFound: true, companionVersion: null }));",
    "",
  ].join("\n"), "utf8");
  await fs.chmod(explicitPython, 0o755);
  const explicitPythonProbe = await PdfProviders.probe({
    provider: "reportlab",
    task: "create-layout",
    policy: {
      installPolicy: "system-only",
      allowedProviders: ["reportlab"],
      providerPython: explicitPython,
    },
  });
  assert.equal(explicitPythonProbe.status, "ready", JSON.stringify(explicitPythonProbe.reason));
  assert.equal(explicitPythonProbe.runtime.evidence.python, explicitPython);
  assert.equal(explicitPythonProbe.runtime.evidence.version, "4.4.9");
  const relativePythonProbe = await PdfProviders.probe({
    provider: "reportlab",
    task: "create-layout",
    policy: {
      installPolicy: "system-only",
      allowedProviders: ["reportlab"],
      providerPython: path.relative(process.cwd(), explicitPython),
    },
  });
  assert.equal(relativePythonProbe.status, "ready", JSON.stringify(relativePythonProbe.reason));
  assert.equal(relativePythonProbe.runtime.evidence.python, explicitPython);

  const policyDirectory = path.join(tempRoot, ".open-office-artifact-tool");
  await fs.mkdir(policyDirectory);
  const policyPath = path.join(policyDirectory, "pdf-providers.json");
  await fs.writeFile(policyPath, JSON.stringify({
    installPolicy: "managed",
    allowedProviders: ["qpdf"],
    allowedPacks: ["qpdf"],
    maxDownloadBytes: 100_000_000,
    maxUnpackedBytes: 100_000_000,
  }), "utf8");
  const fileBackedQpdf = await PdfProviders.resolve({
    task: "repair",
    provider: "qpdf",
    savePolicy: "rewrite",
    inspection: inspectedPdf,
    mutationAuthorized: true,
    invalidateSignaturesAuthorized: true,
    policyPath,
  });
  assert.equal(fileBackedQpdf.policySource, "explicit-file");
  assert.equal(fileBackedQpdf.policyPath, policyPath);
  assert.equal(fileBackedQpdf.cacheRoot, path.join(policyDirectory, "providers"));
  assert.equal(fileBackedQpdf.status, "installable");
  assert.equal(fileBackedQpdf.reason.code, "managed-install-required");

  const normalArchive = tarGz([{ name: "bin/tool", bytes: "#!/bin/sh\necho fixture\n", mode: 0o755 }]);
  const pack = fixturePack(normalArchive);
  const cacheRoot = path.join(tempRoot, "cache");
  const calls = [];
  const first = await installManagedPackForTest({ cacheRoot, pack, fetchImpl: fakeFetch(normalArchive, calls) });
  assert.equal(first.ready, true);
  assert.equal(first.reused, false);
  assert.equal(calls.length, 1);
  const receipt = JSON.parse(await fs.readFile(path.join(first.root, ".receipt.json"), "utf8"));
  assert.equal(receipt.schema, PDF_PROVIDER_RECEIPT_SCHEMA);
  assert.equal(receipt.artifact.sha256, pack.artifacts[0].sha256);
  await fs.access(path.join(first.root, "bin", "tool"));
  assert.ok((await listTree(cacheRoot)).every((entry) => !entry.includes(".fixture-pack.tmp-") && !entry.endsWith(".lock")), "successful install must clean its download staging and lock");
  const second = await installManagedPackForTest({ cacheRoot, pack, fetchImpl: async () => { throw new Error("cache hit must not download"); } });
  assert.equal(second.ready, true);
  assert.equal(second.reused, true);

  const mirrorCalls = [];
  const mirrorCache = path.join(tempRoot, "mirror-cache");
  await installManagedPackForTest({
    cacheRoot: mirrorCache,
    pack,
    enterpriseMirror: "https://mirror.example.test/open-office-artifact-tool/v1.2.3/",
    fetchImpl: fakeFetch(normalArchive, mirrorCalls),
  });
  assert.deepEqual(mirrorCalls, ["https://mirror.example.test/open-office-artifact-tool/v1.2.3/fixture.tar.gz"]);

  const redirectCalls = [];
  const redirectCache = path.join(tempRoot, "redirect-cache");
  const redirected = await installManagedPackForTest({
    cacheRoot: redirectCache,
    pack,
    fetchImpl: async (url, options) => {
      redirectCalls.push({ url: String(url), redirect: options?.redirect });
      if (redirectCalls.length === 1) {
        return new Response(null, {
          status: 302,
          headers: { location: "https://release-storage.example.test/download/fixture.tar.gz?short_lived=token" },
        });
      }
      return new Response(normalArchive, { status: 200 });
    },
  });
  assert.equal(redirected.sourceUrl, pack.artifacts[0].url, "a temporary redirected URL must not enter the receipt result");
  assert.deepEqual(redirectCalls, [
    { url: pack.artifacts[0].url, redirect: "manual" },
    { url: "https://release-storage.example.test/download/fixture.tar.gz?short_lived=token", redirect: "manual" },
  ]);
  await assert.rejects(() => installManagedPackForTest({
    cacheRoot: path.join(tempRoot, "http-redirect-cache"),
    pack,
    fetchImpl: async () => new Response(null, { status: 302, headers: { location: "http://release-storage.example.test/fixture.tar.gz" } }),
  }), /credential-free HTTPS/);
  await assert.rejects(() => installManagedPackForTest({
    cacheRoot: path.join(tempRoot, "missing-redirect-location-cache"),
    pack,
    fetchImpl: async () => new Response(null, { status: 302 }),
  }), /redirect without a location/);
  await assert.rejects(() => installManagedPackForTest({
    cacheRoot: path.join(tempRoot, "redirect-loop-cache"),
    pack,
    fetchImpl: async () => new Response(null, { status: 307, headers: { location: "https://release-storage.example.test/again" } }),
  }), /exceeded 5 HTTPS redirects/);

  const concurrentCache = path.join(tempRoot, "concurrent-cache");
  let concurrentFetches = 0;
  const delayedFetch = async () => {
    concurrentFetches += 1;
    await new Promise((resolve) => setTimeout(resolve, 75));
    return new Response(normalArchive, { status: 200 });
  };
  const concurrent = await Promise.all([
    installManagedPackForTest({ cacheRoot: concurrentCache, pack, fetchImpl: delayedFetch }),
    installManagedPackForTest({ cacheRoot: concurrentCache, pack, fetchImpl: delayedFetch }),
  ]);
  assert.equal(concurrentFetches, 1, "concurrent installers must share one per-pack lock");
  assert.ok(concurrent.some((result) => result.reused === false));
  assert.ok(concurrent.some((result) => result.reused === true));

  const badHash = fixturePack(normalArchive, { artifact: { sha256: "0".repeat(64) } });
  await assert.rejects(() => installManagedPackForTest({ cacheRoot: path.join(tempRoot, "bad-hash"), pack: badHash, fetchImpl: fakeFetch(normalArchive) }), /SHA-256 mismatch/);
  const oversize = fixturePack(normalArchive, { artifact: { downloadBytes: normalArchive.length - 1 } });
  await assert.rejects(() => installManagedPackForTest({ cacheRoot: path.join(tempRoot, "oversize"), pack: oversize, fetchImpl: fakeFetch(normalArchive) }), /exceeds pinned size/);
  await assert.rejects(() => installManagedPackForTest({ cacheRoot: path.join(tempRoot, "wrong-platform"), pack, platform: "win32-x64", fetchImpl: fakeFetch(normalArchive) }), /exactly one artifact/);

  const traversal = tarGz([{ name: "../escape", bytes: "nope", mode: 0o644 }]);
  const traversalRoot = path.join(tempRoot, "traversal");
  await fs.mkdir(traversalRoot);
  await assert.rejects(() => safeExtractTarGz(traversal, traversalRoot, 16 * 1024), /Unsafe|escape/);
  const tinyPayload = tarGz([{ name: "bin/tool", bytes: "two bytes", mode: 0o755 }]);
  const tinyPayloadRoot = path.join(tempRoot, "tiny-payload");
  await fs.mkdir(tinyPayloadRoot);
  await assert.rejects(() => safeExtractTarGz(tinyPayload, tinyPayloadRoot, 1), /declared unpacked limit/);
  const metadataHeavyEntries = Array.from({ length: 2_500 }, (_, index) => ({
    name: `metadata/${String(index).padStart(5, "0")}.txt`,
    bytes: "x",
    mode: 0o644,
  }));
  const metadataHeavyArchive = tarGz(metadataHeavyEntries);
  const metadataHeavyRoot = path.join(tempRoot, "metadata-heavy");
  await fs.mkdir(metadataHeavyRoot);
  const metadataHeavyExtraction = await safeExtractTarGz(metadataHeavyArchive, metadataHeavyRoot, metadataHeavyEntries.length);
  assert.equal(metadataHeavyExtraction.unpackedBytes, metadataHeavyEntries.length, "USTAR headers and padding must not consume the advertised extracted-file budget");
  const hardlink = tarGz([{ name: "bin/tool", bytes: "target", type: "1", mode: 0o755 }]);
  await fs.mkdir(path.join(tempRoot, "hardlink"));
  await assert.rejects(() => safeExtractTarGz(hardlink, path.join(tempRoot, "hardlink"), 16 * 1024), /Unsupported or unsafe/);
  const symlink = tarGz([{ name: "bin/tool", bytes: "target", type: "2", mode: 0o755 }]);
  await fs.mkdir(path.join(tempRoot, "symlink"));
  await assert.rejects(() => safeExtractTarGz(symlink, path.join(tempRoot, "symlink"), 16 * 1024), /Unsupported or unsafe/);

  const interruptedCache = path.join(tempRoot, "interrupted-cache");
  await assert.rejects(() => installManagedPackForTest({ cacheRoot: interruptedCache, pack, fetchImpl: async () => { throw new Error("offline fixture"); } }), /offline fixture/);
  const interruptedTree = await listTree(interruptedCache);
  assert.ok(interruptedTree.every((entry) => !entry.includes(".fixture-pack.tmp-") && !entry.endsWith(".lock")), `interrupted install left temporary state: ${interruptedTree.join(", ")}`);
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log("pdf providers smoke ok");
