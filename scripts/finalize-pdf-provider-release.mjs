#!/usr/bin/env node
/**
 * Combine independently built platform packs into one immutable release
 * manifest. The emitted `catalogFragment` is deliberately data-only: a human
 * must review and commit its pinned digests before a package advertises the
 * release to customers.
 */

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

const RELEASE_SCHEMA = "open-office-artifact-tool.pdf-provider-release.v1";
const PACK_SCHEMA = "open-office-artifact-tool.pdf-provider-pack.v1";
const SUPPORTED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const SHA256 = /^[a-f0-9]{64}$/i;

function fail(message) {
  throw new Error(`PDF capability-pack release: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function safeSegment(value, label) {
  if (!nonEmptyString(value) || value.includes("/") || value.includes("\\") || value === "." || value === "..") fail(`${label} must be one safe path segment.`);
  return value;
}

function safeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\") || value.startsWith("/")) return false;
  const normalized = path.posix.normalize(value);
  return normalized !== "." && normalized !== ".." && !normalized.startsWith("../") && !normalized.includes("/../");
}

function httpsUrl(value) {
  try {
    return new URL(value).protocol === "https:";
  } catch {
    return false;
  }
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function stableJson(value) {
  return `${JSON.stringify(value, null, 2)}\n`;
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
  for (const required of ["pack", "version", "input", "output", "release-base-url", "repository", "workflow"]) {
    if (!nonEmptyString(values[required])) fail(`--${required} is required.`);
  }
  safeSegment(values.pack, "pack");
  safeSegment(values.version, "version");
  const expectedPlatforms = (values["expected-platforms"] || "darwin-arm64,linux-x64").split(",").map((item) => item.trim()).filter(Boolean);
  if (!expectedPlatforms.length || expectedPlatforms.some((platform) => !SUPPORTED_PLATFORMS.has(platform))) {
    fail("--expected-platforms must be a non-empty comma-separated subset of darwin-arm64,linux-x64.");
  }
  let releaseBaseUrl;
  try {
    releaseBaseUrl = new URL(values["release-base-url"]);
  } catch {
    fail("--release-base-url must be an HTTPS URL.");
  }
  if (releaseBaseUrl.protocol !== "https:" || releaseBaseUrl.username || releaseBaseUrl.password || !releaseBaseUrl.pathname.endsWith("/")) {
    fail("--release-base-url must be a credential-free HTTPS URL ending in '/'.");
  }
  if (!/^[A-Za-z0-9_.-]+\/[A-Za-z0-9_.-]+$/.test(values.repository)) fail("--repository must be an owner/repository identifier.");
  if (!values.workflow.startsWith(".github/workflows/") || !values.workflow.endsWith(".yml")) fail("--workflow must name a repository workflow path.");
  return {
    pack: values.pack,
    version: values.version,
    input: path.resolve(values.input),
    output: path.resolve(values.output),
    releaseBaseUrl: releaseBaseUrl.href,
    repository: values.repository,
    workflow: values.workflow,
    expectedPlatforms: [...new Set(expectedPlatforms)].sort(),
  };
}

async function collectManifests(directory) {
  const stat = await fs.lstat(directory);
  if (!stat.isDirectory() || stat.isSymbolicLink()) fail("--input must be a real non-symlink directory.");
  const entries = await fs.readdir(directory, { withFileTypes: true });
  const files = entries.filter((entry) => entry.isFile() && entry.name.endsWith(".manifest.json")).map((entry) => entry.name).sort();
  if (!files.length) fail("--input contains no platform pack manifests.");
  return Promise.all(files.map(async (file) => {
    const absolute = path.join(directory, file);
    const stat = await fs.lstat(absolute);
    if (stat.isSymbolicLink()) fail(`manifest must not be a symlink: ${file}.`);
    let value;
    try {
      value = JSON.parse(await fs.readFile(absolute, "utf8"));
    } catch (error) {
      fail(`manifest is invalid JSON: ${file}: ${error.message}`);
    }
    return { file, value };
  }));
}

function assertPackManifest(value, { pack, version, expectedPlatforms }) {
  if (!value || typeof value !== "object" || value.schema !== PACK_SCHEMA || value.schemaVersion !== 1) fail("input contains an unsupported pack manifest.");
  if (value.pack !== pack || value.version !== version || !expectedPlatforms.includes(value.platform)) fail("input manifest does not match requested pack, version, or platform.");
  const artifact = value.artifact;
  if (!artifact || typeof artifact !== "object" || !safeRelativePath(artifact.asset) || !SHA256.test(artifact.sha256 || "")
    || !Number.isSafeInteger(artifact.downloadBytes) || artifact.downloadBytes <= 0
    || !Number.isSafeInteger(artifact.unpackedBytes) || artifact.unpackedBytes <= 0 || artifact.archiveFormat !== "tar.gz") {
    fail(`input manifest has an invalid artifact for ${value.platform}.`);
  }
  for (const evidenceName of ["sbom", "thirdPartyNotices"]) {
    const evidence = value[evidenceName];
    if (!evidence || !safeRelativePath(evidence.asset) || !SHA256.test(evidence.sha256 || "") || !Number.isSafeInteger(evidence.bytes) || evidence.bytes <= 0) {
      fail(`input manifest has invalid ${evidenceName} evidence for ${value.platform}.`);
    }
  }
  if (!value.source || !httpsUrl(value.source.url) || !SHA256.test(value.source.sha256 || "")) fail(`input manifest has invalid source provenance for ${value.platform}.`);
}

async function verifiedBytes(directory, asset, expectedSha, expectedBytes, label) {
  if (!safeRelativePath(asset)) fail(`${label} has an unsafe asset path: ${asset}.`);
  const target = path.join(directory, asset);
  const stat = await fs.lstat(target);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(`${label} is missing or unsafe: ${asset}.`);
  const bytes = await fs.readFile(target);
  if (bytes.length !== expectedBytes || sha256(bytes) !== expectedSha.toLowerCase()) fail(`${label} hash or size does not match its platform manifest: ${asset}.`);
  return bytes;
}

function releaseAssetUrl(base, asset) {
  if (!safeRelativePath(asset)) fail(`release asset path is unsafe: ${asset}.`);
  return new URL(asset, base).href;
}

function deterministicBomSerial(...values) {
  const digest = sha256(Buffer.from(values.join("\u0000"), "utf8"));
  return `urn:uuid:${digest.slice(0, 8)}-${digest.slice(8, 12)}-5${digest.slice(13, 16)}-${(Number.parseInt(digest[16], 16) & 0x3 | 0x8).toString(16)}${digest.slice(17, 20)}-${digest.slice(20, 32)}`;
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const manifests = await collectManifests(options.input);
  const byPlatform = new Map();
  for (const { value } of manifests) {
    assertPackManifest(value, options);
    if (byPlatform.has(value.platform)) fail(`input has duplicate platform manifest ${value.platform}.`);
    byPlatform.set(value.platform, value);
  }
  if (byPlatform.size !== options.expectedPlatforms.length || options.expectedPlatforms.some((platform) => !byPlatform.has(platform))) {
    fail(`input must contain exactly these platforms: ${options.expectedPlatforms.join(", ")}.`);
  }
  const ordered = options.expectedPlatforms.map((platform) => byPlatform.get(platform));
  const noticeBytes = [];
  for (const manifest of ordered) {
    await verifiedBytes(options.input, manifest.artifact.asset, manifest.artifact.sha256, manifest.artifact.downloadBytes, "archive");
    await verifiedBytes(options.input, manifest.sbom.asset, manifest.sbom.sha256, manifest.sbom.bytes, "platform SBOM");
    noticeBytes.push(await verifiedBytes(options.input, manifest.thirdPartyNotices.asset, manifest.thirdPartyNotices.sha256, manifest.thirdPartyNotices.bytes, "platform notices"));
  }
  if (noticeBytes.some((bytes) => !bytes.equals(noticeBytes[0]))) fail("platform packs must have byte-identical third-party notices.");

  const releaseBaseName = `${options.pack}-${options.version}`;
  const releaseSbomAsset = `${releaseBaseName}.sbom.cdx.json`;
  const releaseNoticesAsset = `${releaseBaseName}.THIRD_PARTY_NOTICES.md`;
  const releaseManifestAsset = `${releaseBaseName}.release-manifest.json`;
  const releaseSbomBytes = Buffer.from(stableJson({
    bomFormat: "CycloneDX",
    specVersion: "1.5",
    serialNumber: deterministicBomSerial(options.pack, options.version, options.repository, options.workflow, ...options.expectedPlatforms),
    version: 1,
    metadata: {
      component: {
        type: "application",
        name: `open-office-artifact-tool-pdf-provider-${options.pack}`,
        version: options.version,
        properties: [
          { name: "open-office-artifact-tool:release-platforms", value: options.expectedPlatforms.join(",") },
          { name: "open-office-artifact-tool:provenance-repository", value: options.repository },
          { name: "open-office-artifact-tool:provenance-workflow", value: options.workflow },
        ],
      },
    },
    components: ordered.map((manifest) => ({
      type: "application",
      name: options.pack,
      version: options.version,
      properties: [{ name: "open-office-artifact-tool:platform", value: manifest.platform }],
      hashes: [{ alg: "SHA-256", content: manifest.artifact.sha256 }],
      externalReferences: [{ type: "distribution", url: releaseAssetUrl(options.releaseBaseUrl, manifest.artifact.asset) }],
    })),
  }), "utf8");
  const releaseNoticesBytes = noticeBytes[0];
  const releaseEvidence = {
    sbom: { asset: releaseSbomAsset, url: releaseAssetUrl(options.releaseBaseUrl, releaseSbomAsset), sha256: sha256(releaseSbomBytes) },
    thirdPartyNotices: { asset: releaseNoticesAsset, url: releaseAssetUrl(options.releaseBaseUrl, releaseNoticesAsset), sha256: sha256(releaseNoticesBytes) },
    provenance: {
      provider: "github-actions-artifact-attestation",
      repository: options.repository,
      workflow: options.workflow,
      verificationCommand: `gh attestation verify <asset> --repo ${options.repository}`,
    },
    verifiedPlatforms: [...options.expectedPlatforms],
  };
  const catalogFragment = {
    state: "published",
    version: options.version,
    artifacts: ordered.map((manifest) => ({
      platform: manifest.platform,
      asset: manifest.artifact.asset,
      version: options.version,
      url: releaseAssetUrl(options.releaseBaseUrl, manifest.artifact.asset),
      sha256: manifest.artifact.sha256,
      downloadBytes: manifest.artifact.downloadBytes,
      unpackedBytes: manifest.artifact.unpackedBytes,
      archiveFormat: "tar.gz",
    })),
    releaseEvidence,
  };
  const releaseManifest = {
    schema: RELEASE_SCHEMA,
    schemaVersion: 1,
    pack: options.pack,
    version: options.version,
    releaseBaseUrl: options.releaseBaseUrl,
    catalogFragment,
    platformManifests: ordered.map((manifest) => ({
      platform: manifest.platform,
      source: manifest.source,
      artifact: manifest.artifact,
      sbom: manifest.sbom,
      thirdPartyNotices: manifest.thirdPartyNotices,
    })),
  };
  await fs.mkdir(options.output, { recursive: true, mode: 0o700 });
  await Promise.all([
    fs.writeFile(path.join(options.output, releaseSbomAsset), releaseSbomBytes, { mode: 0o600 }),
    fs.writeFile(path.join(options.output, releaseNoticesAsset), releaseNoticesBytes, { mode: 0o600 }),
    fs.writeFile(path.join(options.output, releaseManifestAsset), stableJson(releaseManifest), { mode: 0o600 }),
  ]);
  process.stdout.write(stableJson(releaseManifest));
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
