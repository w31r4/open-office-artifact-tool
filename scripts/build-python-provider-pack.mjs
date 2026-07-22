#!/usr/bin/env node
/**
 * Build a self-contained, hash-pinned Python PDF-provider pack.
 *
 * This is release tooling, never a customer-side installer. It starts from an
 * exact python-build-standalone archive plus an exact wheel lock, dereferences
 * every source symlink before packaging, installs only already verified wheels
 * with `pip --no-index --no-deps`, then delegates archive construction to the
 * same strict USTAR+gzip pack builder used by qpdf.
 */

import crypto from "node:crypto";
import { execFile as execFileCallback } from "node:child_process";
import { constants as fsConstants } from "node:fs";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

const execFile = promisify(execFileCallback);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const DEFAULT_INPUTS = path.join(__dirname, "pdf-provider-python-release-inputs.v1.json");
const PACK_BUILDER = path.join(__dirname, "build-pdf-provider-pack.mjs");
const INPUT_SCHEMA = "open-office-artifact-tool.pdf-provider-python-release-inputs.v1";
const SUPPORTED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const SHA256 = /^[a-f0-9]{64}$/i;
const MAX_REDIRECTS = 5;
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_TOTAL_WHEEL_BYTES = 256 * 1024 * 1024;
const MAX_NOTICE_FILE_BYTES = 1024 * 1024;

function fail(message) {
  throw new Error(`Python PDF capability-pack build: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function plainObject(value) {
  return Boolean(value) && typeof value === "object" && !Array.isArray(value);
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function safeSegment(value, label) {
  if (!nonEmptyString(value) || value.includes("/") || value.includes("\\") || value === "." || value === "..") fail(`${label} must be one safe path segment.`);
  return value;
}

function httpsUrl(value, label) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    fail(`${label} must be a credential-free HTTPS URL.`);
  }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password) fail(`${label} must be a credential-free HTTPS URL.`);
  return parsed;
}

function safeWheelName(value, label) {
  if (!nonEmptyString(value) || !value.endsWith(".whl") || value.includes("/") || value.includes("\\") || value === "." || value === "..") fail(`${label} must name one wheel file.`);
  return value;
}

function normalizedPackageName(value) {
  return value.toLowerCase().replace(/[-_.]+/g, "-");
}

function pythonLibraryRelativePath(runtime) {
  const match = /^(\d+)\.(\d+)/.exec(runtime.version || "");
  if (!match) fail("Python runtime version must begin with major.minor.");
  return path.join("lib", `python${match[1]}.${match[2]}`);
}

function wheelNameFromUrl(value, label) {
  const url = httpsUrl(value, label);
  return safeWheelName(path.posix.basename(url.pathname), label);
}

function assertPinnedSource(value, label, { requireBytes = false } = {}) {
  if (!plainObject(value)) fail(`${label} must be an object.`);
  httpsUrl(value.url, `${label}.url`);
  if (!SHA256.test(value.sha256 || "")) fail(`${label}.sha256 must be exactly 64 hexadecimal characters.`);
  if (requireBytes && (!Number.isSafeInteger(value.downloadBytes) || value.downloadBytes <= 0 || value.downloadBytes > MAX_SOURCE_BYTES)) {
    fail(`${label}.downloadBytes must be a positive bounded integer.`);
  }
}

function assertWheel(value, label) {
  assertPinnedSource(value, label);
  if (!nonEmptyString(value.name) || !nonEmptyString(value.version)) fail(`${label} must pin a package name and version.`);
  wheelNameFromUrl(value.url, `${label}.url`);
}

function validateInputs(value) {
  if (!plainObject(value) || value.schema !== INPUT_SCHEMA || value.schemaVersion !== 1) fail("inputs use an unsupported schema.");
  const runtime = value.pythonRuntime;
  if (!plainObject(runtime) || !nonEmptyString(runtime.distribution) || !nonEmptyString(runtime.version) || !nonEmptyString(runtime.license) || !plainObject(runtime.platforms)) {
    fail("inputs must declare one pinned Python runtime.");
  }
  for (const platform of SUPPORTED_PLATFORMS) assertPinnedSource(runtime.platforms[platform], `pythonRuntime.platforms.${platform}`, { requireBytes: true });
  if (!plainObject(value.packs)) fail("inputs must declare packs.");
  for (const [packId, pack] of Object.entries(value.packs)) {
    safeSegment(packId, "pack id");
    if (!plainObject(pack) || !nonEmptyString(pack.license) || !plainObject(pack.directRequirements)
      || !Array.isArray(pack.probes) || pack.probes.some((probe) => !nonEmptyString(probe))
      || !Array.isArray(pack.commonWheels) || !plainObject(pack.platformWheels)) {
      fail(`pack ${packId} is incomplete.`);
    }
    const direct = Object.entries(pack.directRequirements);
    if (!direct.length || direct.some(([name, version]) => !nonEmptyString(name) || !nonEmptyString(version))) fail(`pack ${packId} has invalid direct requirements.`);
    const commonWheelKeys = new Set();
    const inspectWheel = (wheel, label) => {
      assertWheel(wheel, label);
      return `${normalizedPackageName(wheel.name)}@${wheel.version}`;
    };
    pack.commonWheels.forEach((wheel, index) => {
      const key = inspectWheel(wheel, `pack ${packId}.commonWheels[${index}]`);
      if (commonWheelKeys.has(key)) fail(`pack ${packId} declares duplicate common wheel ${key}.`);
      commonWheelKeys.add(key);
    });
    if (Object.keys(pack.platformWheels).some((platform) => !SUPPORTED_PLATFORMS.has(platform))) {
      fail(`pack ${packId} declares wheels for an unsupported platform.`);
    }
    for (const platform of SUPPORTED_PLATFORMS) {
      const wheels = pack.platformWheels[platform];
      if (!Array.isArray(wheels) || !wheels.length) fail(`pack ${packId} lacks ${platform} wheels.`);
      const wheelKeys = new Set(commonWheelKeys);
      wheels.forEach((wheel, index) => {
        const key = inspectWheel(wheel, `pack ${packId}.platformWheels.${platform}[${index}]`);
        if (wheelKeys.has(key)) fail(`pack ${packId} declares duplicate ${platform} wheel ${key}.`);
        wheelKeys.add(key);
      });
      for (const [name, version] of direct) {
        const key = `${normalizedPackageName(name)}@${version}`;
        if (!wheelKeys.has(key)) fail(`pack ${packId} direct requirement ${name}==${version} is absent from its ${platform} wheel lock.`);
      }
    }
  }
  return value;
}

async function loadInputs(inputPath) {
  const absolute = path.resolve(inputPath);
  const stat = await fs.lstat(absolute);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > 512 * 1024) fail("inputs must be a bounded regular non-symlink file.");
  const bytes = await fs.readFile(absolute);
  let parsed;
  try {
    parsed = JSON.parse(bytes);
  } catch (error) {
    fail(`inputs are not valid JSON: ${error.message}`);
  }
  return { absolute, bytes, value: validateInputs(parsed) };
}

function parseArguments(argv) {
  const values = {};
  const flags = new Set();
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    if (name === "verify-lock") {
      if (flags.has(name)) fail(`--${name} may be supplied only once.`);
      flags.add(name);
      continue;
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
    values[name] = value;
    index += 1;
  }
  const verifyLock = flags.has("verify-lock");
  if (!verifyLock) {
    for (const required of ["pack", "version", "platform", "output", "source-url"]) {
      if (!nonEmptyString(values[required])) fail(`--${required} is required.`);
    }
    safeSegment(values.pack, "pack");
    safeSegment(values.version, "version");
    if (!SUPPORTED_PLATFORMS.has(values.platform)) fail(`platform must be one of ${[...SUPPORTED_PLATFORMS].join(", ")}.`);
    httpsUrl(values["source-url"], "--source-url");
  }
  return {
    verifyLock,
    inputs: values.inputs ? path.resolve(values.inputs) : DEFAULT_INPUTS,
    pack: values.pack,
    version: values.version,
    platform: values.platform,
    output: values.output ? path.resolve(values.output) : undefined,
    sourceUrl: values["source-url"],
  };
}

function assertWithin(root, target, label) {
  const normalizedRoot = path.resolve(root);
  const normalizedTarget = path.resolve(target);
  if (normalizedTarget === normalizedRoot || !normalizedTarget.startsWith(`${normalizedRoot}${path.sep}`)) fail(`${label} escapes its root.`);
  return normalizedTarget;
}

async function copyDereferencedTree(source, destination, rootReal, visiting = new Set()) {
  const link = await fs.lstat(source);
  let actual = source;
  let stat = link;
  if (link.isSymbolicLink()) {
    actual = await fs.realpath(source);
    assertWithin(rootReal, actual, "Python runtime symlink");
    stat = await fs.lstat(actual);
  }
  const real = await fs.realpath(actual);
  if (stat.isDirectory()) {
    if (visiting.has(real)) fail(`Python runtime contains a symlink directory cycle at ${source}.`);
    visiting.add(real);
    await fs.mkdir(destination, { recursive: true, mode: 0o755 });
    const entries = await fs.readdir(actual, { withFileTypes: true });
    entries.sort((left, right) => left.name.localeCompare(right.name, "en"));
    for (const entry of entries) {
      if (entry.name === "." || entry.name === "..") fail("Python runtime contains an unsafe directory entry.");
      await copyDereferencedTree(path.join(actual, entry.name), path.join(destination, entry.name), rootReal, visiting);
    }
    visiting.delete(real);
    return;
  }
  if (!stat.isFile()) fail(`Python runtime contains unsupported filesystem entry ${source}.`);
  await fs.copyFile(actual, destination, fsConstants.COPYFILE_EXCL);
  await fs.chmod(destination, stat.mode & 0o777);
}

function isRedirect(status) {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

async function downloadPinned({ url, sha256: expectedSha256, downloadBytes }, destination, label, total) {
  let current = httpsUrl(url, `${label}.url`).href;
  let response;
  for (let redirects = 0; redirects <= MAX_REDIRECTS; redirects += 1) {
    response = await fetch(current, { redirect: "manual" });
    if (!isRedirect(response.status)) break;
    if (redirects === MAX_REDIRECTS) fail(`${label} exceeded ${MAX_REDIRECTS} HTTPS redirects.`);
    const location = response.headers.get("location");
    if (!nonEmptyString(location)) fail(`${label} redirected without a location.`);
    const next = new URL(location, current);
    if (next.protocol !== "https:" || next.username || next.password) fail(`${label} redirect must use credential-free HTTPS.`);
    current = next.href;
  }
  if (!response?.ok || !response.body) fail(`${label} download failed with HTTP ${response?.status ?? "unknown"}.`);
  const handle = await fs.open(destination, "wx", 0o600);
  let bytes = 0;
  const digest = crypto.createHash("sha256");
  try {
    for await (const chunk of response.body) {
      const buffer = Buffer.from(chunk);
      bytes += buffer.length;
      total.value += buffer.length;
      const limit = downloadBytes || MAX_SOURCE_BYTES;
      if (bytes > limit || bytes > MAX_SOURCE_BYTES || total.value > MAX_TOTAL_WHEEL_BYTES + MAX_SOURCE_BYTES) fail(`${label} exceeds its pinned source budget.`);
      digest.update(buffer);
      await handle.write(buffer);
    }
  } finally {
    await handle.close();
  }
  if (downloadBytes !== undefined && bytes !== downloadBytes) fail(`${label} download size does not match the pin.`);
  if (digest.digest("hex") !== expectedSha256.toLowerCase()) fail(`${label} SHA-256 does not match the pin.`);
  return { bytes, sourceUrl: current };
}

async function collectLicenseFiles(root, limit = []) {
  const entries = await fs.readdir(root, { withFileTypes: true }).catch(() => []);
  for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name, "en"))) {
    const target = path.join(root, entry.name);
    if (entry.isDirectory()) {
      if (/\.dist-info$/.test(entry.name) || entry.name === "licenses") await collectLicenseFiles(target, limit);
      continue;
    }
    if (!entry.isFile() || !/(?:^|[-_.])(license|copying|notice)(?:[-_.]|$)/i.test(entry.name)) continue;
    const stat = await fs.lstat(target);
    if (stat.size <= 0 || stat.size > MAX_NOTICE_FILE_BYTES) continue;
    limit.push(target);
  }
  return limit;
}

async function noticesFor({ payload, runtime, wheels }) {
  const parts = [
    "# Python PDF capability-pack notices",
    "",
    "This immutable pack was assembled only from the hash-pinned runtime and wheel sources recorded below.",
    "",
    "## Python runtime",
    `- ${runtime.distribution} ${runtime.version}`,
    `- license expression: ${runtime.license}`,
    "",
    "## Wheel sources",
    ...wheels.map((wheel) => `- ${wheel.name} ${wheel.version}: ${wheel.url} (sha256:${wheel.sha256})`),
    "",
    "## Included license material",
  ];
  const pythonLibrary = path.join(payload, pythonLibraryRelativePath(runtime));
  const runtimeLicense = path.join(pythonLibrary, "LICENSE.txt");
  const material = [runtimeLicense, ...await collectLicenseFiles(path.join(pythonLibrary, "site-packages"))];
  const seen = new Set();
  for (const file of material.sort((left, right) => left.localeCompare(right, "en"))) {
    if (seen.has(file)) continue;
    seen.add(file);
    const relative = path.relative(payload, file).split(path.sep).join("/");
    const bytes = await fs.readFile(file).catch(() => undefined);
    if (!bytes?.length || bytes.length > MAX_NOTICE_FILE_BYTES) continue;
    parts.push("", `### ${relative}`, "", bytes.toString("utf8"));
  }
  return Buffer.from(`${parts.join("\n")}\n`, "utf8");
}

async function removePayloadPath(payload, relativePath) {
  const target = assertWithin(payload, path.join(payload, relativePath), "Python runtime prune path");
  await fs.rm(target, { recursive: true, force: true });
}

async function pruneRuntimePayload(payload, runtime) {
  // The standalone runtime deliberately includes developer tooling and GUI
  // stacks. Capability packs only execute the shipped thin adapters through
  // bin/python3, so retain exactly that interpreter and its import runtime.
  const bin = path.join(payload, "bin");
  const binEntries = await fs.readdir(bin, { withFileTypes: true });
  for (const entry of binEntries) {
    if (entry.name !== "python3") await removePayloadPath(payload, path.join("bin", entry.name));
  }
  const pythonLibrary = pythonLibraryRelativePath(runtime);
  for (const relativePath of [
    "include",
    "share",
    path.join("lib", "pkgconfig"),
    path.join("lib", "tcl9"),
    path.join("lib", "tcl9.0"),
    path.join("lib", "tk9.0"),
    path.join("lib", "itcl4.3.5"),
    path.join("lib", "thread3.0.4"),
    path.join(pythonLibrary, "ensurepip"),
    path.join(pythonLibrary, "idlelib"),
    path.join(pythonLibrary, "tkinter"),
    path.join(pythonLibrary, "turtledemo"),
  ]) await removePayloadPath(payload, relativePath);
  const libraryEntries = await fs.readdir(path.join(payload, pythonLibrary), { withFileTypes: true });
  for (const entry of libraryEntries) {
    if (entry.name.startsWith("config-")) await removePayloadPath(payload, path.join(pythonLibrary, entry.name));
  }
  const sitePackages = path.join(payload, pythonLibrary, "site-packages");
  const siteEntries = await fs.readdir(sitePackages, { withFileTypes: true });
  for (const entry of siteEntries) {
    if (/^pip(?:-|$)/.test(entry.name)) await removePayloadPath(payload, path.join(pythonLibrary, "site-packages", entry.name));
  }
  // `_tkinter` is a native bridge to the GUI stack removed above. Retaining
  // just that extension makes an otherwise headless OCR runtime depend on
  // Tcl/Tk at library-closure time, despite no supported PDF provider using
  // it. Remove the native bridge as well as the Python package so a pack
  // cannot accidentally regain a host Tcl/Tk dependency.
  const dynamicExtensions = path.join(payload, pythonLibrary, "lib-dynload");
  const dynamicEntries = await fs.readdir(dynamicExtensions, { withFileTypes: true }).catch(() => []);
  for (const entry of dynamicEntries) {
    if (/^_tkinter(?:\.|$)/.test(entry.name)) await removePayloadPath(payload, path.join(pythonLibrary, "lib-dynload", entry.name));
  }
}

async function verifyPayloadPython(python, pack) {
  const expected = Object.fromEntries(Object.entries(pack.directRequirements).map(([name, version]) => [name.toLowerCase(), version]));
  const program = [
    "import importlib, importlib.metadata as metadata, json, sys",
    "payload=json.loads(sys.argv[1])",
    "[importlib.import_module(name) for name in payload['probes']]",
    "actual={name.lower(): metadata.version(name) for name in payload['expected']}",
    "assert actual == {name.lower(): version for name, version in payload['expected'].items()}, (actual, payload['expected'])",
    "print(json.dumps(actual, sort_keys=True))",
  ].join("; ");
  await execFile(python, ["-I", "-c", program, JSON.stringify({ probes: pack.probes, expected })], {
    timeout: 30_000,
    maxBuffer: 64 * 1024,
    env: { ...process.env, PYTHONPATH: "", PYTHONHOME: "", PYTHONNOUSERSITE: "1", PIP_NO_INDEX: "1" },
  });
}

async function build(options, loaded) {
  const pack = loaded.value.packs[options.pack];
  if (!pack) fail(`inputs do not define pack ${options.pack}.`);
  const runtimeSource = loaded.value.pythonRuntime.platforms[options.platform];
  const wheels = [...pack.commonWheels, ...pack.platformWheels[options.platform]];
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-python-provider-pack-"));
  try {
    const total = { value: 0 };
    const runtimeArchive = path.join(temporary, "runtime.tar.gz");
    await downloadPinned(runtimeSource, runtimeArchive, "Python runtime", total);
    const runtimeExtract = path.join(temporary, "runtime");
    await fs.mkdir(runtimeExtract, { mode: 0o700 });
    await execFile("tar", ["-xzf", runtimeArchive, "-C", runtimeExtract], { timeout: 120_000, maxBuffer: 64 * 1024 });
    const sourceRoot = path.join(runtimeExtract, "python");
    const sourceStat = await fs.lstat(sourceRoot);
    if (!sourceStat.isDirectory() || sourceStat.isSymbolicLink()) fail("Python runtime archive does not contain a safe python directory.");
    const payload = path.join(temporary, "payload");
    await fs.mkdir(payload, { mode: 0o700 });
    await copyDereferencedTree(sourceRoot, payload, await fs.realpath(sourceRoot));
    const python = path.join(payload, "bin", "python3");
    const pythonStat = await fs.lstat(python);
    if (!pythonStat.isFile() || pythonStat.isSymbolicLink() || (pythonStat.mode & 0o111) === 0) fail("dereferenced Python runtime does not expose executable bin/python3.");

    const wheelhouse = path.join(temporary, "wheelhouse");
    await fs.mkdir(wheelhouse, { mode: 0o700 });
    const wheelPaths = [];
    const seenNames = new Set();
    for (const wheel of wheels) {
      const name = wheelNameFromUrl(wheel.url, `wheel ${wheel.name}`);
      if (seenNames.has(name)) fail(`wheel lock reuses filename ${name}.`);
      seenNames.add(name);
      const destination = path.join(wheelhouse, name);
      await downloadPinned(wheel, destination, `wheel ${wheel.name}==${wheel.version}`, total);
      wheelPaths.push(destination);
    }
    await execFile(python, ["-m", "pip", "install", "--disable-pip-version-check", "--no-index", "--no-deps", "--no-compile", "--no-warn-script-location", ...wheelPaths], {
      timeout: 120_000,
      maxBuffer: 128 * 1024,
      env: { ...process.env, PYTHONPATH: "", PYTHONHOME: "", PYTHONNOUSERSITE: "1", PIP_NO_INDEX: "1", PIP_DISABLE_PIP_VERSION_CHECK: "1" },
    });
    await pruneRuntimePayload(payload, loaded.value.pythonRuntime);
    await verifyPayloadPython(python, pack);
    const notices = path.join(temporary, "THIRD_PARTY_NOTICES.md");
    await fs.writeFile(notices, await noticesFor({ payload, runtime: loaded.value.pythonRuntime, wheels }), { mode: 0o600 });
    const result = await execFile(process.execPath, [PACK_BUILDER,
      "--pack", options.pack,
      "--version", options.version,
      "--platform", options.platform,
      "--payload", payload,
      "--output", options.output,
      "--source-url", options.sourceUrl,
      "--source-sha256", sha256(loaded.bytes),
      "--license", pack.license,
      "--notices", notices,
      "--summary",
    ], { timeout: 120_000, maxBuffer: 1024 * 1024 });
    process.stdout.write(result.stdout);
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  const loaded = await loadInputs(options.inputs);
  if (options.verifyLock) {
    const summary = Object.fromEntries(Object.entries(loaded.value.packs).map(([packId, pack]) => [packId, Object.fromEntries([...SUPPORTED_PLATFORMS].map((platform) => [platform, pack.commonWheels.length + pack.platformWheels[platform].length]))]));
    process.stdout.write(`${JSON.stringify({ schema: loaded.value.schema, sha256: sha256(loaded.bytes), packs: summary }, null, 2)}\n`);
    return;
  }
  await build(options, loaded);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
