#!/usr/bin/env node
/**
 * Build a self-contained, hash-pinned veraPDF capability pack.
 *
 * The published pack contains only the CLI installation and its own Eclipse
 * Temurin JRE. This is release tooling, never a customer-side installer:
 * normal npm use only consumes the already-built, hash-pinned archive through
 * the managed-provider installer.
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
const ROOT = path.resolve(__dirname, "..");
const DEFAULT_INPUTS = path.join(__dirname, "pdf-provider-verapdf-release-inputs.v1.json");
const PACK_BUILDER = path.join(__dirname, "build-pdf-provider-pack.mjs");
const INSTALL_TEMPLATE = path.join(ROOT, ".github", "verapdf-auto-install.xml");
const INPUT_SCHEMA = "open-office-artifact-tool.pdf-provider-verapdf-release-inputs.v1";
const SUPPORTED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const SHA256 = /^[a-f0-9]{64}$/i;
const MAX_INPUT_BYTES = 256 * 1024;
const MAX_SOURCE_BYTES = 128 * 1024 * 1024;
const MAX_TOTAL_SOURCE_BYTES = 256 * 1024 * 1024;
const MAX_REDIRECTS = 5;

function fail(message) {
  throw new Error("veraPDF capability-pack build: " + message);
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
  if (!nonEmptyString(value) || value.includes("/") || value.includes("\\") || value === "." || value === "..") {
    fail(label + " must be one safe path segment.");
  }
  return value;
}

function safeRelativePath(value) {
  if (!nonEmptyString(value) || value.includes("\\") || value.startsWith("/")) return false;
  const normalized = path.posix.normalize(value);
  return normalized !== "." && normalized !== ".." && !normalized.startsWith("../") && !normalized.includes("/../");
}

function httpsUrl(value, label) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    fail(label + " must be a credential-free HTTPS URL.");
  }
  if (parsed.protocol !== "https:" || parsed.username || parsed.password) fail(label + " must be a credential-free HTTPS URL.");
  return parsed;
}

function assertPinnedSource(value, label) {
  if (!plainObject(value)) fail(label + " must be an object.");
  httpsUrl(value.url, label + ".url");
  if (!SHA256.test(value.sha256 || "")) fail(label + ".sha256 must be exactly 64 hexadecimal characters.");
  if (!Number.isSafeInteger(value.downloadBytes) || value.downloadBytes <= 0 || value.downloadBytes > MAX_SOURCE_BYTES) {
    fail(label + ".downloadBytes must be a positive bounded integer.");
  }
}

function assertWithin(root, target, label) {
  const normalizedRoot = path.resolve(root);
  const normalizedTarget = path.resolve(target);
  if (normalizedTarget === normalizedRoot || !normalizedTarget.startsWith(normalizedRoot + path.sep)) fail(label + " escapes its root.");
  return normalizedTarget;
}

function resolveWithin(root, relativePath, label) {
  if (!safeRelativePath(relativePath)) fail(label + " must be a safe relative path.");
  return assertWithin(root, path.resolve(root, ...relativePath.split("/")), label);
}

function validateInputs(value) {
  if (!plainObject(value) || value.schema !== INPUT_SCHEMA || value.schemaVersion !== 1) fail("inputs use an unsupported schema.");
  const verapdf = value.verapdf;
  if (!plainObject(verapdf) || verapdf.packId !== "verapdf" || !nonEmptyString(verapdf.version) || !nonEmptyString(verapdf.license)) {
    fail("inputs must define the immutable verapdf pack, version, and license.");
  }
  safeSegment(verapdf.packId, "pack id");
  safeSegment(verapdf.version, "pack version");
  assertPinnedSource(verapdf.installer, "verapdf.installer");
  for (const field of ["installerJarPath", "cliJarPath"]) {
    if (!safeRelativePath(verapdf.installer[field])) fail("verapdf.installer." + field + " must be a safe relative path.");
  }
  const jre = verapdf.jre;
  if (!plainObject(jre) || !nonEmptyString(jre.distribution) || !nonEmptyString(jre.version)
    || !/^\d+\.\d+\.\d+$/.test(jre.javaVersion || "") || !nonEmptyString(jre.license) || !plainObject(jre.platforms)) {
    fail("inputs must define one pinned JRE and its expected Java version.");
  }
  if (Object.keys(jre.platforms).some((platform) => !SUPPORTED_PLATFORMS.has(platform))) fail("JRE inputs declare an unsupported platform.");
  for (const platform of SUPPORTED_PLATFORMS) {
    const source = jre.platforms[platform];
    assertPinnedSource(source, "verapdf.jre.platforms." + platform);
    if (!safeRelativePath(source.javaHomePath)) fail("verapdf.jre.platforms." + platform + ".javaHomePath must be a safe relative path.");
  }
  if (!Array.isArray(verapdf.licenseMaterial) || verapdf.licenseMaterial.length !== 2) {
    fail("inputs must pin exactly the veraPDF GPL and MPL notice sources.");
  }
  const noticeNames = new Set();
  for (let index = 0; index < verapdf.licenseMaterial.length; index += 1) {
    const source = verapdf.licenseMaterial[index];
    assertPinnedSource(source, "verapdf.licenseMaterial[" + index + "]");
    if (!nonEmptyString(source.name) || noticeNames.has(source.name)) fail("license material names must be non-empty and unique.");
    noticeNames.add(source.name);
  }
  if (!noticeNames.has("veraPDF GPL-3.0-or-later") || !noticeNames.has("veraPDF MPL-2.0")) {
    fail("inputs must pin the veraPDF GPL and MPL notices.");
  }
  return value;
}

async function loadInputs(inputPath) {
  const absolute = path.resolve(inputPath);
  const stat = await fs.lstat(absolute);
  if (!stat.isFile() || stat.isSymbolicLink() || stat.size > MAX_INPUT_BYTES) fail("inputs must be a bounded regular non-symlink file.");
  const bytes = await fs.readFile(absolute);
  let value;
  try {
    value = JSON.parse(bytes);
  } catch (error) {
    fail("inputs are not valid JSON: " + error.message);
  }
  return { absolute, bytes, value: validateInputs(value) };
}

function parseArguments(argv) {
  const values = {};
  const flags = new Set();
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail("unexpected argument " + token + ".");
    const name = token.slice(2);
    if (name === "verify-lock") {
      if (flags.has(name)) fail("--" + name + " may be supplied only once.");
      flags.add(name);
      continue;
    }
    const value = argv[index + 1];
    if (!value || value.startsWith("--")) fail("--" + name + " requires a value.");
    if (Object.hasOwn(values, name)) fail("--" + name + " may be supplied only once.");
    values[name] = value;
    index += 1;
  }
  const verifyLock = flags.has("verify-lock");
  if (!verifyLock) {
    for (const required of ["pack", "version", "platform", "output", "source-url"]) {
      if (!nonEmptyString(values[required])) fail("--" + required + " is required.");
    }
    safeSegment(values.pack, "pack");
    safeSegment(values.version, "version");
    if (!SUPPORTED_PLATFORMS.has(values.platform)) fail("platform must be one of " + [...SUPPORTED_PLATFORMS].join(", ") + ".");
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

function isRedirect(status) {
  return status === 301 || status === 302 || status === 303 || status === 307 || status === 308;
}

async function downloadPinned(source, destination, label, total) {
  let current = httpsUrl(source.url, label + ".url").href;
  let response;
  for (let redirects = 0; redirects <= MAX_REDIRECTS; redirects += 1) {
    response = await fetch(current, { redirect: "manual" });
    if (!isRedirect(response.status)) break;
    if (redirects === MAX_REDIRECTS) fail(label + " exceeded " + MAX_REDIRECTS + " HTTPS redirects.");
    const location = response.headers.get("location");
    if (!nonEmptyString(location)) fail(label + " redirected without a location.");
    const next = new URL(location, current);
    if (next.protocol !== "https:" || next.username || next.password) fail(label + " redirect must use credential-free HTTPS.");
    current = next.href;
  }
  if (!response?.ok || !response.body) fail(label + " download failed with HTTP " + (response?.status ?? "unknown") + ".");
  const handle = await fs.open(destination, "wx", 0o600);
  const digest = crypto.createHash("sha256");
  let downloaded = 0;
  try {
    for await (const chunk of response.body) {
      const bytes = Buffer.from(chunk);
      downloaded += bytes.length;
      total.value += bytes.length;
      if (downloaded > source.downloadBytes || downloaded > MAX_SOURCE_BYTES || total.value > MAX_TOTAL_SOURCE_BYTES) {
        fail(label + " exceeds its pinned source budget.");
      }
      digest.update(bytes);
      await handle.write(bytes);
    }
  } finally {
    await handle.close();
  }
  if (downloaded !== source.downloadBytes) fail(label + " download size does not match the pin.");
  if (digest.digest("hex") !== source.sha256.toLowerCase()) fail(label + " SHA-256 does not match the pin.");
  return { bytes: downloaded, sourceUrl: current };
}

async function requireRegularFile(target, label, { executable = false } = {}) {
  const stat = await fs.lstat(target);
  if (!stat.isFile() || stat.isSymbolicLink()) fail(label + " must be a regular non-symlink file: " + target + ".");
  if (executable && (stat.mode & 0o111) === 0) fail(label + " must be executable: " + target + ".");
  return stat;
}

async function requireDirectory(target, label) {
  const stat = await fs.lstat(target);
  if (!stat.isDirectory() || stat.isSymbolicLink()) fail(label + " must be a real directory: " + target + ".");
  return stat;
}

async function copyDereferencedTree(source, destination, rootReal, visiting = new Set()) {
  const link = await fs.lstat(source);
  let actual = source;
  let stat = link;
  if (link.isSymbolicLink()) {
    actual = await fs.realpath(source).catch(() => fail("source contains a dangling symlink at " + source + "."));
    assertWithin(rootReal, actual, "source symlink");
    stat = await fs.lstat(actual);
  }
  const real = await fs.realpath(actual);
  if (stat.isDirectory()) {
    if (visiting.has(real)) fail("source contains a symlink directory cycle at " + source + ".");
    visiting.add(real);
    try {
      await fs.mkdir(destination, { recursive: true, mode: 0o755 });
      const children = await fs.readdir(actual, { withFileTypes: true });
      children.sort((left, right) => left.name.localeCompare(right.name, "en"));
      for (const child of children) {
        if (child.name === "." || child.name === "..") fail("source contains an unsafe directory entry.");
        await copyDereferencedTree(path.join(actual, child.name), path.join(destination, child.name), rootReal, visiting);
      }
    } finally {
      visiting.delete(real);
    }
    return;
  }
  if (!stat.isFile()) fail("source contains unsupported filesystem entry " + source + ".");
  await fs.mkdir(path.dirname(destination), { recursive: true, mode: 0o755 });
  await fs.copyFile(actual, destination, fsConstants.COPYFILE_EXCL);
  await fs.chmod(destination, stat.mode & 0o777);
}

async function execute(file, arguments_, options = {}) {
  try {
    return await execFile(file, arguments_, options);
  } catch (error) {
    const stderr = typeof error?.stderr === "string" ? error.stderr.trim() : "";
    const detail = stderr ? ": " + stderr.slice(0, 2048) : "";
    fail("command failed " + file + " " + arguments_.join(" ") + detail);
  }
}

async function safeJavaHome(temporary, label) {
  const home = path.join(temporary, label);
  await fs.mkdir(path.join(home, "tmp"), { recursive: true, mode: 0o700 });
  return home;
}

function safeJavaEnvironment(javaHome, home) {
  // The source installer and the final pack must be proved with the pinned
  // JRE, not a Java or dynamic-library choice inherited from the build host.
  // Keep only values the launcher actually needs; no proxy, classpath, or
  // loader-injection variable crosses this boundary.
  return {
    HOME: home,
    JAVA_HOME: javaHome,
    PATH: path.join(javaHome, "bin") + ":/usr/bin:/bin",
    TMPDIR: path.join(home, "tmp"),
  };
}

function launcherText() {
  return [
    "#!/bin/sh",
    "set -eu",
    "ROOT=$(CDPATH= cd \"$(dirname \"$0\")/..\" && pwd -P)",
    "export JAVA_HOME=\"$ROOT/jre\"",
    "exec env -i HOME=\"${HOME:-$ROOT}\" JAVA_HOME=\"$JAVA_HOME\" PATH=\"$JAVA_HOME/bin:/usr/bin:/bin\" \"$ROOT/libexec/verapdf\" \"$@\"",
    "",
  ].join("\n");
}

async function renderInstallerConfig(temporary) {
  const bytes = await fs.readFile(INSTALL_TEMPLATE);
  const text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  const marker = "<installpath>/tmp/open-office-verapdf</installpath>";
  if (text.split(marker).length !== 2) fail("the checked-in veraPDF installer template must contain exactly one install path marker.");
  const destination = path.join(temporary, "installation");
  const rendered = text.replace(marker, "<installpath>" + destination + "</installpath>");
  const target = path.join(temporary, "auto-install.xml");
  await fs.writeFile(target, rendered, { mode: 0o600 });
  return { target, destination };
}

async function verifyJre(javaHome, expectedVersion, temporary) {
  const java = path.join(javaHome, "bin", "java");
  await requireRegularFile(java, "JRE java executable", { executable: true });
  const home = await safeJavaHome(temporary, "jre-probe-home");
  const result = await execute(java, ["-version"], {
    timeout: 30_000,
    maxBuffer: 128 * 1024,
    env: safeJavaEnvironment(javaHome, home),
  });
  const output = String(result.stdout || "") + "\n" + String(result.stderr || "");
  if (!new RegExp("version \\\"" + expectedVersion.replace(/[.*+?^\x24{}()|[\]\\]/g, "\\$&") + "\\\"").test(output)) {
    fail("JRE did not report the pinned Java version " + expectedVersion + ".");
  }
}

async function verifyPayload(payload, version, temporary) {
  const launcher = path.join(payload, "bin", "verapdf");
  await requireRegularFile(launcher, "pack launcher", { executable: true });
  await requireRegularFile(path.join(payload, "jre", "bin", "java"), "pack JRE", { executable: true });
  const environment = {
    HOME: path.join(temporary, "consumer-home"),
    PATH: path.join(payload, "bin") + ":/usr/bin:/bin",
    NO_COLOR: "1",
  };
  await fs.mkdir(environment.HOME, { recursive: true, mode: 0o700 });
  const versionResult = await execute(launcher, ["--version"], {
    timeout: 30_000,
    maxBuffer: 128 * 1024,
    env: environment,
  });
  const plainVersion = version.replace(/-oat\.\d+$/, "");
  if (!new RegExp("veraPDF " + plainVersion.replace(/[.*+?^\x24{}()|[\]\\]/g, "\\$&")).test(String(versionResult.stdout || ""))) {
    fail("pack launcher did not report the expected veraPDF version.");
  }
  const profiles = await execute(launcher, ["--list"], {
    timeout: 30_000,
    maxBuffer: 256 * 1024,
    env: environment,
  });
  if (!/(?:^|\n)\s*ua2\s+-/m.test(String(profiles.stdout || ""))) fail("pack launcher does not expose the required built-in ua2 profile.");
}

function noticesFor(verapdf, jreSource, licenseBytes) {
  const parts = [
    "# veraPDF PDF capability-pack notices",
    "",
    "This immutable pack contains only the veraPDF CLI installation and a hash-pinned Eclipse Temurin JRE.",
    "",
    "## veraPDF Greenfield CLI " + verapdf.version.replace(/-oat\.\d+$/, ""),
    "- installer: " + verapdf.installer.url,
    "- SHA-256: " + verapdf.installer.sha256,
    "- selected payload: CLI only; GUI, samples, and uninstaller are not packaged.",
  ];
  for (let index = 0; index < verapdf.licenseMaterial.length; index += 1) {
    const source = verapdf.licenseMaterial[index];
    const text = new TextDecoder("utf-8", { fatal: true }).decode(licenseBytes[index]).trimEnd();
    parts.push("", "### " + source.name, "", text);
  }
  parts.push(
    "",
    "## " + verapdf.jre.distribution + " JRE " + verapdf.jre.version,
    "- source: " + jreSource.url,
    "- SHA-256: " + jreSource.sha256,
    "- license expression: " + verapdf.jre.license,
    "- complete per-module JRE legal material is retained in jre/legal/ inside this pack.",
    "",
  );
  return Buffer.from(parts.join("\n"), "utf8");
}

async function build(options, loaded) {
  const verapdf = loaded.value.verapdf;
  if (options.pack !== verapdf.packId) fail("inputs do not define pack " + options.pack + ".");
  if (options.version !== verapdf.version) fail("requested version does not match the pinned veraPDF pack version.");
  const jreSource = verapdf.jre.platforms[options.platform];
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-verapdf-provider-pack-"));
  try {
    const total = { value: 0 };
    const installerArchive = path.join(temporary, "verapdf-installer.zip");
    const jreArchive = path.join(temporary, "jre.tar.gz");
    await downloadPinned(verapdf.installer, installerArchive, "veraPDF installer", total);
    await downloadPinned(jreSource, jreArchive, "managed JRE", total);

    const licenseBytes = [];
    for (let index = 0; index < verapdf.licenseMaterial.length; index += 1) {
      const source = verapdf.licenseMaterial[index];
      const destination = path.join(temporary, "license-" + index + ".txt");
      await downloadPinned(source, destination, source.name, total);
      licenseBytes.push(await fs.readFile(destination));
    }

    const jreExtract = path.join(temporary, "jre-extract");
    const installerExtract = path.join(temporary, "installer-extract");
    await fs.mkdir(jreExtract, { mode: 0o700 });
    await fs.mkdir(installerExtract, { mode: 0o700 });
    await execute("tar", ["-xzf", jreArchive, "-C", jreExtract], { timeout: 120_000, maxBuffer: 128 * 1024 });
    await execute("unzip", ["-q", installerArchive, "-d", installerExtract], { timeout: 120_000, maxBuffer: 128 * 1024 });

    const jreHome = resolveWithin(jreExtract, jreSource.javaHomePath, "JRE home");
    await requireDirectory(jreHome, "JRE home");
    await verifyJre(jreHome, verapdf.jre.javaVersion, temporary);
    const installerJar = resolveWithin(installerExtract, verapdf.installer.installerJarPath, "veraPDF installer JAR");
    await requireRegularFile(installerJar, "veraPDF installer JAR");
    const config = await renderInstallerConfig(temporary);
    await fs.mkdir(config.destination, { mode: 0o700 });
    const installerHome = await safeJavaHome(temporary, "installer-home");
    await execute(path.join(jreHome, "bin", "java"), ["-Djava.awt.headless=true", "-jar", installerJar, config.target], {
      timeout: 120_000,
      maxBuffer: 512 * 1024,
      env: safeJavaEnvironment(jreHome, installerHome),
    });

    await requireDirectory(config.destination, "veraPDF CLI installation");
    const installedLauncher = resolveWithin(config.destination, "verapdf", "veraPDF launcher");
    const installedCli = resolveWithin(config.destination, verapdf.installer.cliJarPath, "veraPDF CLI JAR");
    await requireRegularFile(installedLauncher, "veraPDF launcher", { executable: true });
    await requireRegularFile(installedCli, "veraPDF CLI JAR");

    const payload = path.join(temporary, "payload");
    await fs.mkdir(payload, { mode: 0o700 });
    const jreExtractReal = await fs.realpath(jreExtract);
    const installReal = await fs.realpath(config.destination);
    await copyDereferencedTree(jreHome, path.join(payload, "jre"), jreExtractReal);
    await copyDereferencedTree(installedLauncher, path.join(payload, "libexec", "verapdf"), installReal);
    await copyDereferencedTree(path.dirname(installedCli), path.join(payload, "libexec", "bin"), installReal);
    const launcher = path.join(payload, "bin", "verapdf");
    await fs.mkdir(path.dirname(launcher), { recursive: true, mode: 0o755 });
    await fs.writeFile(launcher, launcherText(), { mode: 0o755 });
    await fs.chmod(launcher, 0o755);
    await verifyPayload(payload, verapdf.version, temporary);

    const notices = path.join(temporary, "THIRD_PARTY_NOTICES.md");
    await fs.writeFile(notices, noticesFor(verapdf, jreSource, licenseBytes), { mode: 0o600 });
    const result = await execute(process.execPath, [
      PACK_BUILDER,
      "--pack", options.pack,
      "--version", options.version,
      "--platform", options.platform,
      "--payload", payload,
      "--output", options.output,
      "--source-url", options.sourceUrl,
      "--source-sha256", sha256(loaded.bytes),
      "--license", verapdf.license,
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
    const verapdf = loaded.value.verapdf;
    process.stdout.write(JSON.stringify({
      schema: loaded.value.schema,
      sha256: sha256(loaded.bytes),
      pack: verapdf.packId,
      version: verapdf.version,
      platforms: Object.fromEntries([...SUPPORTED_PLATFORMS].map((platform) => [platform, {
        jreVersion: verapdf.jre.version,
        downloadBytes: verapdf.installer.downloadBytes + verapdf.jre.platforms[platform].downloadBytes,
      }])),
    }, null, 2) + "\n");
    return;
  }
  await build(options, loaded);
}

main().catch((error) => {
  process.stderr.write(String(error?.stack || error) + "\n");
  process.exitCode = 2;
});
