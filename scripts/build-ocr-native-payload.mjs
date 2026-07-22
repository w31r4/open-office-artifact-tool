#!/usr/bin/env node
/**
 * Assemble the native half of the OCR capability pack around an already safe,
 * extracted isolated-Python payload. This is release tooling only: customer
 * installation always consumes the resulting hash-pinned USTAR archive.
 *
 * The script deliberately dereferences every build-machine symlink, copies a
 * closed native library set, writes thin relocatable launchers, and rejects
 * unresolved Homebrew/Linux library references. It never runs a package
 * manager or downloads a runtime itself.
 */

import crypto from "node:crypto";
import { execFile as execFileCallback } from "node:child_process";
import { constants as fsConstants } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";
import { promisify } from "node:util";

const execFile = promisify(execFileCallback);
const SUPPORTED_PLATFORMS = new Set(["darwin-arm64", "linux-x64"]);
const MAX_OUTPUT = 512 * 1024;
// `otool` can report a non-Mach-O input without a failing process status on
// some macOS releases. Never let that turn a Python source file or launcher
// into an `install_name_tool` target. These are every 32/64-bit thin and fat
// Mach-O magic value, read in big-endian byte order.
const MACHO_MAGICS = new Set([
  0xfeedface,
  0xcefaedfe,
  0xfeedfacf,
  0xcffaedfe,
  0xcafebabe,
  0xbebafeca,
  0xcafebabf,
  0xbfbafeca,
]);
const ELF_MAGIC = Buffer.from([0x7f, 0x45, 0x4c, 0x46]);

function fail(message) {
  throw new Error(`OCR native capability-pack build: ${message}`);
}

function nonEmptyString(value) {
  return typeof value === "string" && value.trim().length > 0;
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function parseArguments(argv) {
  const values = {};
  const repeated = { "library-root": [], "resource-root": [] };
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token.startsWith("--")) fail(`unexpected argument ${token}.`);
    const name = token.slice(2);
    const value = argv[index + 1];
    if (!nonEmptyString(value) || value.startsWith("--")) fail(`--${name} requires a value.`);
    if (Object.hasOwn(repeated, name)) {
      repeated[name].push(value);
    } else {
      if (Object.hasOwn(values, name)) fail(`--${name} may be supplied only once.`);
      values[name] = value;
    }
    index += 1;
  }
  for (const required of ["platform", "payload", "notices", "tesseract", "ghostscript", "pdftotext", "ghostscript-root", "tessdata-root"]) {
    if (!nonEmptyString(values[required])) fail(`--${required} is required.`);
  }
  if (!SUPPORTED_PLATFORMS.has(values.platform)) fail(`platform must be one of ${[...SUPPORTED_PLATFORMS].join(", ")}.`);
  if (!repeated["library-root"].length) fail("at least one --library-root is required.");
  return {
    platform: values.platform,
    payload: path.resolve(values.payload),
    notices: path.resolve(values.notices),
    tesseract: path.resolve(values.tesseract),
    ghostscript: path.resolve(values.ghostscript),
    pdftotext: path.resolve(values.pdftotext),
    ghostscriptRoot: path.resolve(values["ghostscript-root"]),
    popplerRoot: values["poppler-root"] ? path.resolve(values["poppler-root"]) : undefined,
    tessdataRoot: path.resolve(values["tessdata-root"]),
    libraryRoots: repeated["library-root"].map((value) => path.resolve(value)),
    resourceRoots: repeated["resource-root"].map((value) => path.resolve(value)),
  };
}

async function realDirectory(value, label) {
  const stat = await fs.lstat(value).catch(() => undefined);
  if (!stat?.isDirectory()) fail(`${label} must be an existing directory.`);
  return fs.realpath(value);
}

async function realFile(value, label) {
  const stat = await fs.lstat(value).catch(() => undefined);
  if (!stat?.isFile() && !stat?.isSymbolicLink()) fail(`${label} must be an existing executable file.`);
  const actual = await fs.realpath(value);
  const actualStat = await fs.lstat(actual);
  if (!actualStat.isFile() || (actualStat.mode & 0o111) === 0) fail(`${label} must resolve to an executable regular file.`);
  return actual;
}

function containedPath(roots, target, label) {
  const normalizedTarget = path.resolve(target);
  for (const root of roots) {
    const normalizedRoot = path.resolve(root);
    if (normalizedTarget === normalizedRoot || normalizedTarget.startsWith(`${normalizedRoot}${path.sep}`)) return normalizedTarget;
  }
  fail(`${label} escapes its declared roots: ${normalizedTarget}.`);
}

async function copyFile(source, destination, { executable = false } = {}) {
  const sourceReal = await fs.realpath(source);
  const sourceStat = await fs.lstat(sourceReal);
  if (!sourceStat.isFile()) fail(`native source is not a regular file: ${source}.`);
  await fs.mkdir(path.dirname(destination), { recursive: true, mode: 0o755 });
  await fs.copyFile(sourceReal, destination, fsConstants.COPYFILE_EXCL);
  await fs.chmod(destination, executable || (sourceStat.mode & 0o111) !== 0 ? 0o755 : 0o644);
}

async function treeCopy(source, destination, label, { resourceRoots = [] } = {}) {
  const root = await realDirectory(source, label);
  const trustedRoots = [root, ...await Promise.all(resourceRoots.map((candidate) => realDirectory(candidate, `${label} resource root`)))];
  const visiting = new Set();
  async function copy(current, target) {
    const link = await fs.lstat(current);
    let actual = current;
    let stat = link;
    if (link.isSymbolicLink()) {
      actual = await fs.realpath(current).catch(() => fail(`${label} contains a dangling symlink: ${current}.`));
      containedPath(trustedRoots, actual, `${label} symlink`);
      stat = await fs.lstat(actual);
    }
    if (stat.isDirectory()) {
      const real = await fs.realpath(actual);
      if (visiting.has(real)) fail(`${label} contains a symlink directory cycle: ${current}.`);
      visiting.add(real);
      try {
        await fs.mkdir(target, { recursive: true, mode: 0o755 });
        const children = await fs.readdir(actual, { withFileTypes: true });
        children.sort((left, right) => left.name.localeCompare(right.name, "en"));
        for (const child of children) {
          if (child.name === "." || child.name === "..") fail(`${label} contains an unsafe entry.`);
          await copy(path.join(actual, child.name), path.join(target, child.name));
        }
      } finally {
        visiting.delete(real);
      }
      return;
    }
    if (!stat.isFile()) fail(`${label} contains an unsupported filesystem entry.`);
    await copyFile(actual, target);
  }
  await copy(root, destination);
}

async function removeTraineddata(root) {
  async function walk(directory) {
    const entries = await fs.readdir(directory, { withFileTypes: true });
    for (const entry of entries) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile() && entry.name.endsWith(".traineddata")) await fs.rm(target, { force: true });
    }
  }
  await walk(root);
}

async function findNamedDirectory(root, name, label) {
  const realRoot = await realDirectory(root, label);
  const queue = [realRoot];
  while (queue.length) {
    const current = queue.shift();
    const entries = await fs.readdir(current, { withFileTypes: true });
    for (const entry of entries.sort((left, right) => left.name.localeCompare(right.name, "en"))) {
      const candidate = path.join(current, entry.name);
      if (entry.isDirectory() && entry.name === name) return candidate;
      if (entry.isDirectory()) queue.push(candidate);
    }
  }
  fail(`${label} has no ${name} resource directory.`);
}

async function listRegularFiles(root) {
  const results = [];
  async function walk(directory) {
    for (const entry of (await fs.readdir(directory, { withFileTypes: true })).sort((left, right) => left.name.localeCompare(right.name, "en"))) {
      const target = path.join(directory, entry.name);
      if (entry.isDirectory()) await walk(target);
      else if (entry.isFile()) results.push(target);
    }
  }
  await walk(root);
  return results;
}

async function copyByBasename(source, destinationDirectory, copied) {
  const actual = await fs.realpath(source);
  // Keep the *referenced* basename even when Homebrew or the system runtime
  // expresses it as a symlink. The final customer archive forbids symlinks,
  // so each ABI name must become a regular copy of the verified target.
  const name = path.basename(source);
  const destination = path.join(destinationDirectory, name);
  const existing = copied.get(name);
  if (existing) {
    const [left, right] = await Promise.all([fs.readFile(existing), fs.readFile(actual)]);
    if (sha256(left) !== sha256(right)) fail(`native library basename collision: ${name}.`);
    return destination;
  }
  await copyFile(actual, destination);
  copied.set(name, destination);
  return destination;
}

async function run(command, args, label, { allowFailure = false } = {}) {
  try {
    return await execFile(command, args, { timeout: 60_000, maxBuffer: MAX_OUTPUT, encoding: "utf8" });
  } catch (error) {
    if (allowFailure) return undefined;
    fail(`${label} failed: ${String(error?.stderr || error?.message || error).trim()}`);
  }
}

async function copyMacLibraries(options, libDirectory) {
  const copied = new Map();
  for (const rootPath of options.libraryRoots) {
    const root = await realDirectory(rootPath, "library root");
    for (const candidate of await listMacLibraryFiles(root)) {
      await copyByBasename(candidate, libDirectory, copied);
    }
  }
  if (!copied.size) fail("no macOS native libraries were collected from the declared library roots.");
  return copied;
}

async function listMacLibraryFiles(root) {
  const results = [];
  async function walk(directory) {
    for (const entry of (await fs.readdir(directory, { withFileTypes: true })).sort((left, right) => left.name.localeCompare(right.name, "en"))) {
      const candidate = path.join(directory, entry.name);
      if (entry.isDirectory()) {
        await walk(candidate);
        continue;
      }
      if ((!entry.isFile() && !entry.isSymbolicLink()) || !entry.name.includes(".dylib")) continue;
      const actual = await fs.realpath(candidate).catch(() => fail(`macOS library link is unresolved: ${candidate}.`));
      containedPath([root], actual, "macOS library link");
      const stat = await fs.lstat(actual);
      if (!stat.isFile()) fail(`macOS library link does not resolve to a regular file: ${candidate}.`);
      results.push(candidate);
    }
  }
  await walk(root);
  return results;
}

function parseMacDependencies(output) {
  return output.split("\n").slice(1).map((line) => line.trim().split(" (")[0]).filter(Boolean);
}

async function isMachOFile(target) {
  const stat = await fs.lstat(target).catch(() => undefined);
  if (!stat?.isFile() || stat.isSymbolicLink() || stat.size < 4) return false;
  const handle = await fs.open(target, "r");
  try {
    const header = Buffer.alloc(4);
    const { bytesRead } = await handle.read(header, 0, header.length, 0);
    return bytesRead === header.length && MACHO_MAGICS.has(header.readUInt32BE(0));
  } finally {
    await handle.close();
  }
}

async function isElfFile(target) {
  const stat = await fs.lstat(target).catch(() => undefined);
  if (!stat?.isFile() || stat.isSymbolicLink() || stat.size < ELF_MAGIC.length) return false;
  const handle = await fs.open(target, "r");
  try {
    const header = Buffer.alloc(ELF_MAGIC.length);
    const { bytesRead } = await handle.read(header, 0, header.length, 0);
    return bytesRead === header.length && header.equals(ELF_MAGIC);
  } finally {
    await handle.close();
  }
}

function loaderPath(target, destination) {
  const relative = path.relative(path.dirname(target), destination).split(path.sep).join("/");
  if (!relative || relative === ".") return `@loader_path/${path.basename(destination)}`;
  if (path.isAbsolute(relative) || relative.startsWith("/")) fail(`cannot create a relocatable loader path for ${target}.`);
  return `@loader_path/${relative}`;
}

function linuxRpath(target, libDirectory) {
  const relative = path.relative(path.dirname(target), libDirectory).split(path.sep).join("/");
  if (!relative || relative === ".") return "$ORIGIN";
  if (path.isAbsolute(relative) || relative.startsWith("/")) fail(`cannot create a relocatable Linux rpath for ${target}.`);
  return `$ORIGIN/${relative}`;
}

async function patchMacBinary(target, libraryNames, { library = false, libDirectory } = {}) {
  if (!await isMachOFile(target)) return false;
  const listed = await run("otool", ["-L", target], `otool ${target}`, { allowFailure: true });
  if (!listed) return false;
  for (const dependency of parseMacDependencies(listed.stdout)) {
    const name = path.basename(dependency);
    if (!libraryNames.has(name)) continue;
    if (dependency.startsWith("/System/") || dependency.startsWith("/usr/lib/")) continue;
    const replacement = loaderPath(target, path.join(libDirectory, name));
    if (dependency !== replacement) await run("install_name_tool", ["-change", dependency, replacement, target], `install_name_tool ${target}`);
  }
  // Every consumer is rewritten to resolve siblings through @loader_path.
  // Keep the copied library's own install name on that same basis: a library
  // may otherwise report its own @rpath alias as an unresolved dependency
  // even though no runtime rpath is present in the relocated payload.
  if (library) await run("install_name_tool", ["-id", `@loader_path/${path.basename(target)}`, target], `install_name_tool id ${target}`);
  return true;
}

async function finalizeMacPayload(payload, libraryNames) {
  const libexec = path.join(payload, "libexec");
  const lib = path.join(payload, "lib");
  const targets = [...await listRegularFiles(libexec), ...await listRegularFiles(lib)];
  for (const target of targets) {
    const library = path.dirname(target) === lib && path.basename(target).includes(".dylib");
    await patchMacBinary(target, libraryNames, { library, libDirectory: lib });
  }
  for (const target of targets) {
    if (!await isMachOFile(target)) continue;
    const listed = await run("otool", ["-L", target], `verify otool ${target}`, { allowFailure: true });
    if (!listed) continue;
    for (const dependency of parseMacDependencies(listed.stdout)) {
      const name = path.basename(dependency);
      if (dependency.startsWith("/opt/homebrew/") || dependency.startsWith("/usr/local/")) fail(`relocatable payload still links build-machine path ${dependency}.`);
      if (dependency.startsWith("@rpath/") && libraryNames.has(name)) fail(`relocatable payload ${target} retains unresolved rpath dependency ${dependency}.`);
    }
    await run("codesign", ["--force", "--sign", "-", "--timestamp=none", target], `codesign ${target}`);
  }
}

function parseLddDependencies(output) {
  const result = [];
  for (const line of output.split("\n")) {
    const mapped = /=>\s+(\/[^\s]+)\s+\(/.exec(line);
    if (mapped) {
      result.push(mapped[1]);
      continue;
    }
    const direct = /^\s*(\/[^\s]+)\s+\(/.exec(line);
    if (direct) result.push(direct[1]);
    if (/=>\s+not found/.test(line)) fail(`native Linux dependency is missing: ${line.trim()}`);
  }
  return result;
}

function hostBaselineLibrary(candidate) {
  const name = path.basename(candidate);
  return /^(?:ld-linux|ld-musl|libc\.so|libm\.so|libdl\.so|libpthread\.so|librt\.so|libresolv\.so|libnsl\.so)/.test(name);
}

async function copyLinuxLibraries(initialTargets, libDirectory) {
  // The isolated Python payload already owns libpython in the top-level lib
  // directory. Seed that direct namespace so an ldd edge back to libpython is
  // checked for byte identity rather than attempting a COPYFILE_EXCL copy
  // onto the same destination.
  const copied = new Map();
  for (const entry of await fs.readdir(libDirectory, { withFileTypes: true })) {
    if (!entry.isFile()) continue;
    const candidate = path.join(libDirectory, entry.name);
    const stat = await fs.lstat(candidate);
    if (stat.isSymbolicLink()) fail(`payload library directory contains an unsafe symlink: ${candidate}.`);
    copied.set(entry.name, candidate);
  }
  const queued = [...initialTargets];
  const seen = new Set();
  while (queued.length) {
    const target = await fs.realpath(queued.shift());
    if (seen.has(target)) continue;
    seen.add(target);
    const ldd = await run("ldd", [target], `ldd ${target}`);
    for (const dependency of parseLddDependencies(ldd.stdout)) {
      if (hostBaselineLibrary(dependency)) continue;
      const copiedPath = await copyByBasename(dependency, libDirectory, copied);
      queued.push(copiedPath);
    }
  }
  if (!copied.size) fail("no Linux native libraries were resolved.");
  return copied;
}

async function finalizeLinuxPayload(payload) {
  const libexec = path.join(payload, "libexec");
  const lib = path.join(payload, "lib");
  for (const target of await listRegularFiles(libexec)) {
    if (!await isElfFile(target)) continue;
    await run("patchelf", ["--set-rpath", linuxRpath(target, lib), target], `patchelf ${target}`);
  }
  for (const target of await listRegularFiles(lib)) {
    if (!await isElfFile(target)) continue;
    await run("patchelf", ["--set-rpath", linuxRpath(target, lib), target], `patchelf ${target}`);
  }
}

async function writeText(target, value, { executable = false } = {}) {
  await fs.writeFile(target, value, { encoding: "utf8", flag: "wx", mode: executable ? 0o755 : 0o644 });
}

function wrapperFor(name, platform) {
  const root = "ROOT=$(CDPATH= cd -- \"$(dirname -- \"$0\")/..\" && pwd -P)";
  const dynamic = platform === "darwin-arm64"
    ? "export DYLD_FALLBACK_LIBRARY_PATH=\"$ROOT/lib${DYLD_FALLBACK_LIBRARY_PATH:+:$DYLD_FALLBACK_LIBRARY_PATH}\""
    : "export LD_LIBRARY_PATH=\"$ROOT/lib${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}\"";
  const gs = name === "gs" ? "export GS_LIB=\"$ROOT/share/ghostscript/Resource:$ROOT/share/ghostscript/lib\"\n" : "";
  return `#!/bin/sh\nset -eu\n${root}\n${dynamic}\n${gs}exec \"$ROOT/libexec/${name}\" \"$@\"\n`;
}

async function writeLaunchers(payload, platform) {
  const bin = path.join(payload, "bin");
  const python = path.join(bin, "python3");
  const pythonStat = await fs.lstat(python).catch(() => undefined);
  if (!pythonStat?.isFile() || pythonStat.isSymbolicLink() || (pythonStat.mode & 0o111) === 0) fail("OCR payload is missing safe isolated bin/python3.");
  await fs.mkdir(bin, { recursive: true, mode: 0o755 });
  await writeText(path.join(bin, "ocrmypdf"), "#!/bin/sh\nset -eu\nROOT=$(CDPATH= cd -- \"$(dirname -- \"$0\")/..\" && pwd -P)\nexec \"$ROOT/bin/python3\" -I -m ocrmypdf \"$@\"\n", { executable: true });
  for (const name of ["tesseract", "gs", "pdftotext"]) await writeText(path.join(bin, name), wrapperFor(name, platform), { executable: true });
  if (platform === "darwin-arm64" && process.platform !== "darwin") fail("darwin payload must be assembled on darwin.");
  if (platform === "linux-x64" && process.platform !== "linux") fail("linux payload must be assembled on linux.");
}

async function captureVersion(command, args) {
  const result = await run(command, args, `version probe ${command}`);
  return `${result.stdout}${result.stderr}`.trim().slice(0, 4096);
}

async function build(options) {
  const payload = await realDirectory(options.payload, "payload");
  const libexec = path.join(payload, "libexec");
  const lib = path.join(payload, "lib");
  const share = path.join(payload, "share");
  await Promise.all([fs.mkdir(libexec, { recursive: true, mode: 0o755 }), fs.mkdir(lib, { recursive: true, mode: 0o755 }), fs.mkdir(share, { recursive: true, mode: 0o755 })]);
  const [tesseract, ghostscript, pdftotext] = await Promise.all([
    realFile(options.tesseract, "tesseract"),
    realFile(options.ghostscript, "ghostscript"),
    realFile(options.pdftotext, "pdftotext"),
  ]);
  await copyFile(tesseract, path.join(libexec, "tesseract"), { executable: true });
  await copyFile(ghostscript, path.join(libexec, "gs"), { executable: true });
  await copyFile(pdftotext, path.join(libexec, "pdftotext"), { executable: true });

  const resource = path.join(share, "ghostscript");
  await fs.mkdir(resource, { recursive: true, mode: 0o755 });
  await treeCopy(
    await findNamedDirectory(options.ghostscriptRoot, "Resource", "ghostscript root"),
    path.join(resource, "Resource"),
    "ghostscript Resource",
    { resourceRoots: options.resourceRoots },
  );
  await treeCopy(await findNamedDirectory(options.ghostscriptRoot, "lib", "ghostscript root"), path.join(resource, "lib"), "ghostscript lib");
  await treeCopy(options.tessdataRoot, path.join(share, "tessdata"), "tessdata root");
  await removeTraineddata(path.join(share, "tessdata"));
  if (options.popplerRoot) await treeCopy(options.popplerRoot, path.join(share, "poppler"), "poppler resource root");

  let copied;
  if (options.platform === "darwin-arm64") {
    copied = await copyMacLibraries(options, lib);
    await finalizeMacPayload(payload, new Set(copied.keys()));
  } else {
    const pythonNativeFiles = [];
    for (const target of await listRegularFiles(lib)) {
      if (await isElfFile(target)) pythonNativeFiles.push(target);
    }
    copied = await copyLinuxLibraries([
      path.join(libexec, "tesseract"),
      path.join(libexec, "gs"),
      path.join(libexec, "pdftotext"),
      ...pythonNativeFiles,
    ], lib);
    await finalizeLinuxPayload(payload);
  }
  await writeLaunchers(payload, options.platform);

  const versions = {
    tesseract: await captureVersion(path.join(payload, "bin", "tesseract"), ["--version"]),
    ghostscript: await captureVersion(path.join(payload, "bin", "gs"), ["--version"]),
    pdftotext: await captureVersion(path.join(payload, "bin", "pdftotext"), ["-v"]),
    ocrmypdf: await captureVersion(path.join(payload, "bin", "ocrmypdf"), ["--version"]),
  };
  const tessdata = path.join(payload, "share", "tessdata");
  const leakedLanguage = (await listRegularFiles(tessdata)).find((file) => file.endsWith(".traineddata"));
  if (leakedLanguage) fail(`OCR core must not include a language data file: ${leakedLanguage}.`);
  const notice = [
    "# OCR native capability-pack build evidence",
    "",
    "The core payload contains isolated OCRmyPDF/Python plus relocatable Tesseract, Ghostscript, and Poppler pdftotext sidecars. Tesseract language data is intentionally excluded and is supplied only by separately verified language packs.",
    "",
    "## Runtime probes",
    ...Object.entries(versions).flatMap(([name, value]) => [`### ${name}`, "", value || "<no version output>", ""]),
    "## Native library closure",
    ...[...copied.keys()].sort((left, right) => left.localeCompare(right, "en")).map((name) => `- ${name}`),
    "",
  ].join("\n");
  await writeText(options.notices, notice);
  process.stdout.write(`${JSON.stringify({ platform: options.platform, payload, nativeLibraries: copied.size, versions }, null, 2)}\n`);
}

async function main() {
  const options = parseArguments(process.argv.slice(2));
  await build(options);
}

main().catch((error) => {
  process.stderr.write(`${error?.stack || error}\n`);
  process.exitCode = 2;
});
