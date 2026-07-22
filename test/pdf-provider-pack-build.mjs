import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

import { safeExtractTarGz } from "../src/pdf/providers/installer.mjs";

const root = path.resolve(import.meta.dirname, "..");
const buildScript = path.join(root, "scripts", "build-pdf-provider-pack.mjs");
const finalizeScript = path.join(root, "scripts", "finalize-pdf-provider-release.mjs");

function run(arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [buildScript, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

function finalize(arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [finalizeScript, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-pdf-pack-build-"));
try {
  const payload = path.join(temporary, "payload");
  const outputA = path.join(temporary, "output-a");
  const outputB = path.join(temporary, "output-b");
  const outputLinux = path.join(temporary, "output-linux");
  const release = path.join(temporary, "release");
  const unpacked = path.join(temporary, "unpacked");
  const notices = path.join(temporary, "notices.md");
  await fs.mkdir(path.join(payload, "bin"), { recursive: true });
  await fs.mkdir(path.join(payload, "share", "data"), { recursive: true });
  await fs.writeFile(path.join(payload, "bin", "tool"), "#!/bin/sh\necho capability-pack\n", { mode: 0o755 });
  await fs.writeFile(path.join(payload, "share", "data", "fixture.txt"), "fixture\n", { mode: 0o644 });
  await fs.writeFile(notices, "fixture notices\n", "utf8");
  const arguments_ = [
    "--pack", "fixture-pack",
    "--version", "1.2.3",
    "--platform", "darwin-arm64",
    "--payload", payload,
    "--source-url", "https://releases.example.test/fixture-1.2.3.tar.gz",
    "--source-sha256", "a".repeat(64),
    "--license", "Apache-2.0",
    "--notices", notices,
  ];
  const first = JSON.parse(run([...arguments_, "--output", outputA]).stdout);
  const second = JSON.parse(run([...arguments_, "--output", outputB]).stdout);
  assert.equal(first.schema, "open-office-artifact-tool.pdf-provider-pack.v1");
  assert.deepEqual(first, second, "same payload and provenance must create identical metadata");
  const archiveA = await fs.readFile(path.join(outputA, first.artifact.asset));
  const archiveB = await fs.readFile(path.join(outputB, second.artifact.asset));
  assert.deepEqual(archiveA, archiveB, "capability pack archive must be deterministic");
  assert.equal(crypto.createHash("sha256").update(archiveA).digest("hex"), first.artifact.sha256);
  await fs.mkdir(unpacked);
  const extraction = await safeExtractTarGz(archiveA, unpacked, first.artifact.unpackedBytes);
  assert.ok(extraction.entries.includes("bin/tool"));
  assert.equal((await fs.stat(path.join(unpacked, "bin", "tool"))).mode & 0o111, 0o111);
  assert.equal(await fs.readFile(path.join(unpacked, "share", "data", "fixture.txt"), "utf8"), "fixture\n");
  assert.match(await fs.readFile(path.join(unpacked, "THIRD_PARTY_NOTICES.md"), "utf8"), /fixture notices/);
  assert.equal(JSON.parse(await fs.readFile(path.join(unpacked, "sbom.cdx.json"), "utf8")).bomFormat, "CycloneDX");

  const linuxArguments = [...arguments_];
  linuxArguments[linuxArguments.indexOf("darwin-arm64")] = "linux-x64";
  const linux = JSON.parse(run([...linuxArguments, "--output", outputLinux]).stdout);
  for (const file of await fs.readdir(outputLinux)) await fs.copyFile(path.join(outputLinux, file), path.join(outputA, file));
  const finalized = JSON.parse(finalize([
    "--pack", "fixture-pack", "--version", "1.2.3", "--input", outputA, "--output", release,
    "--release-base-url", "https://github.com/example/project/releases/download/pdf-provider-fixture-pack-1.2.3/",
    "--repository", "example/project", "--workflow", ".github/workflows/pdf-capability-packs.yml",
  ]).stdout);
  assert.equal(finalized.schema, "open-office-artifact-tool.pdf-provider-release.v1");
  assert.equal(finalized.catalogFragment.state, "published");
  assert.deepEqual(finalized.catalogFragment.artifacts.map((artifact) => artifact.platform).sort(), ["darwin-arm64", "linux-x64"]);
  assert.equal(finalized.catalogFragment.artifacts.find((artifact) => artifact.platform === "linux-x64").sha256, linux.artifact.sha256);
  assert.equal(JSON.parse(await fs.readFile(path.join(release, "fixture-pack-1.2.3.sbom.cdx.json"), "utf8")).bomFormat, "CycloneDX");

  const malicious = path.join(temporary, "malicious");
  await fs.mkdir(malicious);
  await fs.symlink(path.join(payload, "bin", "tool"), path.join(malicious, "tool"));
  const rejected = run([
    "--pack", "fixture-pack", "--version", "1.2.3", "--platform", "darwin-arm64",
    "--payload", malicious, "--output", path.join(temporary, "malicious-output"),
    "--source-url", "https://releases.example.test/fixture-1.2.3.tar.gz", "--source-sha256", "a".repeat(64), "--license", "Apache-2.0", "--notices", notices,
  ], { expect: 2 });
  assert.match(rejected.stderr, /symlink/);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("pdf provider pack build smoke ok");
