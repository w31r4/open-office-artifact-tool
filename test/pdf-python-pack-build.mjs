import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { spawnSync } from "node:child_process";

const root = path.resolve(import.meta.dirname, "..");
const buildScript = path.join(root, "scripts", "build-python-provider-pack.mjs");
const inputPath = path.join(root, "scripts", "pdf-provider-python-release-inputs.v1.json");

function run(arguments_, { expect = 0 } = {}) {
  const result = spawnSync(process.execPath, [buildScript, ...arguments_], { cwd: root, encoding: "utf8" });
  assert.equal(result.status, expect, result.stderr || result.stdout);
  return result;
}

const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-python-pack-build-"));
try {
  const bytes = await fs.readFile(inputPath);
  const source = JSON.parse(bytes);
  const verified = JSON.parse(run(["--verify-lock"]).stdout);
  assert.equal(verified.schema, source.schema);
  assert.equal(verified.sha256, crypto.createHash("sha256").update(bytes).digest("hex"));
  assert.deepEqual(verified.packs, {
    "python-foundation": { "darwin-arm64": 10, "linux-x64": 10 },
    "python-specialists": { "darwin-arm64": 20, "linux-x64": 20 },
  });

  async function rejectMutation(label, mutate, expected) {
    const candidate = structuredClone(source);
    mutate(candidate);
    const candidatePath = path.join(temporary, `${label}.json`);
    await fs.writeFile(candidatePath, `${JSON.stringify(candidate, null, 2)}\n`, "utf8");
    const rejected = run(["--verify-lock", "--inputs", candidatePath], { expect: 2 });
    assert.match(rejected.stderr, expected);
  }

  await rejectMutation("duplicate-platform-wheel", (candidate) => {
    candidate.packs["python-foundation"].platformWheels["darwin-arm64"].push(
      structuredClone(candidate.packs["python-foundation"].commonWheels[0]),
    );
  }, /duplicate darwin-arm64 wheel reportlab@4\.4\.9/);
  await rejectMutation("missing-direct-requirement", (candidate) => {
    candidate.packs["python-foundation"].directRequirements.pypdf = "0.0.0";
  }, /direct requirement pypdf==0\.0\.0 is absent from its darwin-arm64 wheel lock/);
  await rejectMutation("unsafe-source-url", (candidate) => {
    candidate.pythonRuntime.platforms["darwin-arm64"].url = "http://example.test/python.tar.gz";
  }, /credential-free HTTPS URL/);
  await rejectMutation("unsupported-wheel-platform", (candidate) => {
    candidate.packs["python-foundation"].platformWheels["win32-x64"] = [];
  }, /unsupported platform/);
} finally {
  await fs.rm(temporary, { recursive: true, force: true });
}

console.log("python PDF provider pack build smoke ok");
