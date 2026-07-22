import assert from "node:assert/strict";
import crypto from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";


const repoRoot = path.resolve(import.meta.dirname, "..");
const skillRoot = path.join(repoRoot, "skills", "pdf", "skills", "pdf");
const provider = path.join(skillRoot, "scripts", "verapdf_provider.py");
const registry = path.join(skillRoot, "scripts", "pdf_provider.py");
const fixture = path.join(repoRoot, "test", "fixtures", "pdf", "verapdf-pdfa1b-pass.pdf");
const fixtureHash = "66077f449d472a048e3bbf7192aa6d2b0b0ebd6b6d8a6f878f776f69424b6deb";
const python = "python3";


function run(executable, args, options = {}) {
  const result = spawnSync(executable, args, {
    cwd: options.cwd || repoRoot,
    encoding: "utf8",
    env: { ...process.env, PYTHONDONTWRITEBYTECODE: "1", ...options.env },
    maxBuffer: 24 * 1024 * 1024,
  });
  if (options.status !== undefined) {
    assert.equal(
      result.status,
      options.status,
      `${executable} ${args.join(" ")}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`,
    );
  }
  return result;
}


function jsonResult(result, stream = "stdout") {
  const value = result[stream]?.trim();
  assert.ok(value, `expected JSON on ${stream}\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return JSON.parse(value);
}


function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}


const manifest = (await fs.readFile(path.join(skillRoot, "manifest.txt"), "utf8")).split(/\r?\n/).filter(Boolean);
assert.ok(manifest.includes("scripts/verapdf_provider.py"));
const skillText = await fs.readFile(path.join(skillRoot, "SKILL.md"), "utf8");
assert.match(skillText, /verapdf_provider\.py/);
assert.match(skillText, /human review/i);
const accessibilityText = await fs.readFile(path.join(skillRoot, "tasks", "accessibility.md"), "utf8");
assert.match(accessibilityText, /--expected-sha256/);
assert.match(accessibilityText, /--flavour/);
assert.match(accessibilityText, /--require-compliant/);
const fixtureBytes = await fs.readFile(fixture);
assert.equal(sha256(fixtureBytes), fixtureHash, "the attributed veraPDF corpus fixture must remain byte-identical");

const tempRoot = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-verapdf-provider-"));
try {
  const source = path.join(tempRoot, "source.pdf");
  await fs.writeFile(source, fixtureBytes);
  const fakeProvider = path.join(tempRoot, "fake-verapdf.mjs");
  await fs.writeFile(fakeProvider, String.raw`#!/usr/bin/env node
import fs from "node:fs";

const args = process.argv.slice(2);
if (args.length === 1 && args[0] === "-v") process.exit(2);
if (args.includes("--version")) {
  console.log("veraPDF " + (process.env.FAKE_VERAPDF_VERSION || "1.30.2"));
  process.exit(0);
}
if (args.includes("--list")) {
  const profiles = process.env.FAKE_VERAPDF_PROFILES || "1a 1b 2a 2b 2u 3a 3b 3u 4 4e 4f ua1 ua2 wt1r wt1a";
  for (const profile of profiles.split(/\s+/)) console.log("  " + profile + " - fake profile");
  process.exit(0);
}
if (process.env.FAKE_VERAPDF_HANG === "1") Atomics.wait(new Int32Array(new SharedArrayBuffer(4)), 0, 0);
if (process.env.FAKE_VERAPDF_BIG === "1") {
  process.stdout.write("x".repeat(1024 * 1024));
  process.exit(0);
}
if (process.env.FAKE_VERAPDF_BAD_JSON === "1") {
  console.log("not-json");
  process.exit(0);
}
const input = args.at(-1);
if (!args.includes("--format") || args[args.indexOf("--format") + 1] !== "json") process.exit(30);
if (!args.includes("--loglevel") || args[args.indexOf("--loglevel") + 1] !== "0") process.exit(31);
if (!args.includes("--flavour")) process.exit(32);
if (args.includes("--password") || args.includes("--profile") || args.includes("--recurse")) process.exit(33);
const flavour = args[args.indexOf("--flavour") + 1];
const profileNames = {
  "1a": "PDF/A-1a validation profile", "1b": "PDF/A-1b validation profile",
  "2a": "PDF/A-2a validation profile", "2b": "PDF/A-2b validation profile", "2u": "PDF/A-2u validation profile",
  "3a": "PDF/A-3a validation profile", "3b": "PDF/A-3b validation profile", "3u": "PDF/A-3u validation profile",
  "4": "PDF/A-4 validation profile", "4e": "PDF/A-4e validation profile", "4f": "PDF/A-4f validation profile",
  "ua1": "PDF/UA-1 validation profile", "ua2": "PDF/UA-2 + Tagged PDF validation profile",
};
const compliant = process.env.FAKE_VERAPDF_COMPLIANT !== "0";
if (process.env.FAKE_VERAPDF_DELETE_SOURCE) fs.unlinkSync(process.env.FAKE_VERAPDF_DELETE_SOURCE);
if (process.env.FAKE_VERAPDF_MUTATE === "1") {
  fs.chmodSync(input, 0o600);
  fs.appendFileSync(input, "mutation");
}
const failedRule = {
  ruleStatus: "FAILED", specification: "ISO 14289-1:2014", clause: "7.1", testNumber: 1,
  status: "failed", failedChecks: 1, tags: ["structure"], description: "A fake failed rule",
  object: "PDDocument", test: "isTagged == true",
  checks: [{ status: "failed", context: "root/document[0]", errorMessage: "Missing structure", errorArguments: [] }],
};
const batchProblem = process.env.FAKE_VERAPDF_BATCH_PROBLEM === "1" ? 1 : 0;
const profileName = process.env.FAKE_VERAPDF_WRONG_PROFILE === "1" ? "wrong profile" : profileNames[flavour];
const report = {
  report: {
    buildInformation: { releaseDetails: [
      { id: "core", version: "1.30.2" },
      { id: "validation-model", version: "1.30.2" },
      { id: "apps", version: "1.30.2" },
    ] },
    jobs: [{
      itemDetails: { name: input, size: fs.statSync(input).size },
      validationResult: [{
        details: {
          passedRules: compliant ? 10 : 9, failedRules: compliant ? 0 : 1,
          passedChecks: compliant ? 20 : 19, failedChecks: compliant ? 0 : 1,
          tags: compliant ? undefined : ["structure"], ruleSummaries: compliant ? [] : [failedRule],
        },
        jobEndStatus: "normal", profileName,
        statement: compliant ? "PDF file is compliant with Validation Profile requirements." : "PDF file is not compliant with Validation Profile requirements.",
        compliant,
      }],
      processingTime: { duration: "00:00:00.001" },
    }],
    batchSummary: {
      totalJobs: 1, outOfMemory: 0, veraExceptions: batchProblem,
      failedEncryptedJobs: 0, failedParsingJobs: 0,
      validationSummary: { totalJobCount: 1, successfulJobCount: 1, failedJobCount: 0 },
    },
  },
};
console.log(JSON.stringify(report));
process.exit(compliant ? 0 : 1);
`, "utf8");
  await fs.chmod(fakeProvider, 0o755);
  const fakeEnv = { OPEN_OFFICE_PDF_VERAPDF: fakeProvider };

  const probe = jsonResult(run(python, [provider, "probe"], { env: fakeEnv, status: 0 }));
  assert.equal(probe.provider, "verapdf");
  assert.equal(probe.providerVersion, "1.30.2");
  assert.equal(probe.integration, "shipped-thin-script-external-cli");
  assert.equal(probe.customProfilesAccepted, false);
  assert.equal(probe.passwordsAccepted, false);
  assert.ok(probe.profiles.includes("ua2"));

  const registryProbe = jsonResult(run(python, [registry, "check", "--provider", "verapdf", "--require"], {
    env: fakeEnv,
    status: 0,
  }));
  assert.equal(registryProbe.providers[0].available, true);
  assert.equal(registryProbe.providers[0].integration, "shipped-thin-script-external-cli");
  assert.equal(registryProbe.providers[0].evidence.minimumVersion, "1.30.0");
  assert.equal(registryProbe.providers[0].evidence.maximumVersionExclusive, "1.31.0");
  assert.equal(registryProbe.providers[0].evidence.semanticVersions.verapdf, "1.30.2");
  const plan = jsonResult(run(python, [
    registry, "plan", "--task", "validate-conformance", "--provider", "verapdf", "--strategy", "read-only",
    "--input", source, "--require-provider",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(plan.integration, "shipped-thin-script-external-cli");
  assert.equal(plan.silentFallback, false);

  const fakePass = jsonResult(run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b", "--require-compliant",
  ], { env: fakeEnv, status: 0 }));
  assert.equal(fakePass.schema, "open-office-artifact-tool.verapdf-validation.v1");
  assert.equal(fakePass.machineRuleCompliant, true);
  assert.equal(fakePass.validationPolicy.explicitBuiltInProfile, true);
  assert.equal(fakePass.sourceProtected, true);
  assert.equal(fakePass.rawProviderReport.retained, false);
  assert.deepEqual(await fs.readFile(source), fixtureBytes);

  const fakeNoncompliant = jsonResult(run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "ua1",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_COMPLIANT: "0" }, status: 0 }));
  assert.equal(fakeNoncompliant.ok, true, "noncompliance is a completed validation result without a delivery gate");
  assert.equal(fakeNoncompliant.machineRuleCompliant, false);
  assert.equal(fakeNoncompliant.humanReview.required, true);
  assert.equal(fakeNoncompliant.failedRuleSummaries.length, 1);
  const fakeGate = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "ua1", "--require-compliant",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_COMPLIANT: "0" }, status: 2 });
  assert.match(jsonResult(fakeGate, "stderr").policyGates.failures[0], /not compliant/);

  const stale = run(python, [
    provider, "validate", source, "--expected-sha256", "0".repeat(64), "--flavour", "1b",
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(stale, "stderr").error, /source SHA-256 mismatch/);
  const raisedBudget = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b",
    "--max-input-bytes", String(512 * 1024 * 1024 + 1),
  ], { env: fakeEnv, status: 2 });
  assert.match(jsonResult(raisedBudget, "stderr").error, /cannot exceed the hard maximum/);
  const missingFlavour = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash,
  ], { env: fakeEnv, status: 2 });
  assert.match(missingFlavour.stderr, /--flavour/);
  const customProfile = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b", "--profile", "custom.xml",
  ], { env: fakeEnv, status: 2 });
  assert.match(customProfile.stderr, /unrecognized arguments/);

  const oldVersion = run(python, [provider, "probe"], {
    env: { ...fakeEnv, FAKE_VERAPDF_VERSION: "1.28.2" },
    status: 2,
  });
  assert.match(jsonResult(oldVersion, "stderr").error, />= 1\.30\.0 and < 1\.31\.0/);
  const missingProfile = run(python, [provider, "probe"], {
    env: { ...fakeEnv, FAKE_VERAPDF_PROFILES: "1a 1b ua1" },
    status: 2,
  });
  assert.match(jsonResult(missingProfile, "stderr").error, /missing required built-in profiles/);
  const badJson = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_BAD_JSON: "1" }, status: 2 });
  assert.match(jsonResult(badJson, "stderr").error, /valid UTF-8 JSON/);
  const wrongProfile = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_WRONG_PROFILE: "1" }, status: 2 });
  assert.match(jsonResult(wrongProfile, "stderr").error, /validated the wrong profile/);
  const batchProblem = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_BATCH_PROBLEM: "1" }, status: 2 });
  assert.match(jsonResult(batchProblem, "stderr").error, /veraExceptions=1/);
  const oversized = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b", "--max-stdout-bytes", "1024",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_BIG: "1" }, status: 2 });
  assert.match(jsonResult(oversized, "stderr").error, /stdout exceeded the 1024 byte budget/);
  const timedOut = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b", "--timeout-seconds", "1",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_HANG: "1" }, status: 2 });
  assert.match(jsonResult(timedOut, "stderr").error, /timed out after 1 seconds/);
  const mutatedSnapshot = run(python, [
    provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_MUTATE: "1" }, status: 2 });
  assert.match(jsonResult(mutatedSnapshot, "stderr").error, /changed the private read-only source snapshot/);
  assert.deepEqual(await fs.readFile(source), fixtureBytes, "a malicious provider must still be isolated from source bytes");
  const disappearingSource = path.join(tempRoot, "disappearing-source.pdf");
  await fs.writeFile(disappearingSource, fixtureBytes);
  const disappeared = run(python, [
    provider, "validate", disappearingSource, "--expected-sha256", fixtureHash, "--flavour", "1b",
  ], { env: { ...fakeEnv, FAKE_VERAPDF_DELETE_SOURCE: disappearingSource }, status: 2 });
  const disappearedReport = jsonResult(disappeared, "stderr");
  assert.match(disappearedReport.error, /source PDF became unavailable during validation/);
  assert.doesNotMatch(disappeared.stderr, /Traceback/);

  const realProvider = process.env.OPEN_OFFICE_PDF_VERAPDF_TEST || process.env.OPEN_OFFICE_PDF_VERAPDF;
  if (realProvider) {
    const realEnv = { OPEN_OFFICE_PDF_VERAPDF: realProvider };
    const realProbe = jsonResult(run(python, [provider, "probe"], { env: realEnv, status: 0 }));
    assert.equal(realProbe.providerVersion, "1.30.2");
    const realPass = jsonResult(run(python, [
      provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "1b", "--require-compliant",
    ], { env: realEnv, status: 0 }));
    assert.equal(realPass.machineRuleCompliant, true);
    assert.equal(realPass.summary.failedRules, 0);
    assert.equal(realPass.validationPolicy.profileName, "PDF/A-1b validation profile");
    const realUa = jsonResult(run(python, [
      provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "ua1",
      "--max-failures-displayed", "1",
    ], { env: realEnv, status: 0 }));
    assert.equal(realUa.machineRuleCompliant, false);
    assert.equal(realUa.humanReview.required, true);
    assert.ok(realUa.summary.failedRules > 0);
    assert.equal(realUa.failedRuleSummaries.length, realUa.summary.failedRules);
    const realGate = run(python, [
      provider, "validate", source, "--expected-sha256", fixtureHash, "--flavour", "ua1", "--require-compliant",
    ], { env: realEnv, status: 2 });
    assert.equal(jsonResult(realGate, "stderr").machineRuleCompliant, false);
    assert.deepEqual(await fs.readFile(source), fixtureBytes);
  }
} finally {
  await fs.rm(tempRoot, { recursive: true, force: true });
}

console.log(
  process.env.OPEN_OFFICE_PDF_VERAPDF_TEST || process.env.OPEN_OFFICE_PDF_VERAPDF
    ? "veraPDF provider smoke ok"
    : "veraPDF provider smoke ok (real provider skipped: set OPEN_OFFICE_PDF_VERAPDF_TEST)",
);
