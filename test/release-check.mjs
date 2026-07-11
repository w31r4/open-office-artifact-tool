import assert from "node:assert/strict";
import fs from "node:fs";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const result = spawnSync(process.execPath, ["scripts/release-check.mjs", "--json", "--skip-network", "--skip-commands", "--allow-dirty"], {
  cwd: repoRoot,
  encoding: "utf8",
});
assert.equal(result.status, 0, `release-check failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
const report = JSON.parse(result.stdout);
assert.equal(report.package.name, "open-office-artifact-tool");
assert.equal(report.publishReady, true);
assert.ok(report.checks.some((check) => check.name === "package metadata" && check.ok));
assert.ok(report.checks.some((check) => check.name === "third-party license policy" && check.ok));
assert.ok(report.checks.some((check) => check.name === "npm auth" && check.skipped));
assert.match(report.nextPublishCommand, /npm publish/);
const releaseWorkflow = fs.readFileSync(path.join(repoRoot, ".github/workflows/release.yml"), "utf8");
assert.match(releaseWorkflow, /workflow_dispatch/);
assert.match(releaseWorkflow, /publish_npm/);
assert.match(releaseWorkflow, /default: "false"/);
assert.match(releaseWorkflow, /secrets\.NPM_TOKEN/);
assert.match(releaseWorkflow, /gh release create/);

const rejectedPolicyPath = path.join(os.tmpdir(), `open-office-license-policy-${process.pid}.json`);
const normalPolicy = JSON.parse(fs.readFileSync(path.join(repoRoot, "scripts", "license-policy.json"), "utf8"));
fs.writeFileSync(rejectedPolicyPath, JSON.stringify({ ...normalPolicy, allowedLockLicenses: normalPolicy.allowedLockLicenses.filter((license) => license !== "MIT") }));
try {
  const rejected = spawnSync(process.execPath, ["scripts/release-check.mjs", "--json", "--skip-network", "--skip-commands", "--allow-dirty"], {
    cwd: repoRoot,
    encoding: "utf8",
    env: { ...process.env, OFFICE_ARTIFACT_LICENSE_POLICY: rejectedPolicyPath },
  });
  assert.equal(rejected.status, 1);
  const rejectedReport = JSON.parse(rejected.stdout);
  assert.ok(rejectedReport.checks.some((check) => check.name === "third-party license policy" && !check.ok && /unapproved license expression MIT/.test(check.stderr)));
  assert.ok(rejectedReport.blockers.some((blocker) => /license policy/.test(blocker)));
} finally {
  fs.rmSync(rejectedPolicyPath, { force: true });
}

console.log("release check smoke ok");
