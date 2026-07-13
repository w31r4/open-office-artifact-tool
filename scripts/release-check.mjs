import { spawnSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const args = new Set(process.argv.slice(2));
const json = args.has("--json");
const skipNetwork = args.has("--skip-network");
const skipCommands = args.has("--skip-commands");
const allowDirty = args.has("--allow-dirty");

function run(command, commandArgs, options = {}) {
  const result = spawnSync(command, commandArgs, {
    cwd: repoRoot,
    encoding: "utf8",
    shell: false,
    ...options,
  });
  return {
    ok: result.status === 0,
    status: result.status,
    stdout: String(result.stdout || "").trim(),
    stderr: String(result.stderr || "").trim(),
    command: [command, ...commandArgs].join(" "),
  };
}

function commandExists(command) {
  const result = spawnSync(process.platform === "win32" ? "where" : "command", process.platform === "win32" ? [command] : ["-v", command], { encoding: "utf8", shell: process.platform !== "win32" });
  return result.status === 0;
}

function summarizeCheck(name, result, required = true) {
  return { name, required, ...result };
}

const pkg = JSON.parse(fs.readFileSync(path.join(repoRoot, "package.json"), "utf8"));
const lock = JSON.parse(fs.readFileSync(path.join(repoRoot, "package-lock.json"), "utf8"));
const licensePolicyPath = process.env.OFFICE_ARTIFACT_LICENSE_POLICY || path.join(repoRoot, "scripts", "license-policy.json");
const licensePolicy = JSON.parse(fs.readFileSync(licensePolicyPath, "utf8"));
const checks = [];
const blockers = [];

const gitStatus = run("git", ["status", "--short", "--untracked-files=normal"]);
checks.push(summarizeCheck("git status clean", { ...gitStatus, ok: gitStatus.ok && (allowDirty || !gitStatus.stdout), stdout: gitStatus.stdout || "clean" }, !allowDirty));
if (gitStatus.stdout && !allowDirty) blockers.push("Working tree is not clean.");

checks.push(summarizeCheck("package metadata", {
  ok: Boolean(pkg.name && pkg.version && pkg.type === "module" && pkg.exports?.["."] && pkg.files?.includes("src/**")),
  stdout: `${pkg.name}@${pkg.version}`,
  stderr: "",
  command: "read package.json",
}));
if (!checks.at(-1).ok) blockers.push("package.json metadata is incomplete for npm publish.");

const noticesPath = path.join(repoRoot, "THIRD_PARTY_NOTICES.md");
const declaredDependencyNames = [...new Set([...Object.keys(pkg.dependencies || {}), ...Object.keys(pkg.peerDependencies || {})])];
const policyNames = Object.keys(licensePolicy.declaredPackages || {});
const lockLicenseIssues = Object.entries(lock.packages || {}).flatMap(([packagePath, metadata]) => {
  if (!packagePath.startsWith("node_modules/")) return [];
  if (!metadata.license) return [`${packagePath}: missing license metadata`];
  if (!licensePolicy.allowedLockLicenses.includes(metadata.license)) return [`${packagePath}: unapproved license expression ${metadata.license}`];
  return [];
});
const missingPolicy = declaredDependencyNames.filter((name) => !policyNames.includes(name));
const stalePolicy = policyNames.filter((name) => !declaredDependencyNames.includes(name));
const noticesText = fs.existsSync(noticesPath) ? fs.readFileSync(noticesPath, "utf8") : "";
const missingNotices = declaredDependencyNames.filter((name) => !noticesText.toLowerCase().includes(name.toLowerCase()));
const licenseOk = fs.existsSync(noticesPath) && pkg.files?.includes("THIRD_PARTY_NOTICES.md") && !lockLicenseIssues.length && !missingPolicy.length && !stalePolicy.length && !missingNotices.length;
checks.push(summarizeCheck("third-party license policy", {
  ok: licenseOk,
  stdout: licenseOk ? `${Object.keys(lock.packages || {}).filter((name) => name.startsWith("node_modules/")).length} locked packages audited` : "license audit failed",
  stderr: [...lockLicenseIssues, ...missingPolicy.map((name) => `missing policy: ${name}`), ...stalePolicy.map((name) => `stale policy: ${name}`), ...missingNotices.map((name) => `missing notice: ${name}`)].join("\n"),
  command: "audit package-lock.json + THIRD_PARTY_NOTICES.md",
}));
if (!licenseOk) blockers.push("Third-party license policy or notices are incomplete.");

if (!skipCommands) {
  for (const [name, commandArgs] of [
    ["npm test", ["test"]],
    ["npm run docs:api", ["run", "docs:api"]],
    ["npm run test:pack", ["run", "test:pack"]],
  ]) {
    const check = summarizeCheck(name, run("npm", commandArgs));
    checks.push(check);
    if (!check.ok) blockers.push(`${name} failed.`);
  }
  if (fs.existsSync(path.join(repoRoot, "native", "OfficeBridge")) && commandExists("dotnet")) {
    const check = summarizeCheck("dotnet test native/OfficeBridge", run("dotnet", ["test", "native/OfficeBridge"]));
    checks.push(check);
    if (!check.ok) blockers.push("dotnet test native/OfficeBridge failed.");
  } else {
    checks.push(summarizeCheck("dotnet test native/OfficeBridge", { ok: true, stdout: "skipped: dotnet or native/OfficeBridge unavailable", stderr: "", command: "dotnet test native/OfficeBridge" }, false));
  }
  if (fs.existsSync(path.join(repoRoot, "native", "OpenXmlWasm")) && commandExists("dotnet")) {
    const check = summarizeCheck("dotnet test native/OpenXmlWasm", run("dotnet", ["test", "native/OpenXmlWasm/OpenXmlWasm.sln", "--configuration", "Release"]));
    checks.push(check);
    if (!check.ok) blockers.push("dotnet test native/OpenXmlWasm failed.");
  } else {
    checks.push(summarizeCheck("dotnet test native/OpenXmlWasm", { ok: true, stdout: "skipped: dotnet or native/OpenXmlWasm unavailable", stderr: "", command: "dotnet test native/OpenXmlWasm" }, false));
  }
}

let npmAuth = { ok: false, skipped: skipNetwork, stdout: "", stderr: "", command: "npm whoami" };
let npmView = { ok: false, skipped: skipNetwork, stdout: "", stderr: "", command: `npm view ${pkg.name} version --json` };
if (!skipNetwork) {
  npmAuth = run("npm", ["whoami"]);
  npmView = run("npm", ["view", pkg.name, "version", "--json"]);
  checks.push(summarizeCheck("npm auth", npmAuth));
  checks.push(summarizeCheck("npm package lookup", npmView, false));
  if (!npmAuth.ok) blockers.push("npm auth unavailable: run npm adduser or configure an npm token before publishing.");
} else {
  checks.push(summarizeCheck("npm auth", npmAuth, false));
  checks.push(summarizeCheck("npm package lookup", npmView, false));
}

const publishedVersion = npmView.ok ? npmView.stdout.replace(/^"|"$/g, "") : null;
if (publishedVersion === pkg.version) blockers.push(`npm ${pkg.name}@${pkg.version} is already published; bump version before publishing.`);

const result = {
  package: { name: pkg.name, version: pkg.version },
  publishReady: blockers.length === 0,
  npmAuth: npmAuth.ok ? npmAuth.stdout : null,
  publishedVersion,
  checks,
  blockers,
  nextPublishCommand: `npm publish --access public`,
};

if (json) {
  console.log(JSON.stringify(result, null, 2));
} else {
  console.log(`${pkg.name}@${pkg.version} release check`);
  for (const check of checks) console.log(`${check.ok ? "✓" : check.required ? "✗" : "-"} ${check.name}${check.stdout ? ` — ${check.stdout.split("\n")[0]}` : ""}`);
  if (blockers.length) {
    console.log("\nBlockers:");
    for (const blocker of blockers) console.log(`- ${blocker}`);
  } else {
    console.log(`\nPublish-ready. Command: ${result.nextPublishCommand}`);
  }
}

process.exit(result.publishReady ? 0 : 1);
