#!/usr/bin/env node

import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { isDeepStrictEqual } from "node:util";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const referenceRepoRoot = path.join(repoRoot, "reference", "office-artifact-tool");
const referenceSkillsRoot = path.join(referenceRepoRoot, "skills");
const projectSkillsRoot = path.join(repoRoot, "skills");
const snapshotPath = path.join(projectSkillsRoot, "reference-sync.json");

export const REFERENCE_SKILL_BUNDLES = Object.freeze([
  "documents",
  "spreadsheets",
  "presentations",
  "pdf",
  "template-creator",
  "default-template-library",
]);

async function regularFiles(root, relative = "") {
  const entries = await fs.readdir(path.join(root, relative), { withFileTypes: true });
  const output = [];
  for (const entry of entries.sort((left, right) => left.name < right.name ? -1 : left.name > right.name ? 1 : 0)) {
    const child = relative ? `${relative}/${entry.name}` : entry.name;
    if (entry.isDirectory()) output.push(...await regularFiles(root, child));
    else if (entry.isFile()) output.push(child);
    else throw new Error(`Reference Skill source contains a non-regular path: ${child}`);
  }
  return output;
}

function sha256(bytes) {
  return createHash("sha256").update(bytes).digest("hex");
}

function sourceCommit() {
  const result = spawnSync("git", ["-C", referenceRepoRoot, "rev-parse", "HEAD"], {
    encoding: "utf8",
    shell: false,
  });
  if (result.status !== 0) throw new Error(`Cannot resolve reference Skill source commit: ${result.stderr || result.stdout}`);
  const commit = result.stdout.trim();
  if (!/^[0-9a-f]{40}$/.test(commit)) throw new Error(`Invalid reference Skill source commit: ${commit}`);
  return commit;
}

async function bundleSnapshot(bundle) {
  const root = path.join(referenceSkillsRoot, bundle);
  const files = await regularFiles(root);
  let bytes = 0;
  const digest = createHash("sha256");
  for (const relative of files) {
    const contents = await fs.readFile(path.join(root, relative));
    bytes += contents.length;
    digest.update(relative);
    digest.update("\0");
    digest.update(String(contents.length));
    digest.update("\0");
    digest.update(sha256(contents));
    digest.update("\n");
  }
  return { files: files.length, bytes, sha256: digest.digest("hex"), paths: files };
}

async function scanReferenceSkillSource() {
  const bundles = {};
  const pathsByBundle = {};
  let totalFiles = 0;
  let totalBytes = 0;
  for (const bundle of REFERENCE_SKILL_BUNDLES) {
    const snapshot = await bundleSnapshot(bundle);
    bundles[bundle] = { files: snapshot.files, bytes: snapshot.bytes, sha256: snapshot.sha256 };
    pathsByBundle[bundle] = snapshot.paths;
    totalFiles += snapshot.files;
    totalBytes += snapshot.bytes;
  }
  return {
    snapshot: {
      schemaVersion: 1,
      source: {
        commit: sourceCommit(),
        root: "reference/office-artifact-tool/skills",
      },
      totalFiles,
      totalBytes,
      bundles,
    },
    pathsByBundle,
  };
}

export async function createReferenceSkillSnapshot() {
  return (await scanReferenceSkillSource()).snapshot;
}

async function missingProjectPaths(pathsByBundle) {
  const missing = [];
  for (const bundle of REFERENCE_SKILL_BUNDLES) {
    for (const relative of pathsByBundle[bundle]) {
      const projectPath = path.join(projectSkillsRoot, bundle, relative);
      const stat = await fs.lstat(projectPath).catch(() => null);
      if (!stat?.isFile()) missing.push(`${bundle}/${relative}`);
    }
  }
  return missing;
}

export async function checkReferenceSkillSync() {
  const expected = JSON.parse(await fs.readFile(snapshotPath, "utf8"));
  const { snapshot: actual, pathsByBundle } = await scanReferenceSkillSource();
  const missing = await missingProjectPaths(pathsByBundle);
  if (missing.length) {
    throw new Error(`Project Skill surface is missing ${missing.length} reference path(s):\n${missing.map((entry) => `- ${entry}`).join("\n")}`);
  }
  if (!isDeepStrictEqual(actual, expected)) {
    throw new Error([
      "Reference Skill source drifted from skills/reference-sync.json.",
      `Expected commit: ${expected.source?.commit || "unknown"}`,
      `Actual commit: ${actual.source.commit}`,
      "Review the reference diff and project adapters, then regenerate the snapshot explicitly.",
    ].join("\n"));
  }
  return actual;
}

const entry = process.argv[1] ? path.resolve(process.argv[1]) : "";
if (entry === fileURLToPath(import.meta.url)) {
  const command = process.argv[2] || "check";
  if (command === "snapshot") console.log(`${JSON.stringify(await createReferenceSkillSnapshot(), null, 2)}\n`);
  else if (command === "check") {
    const snapshot = await checkReferenceSkillSync();
    console.log(`reference Skill sync ok: ${snapshot.totalFiles} files at ${snapshot.source.commit}`);
  } else {
    console.error("Usage: node scripts/reference-skill-sync.mjs [check|snapshot]");
    process.exitCode = 2;
  }
}
