import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { PdfProviders } from "../src/pdf/providers/index.mjs";

const execFile = promisify(execFileCallback);

function combinedOutput(result) {
  return String(result.stdout || "") + String(result.stderr || "");
}

if (process.env.OPEN_OFFICE_PDF_LIVE_PACK_TEST !== "1") {
  console.log("veraPDF managed release smoke skipped (set OPEN_OFFICE_PDF_LIVE_PACK_TEST=1)");
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-verapdf-managed-release-"));
  try {
    const policyDirectory = path.join(temporary, ".open-office-artifact-tool");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    const home = path.join(temporary, "home");
    await Promise.all([fs.mkdir(policyDirectory), fs.mkdir(home)]);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: ["verapdf"],
      allowedPacks: ["verapdf"],
      acceptedLicenses: [],
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: 96 * 1024 * 1024,
      maxUnpackedBytes: 256 * 1024 * 1024,
    }), "utf8");
    const resolution = await PdfProviders.resolve({
      task: "validate-conformance",
      provider: "verapdf",
      inspection: { summary: { sourceSha256: "b".repeat(64) } },
      savePolicy: "read-only",
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    assert.deepEqual(resolution.installPlan?.packIds, ["verapdf"]);

    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    const runtime = ready.runtime?.managed;
    const verapdf = runtime?.commandPaths?.verapdf;
    assert.ok(verapdf, "managed veraPDF command must be returned by the fresh probe");
    const commandOptions = {
      timeout: 20_000,
      maxBuffer: 64 * 1024,
      env: { ...process.env, HOME: home, ...runtime.environment },
    };
    assert.match(combinedOutput(await execFile(verapdf, ["--version"], commandOptions)), /^veraPDF 1\.30\.2\s*$/m);
    assert.match(combinedOutput(await execFile(verapdf, ["--list"], commandOptions)), /\bua2\b/i);

    const cached = await PdfProviders.probe({ provider: "verapdf", task: "validate-conformance", policyPath });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.equal(cached.runtime?.managed?.commandPaths?.verapdf, verapdf);
    console.log("veraPDF managed release smoke ok (" + ready.installation.installed.verapdf.receipt.artifact.asset + ")");
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
