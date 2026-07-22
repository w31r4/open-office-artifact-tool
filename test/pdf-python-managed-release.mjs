import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { PdfProviders } from "../src/pdf/providers/index.mjs";

const execFile = promisify(execFileCallback);

if (process.env.OPEN_OFFICE_PDF_LIVE_PACK_TEST !== "1") {
  console.log("python foundation managed release smoke skipped (set OPEN_OFFICE_PDF_LIVE_PACK_TEST=1)");
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-python-foundation-managed-release-"));
  try {
    const policyDirectory = path.join(temporary, ".open-office-artifact-tool");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    await fs.mkdir(policyDirectory);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: ["reportlab"],
      allowedPacks: ["python-foundation"],
      acceptedLicenses: [],
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: 100 * 1024 * 1024,
      maxUnpackedBytes: 256 * 1024 * 1024,
    }), "utf8");
    const resolution = await PdfProviders.resolve({
      task: "create-layout",
      provider: "reportlab",
      savePolicy: "rewrite",
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    const python = ready.runtime?.managed?.pythonPath;
    assert.ok(python, "managed Python executable must be returned by the fresh probe");
    const { stdout } = await execFile(python, ["-I", "-c", [
      "import importlib.metadata as m, json, reportlab, pdfplumber, pypdf, PIL",
      "print(json.dumps({name:m.version(name) for name in ['reportlab','pdfplumber','pypdf','pillow']}, sort_keys=True))",
    ].join("; ")], { timeout: 20_000, maxBuffer: 16 * 1024 });
    assert.deepEqual(JSON.parse(stdout), {
      pdfplumber: "0.11.9",
      pillow: "12.2.0",
      pypdf: "6.10.0",
      reportlab: "4.4.9",
    });
    const cached = await PdfProviders.probe({ provider: "reportlab", task: "create-layout", policyPath });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.equal(cached.runtime?.managed?.pythonPath, python);
    console.log(`python foundation managed release smoke ok (${ready.installation.installed["python-foundation"].receipt.artifact.asset})`);
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
