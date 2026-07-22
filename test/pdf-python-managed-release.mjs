import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { PdfProviders } from "../src/pdf/providers/index.mjs";

const execFile = promisify(execFileCallback);
const packId = process.env.OPEN_OFFICE_PDF_LIVE_PACK || "python-foundation";
const PACKS = Object.freeze({
  "python-foundation": {
    provider: "reportlab",
    task: "create-layout",
    allowedPacks: ["python-foundation"],
    acceptedLicenses: [],
    maxDownloadBytes: 100 * 1024 * 1024,
    maxUnpackedBytes: 256 * 1024 * 1024,
    program: [
      "import importlib.metadata as m, json, reportlab, pdfplumber, pypdf, PIL",
      "print(json.dumps({name:m.version(name) for name in ['reportlab','pdfplumber','pypdf','pillow']}, sort_keys=True))",
    ].join("; "),
    versions: {
      pdfplumber: "0.11.9",
      pillow: "12.2.0",
      pypdf: "6.10.0",
      reportlab: "4.4.9",
    },
  },
  "python-specialists": {
    provider: "pymupdf",
    task: "inspect",
    allowedPacks: ["python-specialists", "qpdf"],
    acceptedLicenses: ["agpl"],
    maxDownloadBytes: 128 * 1024 * 1024,
    maxUnpackedBytes: 300 * 1024 * 1024,
    program: [
      "import importlib.metadata as m, json, pymupdf, pikepdf, pyhanko, pyhanko_certvalidator",
      "print(json.dumps({name:m.version(name) for name in ['PyMuPDF','pikepdf','pyHanko','pyhanko-certvalidator']}, sort_keys=True))",
    ].join("; "),
    versions: {
      PyMuPDF: "1.27.2.3",
      pikepdf: "10.10.0",
      pyHanko: "0.35.2",
      "pyhanko-certvalidator": "0.31.1",
    },
  },
});
const profile = PACKS[packId];

if (!profile) throw new Error(`Unsupported Python managed-release pack ${packId}.`);

if (process.env.OPEN_OFFICE_PDF_LIVE_PACK_TEST !== "1") {
  console.log(`${packId} managed release smoke skipped (set OPEN_OFFICE_PDF_LIVE_PACK_TEST=1)`);
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), `open-office-${packId}-managed-release-`));
  try {
    const policyDirectory = path.join(temporary, ".open-office-artifact-tool");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    await fs.mkdir(policyDirectory);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: [profile.provider],
      allowedPacks: profile.allowedPacks,
      acceptedLicenses: profile.acceptedLicenses,
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: profile.maxDownloadBytes,
      maxUnpackedBytes: profile.maxUnpackedBytes,
    }), "utf8");
    const resolution = await PdfProviders.resolve({
      task: profile.task,
      provider: profile.provider,
      savePolicy: profile.task === "create-layout" ? "rewrite" : "read-only",
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    const python = ready.runtime?.managed?.pythonPath;
    assert.ok(python, "managed Python executable must be returned by the fresh probe");
    const { stdout } = await execFile(python, ["-I", "-c", profile.program], { timeout: 20_000, maxBuffer: 16 * 1024 });
    assert.deepEqual(JSON.parse(stdout), profile.versions);
    const cached = await PdfProviders.probe({ provider: profile.provider, task: profile.task, policyPath });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.equal(cached.runtime?.managed?.pythonPath, python);
    console.log(`${packId} managed release smoke ok (${ready.installation.installed[packId].receipt.artifact.asset})`);
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
