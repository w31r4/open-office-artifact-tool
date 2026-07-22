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
  console.log("OCR managed release smoke skipped (set OPEN_OFFICE_PDF_LIVE_PACK_TEST=1)");
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-ocr-managed-release-"));
  try {
    const policyDirectory = path.join(temporary, ".open-office-artifact-tool");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    await fs.mkdir(policyDirectory);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: ["ocrmypdf"],
      allowedPacks: ["qpdf", "ocr-core", "ocr-language-eng", "ocr-language-chi-sim"],
      acceptedLicenses: [],
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: 256 * 1024 * 1024,
      maxUnpackedBytes: 768 * 1024 * 1024,
    }), "utf8");
    const resolution = await PdfProviders.resolve({
      task: "ocr",
      provider: "ocrmypdf",
      inspection: { summary: { sourceSha256: "a".repeat(64) } },
      savePolicy: "rewrite",
      mutationAuthorized: true,
      ocrLanguages: ["eng", "chi_sim"],
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    assert.deepEqual(resolution.installPlan?.packIds, ["qpdf", "ocr-core", "ocr-language-eng", "ocr-language-chi-sim"]);

    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    assert.deepEqual(Object.keys(ready.installation?.installed || {}).sort(), ["ocr-core", "ocr-language-chi-sim", "ocr-language-eng", "qpdf"]);

    const runtime = ready.runtime?.managed;
    assert.ok(runtime?.commandPaths?.ocrmypdf, "managed OCR command must be returned by the fresh probe");
    assert.ok(runtime.commandPaths.tesseract, "managed Tesseract command must be returned by the fresh probe");
    assert.ok(runtime.commandPaths.qpdf, "managed qpdf dependency must be returned by the fresh probe");
    assert.deepEqual(runtime.languagePacks.map(({ language }) => language).sort(), ["chi_sim", "eng"]);
    for (const language of runtime.languagePacks) {
      const stat = await fs.lstat(language.dataPath);
      assert.ok(stat.isFile() && !stat.isSymbolicLink(), "managed " + language.language + " data must be a regular private-cache file");
    }

    const commandOptions = {
      timeout: 20_000,
      maxBuffer: 64 * 1024,
      env: { ...process.env, ...runtime.environment },
    };
    assert.match(combinedOutput(await execFile(runtime.commandPaths.ocrmypdf, ["--version"], commandOptions)), /^17\.8\.1\s*$/m);
    assert.match(combinedOutput(await execFile(runtime.commandPaths.tesseract, ["--version"], commandOptions)), /^tesseract 5\./m);
    assert.match(combinedOutput(await execFile(runtime.commandPaths.qpdf, ["--version"], commandOptions)), /^qpdf version 12\.3\.2\s*$/m);

    const cached = await PdfProviders.probe({
      provider: "ocrmypdf",
      task: "ocr",
      languages: ["eng", "chi_sim"],
      policyPath,
    });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.deepEqual(cached.runtime?.managed?.languagePacks?.map(({ language }) => language).sort(), ["chi_sim", "eng"]);
    console.log("OCR managed release smoke ok (" + ready.installation.installed["ocr-core"].receipt.artifact.asset + ")");
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
