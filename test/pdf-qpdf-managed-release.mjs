import assert from "node:assert/strict";
import { execFile as execFileCallback } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { promisify } from "node:util";

import { PdfProviders } from "../src/pdf/providers/index.mjs";
import { plainPdfBytes } from "./fixtures/plain-pdf.mjs";

const execFile = promisify(execFileCallback);

if (process.env.OPEN_OFFICE_PDF_LIVE_PACK_TEST !== "1") {
  console.log("qpdf managed release smoke skipped (set OPEN_OFFICE_PDF_LIVE_PACK_TEST=1)");
} else {
  const temporary = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-qpdf-managed-release-"));
  try {
    const policyDirectory = path.join(temporary, ".open-office-artifact-tool");
    const policyPath = path.join(policyDirectory, "pdf-providers.json");
    await fs.mkdir(policyDirectory);
    await fs.writeFile(policyPath, JSON.stringify({
      installPolicy: "managed",
      allowedProviders: ["qpdf"],
      allowedPacks: ["qpdf"],
      acceptedLicenses: [],
      allowedOcrLanguages: ["eng", "chi_sim"],
      maxDownloadBytes: 16 * 1024 * 1024,
      maxUnpackedBytes: 32 * 1024 * 1024,
    }), "utf8");
    const source = path.join(temporary, "source.pdf");
    const sourceBytes = Buffer.from(plainPdfBytes([{ text: "managed qpdf pack smoke", width: 612, height: 792 }]));
    await fs.writeFile(source, sourceBytes, { mode: 0o600 });
    const inspection = { summary: { sourceSha256: "a".repeat(64) } };
    const resolution = await PdfProviders.resolve({
      task: "repair",
      provider: "qpdf",
      inspection,
      savePolicy: "rewrite",
      mutationAuthorized: true,
      invalidateSignaturesAuthorized: true,
      policyPath,
    });
    assert.equal(resolution.status, "installable", JSON.stringify(resolution.reason));
    const ready = await PdfProviders.ensure({ resolution, policyPath });
    assert.equal(ready.status, "ready", JSON.stringify(ready.reason));
    const qpdf = ready.runtime?.evidence?.commands?.qpdf?.executable;
    assert.ok(qpdf, "managed qpdf executable must be returned by the fresh probe");
    const version = await execFile(qpdf, ["--version"], { timeout: 5_000, maxBuffer: 16 * 1024 });
    assert.match(version.stdout, /^qpdf version 12\.3\.2\s*$/m);
    await execFile(qpdf, ["--check", source], { timeout: 5_000, maxBuffer: 64 * 1024 });
    assert.deepEqual(await fs.readFile(source), sourceBytes, "managed qpdf check must not mutate its source");
    const cached = await PdfProviders.probe({ provider: "qpdf", task: "repair", policyPath });
    assert.equal(cached.status, "ready", JSON.stringify(cached.reason));
    assert.equal(cached.runtime?.managed?.commandPaths?.qpdf, qpdf);
    console.log(`qpdf managed release smoke ok (${ready.installation.installed.qpdf.receipt.artifact.asset})`);
  } finally {
    await fs.rm(temporary, { recursive: true, force: true });
  }
}
