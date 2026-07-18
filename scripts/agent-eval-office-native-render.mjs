import { spawnSync } from "node:child_process";
import crypto from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

import sharp from "sharp";

const REQUIRED_COMMANDS = ["soffice", "pdfinfo", "pdftoppm"];

function commandExists(command) {
  const versionArgument = command === "soffice" ? "--version" : "-v";
  const result = spawnSync(command, [versionArgument], { encoding: "utf8" });
  return result.status === 0;
}

function commandFailure(label, result) {
  return label + " failed (" + (result.status ?? "signal") + "): "
    + String(result.stderr || result.stdout || result.error?.message || "unknown error").trim();
}

function numericPngOrder(left, right) {
  return Number(/\d+/.exec(left)?.[0]) - Number(/\d+/.exec(right)?.[0]);
}

export function nativeOfficeRenderStatus() {
  const commands = Object.fromEntries(REQUIRED_COMMANDS.map((command) => [command, commandExists(command)]));
  return { available: Object.values(commands).every(Boolean), commands };
}

/**
 * Renders one Office package through LibreOffice and Poppler without exposing
 * format-specific semantics. Family graders own their own semantic assertions;
 * this module owns only the native visual evidence lifecycle.
 */
export async function renderOfficeFile(filePath, label = "artifact") {
  const status = nativeOfficeRenderStatus();
  const missing = Object.entries(status.commands).filter(([, available]) => !available).map(([command]) => command);
  if (missing.length) return { available: false, reason: "missing " + missing.join(", "), commands: status.commands };

  const safeLabel = String(label).replace(/[^a-z0-9_-]+/gi, "-").replace(/^-+|-+$/g, "") || "artifact";
  const root = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-agent-eval-" + safeLabel + "-"));
  try {
    const profile = path.join(root, "profile");
    const converted = spawnSync("soffice", [
      "--headless",
      "-env:UserInstallation=" + pathToFileURL(profile).href,
      "--convert-to",
      "pdf",
      "--outdir",
      root,
      filePath,
    ], {
      encoding: "utf8",
      timeout: 60_000,
      maxBuffer: 8 * 1024 * 1024,
    });
    if (converted.status !== 0) return { available: true, ok: false, reason: commandFailure("soffice", converted) };

    const pdfs = (await fs.readdir(root)).filter((name) => name.toLowerCase().endsWith(".pdf"));
    if (pdfs.length !== 1) return { available: true, ok: false, reason: "soffice produced " + pdfs.length + " PDF files" };
    const pdfPath = path.join(root, pdfs[0]);
    const info = spawnSync("pdfinfo", [pdfPath], { encoding: "utf8", timeout: 30_000, maxBuffer: 1024 * 1024 });
    if (info.status !== 0) return { available: true, ok: false, reason: commandFailure("pdfinfo", info) };
    const pageCount = Number(/^Pages:\s+(\d+)$/m.exec(info.stdout || "")?.[1] || 0);

    const prefix = path.join(root, "page");
    const raster = spawnSync("pdftoppm", ["-png", "-r", "120", pdfPath, prefix], {
      encoding: "utf8",
      timeout: 60_000,
      maxBuffer: 8 * 1024 * 1024,
    });
    if (raster.status !== 0) return { available: true, ok: false, reason: commandFailure("pdftoppm", raster) };

    const pngs = (await fs.readdir(root))
      .filter((name) => /^page-\d+\.png$/i.test(name))
      .sort(numericPngOrder);
    const pages = [];
    for (const png of pngs) {
      const { data, info: image } = await sharp(path.join(root, png)).ensureAlpha().raw().toBuffer({ resolveWithObject: true });
      let nonWhitePixels = 0;
      for (let index = 0; index < data.length; index += 4) {
        if (data[index] < 245 || data[index + 1] < 245 || data[index + 2] < 245) nonWhitePixels += 1;
      }
      pages.push({
        width: image.width,
        height: image.height,
        nonWhitePixels,
        pixelSha256: crypto.createHash("sha256").update(data).digest("hex"),
      });
    }
    return {
      available: true,
      ok: pageCount > 0 && pages.length === pageCount,
      renderer: "libreoffice-poppler",
      pageCount,
      pages,
    };
  } finally {
    await fs.rm(root, { recursive: true, force: true });
  }
}
