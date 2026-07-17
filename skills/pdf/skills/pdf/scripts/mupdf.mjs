#!/usr/bin/env node

import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";

import { PdfFile } from "open-office-artifact-tool";

function fail(message) {
  process.stderr.write(`${message}\n`);
  process.exitCode = 2;
}

function parseArgs(argv) {
  const positional = [];
  const options = {};
  for (let index = 0; index < argv.length; index += 1) {
    const value = argv[index];
    if (!value.startsWith("--")) {
      positional.push(value);
      continue;
    }
    const [rawName, inline] = value.slice(2).split("=", 2);
    const name = rawName.replaceAll("-", "_");
    if (inline !== undefined) options[name] = inline;
    else if (argv[index + 1] && !argv[index + 1].startsWith("--")) options[name] = argv[++index];
    else options[name] = true;
  }
  return { positional, options };
}

function numberOption(value, label) {
  if (value === undefined) return undefined;
  const number = Number(value);
  if (!Number.isFinite(number)) throw new Error(`${label} must be a finite number.`);
  return number;
}

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function writeAtomic(target, bytes) {
  const absolute = path.resolve(target);
  await fs.mkdir(path.dirname(absolute), { recursive: true });
  const temporary = path.join(path.dirname(absolute), `.${path.basename(absolute)}.${process.pid}.${crypto.randomUUID()}.tmp`);
  try {
    await fs.writeFile(temporary, bytes, { flag: "wx" });
    await fs.rename(temporary, absolute);
  } catch (error) {
    await fs.rm(temporary, { force: true });
    throw error;
  }
  return absolute;
}

async function canonicalPath(candidate, { output = false } = {}) {
  const absolute = path.resolve(candidate);
  try {
    return await fs.realpath(absolute);
  } catch (error) {
    if (!output || error?.code !== "ENOENT") throw error;
    const unresolved = [];
    let cursor = absolute;
    while (true) {
      const parent = path.dirname(cursor);
      unresolved.unshift(path.basename(cursor));
      try {
        return path.join(await fs.realpath(parent), ...unresolved);
      } catch (parentError) {
        if (parentError?.code !== "ENOENT" || parent === cursor) throw parentError;
        cursor = parent;
      }
    }
  }
}

async function assertDistinctSourceOutput(source, output) {
  const [sourcePath, outputPath] = await Promise.all([
    canonicalPath(source),
    canonicalPath(output, { output: true }),
  ]);
  if (sourcePath === outputPath) throw new Error("Refusing to overwrite the source PDF, including through a symlink alias; choose a distinct output path.");
}

function print(value) {
  process.stdout.write(`${JSON.stringify(value, null, 2)}\n`);
}

function usage() {
  return [
    "Usage:",
    "  node scripts/mupdf.mjs probe",
    "  node scripts/mupdf.mjs inspect <input.pdf> [--max-bytes N]",
    "  node scripts/mupdf.mjs render <input.pdf> <output.png|jpg> [--page N] [--dpi N] [--format png|jpeg]",
    "  node scripts/mupdf.mjs edit <input.pdf> <operations.json> <output.pdf> [--save-policy rewrite|incremental] [--allow-signed] [--invalidate-signatures]",
  ].join("\n");
}

async function main() {
  const { positional, options } = parseArgs(process.argv.slice(2));
  const [command, ...args] = positional;
  if (!command || options.help) {
    process.stdout.write(`${usage()}\n`);
    return;
  }

  if (command === "probe") {
    const { MUPDF_VERSION } = await import("open-office-artifact-tool/pdf/mupdf");
    print({ available: true, provider: "mupdf", version: MUPDF_VERSION, runtime: "mupdf.js", license: "AGPL-3.0-or-later", lazyLoaded: true });
    return;
  }

  const maxBytes = numberOption(options.max_bytes, "--max-bytes");
  const limits = maxBytes === undefined ? undefined : { maxBytes };
  const password = process.env.OPEN_OFFICE_PDF_PASSWORD;

  if (command === "inspect") {
    if (args.length !== 1) throw new Error(`inspect requires one input PDF.\n${usage()}`);
    const result = await PdfFile.inspectPdf(path.resolve(args[0]), { limits, password, maxChars: Infinity });
    print({ summary: result.summary, records: result.records });
    return;
  }

  if (command === "render") {
    if (args.length !== 2) throw new Error(`render requires an input PDF and output image.\n${usage()}`);
    await assertDistinctSourceOutput(args[0], args[1]);
    const output = await PdfFile.renderPdf(path.resolve(args[0]), {
      limits,
      password,
      page: numberOption(options.page, "--page") ?? 1,
      dpi: numberOption(options.dpi, "--dpi") ?? 144,
      format: options.format || (path.extname(args[1]).toLowerCase() === ".jpg" ? "jpeg" : "png"),
    });
    const outputPath = await writeAtomic(args[1], output.bytes);
    print({ output: outputPath, bytes: output.bytes.byteLength, sha256: sha256(output.bytes), ...output.metadata });
    return;
  }

  if (command === "edit") {
    if (args.length !== 3) throw new Error(`edit requires an input PDF, operations JSON, and output PDF.\n${usage()}`);
    await assertDistinctSourceOutput(args[0], args[2]);
    const plan = JSON.parse(await fs.readFile(path.resolve(args[1]), "utf8"));
    const operations = Array.isArray(plan) ? plan : plan.operations;
    const output = await PdfFile.editPdf(path.resolve(args[0]), {
      ...(Array.isArray(plan) ? {} : plan),
      operations,
      limits,
      password,
      savePolicy: options.save_policy || (Array.isArray(plan) ? undefined : plan.savePolicy),
      allowSigned: options.allow_signed === true || (Array.isArray(plan) ? false : plan.allowSigned),
      invalidateSignatures: options.invalidate_signatures === true || (Array.isArray(plan) ? false : plan.invalidateSignatures),
    });
    const outputPath = await writeAtomic(args[2], output.bytes);
    print({ output: outputPath, bytes: output.bytes.byteLength, sha256: sha256(output.bytes), ...output.metadata });
    return;
  }

  throw new Error(`Unknown command: ${command}.\n${usage()}`);
}

main().catch((error) => fail(error?.stack || String(error)));
