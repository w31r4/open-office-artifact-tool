import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob } from "../shared/file-blob.mjs";

const MIME_BY_FORMAT = {
  png: "image/png",
  ppm: "image/x-portable-pixmap",
  tiff: "image/tiff",
};

function normalizeFormat(format, outputType) {
  const raw = String(format || "").trim().toLowerCase();
  if (raw === "image/png") return "png";
  if (raw === "image/tiff") return "tiff";
  if (raw) return raw;
  return Object.entries(MIME_BY_FORMAT).find(([, mime]) => mime === outputType)?.[0] || "png";
}

async function readBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new Error("Poppler renderer requires PDF binary input.");
}

function runCommand(command, args, options = {}) {
  const timeoutMs = options.timeoutMs ?? 30_000;
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd: options.cwd || process.cwd(), env: { ...process.env, ...(options.env || {}) }, stdio: ["ignore", "pipe", "pipe"], windowsHide: true });
    let stdout = "";
    let stderr = "";
    let settled = false;
    const finish = (fn, value) => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      fn(value);
    };
    const timer = setTimeout(() => {
      child.kill("SIGKILL");
      finish(reject, new Error(`Poppler renderer timed out after ${timeoutMs}ms.`));
    }, timeoutMs);
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => finish(reject, error));
    child.on("close", (exitCode) => {
      if (exitCode !== 0) {
        finish(reject, new Error(`Poppler renderer command exited with code ${exitCode}: ${stderr || stdout}`));
        return;
      }
      finish(resolve, { stdout, stderr });
    });
  });
}

function defaultArgs({ format, page, dpi, inputPath, outputPrefix }) {
  const fmt = format === "tiff" ? "-tiff" : format === "ppm" ? "-r" : "-png";
  const args = [];
  if (page != null) args.push("-f", String(page), "-l", String(page), "-singlefile");
  if (format === "ppm") args.push(String(dpi || 150));
  else args.push(fmt);
  if (format !== "ppm" && dpi) args.push("-r", String(dpi));
  args.push(inputPath, outputPrefix);
  return args;
}

export async function renderWithPoppler(request = {}, defaultOptions = {}) {
  const input = request.input || request.source;
  const inputType = String(request.inputType || input?.type || defaultOptions.inputType || "application/pdf").split(";")[0].trim().toLowerCase();
  if (!input) throw new Error("Poppler renderer requires request.input or request.source.");
  if (inputType !== "application/pdf") throw new Error(`Poppler renderer supports application/pdf input, not ${inputType || "unknown"}.`);
  const options = { ...defaultOptions, ...(request.options?.poppler || {}), ...(request.poppler || {}) };
  const format = normalizeFormat(request.format || options.format, request.outputType);
  const outputType = request.outputType || MIME_BY_FORMAT[format];
  if (!outputType || !MIME_BY_FORMAT[format]) throw new Error(`Poppler renderer cannot produce ${request.format || request.outputType || "unknown"}; supported formats are png, ppm, and tiff.`);
  const command = options.command || process.env.POPPLER_RENDER_COMMAND || "pdftoppm";
  const tempDir = await fs.mkdtemp(path.join(options.tempRoot || os.tmpdir(), "open-office-poppler-"));
  const inputPath = path.join(tempDir, "input.pdf");
  const outputPrefix = path.join(tempDir, "page");
  const page = Number(request.options?.page ?? request.options?.pageIndex ?? request.page ?? request.pageIndex ?? options.page ?? options.pageIndex ?? 0) + (request.options?.page || request.page ? 0 : 1);
  const extension = format === "tiff" ? "tif" : format;
  const outputPath = `${outputPrefix}.${extension}`;

  try {
    await fs.writeFile(inputPath, await readBytes(input));
    const args = options.argsBuilder
      ? options.argsBuilder({ format, page, dpi: options.dpi, inputPath, outputPrefix, outputPath, request })
      : [...(options.args || []), ...defaultArgs({ format, page, dpi: options.dpi, inputPath, outputPrefix })];
    await runCommand(command, args, options);
    const bytes = await fs.readFile(outputPath);
    return new FileBlob(bytes, {
      type: outputType,
      metadata: {
        renderer: "poppler",
        command: path.basename(command),
        artifactKind: request.artifactKind,
        inputType,
        outputType,
        format,
        page,
      },
    });
  } finally {
    if (!options.keepTemp) await fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

export function createPopplerRenderer(defaultOptions = {}) {
  return async function popplerRendererAdapter(request = {}) {
    return renderWithPoppler(request, defaultOptions);
  };
}

export const popplerRenderer = createPopplerRenderer();
