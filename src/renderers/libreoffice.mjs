import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob } from "../index.mjs";

const MIME_BY_FORMAT = {
  pdf: "application/pdf",
  html: "text/html",
  xhtml: "application/xhtml+xml",
  txt: "text/plain",
  csv: "text/csv",
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
};

const EXT_BY_MIME = new Map(Object.entries(MIME_BY_FORMAT).map(([ext, mime]) => [mime, ext]));

function normalizeMime(type = "") {
  return String(type || "").split(";")[0].trim().toLowerCase();
}

function extensionFor(typeOrFormat, fallback = "bin") {
  const raw = String(typeOrFormat || "").trim().toLowerCase();
  if (MIME_BY_FORMAT[raw]) return raw;
  return EXT_BY_MIME.get(normalizeMime(raw)) || fallback;
}

function outputTypeFor(format, fallback = "application/pdf") {
  const raw = String(format || "").trim().toLowerCase();
  return MIME_BY_FORMAT[raw] || (raw.includes("/") ? raw : fallback);
}

async function readBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new Error("LibreOffice renderer requires a FileBlob, Blob, ArrayBuffer, or Uint8Array input.");
}

function runCommand(command, args, options = {}) {
  const timeoutMs = options.timeoutMs ?? 60_000;
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
      finish(reject, new Error(`LibreOffice renderer timed out after ${timeoutMs}ms.`));
    }, timeoutMs);
    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => finish(reject, error));
    child.on("close", (exitCode) => {
      if (exitCode !== 0) {
        finish(reject, new Error(`LibreOffice renderer command exited with code ${exitCode}: ${stderr || stdout}`));
        return;
      }
      finish(resolve, { stdout, stderr });
    });
  });
}

function libreOfficeFormat(format, options = {}) {
  const raw = String(format || options.format || "pdf").trim().toLowerCase();
  if (options.convertTo) return options.convertTo;
  if (raw === "pdf") return "pdf";
  if (raw === "html") return "html";
  if (raw === "xhtml") return "xhtml";
  if (raw === "txt") return "txt";
  if (raw === "csv") return "csv";
  if (raw === "docx") return "docx";
  if (raw === "xlsx") return "xlsx";
  if (raw === "pptx") return "pptx";
  return raw;
}

function expectedOutputPath(outDir, inputPath, format) {
  const base = path.basename(inputPath, path.extname(inputPath));
  const extension = extensionFor(format, String(format || "pdf").split(":")[0] || "pdf");
  return path.join(outDir, `${base}.${extension}`);
}

export async function renderWithLibreOffice(request = {}, defaultOptions = {}) {
  const input = request.input || request.source;
  if (!input) throw new Error("LibreOffice renderer requires request.input or request.source.");
  const options = { ...defaultOptions, ...(request.options?.libreOffice || {}), ...(request.libreOffice || {}) };
  const inputType = normalizeMime(request.inputType || input?.type || options.inputType || "application/octet-stream");
  const format = String(request.format || options.format || "pdf").trim().toLowerCase();
  const outputType = request.outputType || options.outputType || outputTypeFor(format, "application/pdf");
  const command = options.command || process.env.LIBREOFFICE_COMMAND || process.env.SOFFICE_COMMAND || "soffice";
  const tempDir = await fs.mkdtemp(path.join(options.tempRoot || os.tmpdir(), "open-office-libreoffice-"));
  const inputExt = extensionFor(inputType, options.inputExtension || "bin");
  const inputPath = options.inputPath || path.join(tempDir, `input.${inputExt}`);
  const outDir = options.outDir || tempDir;
  const convertTo = libreOfficeFormat(format, options);
  const outputPath = options.outputPath || expectedOutputPath(outDir, inputPath, convertTo);

  try {
    await fs.writeFile(inputPath, await readBytes(input));
    await fs.mkdir(outDir, { recursive: true });
    const args = options.argsBuilder
      ? options.argsBuilder({ request, inputPath, outDir, outputPath, format, convertTo })
      : [
        ...(options.args || []),
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        convertTo,
        "--outdir",
        outDir,
        inputPath,
      ];
    await runCommand(command, args, options);
    let bytes;
    try {
      bytes = await fs.readFile(outputPath);
    } catch (error) {
      const files = await fs.readdir(outDir).catch(() => []);
      throw new Error(`LibreOffice renderer did not produce expected output ${outputPath}. Files: ${files.join(", ")}. ${error.message}`);
    }
    return new FileBlob(bytes, {
      type: outputType,
      metadata: {
        renderer: "libreoffice",
        command: path.basename(command),
        artifactKind: request.artifactKind,
        inputType,
        outputType,
        format,
        convertTo,
        page: request.options?.page ?? request.page,
        pageIndex: request.options?.pageIndex ?? request.pageIndex,
        slide: request.options?.slide ?? request.slide,
        sheet: request.options?.sheet ?? request.sheet,
        range: request.options?.range ?? request.range,
      },
    });
  } finally {
    if (!options.keepTemp) await fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

export function createLibreOfficeRenderer(defaultOptions = {}) {
  return async function libreOfficeRendererAdapter(request = {}) {
    return renderWithLibreOffice(request, defaultOptions);
  };
}

export const libreOfficeRenderer = createLibreOfficeRenderer();
