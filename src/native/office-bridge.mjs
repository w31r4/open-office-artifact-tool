import { spawn } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { FileBlob } from "../index.mjs";

const MIME_BY_FORMAT = {
  docx: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  xlsx: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  pptx: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  pdf: "application/pdf",
  png: "image/png",
  webp: "image/webp",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
};

const EXTENSION_BY_MIME = new Map(Object.entries(MIME_BY_FORMAT).map(([extension, mime]) => [mime, extension]));

export class OfficeBridgeError extends Error {
  constructor(message, details = {}) {
    super(message);
    this.name = "OfficeBridgeError";
    this.code = details.code || "OFFICE_BRIDGE_ERROR";
    this.details = details;
  }
}

function outputTypeFor(format, fallback = "application/octet-stream") {
  const normalized = String(format || "").trim().toLowerCase();
  if (!normalized) return fallback;
  return MIME_BY_FORMAT[normalized] || (normalized.includes("/") ? normalized : fallback);
}

function extensionFor(typeOrFormat, fallback = "bin") {
  const normalized = String(typeOrFormat || "").trim().toLowerCase();
  if (MIME_BY_FORMAT[normalized]) return normalized === "jpeg" ? "jpg" : normalized;
  return EXTENSION_BY_MIME.get(normalized) || fallback;
}

function normalizeCommand(options = {}) {
  const command = options.command || process.env.OFFICE_BRIDGE_COMMAND;
  if (!command) {
    throw new OfficeBridgeError("Native Office bridge command is not configured. Set OFFICE_BRIDGE_COMMAND or pass { command, args }.", { code: "OFFICE_BRIDGE_NOT_CONFIGURED" });
  }
  const args = options.args || (process.env.OFFICE_BRIDGE_ARGS ? JSON.parse(process.env.OFFICE_BRIDGE_ARGS) : []);
  return { command, args };
}

function parseBridgeJson(stdout) {
  const lines = String(stdout || "").split(/\r?\n/).map((line) => line.trim()).filter(Boolean);
  const candidate = lines.at(-1) || "";
  try {
    return JSON.parse(candidate);
  } catch (error) {
    throw new OfficeBridgeError("Native Office bridge did not return valid JSON on stdout.", { code: "OFFICE_BRIDGE_BAD_JSON", stdout, cause: error.message });
  }
}

export function callOfficeBridge(request = {}, options = {}) {
  const { command, args } = normalizeCommand(options);
  const timeoutMs = Number(options.timeoutMs ?? request.timeoutMs ?? 60_000);
  const cwd = options.cwd || process.cwd();
  const payload = JSON.stringify({ ...request, timeoutMs });

  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { cwd, env: { ...process.env, ...(options.env || {}) }, stdio: ["pipe", "pipe", "pipe"], windowsHide: true });
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
      finish(reject, new OfficeBridgeError(`Native Office bridge timed out after ${timeoutMs}ms.`, { code: "OFFICE_BRIDGE_TIMEOUT", timeoutMs, stderr, stdout }));
    }, timeoutMs);

    child.stdout.on("data", (chunk) => { stdout += chunk; });
    child.stderr.on("data", (chunk) => { stderr += chunk; });
    child.on("error", (error) => finish(reject, new OfficeBridgeError(`Failed to start native Office bridge: ${error.message}`, { code: "OFFICE_BRIDGE_SPAWN_FAILED", cause: error.message, command, args })));
    child.on("close", (exitCode) => {
      if (settled) return;
      if (exitCode !== 0) {
        finish(reject, new OfficeBridgeError(`Native Office bridge exited with code ${exitCode}.`, { code: "OFFICE_BRIDGE_EXIT", exitCode, stderr, stdout }));
        return;
      }
      let response;
      try {
        response = parseBridgeJson(stdout);
      } catch (error) {
        finish(reject, error);
        return;
      }
      if (response?.ok === false) {
        finish(reject, new OfficeBridgeError(response.error?.message || "Native Office bridge returned an error.", { code: response.error?.code || "OFFICE_BRIDGE_RESPONSE_ERROR", response, stderr }));
        return;
      }
      finish(resolve, response);
    });
    child.stdin.end(`${payload}\n`);
  });
}

async function inputBytes(input) {
  if (input instanceof FileBlob) return input.bytes;
  if (input && typeof input.arrayBuffer === "function") return new Uint8Array(await input.arrayBuffer());
  if (input instanceof Uint8Array) return input;
  if (input instanceof ArrayBuffer) return new Uint8Array(input);
  if (ArrayBuffer.isView(input)) return new Uint8Array(input.buffer, input.byteOffset, input.byteLength);
  throw new OfficeBridgeError("Native Office bridge input must be a FileBlob, Blob, ArrayBuffer, or Uint8Array.", { code: "OFFICE_BRIDGE_BAD_INPUT" });
}

export async function renderFileWithNativeOffice(input, options = {}) {
  const inputType = options.inputType || input?.type || "application/octet-stream";
  const outputType = options.outputType || outputTypeFor(options.format, "application/pdf");
  const format = options.format || extensionFor(outputType, "pdf");
  const tempRoot = options.tempRoot || os.tmpdir();
  const tempDir = await fs.mkdtemp(path.join(tempRoot, "open-office-bridge-"));
  const inputPath = options.inputPath || path.join(tempDir, `input.${extensionFor(inputType, "bin")}`);
  const outputPath = options.outputPath || path.join(tempDir, `output.${extensionFor(outputType || format, format || "bin")}`);

  try {
    await fs.writeFile(inputPath, await inputBytes(input));
    const response = await callOfficeBridge({
      operation: options.operation || "render",
      artifactKind: options.artifactKind,
      inputPath,
      outputPath,
      inputType,
      outputType,
      format,
      page: options.page,
      pageIndex: options.pageIndex,
      slide: options.slide,
      sheet: options.sheet,
      range: options.range,
      nativeOptions: options.nativeOptions || {},
      timeoutMs: options.timeoutMs,
    }, options.bridge || options);

    const bytes = response.dataBase64
      ? Buffer.from(response.dataBase64, "base64")
      : await fs.readFile(response.outputPath || outputPath);
    return new FileBlob(bytes, {
      type: response.outputType || response.type || outputType,
      metadata: {
        renderer: "native-office",
        bridge: response.bridge || response.adapter || "office-bridge",
        artifactKind: options.artifactKind,
        format,
        inputType,
        outputType: response.outputType || response.type || outputType,
        ...(response.metadata || {}),
      },
    });
  } finally {
    if (!options.keepTemp) await fs.rm(tempDir, { recursive: true, force: true }).catch(() => undefined);
  }
}

export function createNativeOfficeRenderer(defaultOptions = {}) {
  return async function nativeOfficeRendererAdapter(request = {}) {
    return renderFileWithNativeOffice(request.input || request.source, {
      ...defaultOptions,
      ...(request.options?.nativeOffice || {}),
      artifactKind: request.artifactKind || defaultOptions.artifactKind,
      inputType: request.inputType || defaultOptions.inputType,
      outputType: request.outputType || defaultOptions.outputType,
      format: request.format || defaultOptions.format,
      page: request.options?.page ?? request.page ?? defaultOptions.page,
      pageIndex: request.options?.pageIndex ?? request.pageIndex ?? defaultOptions.pageIndex,
      slide: request.options?.slide ?? request.slide ?? defaultOptions.slide,
      sheet: request.options?.sheet ?? request.sheet ?? defaultOptions.sheet,
      range: request.options?.range ?? request.range ?? defaultOptions.range,
    });
  };
}

export async function nativeOfficeStatus(options = {}) {
  try {
    return await callOfficeBridge({ operation: "status", timeoutMs: options.timeoutMs ?? 10_000 }, options);
  } catch (error) {
    if (options.throwOnError) throw error;
    return { ok: false, available: false, error: { code: error.code || "OFFICE_BRIDGE_STATUS_FAILED", message: error.message, details: error.details } };
  }
}

export const nativeOfficeRenderer = createNativeOfficeRenderer();
