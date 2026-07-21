import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.length) throw new TypeError(label + " must be a non-empty string.");
  return value;
}

function boundedIndex(value, label) {
  const index = Number(value);
  if (!Number.isSafeInteger(index) || index < 0) throw new TypeError(label + " must be a non-negative safe integer.");
  return index;
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  return JSON.parse(await fs.readFile(path.join(path.dirname(path.dirname(entry)), "package.json"), "utf8")).version;
}

async function partHashes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const result = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    result.set(name, sha256(await entry.async("uint8array")));
  }
  return result;
}

async function changedParts(source, output) {
  const [before, after] = await Promise.all([partHashes(source), partHashes(output)]);
  if (before.size !== after.size || [...before.keys()].some((name) => !after.has(name))) {
    throw new Error("Source-bound text patch changed the DOCX package part inventory.");
  }
  return [...before.keys()].filter((name) => before.get(name) !== after.get(name)).sort();
}

function locateTarget(document, target) {
  const blockIndex = boundedIndex(target?.blockIndex, "target.blockIndex");
  const block = document.blocks[blockIndex];
  if (!block) throw new Error("Target block is outside the imported document.");
  if (target.kind === "paragraph") {
    if (block.kind !== "paragraph") throw new Error("Target block is not a paragraph.");
    return { blockIndex, block, value: block.text, editable: block.textEditable, patchable: block.textPatchable, rangeId: `${block.id}/text`, snapshot: { kind: "paragraph", blockIndex, blockId: block.id } };
  }
  if (target.kind === "tableCell") {
    if (block.kind !== "table") throw new Error("Target block is not a table.");
    const row = boundedIndex(target.row, "target.row");
    const column = boundedIndex(target.column, "target.column");
    const cell = block.getCell(row, column);
    return {
      blockIndex,
      block: cell,
      value: cell.value,
      editable: cell.editable,
      patchable: cell.textPatchable,
      rangeId: `${cell.id}/text`,
      snapshot: {
        kind: "tableCell",
        blockIndex,
        blockId: block.id,
        row,
        column,
        gridColumn: cell.gridColumn,
        columnSpan: cell.columnSpan,
        rowSpan: cell.rowSpan,
        verticalMerge: cell.verticalMerge,
      },
    };
  }
  throw new TypeError("target.kind must be paragraph or tableCell.");
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
    throw new Error(label + " already exists; refusing to overwrite it.");
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function publishNoReplace(temporaryPath, finalPath) {
  await fs.link(temporaryPath, finalPath);
  await fs.rm(temporaryPath, { force: true });
}

export async function patchImportedText({
  inputPath,
  outputPath,
  auditPath,
  target,
  search,
  replacement,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const oldText = requiredText(search, "search");
  const nextText = typeof replacement === "string" ? replacement : String(replacement ?? "");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const located = locateTarget(document, target);
  if (located.editable) {
    throw new Error("Target advertises whole-text editing; this workflow is reserved for narrower source-bound textPatchable edits.");
  }
  if (!located.patchable) throw new Error("Target does not advertise source-bound textPatchable capability.");
  const first = located.value.indexOf(oldText);
  if (first < 0 || located.value.indexOf(oldText, first + 1) >= 0) {
    throw new Error("Search must occur exactly once in the selected target's visible text.");
  }
  const expected = located.value.slice(0, first) + nextText + located.value.slice(first + oldText.length);
  const range = document.resolve(located.rangeId);
  if (!range) throw new Error("Advertised source-bound text range did not resolve.");
  range.replace(oldText, nextText);

  const temporaryPath = finalPath + ".tmp-" + process.pid + "-" + Date.now();
  const temporaryAuditPath = finalAuditPath + ".tmp-" + process.pid + "-" + Date.now();
  await fs.mkdir(path.dirname(finalPath), { recursive: true });
  await fs.mkdir(path.dirname(finalAuditPath), { recursive: true });
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction.");
    const changed = await changedParts(source, output);
    if (changed.length !== 1 || changed[0] !== "word/document.xml") {
      throw new Error("Source-bound text patch changed unexpected DOCX parts: " + changed.join(", "));
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = locateTarget(reimported, located.snapshot);
    if (roundTrip.value !== expected || !roundTrip.patchable || roundTrip.editable) {
      throw new Error("Reimport did not preserve the requested text or source-bound capability boundary.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error("Document verification failed: " + verification.ndjson);
    const render = await modelRender(reimported);
    const audit = {
      schema: "open-office-artifact-tool.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-literal-text-patch",
        target: located.snapshot,
        sourceTextSha256: sha256(Buffer.from(located.value, "utf8")),
        searchSha256: sha256(Buffer.from(oldText, "utf8")),
        replacementSha256: sha256(Buffer.from(nextText, "utf8")),
      },
      validation: {
        changedParts: changed,
        reimport: { ok: true, valueSha256: sha256(Buffer.from(roundTrip.value, "utf8")), textPatchable: true, textEditable: false },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2), { flag: "wx" });
    await publishNoReplace(temporaryPath, finalPath);
    try {
      await publishNoReplace(temporaryAuditPath, finalAuditPath);
    } catch (error) {
      await fs.rm(finalPath, { force: true });
      throw error;
    }
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, kind, blockIndex, search, replacement = "", row, column] = argv;
  const target = kind === "tableCell"
    ? { kind, blockIndex: boundedIndex(blockIndex, "blockIndex"), row: boundedIndex(row, "row"), column: boundedIndex(column, "column") }
    : { kind, blockIndex: boundedIndex(blockIndex, "blockIndex") };
  return { inputPath, outputPath, auditPath, target, search, replacement };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await patchImportedText(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  }));
}
