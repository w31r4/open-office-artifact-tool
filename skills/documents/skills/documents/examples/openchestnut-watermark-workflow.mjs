import crypto from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

async function refuseExisting(filePath, label) {
  try {
    await fs.access(filePath);
    throw new Error(`${label} already exists; refusing to replace it.`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function packagePartHashes(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const hashes = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (!entry.dir) hashes.set(name, sha256(await entry.async("uint8array")));
  }
  return hashes;
}

async function changedPackageParts(before, after) {
  const [left, right] = await Promise.all([packagePartHashes(before), packagePartHashes(after)]);
  return [...new Set([...left.keys(), ...right.keys()])]
    .filter((name) => left.get(name) !== right.get(name))
    .sort();
}

function watermarkSnapshot(watermark) {
  return {
    id: watermark.id,
    text: watermark.text,
    referenceType: watermark.referenceType,
    sectionIndex: watermark.sectionIndex,
    editable: watermark.editable,
    sourceBound: watermark.sourceBound,
  };
}

function selectWatermark(document, { expectedText, sectionIndex, referenceType }) {
  const matches = document.watermarks.filter((watermark) =>
    watermark.text === expectedText &&
    watermark.sectionIndex === sectionIndex &&
    watermark.referenceType === referenceType);
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one ${referenceType} watermark with the requested text in section ${sectionIndex}; found ${matches.length}.`);
  }
  const watermark = matches[0];
  if (!watermark.sourceBound || !watermark.editable) {
    throw new Error("The selected watermark does not advertise the recognized source-bound edit capability.");
  }
  return watermark;
}

export async function editDocumentWatermark({
  inputPath,
  outputPath,
  auditPath,
  expectedText,
  replacementText,
  sectionIndex = 0,
  referenceType = "default",
  remove = false,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath) throw new Error("outputPath must differ from inputPath so the source remains immutable.");
  if (finalAuditPath === sourcePath || finalAuditPath === finalPath) throw new Error("auditPath must differ from both DOCX paths.");
  const originalText = requiredText(expectedText, "expectedText");
  const nextText = remove ? undefined : requiredText(replacementText, "replacementText");
  const targetSection = Number(sectionIndex);
  if (!Number.isInteger(targetSection) || targetSection < 0) throw new TypeError("sectionIndex must be a non-negative integer.");
  if (!["default", "first", "even"].includes(referenceType)) throw new TypeError("referenceType must be default, first, or even.");
  await Promise.all([refuseExisting(finalPath, "outputPath"), refuseExisting(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectWatermark(document, {
    expectedText: originalText,
    sectionIndex: targetSection,
    referenceType,
  });
  const before = watermarkSnapshot(selected);
  const untouched = document.watermarks.filter((item) => item !== selected).map(watermarkSnapshot);
  if (remove) selected.remove();
  else selected.text = nextText;

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    const output = Buffer.from(exported.bytes);
    const changedParts = await changedPackageParts(source, output);
    if (changedParts.length !== 1 || !/^word\/header\d+\.xml$/.test(changedParts[0])) {
      throw new Error(`Watermark edit changed an unexpected package scope: ${changedParts.join(", ") || "none"}.`);
    }
    const currentSource = await fs.readFile(sourcePath);
    if (sha256(currentSource) !== sourceHash) throw new Error("The source DOCX changed during the transaction; refusing publication.");

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const retained = reimported.watermarks.filter((item) => item.id !== before.id).map(watermarkSnapshot);
    if (JSON.stringify(retained) !== JSON.stringify(untouched)) throw new Error("Unrelated watermark objects changed during export.");
    if (remove) {
      if (reimported.watermarks.some((item) => item.id === before.id)) throw new Error("Removed watermark remains after the second import.");
    } else {
      const roundTrip = selectWatermark(reimported, {
        expectedText: nextText,
        sectionIndex: targetSection,
        referenceType,
      });
      if (roundTrip.id !== before.id) throw new Error("Watermark package-local identity changed during the bounded text edit.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const preview = await reimported.render({ format: "svg" });
    if (!/<svg\b/i.test(await preview.text())) throw new Error("Document model render did not produce SVG evidence.");

    const audit = {
      schema: "open-office-artifact-tool.docx-watermark-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", silentFallback: false },
      operation: {
        type: remove ? "canonical-text-watermark-remove" : "canonical-text-watermark-edit",
        watermark: before,
        replacementTextSha256: nextText == null ? undefined : sha256(Buffer.from(nextText, "utf8")),
        changedParts,
      },
      validation: {
        sourceImmutable: true,
        secondImport: true,
        verify: { ok: true },
        modelRender: { ok: true, renderer: "model-svg", bytes: preview.bytes.length },
        nativeRenderRequiredBeforeDelivery: true,
      },
      warnings: ["Run render_docx.py and inspect every page image before delivering the DOCX."],
    };
    await Promise.all([
      fs.writeFile(temporaryPath, output),
      fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, "utf8"),
    ]);
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const [inputPath, outputPath, auditPath, expectedText, replacementText, section = "0", referenceType = "default", mode = "edit"] = process.argv.slice(2);
  const result = await editDocumentWatermark({
    inputPath,
    outputPath,
    auditPath,
    expectedText,
    replacementText,
    sectionIndex: Number(section),
    referenceType,
    remove: mode === "remove",
  });
  console.log(JSON.stringify({ outputPath: result.outputPath, auditPath: result.auditPath, changedParts: result.audit.operation.changedParts }));
}
