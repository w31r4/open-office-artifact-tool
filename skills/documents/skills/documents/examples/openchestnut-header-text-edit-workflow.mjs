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
  if (typeof value !== "string" || !value.trim()) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

function escapeRegex(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function escapeXmlText(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

async function assertAbsent(filePath, label) {
  try {
    await fs.lstat(filePath);
    throw new Error(`${label} already exists; refusing to overwrite it.`);
  } catch (error) {
    if (error?.code !== "ENOENT") throw error;
  }
}

async function publishNoReplace(temporaryPath, finalPath) {
  await fs.link(temporaryPath, finalPath);
  await fs.rm(temporaryPath, { force: true });
}

async function packageParts(bytes) {
  const zip = await JSZip.loadAsync(bytes);
  const parts = new Map();
  for (const [name, entry] of Object.entries(zip.files)) {
    if (entry.dir) continue;
    parts.set(name, Buffer.from(await entry.async("uint8array")));
  }
  return parts;
}

async function changedParts(source, output) {
  const [before, after] = await Promise.all([packageParts(source), packageParts(output)]);
  if (before.size !== after.size || [...before.keys()].some((name) => !after.has(name))) {
    throw new Error("Source-bound header edit changed the DOCX package part inventory.");
  }
  return [...before.keys()].filter((name) => !before.get(name).equals(after.get(name))).sort();
}

async function headerXml(bytes, partPath) {
  const zip = await JSZip.loadAsync(bytes);
  const entry = zip.file(partPath);
  if (!entry) throw new Error(`Selected header part is missing from the DOCX package: ${partPath}.`);
  return entry.async("text");
}

function normalizeTargetText(xml, expectedText, label) {
  const escaped = escapeXmlText(expectedText);
  const pattern = new RegExp(`(<w:t(?:\\s[^>]*)?>)${escapeRegex(escaped)}(</w:t>)`, "g");
  let matches = 0;
  const normalized = xml.replace(pattern, (_whole, open, close) => {
    matches += 1;
    return `${open}__OPEN_CHESTNUT_TARGET_TEXT__${close}`;
  });
  if (matches !== 1) throw new Error(`${label} must occur in exactly one ordinary w:t node in the selected header part; found ${matches}.`);
  return normalized;
}

function headerSnapshot(header) {
  return {
    id: header.id,
    kind: header.kind,
    name: header.name,
    text: header.text,
    sectionIndex: header.sectionIndex,
    referenceType: header.referenceType,
    variantActive: header.variantActive,
    fieldInstruction: header.fieldInstruction,
    sourceBound: header.sourceBound,
    editable: header.editable,
    relationshipId: header.relationshipId,
    partPath: header.partPath,
  };
}

function footerSnapshot(footer) {
  return headerSnapshot(footer);
}

function stableDocumentProjection(document) {
  return {
    blocks: document.blocks.map((block) => ({ id: block.id, kind: block.kind, text: block.text })),
    headers: document.headers.map(headerSnapshot),
    footers: document.footers.map(footerSnapshot),
  };
}

function selectHeader(document, { sectionIndex, referenceType, expectedText }) {
  const matches = document.headers.filter((header) =>
    header.sectionIndex === sectionIndex
    && header.referenceType === referenceType
    && header.text === expectedText);
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one ${referenceType} header in section ${sectionIndex} with the requested text; found ${matches.length}.`);
  }
  const header = matches[0];
  if (!header.sourceBound || !header.editable) {
    throw new Error("The selected header does not advertise the narrow source-bound ordinary-text edit capability.");
  }
  if (document.resolve(header.id) !== header) throw new Error("The selected header locator did not resolve to the inspected object.");
  if (!/^word\/header\d+\.xml$/i.test(header.partPath || "")) throw new Error("The selected header has no canonical HeaderPart path.");
  return header;
}

async function modelRender(document) {
  const preview = await document.render({ format: "svg" });
  const svg = await preview.text();
  if (!/<svg\b/i.test(svg)) throw new Error("Document model render did not produce SVG.");
  return { renderer: "model-svg", bytes: preview.bytes.length };
}

function equalJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

export async function editImportedHeaderText({
  inputPath,
  outputPath,
  auditPath,
  expectedText,
  replacementText,
  sectionIndex = 0,
  referenceType = "default",
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct.");
  }
  const originalText = requiredText(expectedText, "expectedText");
  const nextText = requiredText(replacementText, "replacementText");
  const targetSection = Number(sectionIndex);
  if (!Number.isSafeInteger(targetSection) || targetSection < 0) throw new TypeError("sectionIndex must be a non-negative safe integer.");
  if (!new Set(["default", "first", "even"]).has(referenceType)) throw new TypeError("referenceType must be default, first, or even.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const sourceHash = sha256(source);
  const document = await DocumentFile.importDocx(new FileBlob(source, { type: DOCX_MIME, name: path.basename(sourcePath) }));
  const selected = selectHeader(document, { sectionIndex: targetSection, referenceType, expectedText: originalText });
  const before = headerSnapshot(selected);
  const beforeProjection = stableDocumentProjection(document);
  const sourceHeaderXml = await headerXml(source, before.partPath);
  const normalizedSourceHeaderXml = normalizeTargetText(sourceHeaderXml, originalText, "expectedText");
  selected.text = nextText;

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await DocumentFile.exportDocx(document);
    await fs.writeFile(temporaryPath, Buffer.from(await exported.arrayBuffer()), { flag: "wx" });
    const output = await fs.readFile(temporaryPath);
    if (sha256(await fs.readFile(sourcePath)) !== sourceHash) throw new Error("Source DOCX changed during the transaction; refusing publication.");
    const changed = await changedParts(source, output);
    if (!equalJson(changed, [before.partPath])) {
      throw new Error(`Source-bound header edit changed an unexpected package scope: ${changed.join(", ") || "none"}.`);
    }
    const outputHeaderXml = await headerXml(output, before.partPath);
    const normalizedOutputHeaderXml = normalizeTargetText(outputHeaderXml, nextText, "replacementText");
    if (normalizedOutputHeaderXml !== normalizedSourceHeaderXml) {
      throw new Error("Header edit changed XML outside the one requested ordinary w:t payload.");
    }

    const reimported = await DocumentFile.importDocx(new FileBlob(output, { type: DOCX_MIME, name: path.basename(finalPath) }));
    const roundTrip = selectHeader(reimported, { sectionIndex: targetSection, referenceType, expectedText: nextText });
    const afterProjection = stableDocumentProjection(reimported);
    const expectedProjection = structuredClone(beforeProjection);
    const expectedHeader = expectedProjection.headers.find((header) => header.id === before.id);
    if (!expectedHeader) throw new Error("Selected header disappeared from the imported document projection.");
    expectedHeader.text = nextText;
    if (!equalJson(afterProjection, expectedProjection)) {
      throw new Error("DOCX export changed body, page-furniture, or source-bound header identity outside the requested text.");
    }
    if (roundTrip.id !== before.id || roundTrip.partPath !== before.partPath || !roundTrip.sourceBound || !roundTrip.editable) {
      throw new Error("Second import did not preserve the selected header identity or edit capability.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Document verification failed: ${verification.ndjson}`);
    const render = await modelRender(reimported);
    const audit = {
      schema: "open-office-artifact-tool.docx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sourceHash, bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite", noReplace: true },
      operation: {
        type: "source-bound-header-text-edit",
        target: {
          id: before.id,
          sectionIndex: before.sectionIndex,
          referenceType: before.referenceType,
          relationshipId: before.relationshipId,
          partPath: before.partPath,
        },
        sourceTextSha256: sha256(Buffer.from(originalText, "utf8")),
        replacementTextSha256: sha256(Buffer.from(nextText, "utf8")),
      },
      validation: {
        changedParts: changed,
        headerXmlResidual: { ok: true, normalizedSha256: sha256(Buffer.from(normalizedSourceHeaderXml, "utf8")) },
        reimport: { ok: true, headerId: roundTrip.id, partPath: roundTrip.partPath, sourceBound: true, editable: true },
        verify: { ok: true },
        modelRender: { ok: true, ...render },
        nativeRenderRequired: true,
      },
      warnings: ["Run render_docx.py and inspect every affected page image before delivery; PAGE/simple fields and other page furniture remain source-owned."],
    };
    await fs.writeFile(temporaryAuditPath, `${JSON.stringify(audit, null, 2)}\n`, { flag: "wx" });
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
  const [
    inputPath,
    outputPath,
    auditPath,
    expectedText = "Northwind | Internal",
    replacementText = "Northwind | Reviewed",
    sectionIndex = "0",
    referenceType = "default",
  ] = argv;
  return { inputPath, outputPath, auditPath, expectedText, replacementText, sectionIndex: Number(sectionIndex), referenceType };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editImportedHeaderText(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    changedParts: result.audit.validation.changedParts,
  }));
}
