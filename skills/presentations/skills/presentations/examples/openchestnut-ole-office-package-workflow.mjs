import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { DocumentFile, FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const require = createRequire(import.meta.url);

function sha256(bytes) {
  return crypto.createHash("sha256").update(bytes).digest("hex");
}

async function packageVersion() {
  const entry = require.resolve("open-office-artifact-tool");
  const packagePath = path.join(path.dirname(path.dirname(entry)), "package.json");
  return JSON.parse(await fs.readFile(packagePath, "utf8")).version;
}

function requiredText(value, label) {
  if (typeof value !== "string" || !value.length) throw new TypeError(`${label} must be a non-empty string.`);
  return value;
}

function requiredString(value, label) {
  if (typeof value !== "string") throw new TypeError(`${label} must be a string.`);
  return value;
}

async function assertAbsent(target, label) {
  try {
    await fs.lstat(target);
  } catch (error) {
    if (error?.code === "ENOENT") return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}`);
}

function officePackageCandidates(presentation, objectName) {
  return presentation.slides.items.flatMap((slide, slideIndex) => slide.nativeObjects.items
    .filter((object) => object.nativeKind === "oleObject" && object.name === objectName && object.oleOfficePackage?.kind === "docx")
    .map((object) => ({ slide, slideIndex, object })));
}

function packageBindingSnapshot(object) {
  const binding = object.oleOfficePackage;
  if (!binding || binding.kind !== "docx" || binding.contentType !== DOCX_MIME) {
    throw new Error("Target does not expose the bounded DOCX Office-package profile.");
  }
  if (object.nativeKind !== "oleObject") throw new Error("Target must remain a top-level native OLE object.");
  return {
    kind: binding.kind,
    partPath: binding.partPath,
    contentType: binding.contentType,
    sourceSha256: binding.sourceSha256,
    relationshipId: binding.relationshipId,
    nativePartPaths: object.parts.map((part) => part.path).sort(),
    relationshipIds: object.rootRelationships.map((relationship) => relationship.id).sort(),
  };
}

function uniqueEditableParagraph(document, expectedText) {
  const matches = document.blocks.filter((block) => block.kind === "paragraph" && block.text === expectedText);
  if (matches.length !== 1) {
    throw new Error(`Expected exactly one ordinary embedded-DOCX paragraph with ${JSON.stringify(expectedText)}; found ${matches.length}.`);
  }
  const paragraph = matches[0];
  if (!Array.isArray(paragraph.runs) || paragraph.runs.length !== 1 || paragraph.runs[0].text !== expectedText) {
    throw new Error("Embedded DOCX paragraph must be the bounded one-run plain-text profile; rich or fragmented text must use a dedicated document workflow.");
  }
  return paragraph;
}

async function assertPackageScope(sourceBytes, outputBytes, packagePartPath, replacementDocx) {
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBytes), JSZip.loadAsync(outputBytes)]);
  const sourcePaths = Object.keys(sourceZip.files).filter((partPath) => !sourceZip.files[partPath].dir).sort();
  const outputPaths = Object.keys(outputZip.files).filter((partPath) => !outputZip.files[partPath].dir).sort();
  if (JSON.stringify(sourcePaths) !== JSON.stringify(outputPaths)) {
    throw new Error("Embedded DOCX replacement changed the PPTX package part topology.");
  }
  const changedPartPaths = [];
  for (const partPath of sourcePaths) {
    const [before, after] = await Promise.all([
      sourceZip.file(partPath).async("uint8array"),
      outputZip.file(partPath).async("uint8array"),
    ]);
    if (Buffer.from(before).equals(Buffer.from(after))) continue;
    if (partPath !== packagePartPath) {
      throw new Error(`Embedded DOCX replacement changed an unowned PPTX package part: ${partPath}`);
    }
    changedPartPaths.push(partPath);
    if (!Buffer.from(after).equals(Buffer.from(replacementDocx.bytes))) {
      throw new Error("Embedded DOCX replacement bytes do not match the validated DocumentFile export.");
    }
  }
  if (JSON.stringify(changedPartPaths) !== JSON.stringify([packagePartPath])) {
    throw new Error("Embedded DOCX replacement did not change exactly its bound package part.");
  }
  return { partCount: sourcePaths.length, changedPartPaths };
}

async function modelSvg(slide) {
  const blob = await slide.export({ format: "svg" });
  const text = await blob.text();
  if (!/<svg\b/i.test(text)) throw new Error("Presentation model render did not produce SVG.");
  return { sha256: sha256(blob.bytes), bytes: blob.bytes.length };
}

export async function editPptxEmbeddedDocxPackage({
  inputPath,
  outputPath,
  auditPath,
  objectName,
  expectedText,
  replacementText,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const targetName = requiredText(objectName, "objectName");
  const originalText = requiredString(expectedText, "expectedText");
  const nextText = requiredString(replacementText, "replacementText");
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct so the source remains immutable.");
  }
  if (originalText === nextText) throw new Error("replacementText must differ from expectedText.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const candidates = officePackageCandidates(presentation, targetName);
  if (candidates.length !== 1) {
    throw new Error(`Expected exactly one source-bound DOCX OLE object named ${JSON.stringify(targetName)}; found ${candidates.length}.`);
  }
  const target = candidates[0];
  const before = packageBindingSnapshot(target.object);
  const sourceDocx = target.object.getEmbeddedOfficePackage();
  if (sourceDocx.type !== DOCX_MIME || sourceDocx.metadata.officePackageKind !== "docx") {
    throw new Error("Target FileBlob does not retain the expected DOCX MIME type and package-kind metadata.");
  }
  const document = await DocumentFile.importDocx(sourceDocx);
  const paragraph = uniqueEditableParagraph(document, originalText);
  paragraph.text = nextText;
  paragraph.runs[0].text = nextText;
  const replacementDocx = await DocumentFile.exportDocx(document);
  if (replacementDocx.type !== DOCX_MIME) throw new Error("DocumentFile did not produce a DOCX FileBlob.");
  target.object.replaceEmbeddedOfficePackage(replacementDocx);
  if (target.object.getEmbeddedOfficePackage().metadata.pendingReplacement !== true) {
    throw new Error("Expected a pending embedded Office-package replacement.");
  }
  const sourceRender = await modelSvg(target.slide);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, before.partPath, replacementDocx);
    const sourceAfter = await fs.readFile(sourcePath);
    if (!Buffer.from(sourceAfter).equals(Buffer.from(source))) throw new Error("Source PPTX changed during the replacement transaction.");
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const reboundCandidates = officePackageCandidates(reimported, targetName);
    if (reboundCandidates.length !== 1) throw new Error("Saved PPTX does not expose exactly one rebound DOCX OLE object.");
    const rebound = reboundCandidates[0];
    const after = packageBindingSnapshot(rebound.object);
    if (after.kind !== before.kind || after.partPath !== before.partPath || after.contentType !== before.contentType ||
        after.relationshipId !== before.relationshipId || after.sourceSha256 === before.sourceSha256 ||
        JSON.stringify(after.nativePartPaths) !== JSON.stringify(before.nativePartPaths) ||
        JSON.stringify(after.relationshipIds) !== JSON.stringify(before.relationshipIds)) {
      throw new Error("Saved DOCX OLE graph did not retain its source-bound shell/relationship contract.");
    }
    const reboundDocx = rebound.object.getEmbeddedOfficePackage();
    if (reboundDocx.metadata.pendingReplacement === true || sha256(reboundDocx.bytes) !== after.sourceSha256) {
      throw new Error("Rebound DOCX package does not expose a settled source digest.");
    }
    const reboundDocument = await DocumentFile.importDocx(reboundDocx);
    if (uniqueEditableParagraph(reboundDocument, nextText).text !== nextText) {
      throw new Error("Rebound embedded DOCX does not retain the requested paragraph replacement.");
    }
    const outputRender = await modelSvg(rebound.slide);
    if (outputRender.sha256 !== sourceRender.sha256) {
      throw new Error("Embedded DOCX replacement changed the containing slide's model SVG render.");
    }
    const verification = reimported.verify({ visualQa: true });
    if (!verification.ok) throw new Error(`Presentation verification failed: ${verification.ndjson}`);
    const audit = {
      schema: "open-office-artifact-tool.pptx-audit.v1",
      status: "succeeded",
      source: { path: sourcePath, sha256: sha256(source), bytes: source.length },
      output: { path: finalPath, sha256: sha256(output), bytes: output.length },
      provider: { actual: "open-chestnut", version: await packageVersion(), silentFallback: false },
      savePolicy: { strategy: "rewrite" },
      operation: {
        type: "source-bound-ole-docx-package-paragraph-edit",
        slideIndex: target.slideIndex,
        objectName: targetName,
        expectedText: originalText,
        replacementText: nextText,
        package: before,
      },
      warnings: ["The preserved OLE preview image is not regenerated from the replacement DOCX."],
      validation: {
        package: { ok: true, ...packageScope, nonTargetPartsByteIdentical: true, exactReplacementBytes: true },
        reimport: { ok: true, sourceBindingReproved: true, replacementTextVerified: true },
        modelRender: { ok: true, renderer: "model-svg", sourceSha256: sourceRender.sha256, outputSha256: outputRender.sha256, byteIdentical: true },
        verify: { ok: verification.ok },
      },
    };
    await fs.writeFile(temporaryAuditPath, JSON.stringify(audit, null, 2));
    await fs.rename(temporaryPath, finalPath);
    await fs.rename(temporaryAuditPath, finalAuditPath);
    return { outputPath: finalPath, auditPath: finalAuditPath, audit };
  } catch (error) {
    await Promise.all([fs.rm(temporaryPath, { force: true }), fs.rm(temporaryAuditPath, { force: true })]);
    throw error;
  }
}

function parseCli(argv) {
  const [inputPath, outputPath, auditPath, objectName, expectedText, replacementText] = argv;
  return { inputPath, outputPath, auditPath, objectName, expectedText, replacementText };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxEmbeddedDocxPackage(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    packagePart: result.audit.operation.package.partPath,
  }));
}
