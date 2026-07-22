import crypto from "node:crypto";
import fs from "node:fs/promises";
import { createRequire } from "node:module";
import path from "node:path";
import { pathToFileURL } from "node:url";

import JSZip from "jszip";
import { FileBlob, PresentationFile } from "open-office-artifact-tool";

const PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
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

function isMissing(error) {
  return error?.code === "ENOENT";
}

async function assertAbsent(target, label) {
  try {
    await fs.lstat(target);
  } catch (error) {
    if (isMissing(error)) return;
    throw error;
  }
  throw new Error(`${label} already exists: ${target}`);
}

function nodeById(diagramText, nodeId) {
  const nodes = diagramText?.nodes;
  if (!Array.isArray(nodes)) return undefined;
  return nodes.find((node) => node.id === nodeId);
}

function targetCandidates(presentation, objectName, nodeId, expectedText) {
  return presentation.slides.items.flatMap((slide, slideIndex) => slide.nativeObjects.items
    .filter((object) => object.nativeKind === "diagram" && object.name === objectName)
    .map((object) => ({ slide, slideIndex, object, node: nodeById(object.diagramText, nodeId) }))
    .filter((candidate) => candidate.node?.text === expectedText));
}

function graphSnapshot(object) {
  const diagram = object.diagramText;
  if (!diagram || !Array.isArray(diagram.nodes) || !diagram.nodes.length) {
    throw new Error("Target SmartArt does not expose the bounded plain node-text capability.");
  }
  if (object.nativeKind !== "diagram" || object.parts.length !== 4 || object.parts.some((part) => part.relationships.length !== 0)) {
    throw new Error("Target SmartArt is not the required closed four-part native graph.");
  }
  return {
    nativeKind: object.nativeKind,
    partPath: diagram.partPath,
    contentType: diagram.contentType,
    sourceSha256: diagram.sourceSha256,
    relationshipId: diagram.relationshipId,
    nodes: diagram.nodes.map((node) => ({ id: node.id, text: node.text })),
    nativePartPaths: object.parts.map((part) => part.path).sort(),
    relationshipIds: object.rootRelationships.map((relationship) => relationship.id).sort(),
  };
}

async function assertPackageScope(sourceBytes, outputBytes, dataPartPath) {
  const [sourceZip, outputZip] = await Promise.all([JSZip.loadAsync(sourceBytes), JSZip.loadAsync(outputBytes)]);
  const sourcePaths = Object.keys(sourceZip.files).filter((path) => !sourceZip.files[path].dir).sort();
  const outputPaths = Object.keys(outputZip.files).filter((path) => !outputZip.files[path].dir).sort();
  if (JSON.stringify(sourcePaths) !== JSON.stringify(outputPaths)) {
    throw new Error("SmartArt text edit changed the PPTX package part topology.");
  }
  let dataPartChanged = false;
  for (const partPath of sourcePaths) {
    const [before, after] = await Promise.all([
      sourceZip.file(partPath).async("uint8array"),
      outputZip.file(partPath).async("uint8array"),
    ]);
    if (Buffer.from(before).equals(Buffer.from(after))) continue;
    if (partPath !== dataPartPath) {
      throw new Error(`SmartArt text edit changed an unowned package part: ${partPath}`);
    }
    dataPartChanged = true;
  }
  if (!dataPartChanged) throw new Error("SmartArt text edit did not change its bound DiagramDataPart.");
  return { partCount: sourcePaths.length, changedPartPaths: [dataPartPath] };
}

export async function editPptxSmartArtNodeText({
  inputPath,
  outputPath,
  auditPath,
  objectName,
  nodeId,
  expectedText,
  replacementText,
}) {
  const sourcePath = path.resolve(requiredText(inputPath, "inputPath"));
  const finalPath = path.resolve(requiredText(outputPath, "outputPath"));
  const finalAuditPath = path.resolve(requiredText(auditPath, "auditPath"));
  const targetName = requiredText(objectName, "objectName");
  const targetNodeId = requiredText(nodeId, "nodeId");
  const originalText = requiredString(expectedText, "expectedText");
  const nextText = requiredString(replacementText, "replacementText");
  if (sourcePath === finalPath || sourcePath === finalAuditPath || finalPath === finalAuditPath) {
    throw new Error("inputPath, outputPath, and auditPath must be distinct so the source remains immutable.");
  }
  if (originalText === nextText) throw new Error("replacementText must differ from expectedText.");
  await Promise.all([assertAbsent(finalPath, "outputPath"), assertAbsent(finalAuditPath, "auditPath")]);

  const source = await fs.readFile(sourcePath);
  const presentation = await PresentationFile.importPptx(new FileBlob(source, { type: PPTX_MIME, name: path.basename(sourcePath) }));
  const candidates = targetCandidates(presentation, targetName, targetNodeId, originalText);
  if (candidates.length !== 1) {
    throw new Error(`Expected exactly one source-bound SmartArt ${JSON.stringify(targetName)} node ${JSON.stringify(targetNodeId)} with the requested text; found ${candidates.length}.`);
  }
  const target = candidates[0];
  const before = graphSnapshot(target.object);
  target.object.setDiagramNodeText(targetNodeId, nextText);

  const temporaryPath = `${finalPath}.tmp-${process.pid}-${Date.now()}`;
  const temporaryAuditPath = `${finalAuditPath}.tmp-${process.pid}-${Date.now()}`;
  await Promise.all([fs.mkdir(path.dirname(finalPath), { recursive: true }), fs.mkdir(path.dirname(finalAuditPath), { recursive: true })]);
  try {
    const exported = await PresentationFile.exportPptx(presentation);
    await exported.save(temporaryPath);
    const output = await fs.readFile(temporaryPath);
    const packageScope = await assertPackageScope(source, output, before.partPath);
    const reimported = await PresentationFile.importPptx(new FileBlob(output, { type: PPTX_MIME, name: path.basename(finalPath) }));
    const rebound = targetCandidates(reimported, targetName, targetNodeId, nextText);
    if (rebound.length !== 1) throw new Error("Saved PPTX does not expose exactly one edited SmartArt node after reimport.");
    const after = graphSnapshot(rebound[0].object);
    if (after.partPath !== before.partPath || after.contentType !== before.contentType || after.relationshipId !== before.relationshipId ||
        JSON.stringify(after.nativePartPaths) !== JSON.stringify(before.nativePartPaths) ||
        JSON.stringify(after.relationshipIds) !== JSON.stringify(before.relationshipIds) ||
        after.sourceSha256 === before.sourceSha256) {
      throw new Error("Saved SmartArt graph did not retain its source-bound part/relationship contract.");
    }
    const expectedNodes = before.nodes.map((node) => node.id === targetNodeId ? { ...node, text: nextText } : node);
    if (JSON.stringify(after.nodes) !== JSON.stringify(expectedNodes)) throw new Error("Saved SmartArt node list changed outside the requested text edit.");
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
        type: "source-bound-smartart-node-text-edit",
        slideIndex: target.slideIndex,
        objectName: targetName,
        nodeId: targetNodeId,
        expectedText: originalText,
        replacementText: nextText,
        dataPart: before.partPath,
        relationshipId: before.relationshipId,
      },
      warnings: [],
      validation: {
        package: { ok: true, ...packageScope, nonTargetPartsByteIdentical: true },
        reimport: { ok: true, graphContractPreserved: true, nodeTopologyPreserved: true },
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
  const [inputPath, outputPath, auditPath, objectName, nodeId, expectedText, replacementText] = argv;
  return { inputPath, outputPath, auditPath, objectName, nodeId, expectedText, replacementText };
}

const entry = process.argv[1] ? pathToFileURL(path.resolve(process.argv[1])).href : "";
if (entry === import.meta.url) {
  const result = await editPptxSmartArtNodeText(parseCli(process.argv.slice(2)));
  console.log(JSON.stringify({
    outputPath: result.outputPath,
    auditPath: result.auditPath,
    outputSha256: result.audit.output.sha256,
    dataPart: result.audit.operation.dataPart,
  }));
}
