import assert from "node:assert/strict";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";

import { DocumentModel, renderArtifact } from "open-office-artifact-tool";
import {
  callOfficeBridge,
  createNativeOfficeRenderer,
  nativeOfficeStatus,
  OfficeBridgeError,
  renderFileWithNativeOffice,
} from "open-office-artifact-tool/native/office-bridge";

const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "open-office-bridge-test-"));
const mockBridge = path.join(tempDir, "mock-bridge.mjs");
await fs.writeFile(mockBridge, `
import fs from 'node:fs/promises';
let input = '';
process.stdin.setEncoding('utf8');
for await (const chunk of process.stdin) input += chunk;
const request = JSON.parse(input);
if (request.operation === 'status') {
  console.log(JSON.stringify({ ok: true, available: true, bridge: 'mock-office', officeInstalled: false }));
  process.exit(0);
}
if (request.operation === 'render') {
  const bytes = request.outputType === 'image/png' ? new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a]) : new TextEncoder().encode('%PDF-mock');
  await fs.writeFile(request.outputPath, bytes);
  console.log(JSON.stringify({ ok: true, outputPath: request.outputPath, outputType: request.outputType, bridge: 'mock-office', metadata: { operation: request.operation, inputType: request.inputType } }));
  process.exit(0);
}
console.log(JSON.stringify({ ok: false, error: { code: 'UNSUPPORTED', message: 'unsupported operation' } }));
`, "utf8");

const bridgeOptions = { command: process.execPath, args: [mockBridge], timeoutMs: 10_000 };
const status = await nativeOfficeStatus(bridgeOptions);
assert.equal(status.ok, true);
assert.equal(status.available, true);

const direct = await callOfficeBridge({ operation: "status" }, bridgeOptions);
assert.equal(direct.bridge, "mock-office");

const document = DocumentModel.create({ paragraphs: ["Native bridge wrapper smoke"] });
const svg = await document.render();
const rendered = await renderFileWithNativeOffice(svg, { ...bridgeOptions, artifactKind: "document", inputType: svg.type, outputType: "application/pdf", format: "pdf" });
assert.equal(rendered.type, "application/pdf");
assert.equal(rendered.metadata.renderer, "native-office");
assert.equal(rendered.metadata.bridge, "mock-office");
assert.match(await rendered.text(), /%PDF-mock/);

const renderer = createNativeOfficeRenderer(bridgeOptions);
const png = await renderArtifact(document, { format: "png", renderer });
assert.equal(png.type, "image/png");
assert.equal(png.metadata.renderer, "native-office");
assert.deepEqual([...png.bytes.slice(0, 4)], [0x89, 0x50, 0x4e, 0x47]);

await assert.rejects(
  async () => callOfficeBridge({ operation: "status" }, { timeoutMs: 10 }),
  (error) => error instanceof OfficeBridgeError && error.code === "OFFICE_BRIDGE_NOT_CONFIGURED",
);

await fs.rm(tempDir, { recursive: true, force: true });
console.log("office bridge wrapper smoke ok");
