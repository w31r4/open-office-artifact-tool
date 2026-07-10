import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { FileBlob, PdfArtifact, PdfFile } from "../src/index.mjs";

const pdf = PdfArtifact.create({ text: "PDF research artifact" });
assert.match(pdf.inspect().ndjson, /PDF research artifact/);
const preview = await pdf.render();
assert.equal(preview.type, "image/svg+xml");
assert.match(await preview.text(), /PDF research artifact/);

const blob = await PdfFile.exportPdf(pdf);
assert.equal(blob.type, "application/pdf");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pdf`);
await blob.save(out);
const loaded = await PdfFile.importPdf(await FileBlob.load(out));
assert.match(loaded.inspect().ndjson, /PDF research artifact/);
console.log("pdf smoke ok");
