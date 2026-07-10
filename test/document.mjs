import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { DocumentFile, DocumentModel, FileBlob } from "../src/index.mjs";

const document = DocumentModel.create({
  name: "Research memo",
  paragraphs: ["Research memo", "This document exercises the clean-room DOCX facade."],
});
assert.match(document.inspect().ndjson, /Research memo/);
assert.match(document.help("export").ndjson, /DocumentFile.exportDocx/);

const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.type, "application/vnd.openxmlformats-officedocument.wordprocessingml.document");
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.docx`);
await docx.save(out);
const loaded = await DocumentFile.importDocx(await FileBlob.load(out));
assert.match(loaded.inspect().ndjson, /clean-room DOCX facade/);
console.log("document smoke ok");
