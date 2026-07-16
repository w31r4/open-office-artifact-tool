import assert from "node:assert/strict";
import {
  DocumentFile,
  DocumentModel,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

const workbook = Workbook.create();
workbook.worksheets.add("Facade").getRange("A1:B2").values = [["Path", "Value"], ["OpenChestnut", 7]];

for (const options of [
  { codec: "open-chestnut" },
  { codec: "javascript" },
  { allowLossy: true },
  { preferNative: true },
  { relativeDateAsOf: "2026-07-16" },
]) {
  await assert.rejects(
    SpreadsheetFile.exportXlsx(workbook, options),
    /does not accept option|only Office codec|lossy fallback/i,
  );
}
await assert.rejects(PresentationFile.importPptx(new Uint8Array(), "open-chestnut"), /options must be an object/i);

const xlsx = await SpreadsheetFile.exportXlsx(workbook, { recalculate: false });
assert.equal(xlsx.metadata.codec, "open-chestnut");
assert.equal((await SpreadsheetFile.importXlsx(xlsx)).worksheets.getItem("Facade").getRange("B2").values[0][0], 7);
await assert.rejects(SpreadsheetFile.importXlsx(xlsx, { limits: { maxInputBytes: 1 } }), /input.*bytes|limit|max_input_bytes/i);

const document = DocumentModel.create({ paragraphs: ["Single Office codec"] });
const docx = await DocumentFile.exportDocx(document);
assert.equal(docx.metadata.codec, "open-chestnut");
assert.equal((await DocumentFile.importDocx(docx)).blocks[0].text, "Single Office codec");
await assert.rejects(DocumentFile.exportDocx(document, { allowLossy: true }), /does not accept option|lossy fallback/i);
await assert.rejects(DocumentFile.importDocx(docx, { preferNative: true }), /does not accept option|only Office codec/i);

const presentation = Presentation.create();
presentation.slides.add({ name: "Facade" }).shapes.add({
  name: "Policy",
  text: "Single Office codec",
  position: { left: 40, top: 40, width: 400, height: 80 },
});
const pptx = await PresentationFile.exportPptx(presentation);
assert.equal(pptx.metadata.codec, "open-chestnut");
assert.equal((await PresentationFile.importPptx(pptx)).slides.getItem(0).shapes.items[0].text.value, "Single Office codec");
await assert.rejects(PresentationFile.exportPptx(presentation, { codec: "javascript" }), /does not accept option|only Office codec/i);
await assert.rejects(PresentationFile.importPptx(pptx, { allowLossy: true }), /does not accept option|lossy fallback/i);

console.log("single OpenChestnut Office facade smoke ok");
