import assert from "node:assert/strict";
import {
  DocumentFile,
  DocumentModel,
  OFFICE_CODEC_IDS,
  Presentation,
  PresentationFile,
  SpreadsheetFile,
  Workbook,
} from "../src/index.mjs";

assert.deepEqual(OFFICE_CODEC_IDS, ["javascript", "open-chestnut"]);
assert.ok(Object.isFrozen(OFFICE_CODEC_IDS));

const workbook = Workbook.create();
workbook.worksheets.add("Facade").getRange("A1:B2").values = [["Codec", "Value"], ["policy", 7]];

await assert.rejects(
  SpreadsheetFile.exportXlsx(workbook, { codec: "walnut" }),
  /codec must be javascript or open-chestnut/i,
);
await assert.rejects(
  SpreadsheetFile.exportXlsx(workbook, { codec: "openxml-wasm" }),
  /deprecated openxml-wasm name is available only through its compatibility subpath/i,
);
await assert.rejects(
  PresentationFile.importPptx(new Uint8Array(), "open-chestnut"),
  /options must be an object/i,
);

const javascriptXlsx = await SpreadsheetFile.exportXlsx(workbook, { codec: "javascript" });
assert.equal(javascriptXlsx.metadata.codec, "javascript");
assert.equal((await SpreadsheetFile.importXlsx(javascriptXlsx, { codec: "javascript" })).worksheets.getItem("Facade").getRange("B2").values[0][0], 7);

const openChestnutXlsx = await SpreadsheetFile.exportXlsx(workbook, { codec: "open-chestnut" });
assert.equal(openChestnutXlsx.metadata.codec, "open-chestnut");
assert.equal((await SpreadsheetFile.importXlsx(openChestnutXlsx, { codec: "open-chestnut" })).worksheets.getItem("Facade").getRange("B2").values[0][0], 7);
await assert.rejects(
  SpreadsheetFile.importXlsx(openChestnutXlsx, { codec: "open-chestnut", limits: { maxInputBytes: 1 } }),
  /input.*bytes|limit|max_input_bytes/i,
);

const document = DocumentModel.create({ paragraphs: ["Facade codec policy"] });
const javascriptDocx = await DocumentFile.exportDocx(document, { codec: "javascript" });
assert.equal(javascriptDocx.metadata.codec, "javascript");
assert.equal((await DocumentFile.importDocx(javascriptDocx, { codec: "javascript" })).blocks[0].text, "Facade codec policy");
const openChestnutDocx = await DocumentFile.exportDocx(document, { codec: "open-chestnut" });
assert.equal(openChestnutDocx.metadata.codec, "open-chestnut");
assert.equal((await DocumentFile.importDocx(openChestnutDocx, { codec: "open-chestnut" })).blocks[0].text, "Facade codec policy");

const presentation = Presentation.create();
presentation.slides.add({ name: "Facade" }).shapes.add({
  name: "Policy",
  text: "Facade codec policy",
  position: { left: 40, top: 40, width: 400, height: 80 },
});
const javascriptPptx = await PresentationFile.exportPptx(presentation, { codec: "javascript" });
assert.equal(javascriptPptx.metadata.codec, "javascript");
assert.equal((await PresentationFile.importPptx(javascriptPptx, { codec: "javascript" })).slides.getItem(0).shapes.items[0].text.value, "Facade codec policy");
const openChestnutPptx = await PresentationFile.exportPptx(presentation, { codec: "open-chestnut" });
assert.equal(openChestnutPptx.metadata.codec, "open-chestnut");
assert.equal((await PresentationFile.importPptx(openChestnutPptx, { codec: "open-chestnut" })).slides.getItem(0).shapes.items[0].text.value, "Facade codec policy");

console.log("Office facade codec policy smoke ok");
