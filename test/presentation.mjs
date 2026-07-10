import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import { FileBlob, Presentation, PresentationFile } from "../src/index.mjs";

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = presentation.slides.add();
const shape = slide.shapes.add({
  geometry: "roundRect",
  name: "summary-surface",
  position: { left: 72, top: 96, width: 420, height: 140 },
  fill: "white",
  line: { fill: "#cbd5e1", width: 1 },
});
shape.text = "Revenue outlook";
shape.text.style = { fontSize: 32, bold: true, color: "#0f172a" };

const snapshot = presentation.inspect({ kind: "deck,slide,textbox,layout", maxChars: 4000 }).ndjson;
assert.match(snapshot, /"kind":"textbox"/);
assert.match(snapshot, /Revenue outlook/);
const target = presentation.resolve(shape.id);
assert.equal(target.text.value, "Revenue outlook");
target.text.replace("outlook", "plan");
assert.match(presentation.inspect({ kind: "textbox" }).ndjson, /Revenue plan/);

const layout = await slide.export({ format: "layout" });
assert.equal(layout.type, "application/vnd.open-office-artifact.layout+json");
assert.match(await layout.text(), /summary-surface/);
const preview = await presentation.export({ slide, format: "svg" });
assert.equal(preview.type, "image/svg+xml");

const pptx = await PresentationFile.exportPptx(presentation);
const out = path.join(os.tmpdir(), `open-office-artifact-${process.pid}.pptx`);
await pptx.save(out);
const loaded = await PresentationFile.importPptx(await FileBlob.load(out));
assert.match(loaded.inspect({ kind: "textbox" }).ndjson, /Revenue plan/);
console.log("presentation smoke ok");
