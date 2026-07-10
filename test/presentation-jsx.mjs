import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";

import { FileBlob, Presentation, PresentationFile } from "open-office-artifact-tool";
import { Fragment, grid as gridNode, paragraph as paragraphNode } from "open-office-artifact-tool/presentation-jsx";
import { jsx, jsxs } from "open-office-artifact-tool/presentation-jsx/jsx-runtime";
import { jsxDEV } from "open-office-artifact-tool/presentation-jsx/jsx-dev-runtime";

function MetricCard({ name, label, children }) {
  return jsxs("box", {
    name,
    width: "fill",
    height: "fill",
    fill: "slate-50",
    padding: { x: 12, y: 10 },
    children: [
      jsx("paragraph", {
        name: `${name}-label`,
        className: "text-slate-700 text-lg font-bold",
        children: label,
      }),
      jsx("paragraph", {
        name: `${name}-value`,
        className: "text-slate-950 text-3xl font-bold",
        children,
      }),
    ],
  });
}

const presentation = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = presentation.slides.add();

const tree = jsxs(Fragment, {
  children: [
    jsxs("column", {
      name: "jsx-content-frame",
      width: "fill",
      height: "fill",
      gap: 18,
      padding: { x: 32, y: 28 },
      children: [
        jsxs("paragraph", {
          id: "sh/jsx-headline",
          name: "jsx-headline",
          className: "text-slate-950 text-4xl font-bold leading-tight",
          children: ["JSX ", jsx("run", { textStyle: { bold: true, color: "sky-600" }, children: "runtime" })],
        }),
        jsxs("row", {
          name: "jsx-metrics-row",
          width: "fill",
          height: 140,
          gap: 14,
          children: [
            jsx(MetricCard, { name: "jsx-card-a", label: "Pipeline", children: "$4.2M" }),
            jsx(MetricCard, { name: "jsx-card-b", label: "Coverage", children: "3.1x" }),
          ],
        }),
        jsxDEV("shape", {
          name: "jsx-accent-pill",
          width: 180,
          height: 44,
          geometry: "roundRect",
          fill: "sky-100",
          className: "text-slate-950 text-lg font-bold",
          children: "Inspectable",
        }),
      ],
    }),
  ],
});

const materialized = slide.compose(tree, { frame: { left: 72, top: 84, width: 780, height: 400 } });
assert.ok(materialized.length >= 7);

const inspect = presentation.inspect({ kind: "textbox,shape", maxChars: 10000 }).ndjson;
assert.match(inspect, /jsx-headline/);
assert.match(inspect, /JSX runtime/);
assert.match(inspect, /jsx-card-a-value/);
assert.match(inspect, /Inspectable/);
assert.equal(presentation.resolve("sh/jsx-headline").text.value, "JSX runtime");

const layout = JSON.parse(await (await slide.export({ format: "layout" })).text());
assert.ok(layout.elements.some((element) => element.name === "jsx-accent-pill"));

const gridSlide = presentation.slides.add();
gridSlide.compose(
  gridNode({
    name: "grid-root",
    columns: [{ mode: "fixed", value: 100 }, { mode: "fr", value: 1 }, { mode: "fixed", value: 100 }],
    rows: [{ mode: "fixed", value: 50 }, { mode: "fr", value: 1 }],
    columnGap: 10,
    rowGap: 20,
  }, [
    paragraphNode({ name: "grid-a", columnSpan: 2 }, ["A"]),
    jsx("paragraph", { name: "grid-b", column: 2, row: 0, children: "B" }),
    jsx("paragraph", { name: "grid-c", column: 0, row: 1, columnSpan: 3, children: "C" }),
  ]),
  { frame: { left: 10, top: 20, width: 420, height: 220 } },
);
const gridLayout = JSON.parse(await (await gridSlide.export({ format: "layout" })).text());
const gridA = gridLayout.elements.find((element) => element.name === "grid-a");
const gridB = gridLayout.elements.find((element) => element.name === "grid-b");
const gridC = gridLayout.elements.find((element) => element.name === "grid-c");
assert.deepEqual(gridA.frame, { left: 10, top: 20, width: 310, height: 50 });
assert.deepEqual(gridB.frame, { left: 330, top: 20, width: 100, height: 50 });
assert.deepEqual(gridC.frame, { left: 10, top: 90, width: 420, height: 150 });
assert.match(presentation.inspect({ kind: "textbox", maxChars: 10000 }).ndjson, /grid-c/);

const out = path.join(os.tmpdir(), `open-office-artifact-jsx-${process.pid}.pptx`);
await (await PresentationFile.exportPptx(presentation)).save(out);
const loaded = await PresentationFile.importPptx(await FileBlob.load(out));
assert.match(loaded.inspect({ kind: "textbox", maxChars: 10000 }).ndjson, /JSX runtime/);
console.log("presentation-jsx smoke ok");
