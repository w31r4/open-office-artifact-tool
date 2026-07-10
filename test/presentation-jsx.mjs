import assert from "node:assert/strict";
import os from "node:os";
import path from "node:path";
import JSZip from "jszip";

import { FileBlob, Presentation, PresentationFile } from "open-office-artifact-tool";
import { Fragment, chart as chartNode, grid as gridNode, image as imageNode, paragraph as paragraphNode, table as tableNode } from "open-office-artifact-tool/presentation-jsx";
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

const dataSlide = presentation.slides.add();
const directTable = dataSlide.tables.add({
  name: "direct-kpi-table",
  rows: 2,
  columns: 3,
  position: { left: 40, top: 40, width: 360, height: 96 },
  values: [["Metric", "Now", "Delta"], ["Revenue", "$12M", "+18%"]],
  styleOptions: { headerRow: true, bandedRows: true },
});
directTable.cells.set(1, 1, "$12.4M");
const directChart = dataSlide.charts.add("bar", {
  name: "direct-arr-chart",
  title: "ARR trend",
  position: { left: 430, top: 40, width: 320, height: 200 },
  categories: ["Q1", "Q2", "Q3"],
  series: [{ name: "ARR", values: [12, 18, 24] }],
});
const directImage = dataSlide.images.add({
  name: "direct-product-image",
  alt: "Product screenshot placeholder",
  dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+/p9sAAAAASUVORK5CYII=",
  prompt: "dashboard product screenshot",
  position: { left: 40, top: 180, width: 300, height: 140 },
  fit: "cover",
  borderRadius: "rounded-xl",
});
assert.equal(presentation.resolve(directTable.id).getCell(1, 1).value, "$12.4M");
assert.equal(presentation.resolve(directChart.id).title, "ARR trend");
assert.equal(presentation.resolve(directImage.id).alt, "Product screenshot placeholder");

dataSlide.compose(
  gridNode({
    columns: [{ mode: "fr", value: 1 }, { mode: "fr", value: 1 }],
    rows: [{ mode: "fixed", value: 130 }, { mode: "fixed", value: 170 }],
    columnGap: 18,
    rowGap: 18,
  }, [
    tableNode({ name: "jsx-table", rows: 2, columns: 2, values: [["A", "B"], [1, 2]], styleOptions: { headerRow: true } }),
    chartNode({ name: "jsx-chart", chartType: "bar", title: "Pipeline", categories: ["New", "Won"], series: [{ name: "Deals", values: [7, 4] }] }),
    imageNode({ name: "jsx-image", row: 1, columnSpan: 2, alt: "Generated hero", prompt: "abstract hero image", fit: "cover" }),
  ]),
  { frame: { left: 40, top: 340, width: 720, height: 320 } },
);
const objectInspect = presentation.inspect({ kind: "table,chart,image", maxChars: 12000 }).ndjson;
assert.match(objectInspect, /direct-kpi-table/);
assert.match(objectInspect, /jsx-table/);
assert.match(objectInspect, /direct-arr-chart/);
assert.match(objectInspect, /jsx-chart/);
assert.match(objectInspect, /direct-product-image/);
assert.match(objectInspect, /jsx-image/);
const chartSearch = presentation.inspect({ kind: "table,chart,image", search: "ARR", maxChars: 4000 }).ndjson;
assert.match(chartSearch, /direct-arr-chart/);
assert.doesNotMatch(chartSearch, /direct-kpi-table/);
const imageSearch = presentation.inspect({ kind: "image", search: "hero", maxChars: 4000 }).ndjson;
assert.match(imageSearch, /jsx-image/);
const objectLayout = JSON.parse(await (await dataSlide.export({ format: "layout" })).text());
assert.ok(objectLayout.elements.some((element) => element.kind === "table" && element.name === "jsx-table"));
assert.ok(objectLayout.elements.some((element) => element.kind === "chart" && element.name === "jsx-chart"));
assert.ok(objectLayout.elements.some((element) => element.kind === "image" && element.name === "jsx-image"));
assert.match(await (await dataSlide.export({ format: "svg" })).text(), /ARR trend/);

const out = path.join(os.tmpdir(), `open-office-artifact-jsx-${process.pid}.pptx`);
const pptx = await PresentationFile.exportPptx(presentation);
const zip = await JSZip.loadAsync(new Uint8Array(await pptx.arrayBuffer()));
assert.ok(zip.file("ppt/media/image1.png"));
const slide3Xml = await zip.file("ppt/slides/slide3.xml").async("text");
assert.match(slide3Xml, /<p:pic>/);
assert.match(slide3Xml, /name="direct-product-image"/);
assert.match(slide3Xml, /descr="Product screenshot placeholder"/);
assert.match(slide3Xml, /r:embed="rId1"/);
const slide3RelsXml = await zip.file("ppt/slides/_rels/slide3.xml.rels").async("text");
assert.match(slide3RelsXml, /Target="\.\.\/media\/image1\.png"/);
const contentTypesXml = await zip.file("[Content_Types].xml").async("text");
assert.match(contentTypesXml, /Default Extension="png" ContentType="image\/png"/);
await pptx.save(out);
const loaded = await PresentationFile.importPptx(await FileBlob.load(out));
assert.match(loaded.inspect({ kind: "textbox", maxChars: 10000 }).ndjson, /JSX runtime/);
console.log("presentation-jsx smoke ok");
