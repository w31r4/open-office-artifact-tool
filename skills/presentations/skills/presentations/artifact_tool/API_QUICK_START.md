# Presentation Quick Start (TypeScript)

## API Docs

API docs: `api/API_DOCS.md`

`Presentation` is the in-memory API object. A deck is the exported presentation
file. A slide is one page in the deck.

## Imports

```ts
import fs from "node:fs/promises";
import { FileBlob, Presentation, PresentationFile } from "open-office-artifact-tool";
```

## Create, Render, Export

This script creates an editable deck, renders each slide PNG, writes each slide
layout JSON, writes a deck montage with `montage: true`, and exports PPTX.

```ts
import fs from "node:fs/promises";
import { Presentation, PresentationFile } from "open-office-artifact-tool";

async function writeBlob(path: string, blob: Blob): Promise<void> {
  await fs.writeFile(path, new Uint8Array(await blob.arrayBuffer()));
}

async function main(): Promise<void> {
  await fs.mkdir("output", { recursive: true });

  const presentation = Presentation.create({
    slideSize: { width: 1280, height: 720 },
  });

  const slide = presentation.slides.add();
  slide.background.fill = "slate-50";

  const page = { left: 72, top: 64, width: 1136, height: 592 };
  const gutter = 28;
  const leftCol = 430;
  const rightCol = page.width - leftCol - gutter;

  const eyebrow = slide.shapes.add({
    geometry: "textbox",
    position: { left: page.left, top: page.top, width: 280, height: 28 },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  eyebrow.text = "EDITABLE PRESENTATION";
  eyebrow.text.style = { fontSize: 12, bold: true, color: "slate-500" };

  const title = slide.shapes.add({
    geometry: "textbox",
    position: {
      left: page.left,
      top: page.top + 92,
      width: leftCol,
      height: 184,
    },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  title.text = "Build decks with editable objects";
  title.text.style = { fontSize: 42, bold: true, color: "slate-950" };

  const subtitle = slide.shapes.add({
    geometry: "textbox",
    position: {
      left: page.left,
      top: page.top + 292,
      width: leftCol,
      height: 96,
    },
    fill: "none",
    line: { style: "solid", fill: "none", width: 0 },
  });
  subtitle.text =
    "Rows, columns, grids, charts, tables, images, and text remain editable.";
  subtitle.text.style = { fontSize: 20, color: "slate-600" };

  const chartFrame = slide.shapes.add({
    geometry: "roundRect",
    name: "chart-frame",
    position: {
      left: page.left + leftCol + gutter,
      top: page.top + 92,
      width: rightCol,
      height: 388,
    },
    fill: "white",
    line: { style: "solid", fill: "slate-200", width: 1 },
    borderRadius: "rounded-2xl",
    shadow: "shadow-sm",
  });

  slide.charts.add("bar", {
    position: {
      left: chartFrame.position.left + 36,
      top: chartFrame.position.top + 52,
      width: chartFrame.position.width - 72,
      height: 280,
    },
    categories: ["Rows", "Grid", "Tokens"],
    series: [{ name: "Coverage", values: [3, 4, 5], fill: "accent1" }],
    hasLegend: false,
    dataLabels: { showValue: true, position: "outEnd" },
    yAxis: {
      majorGridlines: { style: "solid", fill: "slate-200", width: 1 },
    },
  });

  for (const [index, slide] of presentation.slides.items.entries()) {
    const stem = `slide-${String(index + 1).padStart(2, "0")}`;
    const png = await presentation.export({ slide, format: "png", scale: 1 });
    await writeBlob(`output/${stem}.png`, png);

    const layout = await slide.export({ format: "layout" });
    await fs.writeFile(`output/${stem}.layout.json`, await layout.text());
  }

  const montage = await presentation.export({
    format: "webp",
    montage: true,
    scale: 1,
  });
  await writeBlob("output/deck-montage.webp", montage);

  const pptx = await PresentationFile.exportPptx(presentation);
  await pptx.save("output/deck.pptx");
}

main().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
```

## Bounded Native Combo Chart

`combo` is native PPTX output only for one intentionally narrow profile: literal
clustered columns plus literal lines, at least one primary bar and one line, and
no topology changes after import. There are two canonical axis variants:

- Leave every line on the default primary axis group to use one shared
  category/value pair.
- Put **every** line at `axisGroup: "secondary"`, keep every bar primary, and
  provide `axes.secondary.category` plus `axes.secondary.value`. OpenChestnut
  writes that second pair at the top and right of the chart.

```ts
slide.charts.add("combo", {
  name: "revenue-margin",
  title: "Revenue and margin",
  position: { left: 90, top: 120, width: 1080, height: 420 },
  categories: ["Q1", "Q2", "Q3"],
  series: [
    { name: "Revenue", chartType: "bar", values: [42, 48, 57], color: "#2563EB" },
    {
      name: "Margin",
      chartType: "line",
      axisGroup: "secondary",
      values: [12, 15, 18],
      line: { fill: "#16A34A", width: 2 },
      marker: { symbol: "circle", size: 7 },
    },
  ],
  axes: {
    category: { title: "Quarter" },
    value: { title: "Revenue ($M)" },
    secondary: {
      category: { title: "Quarter" },
      value: { title: "Margin (%)", min: 0, max: 25 },
    },
  },
  dataLabels: { showValue: true, position: "top" },
});
```

Do not mix primary and secondary line series, put a bar on `axisGroup:
"secondary"`, omit either secondary axis when using secondary lines, or use
external/embedded workbook data, smooth lines, point overrides, per-series
labels, trendlines, or error bars. Those combinations fail closed rather than
being flattened or silently rebuilt.

## JSX Compose Equivalent

Use JSX when the slide is naturally rows, columns, grids, or overlays.

```tsx
/** @jsxRuntime automatic */
/** @jsxImportSource open-office-artifact-tool/presentation-jsx */

const frame = { left: 72, top: 64, width: 1136, height: 592 };

slide.compose(
  <column name="content-frame" width="fill" height="fill" gap={28}>
    <row width="fill" height={72} align="center" justify="between">
      <paragraph name="eyebrow" className="text-slate-500 text-sm font-bold">
        EDITABLE PRESENTATION
      </paragraph>
      <paragraph className="text-slate-400 text-sm">Q2 planning</paragraph>
    </row>
    <row width="fill" height="fill" gap={28} align="stretch">
      <column width={430} height="fill" gap={18}>
        <paragraph
          name="primary-heading"
          className="text-slate-950 text-5xl font-bold leading-tight"
        >
          Build decks with editable objects
        </paragraph>
        <paragraph className="text-slate-600 text-xl leading-relaxed">
          Rows, columns, grids, charts, tables, images, and text remain
          editable.
        </paragraph>
      </column>
      <box
        name="chart-frame"
        width="fill"
        height="fill"
        className="bg-white rounded-2xl shadow-sm"
        line={{ style: "solid", fill: "slate-200", width: 1 }}
      >
        <chart
          name="coverage-chart"
          chartType="bar"
          categories={["Rows", "Grid", "Tokens"]}
          series={[{ name: "Coverage", values: [3, 4, 5], fill: "accent1" }]}
          hasLegend={false}
          width="fill"
          height="fill"
        />
      </box>
    </row>
  </column>,
  { frame, baseUnit: 8 },
);
```

## Import And Edit PPTX

Load a PPTX, inspect for stable ids, render before/after evidence, make a
focused edit, re-inspect, and export the edited PPTX.

```ts
const presentation = await PresentationFile.importPptx(
  await FileBlob.load("input.pptx"),
);

const before = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,thread,layout",
  search: "Revenue",
  maxChars: 8000,
});
console.log(before.ndjson);

const slide = presentation.resolve(slideIdFromInspect);
await writeBlob(
  "output/before-slide.png",
  await presentation.export({ slide, format: "png", scale: 1 }),
);
await fs.writeFile(
  "output/before-slide.layout.json",
  await (await slide.export({ format: "layout" })).text(),
);
await writeBlob(
  "output/before-montage.webp",
  await presentation.export({ format: "webp", montage: true, scale: 1 }),
);

const target = presentation.resolve(anchorIdFromInspect);
target.text.replace("Revenue", "Updated revenue outlook");

await writeBlob(
  "output/after-slide.png",
  await presentation.export({ slide, format: "png", scale: 1 }),
);
await fs.writeFile(
  "output/after-slide.layout.json",
  await (await slide.export({ format: "layout" })).text(),
);
await writeBlob(
  "output/after-montage.webp",
  await presentation.export({ format: "webp", montage: true, scale: 1 }),
);

const after = await presentation.inspect({
  target: { id: anchorIdFromInspect, beforeLines: 2, afterLines: 2 },
  kind: "textbox,shape,image,table,chart",
  maxChars: 3000,
});
console.log(after.ndjson);

const pptx = await PresentationFile.exportPptx(presentation);
await pptx.save("output/edited-deck.pptx");
```

## Bounded Imported Title And Speaker-Notes Edit

When the source deck has one uniquely named slide, one uniquely named title
shape, and one canonical plain-text Notes part, use the shipped workflow rather
than reconstructing the slide or editing OOXML directly:

```ts
import { editPptxTitleAndNotes } from "../examples/openchestnut-title-notes-edit-workflow.mjs";

await editPptxTitleAndNotes({
  inputPath: "input.pptx",
  outputPath: "output/edited.pptx",
  auditPath: "output/audit.json",
  slideName: "Go-no-go decision",
  titleShapeName: "approval-title",
  expectedTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  expectedNotes: "Lead with the pending legal condition.\nClose with the accountable owner.",
  replacementNotes: "Lead with the approved controls.\nClose with the accountable rollout owner.",
});
```

It requires a distinct output path and fails closed if the target is ambiguous,
the source text is not exact, the notes part is missing/rich, or reimport shows
an identity, geometry, background, slide-order, or slide-name change. It emits
only model-render evidence; use the native render route for final visual QA.

## Bounded Legacy Slide Comments

For a simple slide-level review annotation, use the standard legacy PPTX
profile directly. It has one author, one text item, and one explicit slide
coordinate; it is not a substitute for modern threaded comments.

```ts
const review = slide.comments.addThread(
  undefined,
  "Confirm the source before delivery.",
  {
    author: "Presentation Reviewer",
    created: "2026-07-18T09:30:00.000Z",
    position: { x: 1040, y: 84, unit: "px" },
  },
);

const evidence = presentation.inspect({
  kind: "comment",
  target: review.id,
  maxChars: 2000,
});
console.log(evidence.ndjson);
```

Use `undefined` for the target. Replies, resolution state, reactions, and
element/text-range anchors cannot be represented by legacy PresentationML and
therefore fail closed on canonical export. Recognized imported legacy comments
are visible for inspection but must remain unchanged; modern comment graphs
remain opaque and source-bound. An unchanged canonical legacy comments leaf may
travel with `slide.duplicate()` through one export/reimport boundary: its
clone-local `SlideCommentsPart` is byte-copied while the verified immutable
presentation-wide author catalog is shared. This is not an in-place comment
edit; both pending clone comments and reimported legacy comments remain
source-bound read-only. See `api/references/comments.md` for the complete
boundary.

## Local Image Bytes

Use byte-backed images for embedded PPTX assets.

```ts
async function readImageBlob(imagePath: string): Promise<ArrayBuffer> {
  const bytes = await fs.readFile(imagePath);
  return bytes.buffer.slice(
    bytes.byteOffset,
    bytes.byteOffset + bytes.byteLength,
  );
}

const imageBytes = await readImageBlob("assets/product.png");
slide.images.add({
  blob: imageBytes,
  contentType: "image/png",
  alt: "Product screenshot",
  fit: "cover",
  position: { left: 720, top: 96, width: 420, height: 280 },
  geometry: "roundRect",
  borderRadius: "rounded-xl",
});
```
