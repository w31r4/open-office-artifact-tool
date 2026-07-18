# Slide Facade

## Content Collections

```ts
const shape = slide.shapes.add(shapeConfig);
const image = slide.images.add(imageConfig);
const table = slide.tables.add(tableConfig);
const chart = slide.charts.add(chartType, chartConfig);
```

## Content Config Pointers

```ts
type ShapeConfig = PresetShapeConfig | CustomShapeConfig | ConnectorConfig; // see shapes.spec.md
type ImageConfig = ImageAddOptions; // see images.spec.md
type TableConfig = TableAddOptions | TableCellValue[][]; // see tables.spec.md
type ChartConfig = {
  position?: { left?: number; top?: number; width?: number; height?: number };
  title?: string;
  categories?: string[];
  series?: Array<{ name: string; values?: number[] }>;
}; // see charts.spec.md for axes, legend, labels, and series styling
```

## Background

```ts
slide.setBackground({ fill: "#f8fafc", mode: "solid" });
slide.setBackground({ fill: "accent2", mode: "reference", index: 1002 });
slide.clearBackground();
```

The canonical OpenChestnut PPTX boundary owns only a direct `p:cSld/p:bg` with
one six-digit RGB/theme color. `mode: "solid"` authors a direct solid fill;
`mode: "reference"` authors a native background-style reference with an
unsigned 32-bit `index`. Passing `background` to `presentation.slides.add(...)`
uses the same shape.

`slide.background` is the direct slide override. `slide.effectiveBackground()`
resolves that override or the preserved Layout/Master inheritance chain.
`clearBackground()` removes only the direct override and never flattens the
inherited color into the slide part.

Recognized imported direct backgrounds are source/hash-bound and may be
changed or removed. A slide with no direct background may gain one. Gradient,
pattern, image, transform-bearing, effect-bearing, or otherwise irregular
imported background graphs are opaque-preserved; leave them unchanged or
OpenChestnut fails closed. Do not replace an advanced imported background with
a simple solid and describe that as a faithful edit.

## Slide Order, Constrained Deletion, And Bounded Clone

```ts
slide.moveTo(0); // existing 0-based destination index
const clone = slide.duplicate(); // narrow imported shape/inline-table/image/closed-notes profile only
slide.delete(); // removes a non-final source-free slide, or a qualified imported source slide
```

`moveTo` returns the same slide facade. For an imported PPTX, it is a narrow
source-preserving reorder: export changes only the presentation `p:sldIdLst`
order of the retained source parts. The slide parts, their relationship graphs,
and opaque package content are not copied or reconstructed.

`delete()` returns `undefined` and refuses to remove the final slide. For a
source-free deck, it removes the selected slide normally. For an imported PPTX,
it is a real OPC transaction only when the source `SlidePart` is an isolated
layout-only leaf: it has no media, notes, comments, charts, OLE, hyperlinks,
data parts, or other child relationship; no inbound relationship; and no
custom-show, section, extension, or presentation-level identity reference. The
codec removes the slide part and its relationship part, updates
`ppt/presentation.xml` plus its relationships/content types, and keeps the
surviving source slides byte-identical. Any other imported delete fails closed.

`duplicate()` is not a visual-only copy and never creates a second
`p:sldId` reference to one source part. It is available only on an **original
imported** source slide with an unchanged body of canonical simple shapes,
canonical inline fixed-grid tables, plus canonical embedded rectangular images, exactly one internal
`SlidePart -> SlideLayoutPart` relationship, and image relationships bound only
by those pictures. It may also own one closed
`NotesSlide -> NotesMaster` / back-to-source-slide leaf and one canonical
legacy `SlideCommentsPart` leaf. The notes part must have exactly those two
relationships; the comments part must have no child, external,
hyperlink, or data relationship; and every legacy `p:cm/@authorId` must resolve
through one immutable presentation-wide `CommentAuthorsPart`. Export allocates
a distinct new `SlidePart` and a new presentation relationship, intentionally
reuses the verified layout and immutable ImageParts through fresh clone-local
image relationships, shares the verified NotesMaster and CommentAuthorsPart,
copies accepted NotesSlide and SlideComments XML byte-for-byte, and points only
the copied notes part back to the new slide. It preserves the required original
source part. Accepted tables are inline-only and cannot introduce a fill, link,
or any other package relationship. The new clone, its notes, and its legacy comments must remain
semantically unchanged and cannot replace or delete its origin until that
export has completed and the resulting PPTX has been imported again; afterward
the slide may use the supported fixed-topology edit path, while imported legacy
comments remain source-bound read-only.

Source-free slides, already-cloned slides, rich or connected comments, charts,
OLE, hyperlinks, external/data relationships, external/unbound/irregular
images, non-shape/table/image elements, renames/layout/background changes, note
or comment mutation before the boundary, immediate clone edits, and all broader
relationship graphs fail closed. Adding imported slides still fails closed. A
template-derived deck still needs a broader explicit OPC graph-clone transaction
rather than this bounded leaf operation.

## Layouts And Placeholders

```ts
// Source-free authoring only: either supply the layout on creation or bind it
// explicitly before filling direct-frame p:ph shapes.
const slide = presentation.slides.add({ layout });

const placeholder = slide.placeholders.getItem("title");
placeholder.text.set("Updated section title");
```

`slide.setLayout(layout)` / `slide.applyLayout(layout)` provide the same
materialization after creation and are idempotent for the same layout. This is
a bounded authoring path, not a generic template editor. The layout
must be one of the source-free canonical layouts described in
[`layout.spec.md`](./layout.spec.md), and source-free `title`, `body`,
`ctrTitle`, and `subTitle` placeholders need a direct frame. Existing
placeholders imported from a PPTX remain inspectable but are a source-bound,
read-only inherited projection: do not change their text, geometry, identity,
or slide/layout binding through this API.

## Frame And Export

```ts
const frame = slide.frame;
const preview = await slide.export({ format, scale });
```

## Compose Layout

```tsx
/** @jsxRuntime automatic */

slide.compose(
  <row width="fill" height="fill" gap={20}>
    <column width={340} height="fill" gap={12}>
      <paragraph name="summary" className="text-slate-950 text-3xl font-bold">
        Summary
      </paragraph>
      <paragraph className="text-slate-600 text-sm">
        Compose-first layouts stay readable and export cleanly.
      </paragraph>
    </column>
    <box
      name="surface"
      width="fill"
      height="fill"
      padding={16}
      fill="#ffffff"
    >
      <column height="fill" gap={10}>
        <rule stroke="#0f172a" weight={2} />
      </column>
    </box>
  </row>,
  {
    frame: { left: 48, top: 40, width: 864, height: 460 },
    baseUnit: 8,
  },
);
```

`slide.compose(...)` accepts either compose helper nodes or a JSX-authored tree.
It still returns the materialized elements. Use `name` on materializing nodes
when the slide needs stable inspect or layout-export anchors.

## Export Inline Type

```ts
type SlideExportOptions = {
  format?: "png" | "jpeg" | "webp" | "layout";
  width?: number;
  height?: number;
  scale?: number;
  quality?: number;
};
```

`format: "layout"` exports the composed layout tree alongside the materialized slide elements. That is the most useful export when you need stable names, hierarchy paths, and resolved text for follow-up targeting.

## Auto Layout

```ts
slide.autoLayout(shapes, {
  direction,
  frame,
  align,
  horizontalGap,
  verticalGap,
  horizontalPadding,
  verticalPadding,
});
```

## Auto Layout Inline Type

```ts
type AutoLayoutOptions = {
  direction?: "horizontal" | "vertical";
  frame?:
    | "slide"
    | Shape
    | { left: number; top: number; width: number; height: number };
  align?:
    | "center"
    | "topLeft"
    | "topCenter"
    | "topRight"
    | "left"
    | "right"
    | "bottomLeft"
    | "bottomCenter"
    | "bottomRight";
  horizontalGap?: number | "auto";
  verticalGap?: number | "auto";
  horizontalPadding?: number;
  verticalPadding?: number;
};
```

## Speaker Notes

```ts
slide.speakerNotes.textFrame.setText(notesText);
slide.speakerNotes.setVisible(visible);
```

## Visual QA Loop

```ts
const preview = await slide.export({ format: "png", scale: 2 });
const layout = await slide.export({ format: "layout" });
const after = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes",
  target: { id: slideAnchorId, beforeLines: 2, afterLines: 8 },
  maxChars: 4000,
});
```

Check the PNG for clipping, overlap, weak contrast, illegible text, bad image
crops, accidental rasterized labels, and off-canvas elements. Check the layout
export or inspect output for stable `name` values on important nodes.

## Cookbook

```ts
// Canonical direct background plus editable content.
slide.setBackground({ fill: "#0f172a", mode: "solid" });
const title = slide.shapes.add({
  geometry: "textbox",
  name: "title",
  position: { left: 72, top: 72, width: 760, height: 120 },
  fill: "none",
  line: { style: "solid", fill: "none", width: 0 },
});
title.text = "Editable headline";
title.text.style = { fontSize: 44, bold: true, color: "white" };
```

```ts
// Imported deck: inspect direct versus inherited state before mutation.
const record = JSON.parse((await deck.inspect({ kind: "slide", target: slide.id })).ndjson);
if (record.background) slide.setBackground({ fill: "accent1", mode: "reference", index: 1001 });
else slide.setBackground("#f8fafc");

// Restore inheritance when the direct override is no longer wanted.
slide.clearBackground();
```

```ts
// Reorder an existing source-bound slide without changing the source-part set.
const existing = deck.slides.getItem(2);
existing.moveTo(0);
```
