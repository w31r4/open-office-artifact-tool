# Layouts

`open-office-artifact-tool` can author a deliberately small native PresentationML
layout profile. It is useful for repeated agent-generated title/body slides; it
is not a generic PowerPoint template editor.

The source-free profile has one canonical master, layouts with type `blank`,
`title`, `titleOnly`, or `obj` (aliases `object`, `content`, and
`titleAndContent` normalize to `obj`), and direct-frame text placeholders of
type `title`, `body`, `ctrTitle`, or `subTitle`.

## Create and apply

```js
import { Presentation, PresentationFile } from "open-office-artifact-tool";

const presentation = Presentation.create({
  master: {
    name: "Brand master",
    placeholders: [{
      type: "title",
      index: 0,
      name: "Title",
      position: { left: 72, top: 56, width: 1136, height: 80 },
      style: { fontSize: 30, bold: true, color: "#0F172A" },
    }],
  },
});

const layout = presentation.layouts.add({
  name: "Title and body",
  type: "titleAndContent",
});

layout.placeholders.add({
  type: "body",
  index: 1,
  name: "Body",
  position: { left: 72, top: 168, width: 1136, height: 460 },
  style: { fontSize: 18, color: "#334155" },
});

const slide = presentation.slides.add({ name: "Overview", layout });
slide.placeholders.getItem("title").text.set("Q3 operating plan");
slide.placeholders.getItem("body").text.set("Three concise, auditable points.");

const pptx = await PresentationFile.exportPptx(presentation);
```

`slides.add({ layout })` resolves the layout transactionally and materializes
ordinary editable slide shapes with native `p:ph` identity and a direct
`a:xfrm`. `slide.setLayout(layout)` / `slide.applyLayout(layout)` provide the
same binding after creation. Reapplying the same layout is idempotent; changing
an already materialized layout fails closed rather than leaving an ambiguous
placeholder topology.

## Placeholder config

```ts
type SourceFreeTextPlaceholder = {
  id?: string;
  name?: string;
  type: "title" | "body" | "ctrTitle" | "subTitle";
  idx?: number; // `index` is an alias
  index?: number;
  position: { left: number; top: number; width: number; height: number };
  text?: string | Paragraph[];
  style?: TextStyle;
  paragraphStyles?: Record<number, ParagraphStyle>;
  textBodyProperties?: TextBodyProperties;
};
```

All source-free placeholders need a complete direct position. There is no
source-free inherited-geometry mode: a missing frame fails before writing a
PPTX. The helper accepts common input aliases such as `subtitle` and
`centeredTitle`, but inspect/import expose the native tokens `subTitle` and
`ctrTitle`.

Use `presentation.layouts.getById(layout.id)` when an exact stable lookup is
needed. `layouts.getItem(...)` also accepts a name or type and is convenient
only when the result cannot be ambiguous.

Use a defensive snapshot before an agent decides which text slots to fill:

```js
const summary = layout.placeholders.summary();
// { ownerId, count, requiredCount, types, items }
```

Each `items` entry has copied `id`, `name`, native `type`, `idx`/`index`,
`required`, `hasDirectPosition`, and (when direct) `position` fields. Mutating
the returned object never mutates the layout. Imported inherited placeholders
are visible with `hasDirectPosition: false`; that is discovery evidence, not an
invitation to reconstruct their geometry.

## Imported decks and boundaries

Imported Master/Layout graphs are source-bound. They can be inspected, guide
metadata remains visible, and unchanged export preserves their package graph;
their semantic properties, placeholder topology, and layout binding cannot be
rewritten through this API.

Imported layouts also expose read-only `slideGuides` definitions. Use
`presentation.view` only for local guide/grid visibility; it does not mutate
the source-bound `viewProps.xml` guide graph.

Multiple masters, master-specific themes, chart/table/media/object
placeholders, arbitrary inherited placeholder geometry, custom-template graph
editing, and layout/slide placeholder topology changes remain explicit
advanced-package or native-host work. Do not emulate them by flattening an
imported template into ordinary slide shapes.
