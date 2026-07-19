# Imported Deck Cookbook

Use these recipes when editing an existing presentation. The goal is targeted,
type-aware edits with a preview and a small verification snapshot.

For imported decks, render the affected slide before and after focused edits,
export layout JSON before and after when object placement matters, and use a
deck montage before/after when a change may affect multiple slides.

## Edit Loop

```ts
const before = await presentation.inspect({
  kind: "slide,textbox,shape,image,table,chart,notes,thread,layout",
  search: "Revenue",
  maxChars: 8000,
});

// Pick an exact anchor id from `before`; use inspect ids instead of slide indexes.
const target = presentation.resolve(anchorId);

// slideAnchorId is the `sl/...` id from inspect for the affected slide.
const slide = presentation.resolve(slideAnchorId);
const previewBefore = await slide.export({ format: "png", scale: 2 });
const layoutBefore = await slide.export({ format: "layout" });
const montageBefore = await presentation.export({
  format: "webp",
  montage: true,
  scale: 1,
});

target.text.replace("Revenue", "Revenue outlook");

const previewAfter = await slide.export({ format: "png", scale: 2 });
const layoutAfter = await slide.export({ format: "layout" });
const montageAfter = await presentation.export({
  format: "webp",
  montage: true,
  scale: 1,
});

const after = await presentation.inspect({
  target: { id: anchorId, beforeLines: 2, afterLines: 2 },
  kind: "textbox,shape",
  maxChars: 2000,
});
```

## Placeholder Inspection And Source-Free Authoring

```ts
const layoutSnapshot = await presentation.inspect({
  kind: "slide,layout,textbox,shape",
  search: "Click to add title",
  maxChars: 6000,
});

// slideAnchorId is the `sl/...` id from inspect for this slide.
const slide = presentation.resolve(slideAnchorId);
const title = slide.placeholders.getItem("title");
const titleEvidence = title && {
  id: title.id,
  frame: title.position,
  text: title.text.value,
  textEditable: title.placeholder?.textEditable === true,
};
```

For an imported deck, distinguish a concrete placeholder shape owned by the
SlidePart from an inherited Master/Layout projection. When the concrete
`p:sp/p:ph` owns a fully recognized local `p:txBody`, its existing text content
is a bounded editable component even though the surrounding shape remains
source-bound:

```ts
if (title.placeholder?.textEditable !== true) {
  throw new Error("The imported placeholder has no verified local text capability.");
}

// Fixed-topology, in-run replacement keeps the native paragraph/run style.
title.text.replace("Quarterly Review", "Quarterly Outlook");

// Whole-text replacement is also supported when the source newline topology
// is retained and each newline-delimited span maps to one native text run.
title.text.set(title.text.value.replace("FY25", "FY26"));
```

Both routes preserve placeholder type/index, name, geometry, formatting, layout
binding, and unmodeled XML. A changed newline count, ambiguous multi-run span,
field, unsupported local text graph, or any non-text mutation fails closed.
Master/Layout placeholder collections and placeholder shapes that exist only as
inherited projections remain inspection evidence, not a template-rewrite API.
Use an ordinary overlay, a native host, or an explicit source-free rebuild when
the requested change exceeds this local text boundary.

For a new source-free deck, create the canonical master/layout and use
`slides.add({ layout })` (or call `slide.setLayout(layout)` after creation)
before filling the direct-frame `title`, `body`, `ctrTitle`, or `subTitle`
placeholders; see
[`layout.spec.md`](../layout.spec.md). This broader source-free materialization
path is deliberately separate from the bounded imported local-text operation.

Inspect records use 1-based `slide` numbers for display; `slides.getItem(index)`
is 0-based. Prefer resolving the `sl/...` anchor from inspect.

## Source-Bound Duplicate

Use `slide.duplicate()` only for the explicit closed clone profile; it is not a
generic “copy this slide” shortcut. A present straight/elbow connector endpoint
must remain inside the copied source SlidePart tree. Pending clone IDs are new
model IDs, and the clone must export and reimport before any edit.
Recognized literal-data charts may travel only when every frame owns one unique
internal relationship to a closed numbered ChartPart. Export creates a
distinct byte-identical ChartPart for the clone. After reimport, a chart that
advertises the ordinary fixed-topology edit capability can use it without
altering the origin. Formula/external-data/embedded-workbook charts,
connected ChartParts, and orphan or duplicate chart relationships fail closed.
Canonical run-level click links may target one external absolute URI, one
retained internal SlidePart, or a supported relative slide action. Their exact
relationship IDs and targets are copied onto the new SlidePart; shape-level or
hover clicks, unknown actions, orphan relationships, and jumps to a removed
slide remain fail-closed.

For an Agent-facing transaction over the bare profile (no notes/comments leaf),
run the shipped workflow with one unique explicit source name:

```sh
node "$SKILL_DIR/examples/openchestnut-slide-duplicate-workflow.mjs" \
  input/source.pptx output/source-with-copy.pptx output/clone-audit.json \
  "Unique source slide name"
```

The audit binds input/output hashes, source and clone part paths, adjacent
insertion, allowed new package parts, retained-source byte preservation,
exact source/clone ChartPart bytes and relationship IDs, exact run-link
relationship IDs and targets, reimported structural
equivalence, and model-SVG visual equivalence. The SVG
check removes only fresh `data-*-id` locator attributes; a new SlidePart XML
may be canonically serialized and is not promised to be lexically identical.
Missing/duplicate names, notes/comments, unsupported link markup, nonliteral
or connected charts, other unsupported leaves, or an unexpected package delta
fail closed without output promotion.

## Master/Layout Blast Radius

```ts
// layoutId comes from inspect/layout export metadata; use it for search and comparison.
const affected = await presentation.inspect({
  kind: "slide,layout",
  search: layoutId,
  maxChars: 6000,
});
```

Use layout ids for comparison/search and resolve affected slides through their
`sl/...` ids before deciding on an operation. Imported Master/Layout/theme and
placeholder graphs are source-bound and cannot be edited through this API.
For the small source-free authoring profile, configure the one canonical
master/layout before materializing slide placeholders; it is not a way to
retroactively rewrite an arbitrary imported template.

## Preserve Imported Image Placement

```ts
const image = presentation.resolve(imageAnchorId);
const oldFrame = image.frame;
const oldCrop = image.crop;
const oldFit = image.fit;
const oldAlt = image.alt;
const oldPrompt = image.prompt;
const oldGeometry = image.geometry;
const oldBorderRadius = image.borderRadius;
const oldRotation = image.rotation;
const oldFlipHorizontal = image.flipHorizontal;
const oldFlipVertical = image.flipVertical;
const oldLockAspectRatio = image.lockAspectRatio;

image.replace({
  blob: replacementBytes,
  contentType: "image/png",
  alt: oldAlt ?? "Updated product screenshot",
  ...(oldFit ? { fit: oldFit } : {}),
  ...(oldPrompt ? { prompt: oldPrompt } : {}),
});
image.frame = oldFrame;
image.crop = oldCrop;
image.geometry = oldGeometry;
image.borderRadius = oldBorderRadius;
image.rotation = oldRotation;
image.flipHorizontal = oldFlipHorizontal;
image.flipVertical = oldFlipVertical;
image.lockAspectRatio = oldLockAspectRatio;
```

Render the slide after replacement and verify subject crop, aspect ratio,
rounded mask, and legibility. Preserve placement-affecting properties unless the
edit explicitly changes them. Concrete source replacements produce concrete
images; pass `prompt` when it should remain available as regeneration metadata.

## Existing Table And Chart Edits

```ts
const table = presentation.resolve(tableAnchorId);
table.cells.set(1, 2, "$4.2M");
table.getCell(1, 2).text.style = "Body Small";

const chart = presentation.resolve(chartAnchorId);
chart.xAxis = { title: "Quarter" };
chart.yAxis = { numberFormatCode: "$#,##0M" };
chart.series.getItemAt(0).values = [3.1, 3.7, 4.2, 4.8];

await presentation.inspect({
  kind: "table,chart",
  target: { id: chartAnchorId, beforeLines: 3, afterLines: 3 },
  maxChars: 3000,
});
```

If an imported chart or table resolves as an image, preserve it as an image or
rebuild it as a native chart/table intentionally.

## Comments And Speaker Notes

```ts
const review = await presentation.inspect({
  kind: "thread,notes,slide,textbox",
  search: "TODO",
  maxChars: 8000,
});

const self = presentation.comments.setSelf({
  displayName: "Presentation Editor",
  initials: "PE",
  email: "presentation@example.com",
});

const thread = presentation.resolve(commentThreadAnchorId);
thread.addReply("Addressed in this revision.", { author: self });
thread.resolve(self);

const notes = presentation.resolve(notesAnchorId);
const existing = notes.text.trim();
notes.setText([
  existing,
  "Opening: summarize the revised revenue outlook.",
  "Call out risk on enterprise timing.",
].filter(Boolean).join("\n"));

const verified = await presentation.inspect({
  kind: "thread,notes",
  target: { id: slideAnchorId, beforeLines: 1, afterLines: 4 },
  maxChars: 3000,
});
```

Keep unrelated threads and speaker notes intact unless the task explicitly asks
to clear them. Speaker-note edits use the bounded plain-text OpenChestnut
contract; rich or irregular imported notes remain preservation-only and fail
closed on mutation.
