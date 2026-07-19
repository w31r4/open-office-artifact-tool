# Source-bound SmartArt clone

OpenChestnut imports SmartArt as an opaque `nativeObject` with
`nativeKind === "diagram"`. It does not reconstruct SmartArt into ordinary
shapes and does not expose semantic diagram editing.

One unchanged SmartArt frame may travel through the bounded imported
`slide.duplicate()` transaction when all of these preconditions hold:

- the object is one top-level `p:graphicFrame`, never nested in `p:grpSp`;
- it has exactly one `dgm:relIds` root;
- `r:dm`, `r:lo`, `r:qs`, and `r:cs` are present once each and use four unique
  slide-local relationship IDs;
- those relationships are internal and have the standard diagram-data,
  layout, quick-style, and colors relationship types;
- each target has the matching standard OOXML content type, non-empty bytes,
  and no child, external, hyperlink, or data relationship;
- the native object and independent OPC inspection agree on all four IDs,
  target paths, content types, and SHA-256 digests.

```ts
const source = presentation.slides.getItem(0);
const diagram = source.nativeObjects.items.find(
  (object) => object.nativeKind === "diagram",
);

if (!diagram || diagram.parts.length !== 4) {
  throw new Error("Source does not expose one closed SmartArt graph.");
}

const clone = source.duplicate();
const output = await PresentationFile.exportPptx(presentation);
const rebound = await PresentationFile.importPptx(output);
```

The first export keeps the source SlidePart and its relationship part
byte-identical. It preserves the four slide-local relationship IDs but creates
four distinct typed diagram parts for the clone. Each new part is a byte copy
of its corresponding source part. After reimport, source and clone have
disjoint part paths and matching per-role hashes.

Use
`examples/openchestnut-slide-duplicate-workflow.mjs` for an Agent-facing
transaction. Its audit records the source and clone SlideParts, all four
relationship IDs, source and clone part paths, content types, hashes, exact
allowed package delta, second-import evidence, and model-render equivalence.
It writes neither output nor audit when any precondition fails.

This boundary prevents source/clone package coupling; it is not SmartArt
authoring. `setName`, `setPosition`, raw XML mutation, diagram-data editing,
and arbitrary graph cloning remain unsupported. Incomplete, duplicated,
mistyped, external, nested, relationship-bearing, or otherwise noncanonical
SmartArt graphs fail closed. Preserve such objects unchanged or use a separate
explicit OPC operation whose scope is independently reviewed.
