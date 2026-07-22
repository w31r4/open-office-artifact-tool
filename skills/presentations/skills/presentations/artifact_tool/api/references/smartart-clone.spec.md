# Source-bound SmartArt clone and plain-node text update

OpenChestnut imports SmartArt as a native `nativeObject` with
`nativeKind === "diagram"`. It never reconstructs a diagram as ordinary
shapes. There are two deliberately separate, source-bound contracts:

1. an unchanged closed graph may travel through `slide.duplicate()`; and
2. one narrow DiagramDataPart profile may replace text in existing document
   nodes through `nativeObject.setDiagramNodeText()`.

Neither contract is SmartArt authoring, layout editing, graph editing, or raw
XML access.

## Closed graph cloning

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
disjoint part paths and per-role hashes match.

Use `examples/openchestnut-slide-duplicate-workflow.mjs` for an Agent-facing
transaction. Its audit records the source and clone SlideParts, all four
relationship IDs, source and clone part paths, content types, hashes, exact
allowed package delta, second-import evidence, and model-render equivalence.
It writes neither output nor audit when any precondition fails.

## Canonical plain-node text profile

An imported graph that passes the closed four-part contract gains a separate
text capability only when its DiagramDataPart additionally proves all of the
following:

- its root is `dgm:dataModel` and it has exactly one direct `dgm:ptLst`;
- every exposed `dgm:pt` has `type="doc"`, a unique non-empty `modelId`, and
  exactly one direct `dgm:t > a:p > a:r > a:t` text chain; optional
  `a:bodyPr`, `a:lstStyle`, `a:pPr`, `a:rPr`, and `a:endParaRPr` may remain;
- text is XML-safe and at most 32,767 characters. Rich/multiple runs, fields,
  breaks, unknown child markup, disconnected parts, or any topology ambiguity
  withhold the capability rather than being simplified.

`nativeObject.editable` remains `false`: this is a typed exception, not general
write authority. `nativeObject.diagramText` is a defensive snapshot containing
the source data part and the eligible node IDs. The only mutation is one
existing node's text:

```ts
const diagram = presentation.slides.getItem(0).nativeObjects.items.find(
  (object) => object.nativeKind === "diagram" && object.diagramText,
);
if (!diagram) throw new Error("No canonical plain-node SmartArt target.");

const node = diagram.diagramText.nodes.find((item) => item.text === "Before");
if (!node) throw new Error("Expected source text is not unique.");
diagram.setDiagramNodeText(node.id, "After");

const output = await PresentationFile.exportPptx(presentation);
```

Export re-proves the original graph, source digest, node IDs/order, and the
plain-node profile. It may rewrite only the bound DiagramDataPart; it preserves
the graphic frame, `dm/lo/qs/cs` relationship IDs, layout, quick-style,
colors, geometry, and every non-data package part. The output is reimported and
must expose the exact requested node list. Leading or trailing replacement
whitespace is serialized with `xml:space="preserve"`.

Use `examples/openchestnut-smartart-text-edit-workflow.mjs` for a no-overwrite
Agent transaction:

```sh
node "$SKILL_DIR/examples/openchestnut-smartart-text-edit-workflow.mjs" \
  input/source.pptx output/edited.pptx output/edited.audit.json \
  "Closed SmartArt" "{B31B1833-2B65-4D6B-B3D4-9B3988427B21}" "Before" "After"
```

It protects the input bytes, resolves exactly one object/node/expected text,
checks that only the DiagramDataPart changed, reimports the graph, and writes a
source/output-bound audit. Its model verification is structural evidence; run
the normal LibreOffice/Poppler render review when a native rendering result is
required.

Node creation/removal/reordering, `modelId` changes, presentation of arbitrary
diagram text, raw XML mutation, layout/style/color edits, geometry edits,
cross-diagram changes, clone-before-export after a pending text edit, and
arbitrary graph cloning remain unsupported. Incomplete, duplicated, mistyped,
external, nested, relationship-bearing, rich, or otherwise noncanonical
SmartArt graphs fail closed. Preserve such objects unchanged or use a separate
explicit OPC operation whose scope is independently reviewed.
