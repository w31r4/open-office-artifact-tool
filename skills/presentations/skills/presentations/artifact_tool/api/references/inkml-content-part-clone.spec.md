# Source-bound InkML content-part clone

OpenChestnut imports a PresentationML `p:contentPart` as an opaque
`nativeObject` with `nativeKind === "contentPart"`. The bounded profile below
preserves one standard InkML payload while cloning its slide. It does not expose
ink authoring, stroke editing, or arbitrary Custom XML mutation.

An unchanged object may travel through `slide.duplicate()` only when all of
these preconditions hold:

- the object is a direct, top-level child of the source slide's `p:spTree`,
  never nested in `p:grpSp`;
- it has one bounded non-visual property record, one transform with positive
  extents, no extension list, and exactly one relationship attribute;
- that attribute uniquely consumes one internal standard or strict OOXML
  `customXml` relationship owned by the SlidePart;
- the target is a non-empty, fully well-formed `CustomXmlPart` with content type
  `application/inkml+xml` and an `ink` document element in the
  `http://www.w3.org/2003/InkML` namespace;
- the InkML part has no child, external, hyperlink, or data relationship; and
- the imported native object and independent OPC inspection agree on the
  relationship ID, target path, content type, and SHA-256 digest.

```ts
const source = presentation.slides.getItem(0);
const ink = source.nativeObjects.items.find(
  (object) => object.nativeKind === "contentPart",
);

if (!ink || ink.parts.length !== 1) {
  throw new Error("Source does not expose one closed InkML content part.");
}

const clone = source.duplicate();
const output = await PresentationFile.exportPptx(presentation);
const rebound = await PresentationFile.importPptx(output);
```

Export retains the source SlidePart and relationship part byte-for-byte. The
clone keeps the same slide-local relationship ID but receives a distinct Open
XML SDK `CustomXmlPart` under `ppt/customXml/itemN.xml`. Its InkML bytes and
content type exactly match the source. After second import, the source and clone
objects have disjoint part paths and equal payload hashes.

For an Agent-facing transaction, use
`examples/openchestnut-slide-duplicate-workflow.mjs`. Its audit records
`operation.inkContentParts`, the exact allowed package delta, each source/clone
part path and digest, `validation.package.inkContentParts`, second-import
binding independence, model-render equivalence, and native source/clone pixel
equality when LibreOffice and Poppler are available. A failed precondition
publishes neither PPTX nor audit.

The object remains opaque and read-only after reimport. `setName`,
`setPosition`, raw XML replacement, stroke editing, and arbitrary Custom XML
cloning remain unsupported. Self-closing, nested, extension-bearing, ambiguous,
shared-relationship, external, mistyped, malformed/multi-root, non-InkML-root, or connected content
parts fail closed. Preserve those objects unchanged or route a separately
reviewed explicit OPC operation; do not flatten them into ordinary shapes.
