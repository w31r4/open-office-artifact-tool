# Custom Shows

Custom shows are ordered, named subsets of a deck. OpenChestnut writes them as
native `p:custShowLst` data in `ppt/presentation.xml`; they are not preview-only
metadata and do not create a second copy of any SlidePart.

## Source-Free Authoring

Create every slide first, then add the shows that reference those slide
facades. A slide may appear more than once when the requested playback route
requires it.

```ts
const overview = presentation.slides.add({ name: "Overview" });
const evidence = presentation.slides.add({ name: "Evidence" });
const appendix = presentation.slides.add({ name: "Appendix" });

presentation.customShows.add({
  name: "Board route",
  slides: [overview, appendix],
});

presentation.customShows.add("Review route", [evidence, overview]);
```

Names are trimmed, non-empty, at most 255 characters, and
case-insensitively unique. A show contains 1 through 16,384 slide references;
a deck contains at most 4,096 shows. Omit `nativeId` for new content and let
the model assign a unique unsigned 32-bit identity during verification/export.

Use `getItem(indexOrIdOrName)`, `inspect({ kind: "customShow" })`, or
`resolve(customShow.id)` to find a show. Use `show.setSlides(slides)` to replace
its ordered slide membership.

## Imported Canonical Shows

A canonical imported list exposes semantic shows with hash-bound source
evidence. OpenChestnut permits two edits in place:

- change an existing show's `name` without colliding with another name;
- call `setSlides(...)` with an ordered list of retained slides from the same
  imported presentation.

The show count/order, facade ID, and native `p:custShow/@id` remain fixed.
Adding, removing, reordering, or replacing show objects fails closed. Slide
deletion and source-bound slide cloning still reject a deck containing custom
shows because those operations change presentation identity topology;
`slide.moveTo(...)` is safe because show entries reference retained
presentation relationships rather than display indexes.

If a list contains extensions, unknown attributes/children, an unresolved
relationship, duplicate names/native IDs, an empty show, or another graph
outside this profile, OpenChestnut exposes no incomplete semantic facade. It
marks the list opaque, preserves its exact source XML, and rejects attempts to
replace it.

Custom-show hyperlinks are a separate action graph. The model can describe
them, but the current OpenChestnut run-hyperlink slice still fails closed
instead of guessing the native `ppaction` encoding.

## Audited Imported Edit

Use the shipped workflow when one exact imported show must be renamed and its
ordered slide route changed:

```ts
import { editPptxCustomShow } from "../../../examples/openchestnut-custom-show-workflow.mjs";

await editPptxCustomShow({
  inputPath: "input.pptx",
  outputPath: "output.pptx",
  auditPath: "custom-show-audit.json",
  expectedName: "Board route",
  replacementName: "Executive route",
  orderedSlideNames: ["Appendix", "Overview", "Appendix"],
});
```

The transaction requires one exact show and one exact slide for every supplied
name, keeps the source immutable, proves that only `ppt/presentation.xml`
changed, verifies native ID and non-target shows, reimports the result, compares
all model SVGs, and emits a source/output-bound audit. Run LibreOffice/Poppler
review as the final native-host QA when those tools are available.
