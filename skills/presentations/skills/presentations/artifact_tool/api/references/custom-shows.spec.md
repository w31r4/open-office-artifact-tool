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

overview.shapes.add({
  name: "open-board-route",
  position: { left: 80, top: 80, width: 320, height: 48 },
  text: [{
    runs: [{
      text: "Open board route",
      link: { customShow: "Board route", returnToSlide: true },
    }],
  }],
});
```

Names are trimmed, non-empty, at most 255 characters, and
case-insensitively unique. A show contains 1 through 16,384 slide references;
a deck contains at most 4,096 shows. Omit `nativeId` for new content and let
the model assign a unique unsigned 32-bit identity during verification/export.

Use `getItem(indexOrIdOrName)`, `inspect({ kind: "customShow" })`, or
`resolve(customShow.id)` to find a show. Use `show.setSlides(slides)` to replace
its ordered slide membership.

## Run Hyperlinks

A canonical text run may select exactly one existing custom show by exact
name. `returnToSlide` is optional and presence-aware; `true` writes
`&return=true`, `false` writes `&return=false`, and omission writes neither.
OpenChestnut emits a relationship-free DrawingML click action:

```xml
<a:hlinkClick r:id="" action="ppaction://customshow?id=7&amp;return=true"/>
```

The protobuf wire carries the stable custom-show facade ID rather than the
mutable display name. Consequently, renaming a canonical imported show keeps
existing links bound to the same native `p:custShow/@id`; a second import
projects the new show name into `run.link.customShow`. Explicitly changing the
run's target name retargets it only to another existing canonical show.

Missing names, opaque show lists, dangling native IDs, malformed action URIs,
or custom-show actions that carry a relationship ID are not guessed. Imported
unknown actions remain exact-source-preserved and replacing them fails closed.
The bounded slide clone profile accepts a canonical relationship-free action
only after resolving its native ID through this catalog. It copies the action
onto the new SlidePart, creates no hyperlink/slide relationship, and proves the
presentation-wide show identity and ordered membership are unchanged; the
clone is not implicitly added to the show. Slide deletion still rejects
custom-show identity references.

## Imported Canonical Shows

A canonical imported list exposes semantic shows with hash-bound source
evidence. OpenChestnut permits two edits in place:

- change an existing show's `name` without colliding with another name;
- call `setSlides(...)` with an ordered list of retained slides from the same
  imported presentation.

The show count/order, facade ID, and native `p:custShow/@id` remain fixed.
Adding, removing, reordering, or replacing show objects fails closed. Slide
deletion still rejects a deck when the target participates in custom-show
identity topology. Source-bound slide cloning is allowed only for the canonical
run-action profile above and leaves every show member list fixed.
`slide.moveTo(...)` is safe because show entries reference retained
presentation relationships rather than display indexes.

If a list contains extensions, unknown attributes/children, an unresolved
relationship, duplicate names/native IDs, an empty show, or another graph
outside this profile, OpenChestnut exposes no incomplete semantic facade. It
marks the list opaque, preserves its exact source XML, and rejects attempts to
replace it.

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
normalized model SVG visual content, inventories run links bound to the fixed
show identity, and emits a source/output-bound audit. Run LibreOffice/Poppler
review as the final native-host QA when those tools are available.
