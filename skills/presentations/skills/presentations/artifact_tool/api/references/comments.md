# Comments

OpenChestnut exposes two native PPTX comment wire families. Select one for the
whole presentation with `Presentation.create({ commentFormat })`; never mix
them in one artifact.

## Legacy slide annotations

`commentFormat: "legacy"` is the default. It writes standard legacy
PresentationML with one author, one text item, and one explicit slide coordinate:

```js
const presentation = Presentation.create({ commentFormat: "legacy" });
const slide = presentation.slides.add();
slide.comments.addThread(undefined, "Confirm the source before delivery.", {
  author: "Presentation Reviewer",
  created: "2026-07-18T09:30:00.000Z",
  position: { x: 1040, y: 84, unit: "px" },
});
```

Pass `undefined` as the target. Legacy `p:cm` has no element/text-range anchor,
reply graph, reaction, or resolved state.

For an imported presentation, inspect the collection before mutation:

```js
const capability = slide.comments.capability;
// { sourceBound, format, partPresent, addable }
```

`addable` is true only when the complete source presentation has no legacy
`CommentAuthorsPart`/`SlideCommentsPart` and no Office 2021 author/comment Part.
It is a defensive preflight snapshot, not authorization stored in mutable JS.
OpenChestnut re-proves the source bytes on export. The first bounded
source-bound authoring transaction may create one canonical shared legacy
author catalog and one closed slide-local comments Part for each requested
original slide. Relationship IDs are allocated against internal, external,
hyperlink, and data relationships, so an unrelated `rIdComments1` or
`rIdCommentAuthors1` cannot collide.

Use the audited single-comment workflow for the common Agent review task:

```js
import { addPptxLegacyReviewComment } from "../../../examples/openchestnut-legacy-comment-add-workflow.mjs";

await addPptxLegacyReviewComment({
  inputPath: "input.pptx",
  outputPath: "output/with-review.pptx",
  auditPath: "output/with-review.audit.json",
  slideName: "Imported review target",
  text: "Confirm the imported evidence before delivery.",
  author: "Review Owner",
  created: "2026-07-20T03:04:05Z",
  position: { x: 360, y: 240, unit: "px" },
});
```

The workflow protects the source bytes, requires a unique target, checks the
capability, writes through a temporary file, and publishes without overwrite.
Its independent OPC audit permits exactly the new `ppt/commentAuthors.xml`, one
numbered `ppt/comments/commentN.xml`, `[Content_Types].xml`, and the two owner
relationship Parts. It then reimports exact semantics, compares model SVG, and
emits a source/output-bound audit. Final LibreOffice/Poppler pages must remain
pixel-identical because review annotations are not slideshow content.

After export and reimport, `partPresent` becomes true and `addable` becomes
false for every slide. Recognized imported legacy comments are inspectable but
unchanged-only; adding a second comment, editing text/author/time/coordinate, or
forging the capability fails closed. The bounded `slide.duplicate()` workflow
may byte-copy one closed legacy comments leaf and share its verified immutable
author catalog; that is preservation, not in-place editing.

## Office 2021 modern threads

`commentFormat: "modern"` writes a native Office 2021 `p188:authorLst` plus one
`p188:cmLst` part per commented slide. The bounded profile supports:

- one root and direct replies;
- independent GUID/person metadata and ISO-8601 time per comment;
- `active`, `resolved`, or `closed` status;
- one top-level shape, image, table/chart, connector, or group anchor;
- one whole-shape or exact `textMatch` range anchor;
- explicit DrawingML comment coordinates.

Use stable brace-delimited GUIDs when deterministic output matters:

```js
const presentation = Presentation.create({ commentFormat: "modern" });
const slide = presentation.slides.add({ name: "Decision review" });
const title = slide.shapes.add({
  id: "decision-title",
  text: "Customer evidence is ready",
  position: { left: 80, top: 80, width: 560, height: 96 },
});

const thread = slide.comments.addThread({
  textMatch: { element: title, query: "Customer evidence", occurrence: 0 },
}, "Confirm the evidence before delivery.", {
  id: "{11111111-1111-4111-8111-111111111111}",
  nativeFormat: "modern",
  position: { x: 1234500, y: 2345600, unit: "emu" },
  comments: [{
    nativeId: "{11111111-1111-4111-8111-111111111111}",
    author: "Review Owner",
    person: {
      id: "{AAAAAAAA-AAAA-4AAA-8AAA-AAAAAAAAAAAA}",
      name: "Review Owner",
      initials: "RO",
      userId: "review.owner@example.test",
      providerId: "None",
    },
    text: "Confirm the evidence before delivery.",
    created: "2026-07-19T02:55:00Z",
    status: "active",
  }],
});

thread.addReply("Evidence is attached.", {
  nativeId: "{22222222-2222-4222-8222-222222222222}",
  author: "Evidence Owner",
  person: {
    id: "{BBBBBBBB-BBBB-4BBB-8BBB-BBBBBBBBBBBB}",
    name: "Evidence Owner",
    initials: "EO",
    userId: "evidence.owner@example.test",
    providerId: "None",
  },
  created: "2026-07-19T03:05:00Z",
  status: "active",
});
```

Direct target forms are also accepted:

```js
slide.comments.addThread({ element: image }, "Check the crop.", options);
slide.comments.addThread({ textRange: slide.resolve(`${title.id}/text`) }, "Review all title text.", options);
slide.comments.addThread(title, "Review this shape.", options);
```

Modern slide-level anchors and nested group-child moniker chains are not in the
bounded profile. Use a supported top-level element or text range.

## Imported edit boundary

Recognized imported modern threads expose their original root and direct
replies. Only existing comment text and status are mutable:

```js
const presentation = await PresentationFile.importPptx(input);
const records = presentation.inspect({ kind: "comment", maxChars: 4000 });
console.log(records.ndjson);

const thread = presentation.resolve(threadId);
thread.comments[0].text = "Evidence confirmed for delivery.";
thread.comments[1].text = "Recorded in the decision log.";
thread.resolve(); // root status -> resolved

const output = await PresentationFile.exportPptx(presentation);
const roundTrip = await PresentationFile.importPptx(output);
```

Export re-proves the original author catalog, comment-part hash, thread/comment
GUIDs, author/person metadata, timestamps, target moniker, text range, position,
reply count/order, relationships, and fixed-topology hash. Changing any of
those fields, adding/removing a reply, or moving the thread fails closed before
output. `thread.reopen()` changes the root status back to `active` under the same
contract.

## Unsupported graphs

Reactions/likes, task fields, extensions, rich text, nested replies, unknown
anchors, multiple comment parts for one slide, mixed legacy/modern parts, or
comment/author parts with child, external, hyperlink, or data relationships are
not flattened. They remain opaque/source-bound; a modeled replacement or edit
is rejected. Existing legacy catalogs are intentionally not extended in this
slice, even when they look simple: author/index state remains package-local and
unchanged-only until a separately modeled edit contract exists.

Run the shipped end-to-end example for a source-free root/reply thread followed
by an imported fixed-topology text/status edit, second import, package inspect,
model render, and byte-bound audit:

```bash
node examples/openchestnut-modern-comment-workflow.mjs \
  output/decision-review.pptx output/modern-comments-audit.json
```
