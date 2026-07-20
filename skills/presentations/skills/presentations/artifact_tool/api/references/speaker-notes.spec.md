# Speaker Notes

Speaker notes live on `slide.speakerNotes`. The public OpenChestnut boundary is
plain text: the notes object is stable and resolvable, while rich notes layout,
multiple text styles, and arbitrary notes shapes remain source-preserved.

## Create Or Replace Text

```js
const notes = slide.addNotes([
  "Open with the customer problem, not the feature list.",
  "Call out the chart source before discussing the slope.",
].join("\n"));

notes.setText("Opening line\nSecond talking point");
notes.textFrame.setText("The same plain-text contract through textFrame");
```

`slide.addNotes(text)`, `notes.setText(text)`, and
`notes.textFrame.setText(text)` replace the full notes text. OpenChestnut creates
the canonical Notes Master/Notes Slide graph when authoring a new deck.

## Append And Clear

```js
slide.speakerNotes.append("\nFollow-up: confirm launch date with PM.");
slide.speakerNotes.clear();
```

## Inspect And Resolve

```js
const result = presentation.inspect({ kind: "slide,notes", maxChars: 8000 });
const notes = presentation.resolve(`${slide.id}/notes`);
console.log(notes.text, notes.capability);
```

`notes.capability` returns a defensive record with four booleans:

- `sourceBound`: the slide came from an imported PPTX.
- `partPresent`: the source SlidePart already has a NotesSlide.
- `editable`: that existing NotesSlide has the bounded plain-text body profile.
- `addable`: the source has no NotesSlide, but OpenChestnut has re-proved that
  the presentation can safely reuse or create the canonical NotesMaster graph.

The capability is preflight evidence for an Agent, not authority granted by a
mutable flag. Export opens the original package and independently re-proves the
NotesMaster, SlideMaster/Theme, slide ownership, relationships, and package
budgets before writing anything.

## Add Notes To An Imported Notes-Absent Slide

Use the shipped transaction rather than patching relationships by hand:

```js
import { addPptxSpeakerNotes } from "../../../examples/openchestnut-speaker-notes-add-workflow.mjs";

await addPptxSpeakerNotes({
  inputPath: "input.pptx",
  outputPath: "output/with-notes.pptx",
  auditPath: "output/with-notes.audit.json",
  slideName: "Speaker notes target",
  notes: "Lead with the evidence.\nClose with the requested decision.",
});
```

The workflow requires exactly one named imported slide with
`{ sourceBound: true, partPresent: false, editable: false, addable: true }`.
It keeps the input immutable, writes through a temporary output, reimports,
checks exact text and unchanged visible slide semantics, renders model SVG,
and audits the OPC graph. If the presentation already has one existing reusable
NotesMaster, it is reused byte-for-byte. Otherwise OpenChestnut creates one
canonical NotesMaster that shares the first ordered SlideMaster's existing
ThemePart. The new NotesSlide has exactly one NotesMaster relationship and one
back-reference to its owning SlidePart. There is no silent reconstruction or
fallback codec.

## Imported Deck Boundary

- An existing notes part with one bounded body placeholder is hash-bound and
  editable through the plain-text methods above.
- Rich or irregular imported notes are inspectable and preserved, but mutation
  fails closed instead of flattening their formatting.
- A source-bound slide with no notes part can gain one only when
  `speakerNotes.capability.addable` is true. Missing/inconsistent NotesMaster
  lists, multiple masters, missing slide ownership, or an unusable SlideMaster
  ThemePart remain fail-closed.
- Notes visibility, notes-page layout, Notes Master styling, and arbitrary
  notes shapes are not part of the current semantic contract.
