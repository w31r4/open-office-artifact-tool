# Speaker Notes

Speaker notes live on `slide.speakerNotes`. Their public text is always
available through `notes.text`; source-free notes may additionally use the same
paragraph/run data as slide text for a relationship-free talk track. Notes-page
layout, Notes Master styling, note-local links/fields/picture bullets, list
styles, and arbitrary notes shapes remain source-preserved.

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

## Rich Paragraph And Run Notes

Pass structured paragraphs to `slide.addNotes(...)`, or assign the same data to
`notes.textFrame.paragraphs`. This is the regular Presentation text subset:
ordinary runs, direct run styles, paragraph alignment/spacing, character or
auto-number bullets, and line breaks. `notes.text` is the exact flattened
string, joined by `\n` between paragraphs.

```js
slide.addNotes([
  {
    bulletCharacter: "•",
    runs: [
      { text: "Open with ", style: { bold: true, fontSize: 18, color: "#0F172A" } },
      { text: "the customer outcome.", style: { italic: true, fontSize: 18 } },
    ],
  },
  { bulletNone: true, runs: [{ text: "Then explain the operating model." }] },
]);

const paragraphs = slide.speakerNotes.textFrame.paragraphs;
paragraphs[0].runs[1].text = "the operating decision.";
slide.speakerNotes.textFrame.paragraphs = paragraphs;
```

For an imported rich notes body, paragraph count, run count, and inline kind
(`text` versus line break) are fixed. Update the structured paragraphs and let
their flattened value become `notes.text`; do not set `notes.text` alone to
replace a multi-run body. That would flatten its source topology and fails
closed. The legacy text-only profile remains compatible with full-text
replacement.

## Bounded Imported Rich-Run Transaction

When a task needs one known ordinary rich-notes run changed alongside one known
title, use the shipped source-bound transaction rather than editing package XML
or replacing the entire notes string:

```js
import { editPptxRichSpeakerNotes } from "../../../examples/openchestnut-rich-speaker-notes-edit-workflow.mjs";

await editPptxRichSpeakerNotes({
  inputPath: "input.pptx",
  outputPath: "output/edited.pptx",
  auditPath: "output/audit.json",
  slideName: "Go-no-go decision",
  titleShapeName: "approval-title",
  expectedTitle: "Decision: hold for legal review",
  replacementTitle: "Decision: approve controlled rollout",
  paragraphIndex: 0,
  runIndex: 1,
  expectedRunText: "the pending legal condition.",
  replacementRunText: "the approved control set.",
});
```

The workflow is intentionally fixture-shaped: it requires an imported,
editable, relationship-free NotesSlide; snapshots the complete paragraph/run
tree; validates the selected run's exact text and direct style; changes only
that run; and reimports to prove the paragraph/run topology and every sibling
run remain exact. Its audit binds source/output hashes, slide/title/notes IDs,
paragraph/run indices, expected/replacement styles, source-bound capability,
second import, semantic verification, and model SVG evidence. It fails closed
for fields, hyperlinks, picture bullets, list styles/body properties, any
different inline topology, missing or ambiguous target, changed source run, or
identity/geometry/background/order drift. It is not a general NotesSlide
reflow, rich-text replacement, or template editor.

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
- `editable`: that existing NotesSlide has either the legacy text-only profile
  or the bounded relationship-free paragraph/run profile.
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

- An existing notes part with one bounded body placeholder is hash-bound. The
  legacy one-run-per-paragraph form accepts full-text replacement; a recognized
  rich paragraph/run form accepts fixed-topology paragraph/run edits.
- Imported fields, hyperlinks, picture bullets, notes-body list styles/body
  properties, notes-page layout, and irregular notes are inspectable and
  preserved, but mutation fails closed instead of flattening their formatting
  or graph.
- A source-bound slide with no notes part can gain one only when
  `speakerNotes.capability.addable` is true. Missing/inconsistent NotesMaster
  lists, multiple masters, missing slide ownership, or an unusable SlideMaster
  ThemePart remain fail-closed.
- Notes visibility, notes-page layout, Notes Master styling, and arbitrary
  notes shapes are not part of the current semantic contract.
