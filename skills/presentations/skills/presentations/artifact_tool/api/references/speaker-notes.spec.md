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
console.log(notes.text);
```

## Imported Deck Boundary

- An existing notes part with one bounded body placeholder is hash-bound and
  editable through the plain-text methods above.
- Rich or irregular imported notes are inspectable and preserved, but mutation
  fails closed instead of flattening their formatting.
- A source-bound slide with no notes part cannot gain one in place. Rebuild a
  new deck intentionally if the task requires adding notes to such a slide.
- Notes visibility, notes-page layout, Notes Master styling, and arbitrary
  notes shapes are not part of the current semantic contract.
