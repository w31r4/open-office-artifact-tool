# True footnotes and endnotes

## Goal

Create, edit, inspect, or audit **true** DOCX footnotes/endnotes and prove that
the references, note parts, semantics, and rendered pages agree. Never emulate a
footnote with footer text or a superscript character.

## Ordinary bounded workflow: public API

Use `DocumentModel` and bundled OpenChestnut when each note has one plain-text
body and is anchored at the end of one paragraph or list item. The bounded
profile permits at most one note per target block.

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({ name: "Research note", blocks: [] });
const claim = document.addParagraph("The pilot met its release threshold.");
const provenance = document.addParagraph("The evidence snapshot is archived.");

document.addFootnote(claim, "Pilot report, section 4.2.");
document.addEndnote(provenance, "Evidence snapshot dated 2026-07-17.");

const first = await DocumentFile.exportDocx(document);
const imported = await DocumentFile.importDocx(first);
imported.notes.find((note) => note.kind === "footnote").text =
  "Pilot report, section 4.2, independently reviewed.";
const output = await DocumentFile.exportDocx(imported);
await output.save("notes.docx");
```

After import, resolve note objects again from `document.notes`, `inspect()`, or
`resolve(note.id)`. A recognized note body may change text, but its kind,
target, native ID, reference position, and topology are source-bound. The
target paragraph/list item itself is read-only because moving or rebuilding its
reference run would change the native note graph.

## Native package contract

Footnotes and endnotes live in separate package parts:

- `word/footnotes.xml`
- `word/endnotes.xml`

The body points to them with `w:footnoteReference` or `w:endnoteReference`.
Source-free note parts include the required separator (`w:id=-1`) and
continuation-separator (`w:id=0`) entries. OpenChestnut allocates positive
native IDs independently for footnotes and endnotes.

## Inspect and audit

Use semantic inspection first:

```js
const document = await DocumentFile.importDocx(input);
console.log(document.inspect({ kind: "document,note,footnote,endnote" }).ndjson);
for (const note of document.notes) {
  console.log(document.resolve(note.id));
}
```

Then use the package reporter to inventory all references and note IDs,
including irregular graphs that remain opaque to the public model:

```bash
python scripts/footnotes_report.py input.docx
```

## Explicit advanced package workflow

Use `insert_note.py` only when the requested operation is deliberately
package-level and cannot fit the bounded public model—for example, inserting at
an exact in-paragraph marker in a controlled template. It is not an automatic
fallback and must not be used to conceal an unsupported imported graph.

1. Put `[[FN]]` or `[[EN]]` at the exact controlled insertion point.
2. Patch a copy, never the source:

```bash
python scripts/insert_note.py input.docx --kind footnote --marker "[[FN]]" --text "Footnote text" --out with_fn.docx
python scripts/insert_note.py input.docx --kind endnote  --marker "[[EN]]" --text "Endnote text"  --out with_en.docx
```

3. Run `footnotes_report.py`, inspect the package, render every page, and keep
an audit record of the explicit low-level operation.

Multi-paragraph/rich note bodies, reused references, multiple notes on one
target, custom separator/numbering/restart graphs, anchor movement, or other
irregular topologies remain opaque/source-bound through the public codec. If
the narrow helper cannot prove a safe transformation, fail closed.

## Render and verification gate

```bash
python render_docx.py notes.docx --output_dir rendered_notes
```

Verify all of the following:

- semantic re-import contains the expected note kind, target, and body text;
- `document.xml` contains the expected reference IDs and the matching note
  part contains each positive ID exactly once;
- separators `-1` and `0` exist for a source-free note part;
- footnotes render at the expected page bottom and endnotes in the note section;
- numbering is unique and ordered as intended;
- long note text wraps without clipping or overlap;
- unrelated pages/content remain unchanged for an imported-file edit.

For high-stakes delivery, add Microsoft Word application validation when the
environment is available; LibreOffice rendering remains required local visual
evidence, not proof of every Word-specific behavior.
