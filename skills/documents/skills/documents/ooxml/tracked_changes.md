# OOXML: Tracked changes (true redlines)

## Choose the narrowest safe route

Use the public model for standalone whole-paragraph redlines with one text run:

```js
document.addInsertion("Added wording", {
  author: "Reviewer",
  date: "2026-07-17T08:00:00Z",
});
document.addDeletion("Removed wording", {
  author: "Reviewer",
  date: "2026-07-17T08:05:00Z",
});
```

OpenChestnut writes these as native `<w:ins>` or `<w:del>` markup, re-imports them semantically, and permits fixed-topology text/author/date edits. It keeps `w:id` and unmodeled run/paragraph formatting source-bound. Mixed normal/revision runs, in-paragraph replacements, nested changes, moves, property changes, and accept/reject finalization require the explicit OOXML helpers below.

Do not silently rebuild a complex revision graph through the public model. Unsupported imported topologies are visible but read-only and must be preserved unchanged or handled by an explicit package workflow.

`python-docx` does **not** provide a first-class API for tracked changes.

## Minimum wiring
Tracked-change workflows involve:
- `word/settings.xml`: optional `<w:trackRevisions/>` enables Word to track subsequent user edits; it is distinct from explicitly authored revision markup
- `word/document.xml`: wrap inserted runs with `<w:ins ...>` and deletions with `<w:del ...>`

## Key rules (to avoid broken docs)
- IDs: `w:id` should be an integer string and **must not collide** with existing ids in the document
- `w:author` and `w:date` are strongly recommended
- Deletions must use `<w:delText>` (not `<w:t>`) inside `<w:del>`
- Word can split text into many runs; operate at run granularity

## Example pattern: replace a word via tracked delete + tracked insert
Pseudo-structure:

```xml
<w:del w:id="202" w:author="Reviewer" w:date="...">
  <w:r><w:delText> old text </w:delText></w:r>
</w:del>
<w:ins w:id="203" w:author="Reviewer" w:date="...">
  <w:r><w:t> new text </w:t></w:r>
</w:ins>
```

## Advanced route: use the helper script
See `scripts/docx_ooxml_patch.py` for a runnable patcher that:
- enables `<w:trackRevisions/>`
- converts an existing `<w:ins>` to `<w:del>` and inserts a new `<w:ins>`

The CLI defaults to auto-generated `w:id` values (`--del-id auto --ins-id auto`) by scanning existing ids and choosing new ones.

## Verification
- Render to PDF/PNG for layout sanity (`tasks/verify_render.md`)
- Confirm Word shows the change as tracked
- Re-import through `DocumentFile.importDocx`; supported whole-paragraph revisions should inspect as `kind: "change"`, while complex graphs must remain source-bound
- Be aware: renders usually show redlines, but always verify the OOXML is correct too
