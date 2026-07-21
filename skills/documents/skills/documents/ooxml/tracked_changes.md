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

OpenChestnut writes these as native `<w:ins>` or `<w:del>` markup, re-imports them semantically, and permits fixed-topology text/author/date edits. It keeps `w:id` and unmodeled run/paragraph formatting source-bound. The same bounded direct whole-paragraph profile supports native accept/reject finalization; mixed normal/revision runs, in-paragraph replacements, nested changes, moves, property changes, and other story parts require an explicit OOXML or Office-host route.

The same public boundary now supports the native future-edit setting and file-level finalization:

```js
document.setSettings({ trackRevisions: true });

const finalized = await DocumentFile.finalizeRevisions(sourceDocx, {
  mode: "accept", // or "reject"
  expectedSourceSha256,
  keepTracking: false,
});
```

`finalizeRevisions` accepts the original DOCX bytes directly, rechecks the exact source SHA-256 inside OpenChestnut, rewrites only `word/document.xml` and—when the existing tracking flag changes—`word/settings.xml`, and returns source/output hashes, revision counts, tracking state, and the exact changed-part list in `metadata.revisionFinalization`. It supports only direct whole-paragraph `w:ins`/`w:del` wrappers with one recognized run. Any other revision element, story part, mixed or nested graph, move, property change, or malformed wrapper fails closed before an output is published. `keepTracking` preserves an existing tracking flag; it does not silently enable one that was absent.

Do not silently rebuild a complex revision graph through the public model. Unsupported imported topologies are visible but read-only and must be preserved unchanged, handled by an explicit package workflow, or finalized in a real Word host.

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

For a bounded accept/reject transaction, prefer `examples/openchestnut-revision-finalization-workflow.mjs`. It inspects the revisions, binds the source hash, calls the typed OpenChestnut primitive, re-imports the output, proves that no revisions remain, refuses overwrite, and writes a byte-bound audit. The Python `accept_tracked_changes.py` helper is an explicit broader package route, not a silent fallback from the public API.

## Verification
- Render to PDF/PNG for layout sanity (`tasks/verify_render.md`)
- Confirm Word shows the change as tracked
- Re-import through `DocumentFile.importDocx`; supported whole-paragraph revisions should inspect as `kind: "change"`, while complex graphs must remain source-bound
- Be aware: renders usually show redlines, but always verify the OOXML is correct too
