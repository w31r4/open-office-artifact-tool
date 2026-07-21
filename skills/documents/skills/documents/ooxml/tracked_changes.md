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

OpenChestnut writes these as native `<w:ins>` or `<w:del>` markup, re-imports them semantically, and permits fixed-topology text/author/date edits. It keeps `w:id` and unmodeled run/paragraph formatting source-bound.

For one exact replacement inside an existing ordinary paragraph, use the source-bound file primitive instead of rebuilding the document model:

```js
import crypto from "node:crypto";
import fs from "node:fs/promises";
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const bytes = await fs.readFile("input.docx");
const source = new FileBlob(bytes);
const document = await DocumentFile.importDocx(source);
const targetBlockIndex = document.blocks.findIndex(
  (block) => block.kind === "paragraph" && block.text === "The term is 30 days.",
);
const reviewed = await DocumentFile.addTrackedReplacement(source, {
  expectedSourceSha256: crypto.createHash("sha256").update(bytes).digest("hex"),
  target: { kind: "paragraph", blockIndex: targetBlockIndex },
  expectedText: "The term is 30 days.",
  search: "30 days",
  replacement: "45 days",
  author: "Reviewer",
  date: "2026-07-21T09:30:00Z",
});
await reviewed.save("reviewed.docx");
```

The same transaction accepts `{ kind: "tableCell", blockIndex, row, column }` when the exact imported block is a direct body table with a stable physical grid and the selected non-continuation cell contains exactly one direct paragraph. `targetBlockIndex` remains a paragraph-only compatibility selector and cannot be combined with `target`.

The structured selector and full paragraph/cell snapshot bind the target in the exact source bytes. `search` must occur once inside one direct ordinary `w:r/w:t`; OpenChestnut clones that run's formatting into one adjacent `w:del` + `w:ins` pair, uses `w:delText` for the old text, allocates collision-free package-local IDs, and permits only `word/document.xml` to change. Stale text, duplicate/cross-run matches, multi-paragraph/nested/continuation/irregular table cells, hyperlinks, fields, controls, drawings, existing revisions, and other native topologies fail closed. The operation returns the re-proved target plus source/output, paragraph-element, deleted/inserted-text, native-ID, block/body-index, and changed-part evidence in `metadata.trackedReplacement`. Prefer `examples/openchestnut-tracked-replacement-workflow.mjs` when publishing because it also discovers one unique paragraph/table cell, protects the source, refuses overwrite, reimports, renders, and writes an audit.

The same public boundary now supports the native future-edit setting and file-level finalization:

```js
document.setSettings({ trackRevisions: true });

const finalized = await DocumentFile.finalizeRevisions(sourceDocx, {
  mode: "accept", // or "reject"
  expectedSourceSha256,
  keepTracking: false,
});
```

`finalizeRevisions` accepts the original DOCX bytes directly, rechecks the exact source SHA-256 inside OpenChestnut, rewrites only `word/document.xml` and—when the existing tracking flag changes—`word/settings.xml`, and returns source/output hashes, revision counts, tracking state, and the exact changed-part list in `metadata.revisionFinalization`. It supports direct body whole-paragraph `w:ins`/`w:del` wrappers with one recognized run plus exact adjacent direct-run pairs in direct body paragraphs or bounded table cells authored above. Any other revision element, story part, mixed or nested graph, move, property change, irregular table target, or malformed wrapper fails closed before an output is published. `keepTracking` preserves an existing tracking flag; it does not silently enable one that was absent.

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
Use this only for graphs outside the typed one-node replacement. See `scripts/docx_ooxml_patch.py` for a runnable patcher that:
- enables `<w:trackRevisions/>`
- converts an existing `<w:ins>` to `<w:del>` and inserts a new `<w:ins>`

The CLI defaults to auto-generated `w:id` values (`--del-id auto --ins-id auto`) by scanning existing ids and choosing new ones.

For a bounded whole-block accept/reject transaction, prefer `examples/openchestnut-revision-finalization-workflow.mjs`. It inspects the modeled revisions, binds the source hash, calls the typed OpenChestnut primitive, re-imports the output, proves that no revisions remain, refuses overwrite, and writes a byte-bound audit. For an inline pair, call the same `DocumentFile.finalizeRevisions` API against the tracked output and retain both operation audits. The Python `accept_tracked_changes.py` helper is an explicit broader package route, not a silent fallback from the public API.

## Verification
- Render to PDF/PNG for layout sanity (`tasks/verify_render.md`)
- Confirm Word shows the change as tracked
- Re-import through `DocumentFile.importDocx`; supported whole-paragraph revisions inspect as `kind: "change"`, while the exact inline pair exposes its accepted-view paragraph or table-cell value as source-bound and read-only until finalization
- Inspect `word/document.xml`: the old text must be one `w:delText` and the replacement one adjacent `w:t`; after finalization no revision element may remain
- Be aware: renders usually show redlines, but always verify the OOXML is correct too
