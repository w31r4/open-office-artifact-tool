# Comments: Extract, Remove, or Preserve for Review

## Goal
Handle reviewer comments in a `.docx` without confusing intermediate artifacts.

Common situations:
- **Review mode**: keep comments and deliver a commented `.docx`.
- **Final mode**: remove comments (and optionally accept tracked changes) and deliver a clean `.docx`.
- **Triage mode**: extract comments into a machine-readable report (JSON/Markdown) for summarization.

> Use `DocumentModel` for classic comments and the bounded modern profile: one root with direct replies, optional durable/UTC/person metadata, and resolved state. Imported modern identity, people metadata, anchors, and topology remain source-bound.

## Add an ordinary classic Word comment

```js
import { DocumentFile, FileBlob } from "open-office-artifact-tool";

const document = await DocumentFile.importDocx(await FileBlob.load("input.docx"));
const target = document.blocks.find(
  (block) => block.kind === "paragraph" && block.text.includes("Payment Terms"),
);
if (!target) throw new Error("Comment target was not found.");

document.addComment(target, "Please confirm Net 45 is acceptable.", {
  author: "Reviewer",
  initials: "RV",
});
await (await DocumentFile.exportDocx(document)).save("reviewed.docx");
```

Re-import and assert the comment text, author, and target before rendering.

## Add a bounded modern thread

```js
const root = document.addComment(target, "Please confirm the evidence.", {
  author: "Lead reviewer",
  resolved: false,
  dateUtc: "2026-07-19T08:00:00Z",
  person: { providerId: "directory", userId: "lead@example.test" },
});
document.replyToComment(root, "Evidence confirmed.", {
  author: "Release reviewer",
  dateUtc: "2026-07-19T08:05:00Z",
  person: { providerId: "directory", userId: "release@example.test" },
});
```

OpenChestnut authors `commentsExtended.xml` and any required IDs, extensible,
and people parts. Only direct replies are supported. A reply to a reply,
cross-target parent, mention/rich body, or irregular imported graph fails closed.

## Add comments at scale (review mode)
For an explicit package-level batch over imported tracked/deleted text, use:
```bash
python scripts/comments_add.py input.docx --out reviewed.docx --author "Reviewer"   --add "Payment Terms=Please confirm Net 45 is acceptable."   --add "Governing Law=Prefer Delaware; any constraints?"   --ignore_case
```
Notes:
- Matching looks across normal text **and** deleted text (`w:delText`), so it can still find anchors in docs with tracked changes.
- The script warns on patterns with no matches; add `--require_all` to fail fast.

## Patch / resolve existing comments
For one uniquely located imported classic comment, prefer the shipped
`examples/openchestnut-classic-comment-edit-workflow.mjs`: it refuses ambiguous
anchors and modern/reply/resolved/presence metadata, changes only text, then
re-imports, verifies, model-renders, atomically writes the DOCX, and records a
source/output-hash-bound audit. For ordinary imported classic comment text,
edit `document.comments[index].text` only after the same uniqueness and
source-bound identity checks.

For one recognized modern root plus one direct reply, use
`examples/openchestnut-modern-comment-thread-workflow.mjs`. It changes only the
two texts and root resolved state through `.resolve()`/`.reopen()`, then proves
the same paragraph/durable/person identities and fixed topology after re-import.
Adding/removing/reparenting an imported comment remains unsupported.

Use the explicit package helper only for comment shapes outside these public
profiles:
```bash
python scripts/comments_extract.py reviewed.docx --out comments.json

# Create a separate patch file (JSON). Example:
# {
#   "ops": [
#     {"id": 0, "append": "Follow-up note"},
#     {"id": 0, "replace": "Full replacement text"},
#     {"id": 0, "resolved": true}
#   ]
# }
# (Set "resolved": false to clear the resolved state.)

python scripts/comments_apply_patch.py reviewed.docx patch.json --out reviewed_v2.docx
```



## Extract comments (triage)
Produces JSON with comment text, author, date (if present), and the anchored snippet.
```bash
python scripts/comments_extract.py input.docx --out comments.json
```

## Remove all comments (final mode)
This removes:
- comment ranges and references in story parts (main doc + headers/footers)
- `word/comments.xml` and any comment-related relationships / content type overrides

```bash
python scripts/comments_strip.py input.docx --out no_comments.docx
```

## Recommended finalize workflow
If the requested deliverable is a **clean final DOCX**:
```bash
python scripts/accept_tracked_changes.py input.docx --mode accept --out accepted.docx
python scripts/comments_strip.py accepted.docx --out final_clean.docx
python render_docx.py final_clean.docx --output_dir out_final_clean
```

## Pitfalls
- Comments can be anchored in headers/footers too; always strip across all story parts.
- Some docs include `commentsExtended.xml` (newer Word). This script removes it if present.
- After stripping, render PNGs and verify nothing disappeared around comment anchors.
