# OOXML: Word comments (true comments)

## Important reality
**PDF/image rendering is not a reliable way to verify comments.** Comments often do not render at all in headless LibreOffice.

If the task requires verifying comments, do a structural check.

## Minimum wiring for a comment
A comment requires three cooperating pieces:

1) `word/comments.xml` exists and contains a `<w:comment w:id="...">` with the comment text.
2) `word/document.xml` contains anchors:
   - `<w:commentRangeStart w:id="..."/>`
   - `<w:commentRangeEnd w:id="..."/>`
   - a `<w:commentReference w:id="..."/>` in a run after the range
3) Relationships + content-types:
   - `word/_rels/document.xml.rels`: add a relationship of Type `http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments` targeting `comments.xml`
   - `[Content_Types].xml`: add an Override for `/word/comments.xml` with ContentType `application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml`

## Typical insertion strategy
- Identify the target paragraph or run range in `document.xml`.
- Insert `commentRangeStart` before the first run you want covered.
- Insert `commentRangeEnd` after the last run you want covered.
- Append a run containing `commentReference`.
- Create or append the comment body in `comments.xml`.

## Bounded modern threads

OpenChestnut's public modern profile keeps the classic root anchor above and
adds a paragraph identity to every comment body. It then connects only direct
replies through these support parts:

- `word/commentsExtended.xml`: `w15:commentEx` records keyed by `w14:paraId`,
  with optional `w15:paraIdParent` and `w15:done`.
- `word/commentsIds.xml`: optional paragraph-to-durable-ID mapping.
- `word/commentsExtensible.xml`: optional UTC and intelligent-placeholder
  metadata keyed by durable ID.
- `word/people.xml`: optional author presence records with provider/user IDs.

A reply shares its root paragraph target but does not create a second range in
`document.xml`. Source-free IDs are deterministic when omitted. On import,
thread identity, parentage, support-part graphs, people data, timestamps, and
anchors are source-bound; only comment text and recognized resolved state may
change. Nested replies, orphan/cross-target parents, extensions, connected
support-part relationships, or irregular mappings fail closed instead of being
flattened.

## Recommended: use the helper script
See `scripts/docx_ooxml_patch.py` (`--add-comment`). It:
- auto-picks a non-colliding comment id by scanning the DOCX
- **appends** to `word/comments.xml` if it already exists (does not overwrite existing comments)
- reuses an existing comments relationship if present (avoids duplicate rels)

## Structural verification checklist
- `comments.xml` present in the ZIP
- `document.xml.rels` has a comments relationship
- `[Content_Types].xml` includes the Override
- For each comment id, `document.xml` has start/end/reference anchors
- For a modern graph, every comment body paragraph has one `w14:paraId` and one
  matching `w15:commentEx`
- Every direct reply's `w15:paraIdParent` names its root paragraph ID
- IDs/extensible/people parts, when present, contain no unknown child graph
