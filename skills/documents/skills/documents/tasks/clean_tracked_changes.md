# Accept or reject tracked changes

## Goal

Given a DOCX with tracked changes, produce a distinct accepted or rejected copy, prove which source bytes were used, and review the native render.

## Preferred bounded workflow

For direct whole-paragraph `w:ins` / `w:del` revisions containing one recognized run, use the shipped OpenChestnut workflow:

```bash
node examples/openchestnut-revision-finalization-workflow.mjs \
  input.docx accepted.docx accepted.audit.json accept

# Or reject the revisions and preserve an existing trackRevisions setting:
node examples/openchestnut-revision-finalization-workflow.mjs \
  input.docx rejected.docx rejected.audit.json reject --keep-tracking
```

The workflow:

1. reads the original bytes and computes SHA-256;
2. imports and inspects the modeled revisions;
3. calls `DocumentFile.finalizeRevisions` with that expected source hash;
4. accepts or rejects through OpenChestnut without reconstructing the DOCX in JavaScript;
5. admits changes only to `word/document.xml` and, when the existing tracking flag changes, `word/settings.xml`;
6. re-imports the output, requires zero remaining revisions, verifies the expected paragraph projection, and checks the tracking state;
7. leaves the source immutable, refuses to overwrite output/audit paths, and writes a byte-bound audit.

`--keep-tracking` preserves an existing `<w:trackRevisions/>`; it does not enable tracking when the source did not have it. Without the flag, finalization removes the setting so later manual edits are not automatically tracked.

`DocumentFile.finalizeRevisions` also accepts the exact adjacent direct-run `w:del` + `w:ins` pair produced by `DocumentFile.addTrackedReplacement` and `examples/openchestnut-tracked-replacement-workflow.mjs`. The whole-block workflow above performs a richer model projection and therefore expects modeled `kind: "change"` blocks. For an inline pair, hash the tracked output, call the same API directly, retain both `metadata.trackedReplacement` and `metadata.revisionFinalization`, reimport the accepted or rejected text, inspect the package for zero remaining revision elements, and native-render every page.

## Capability boundary

The native primitive deliberately fails closed for:

- mixed ordinary and revision runs other than one exact adjacent bounded deletion/insertion pair;
- more than one run inside a revision wrapper;
- nested revisions;
- `w:moveFrom` / `w:moveTo`;
- run, paragraph, table, row, or cell property revisions;
- revisions in tables, headers, footers, notes, comments, or any story other than a direct body paragraph in `word/document.xml`;
- malformed wrappers, a wrong source hash, no revisions, or an unsupported mode.

Do not flatten those graphs through the public model. Use the explicit Python helper for a reviewed package transformation, use Microsoft Word's review commands when its semantics are required, or report the unsupported boundary.

## Explicit broader package helper

```bash
python scripts/accept_tracked_changes.py input.docx --mode report
python scripts/accept_tracked_changes.py input.docx --mode accept --out accepted.docx
python scripts/accept_tracked_changes.py input.docx --mode reject --out rejected.docx
```

This is a pragmatic OOXML helper, not a silent fallback or a complete Word revision engine. Re-run `--mode report` on its output and retain a source/output audit.

## Native render review

```bash
python render_docx.py accepted.docx --output_dir out_accept
```

Inspect every `out_accept/page-*.png` at 100%:

- no redlines or strikethrough remain;
- accepted insertions and rejected deletions remain, while their opposites are gone;
- formatting around retained runs is intact;
- headers, footers, tables, pagination, and spacing did not drift.

Model SVG verification is useful but does not replace this native DOCX render.
