# Troubleshooting: run splitting ("why isn't my replace working?")

## Reality
Word splits text into runs unpredictably (style changes, proofing boundaries, fields, etc.).
So searching for a substring and replacing it "as text" often fails.

## Practical strategies
- Import with `DocumentFile.importDocx(...)` and inspect `textEditable` versus
  `textPatchable` before choosing an operation. Do not infer editability from
  concatenated paragraph text.
- The public source-bound `resolve(.../text).replace(old, next)` route handles a
  unique literal inside one ordinary `<w:r>/<w:t>` or across adjacent non-empty
  ordinary runs only when their exact `<w:rPr>` markup matches. It retains the
  native run nodes and changes text payloads only.
- Mixed run properties, empty-run gaps, paragraph boundaries, fields,
  hyperlinks, controls, revisions, and duplicate visible matches fail closed.
  Do not flatten the paragraph to make the replacement pass.
- Use `../examples/openchestnut-source-text-patch-workflow.mjs` for the full
  immutable-input, changed-part, reimport, audit, and render-review workflow.
- When you must replace a token, consider inserting a hidden marker run first (during `python-docx` authoring) so you can reliably locate the target later when patching OOXML.
- For tracked changes replacements, wrap **exact runs** you want deleted as `<w:del>`, then insert new `<w:ins>` adjacent.

## Helper script
`scripts/docx_ooxml_patch.py` contains utilities that:
- find paragraphs by simple predicates (e.g., indentation)
- replace the Nth tracked insertion inside a paragraph
