---
name: artifact-template-strategy-memorandum
description: Create a source-free Strategy Memorandum DOCX with a decision frame, evidence table, constraints, recommendation, and accountable next actions. Use when the user asks for a strategy memo or decision memorandum and does not supply a reference file to preserve.
---

# Strategy Memorandum

Generate a new document from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-strategy-memorandum \
  --output /absolute/path/strategy-memorandum.docx \
  --audit /absolute/path/strategy-memorandum.audit.json
```

Then use the Documents workflow to replace only the requested content. Preserve the deliberate sections unless the user asks for a structural change: decision, context, options, evidence, constraints, recommendation, and next actions.

Before delivery, import the generated or edited DOCX, inspect the named blocks, verify it, export to a distinct final path, reimport, render it, and report the audit location. Do not claim that the document preserves an external reference: this template has no retained reference or preview asset.
