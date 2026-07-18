# Office Template Library

This plugin is a clean-room, source-free Office template catalog. It is deliberately not a mirror of any retained third-party Office file, PNG preview, XML package graph, or aggregate binary fingerprint.

## Ready templates

- Strategy Memorandum (`.docx`)
- Design Report (`.docx`)
- Project Kickoff (`.pptx`)
- Operating Review (`.pptx`)
- Financial Budget (`.xlsx`)
- Project Tracker (`.xlsx`)

Each ready entry is generated from reviewed JavaScript source through the public `open-office-artifact-tool` API and bundled OpenChestnut codec. The generator verifies the model, exports it, imports it again, renders a model preview, and records a hash-bound audit. It creates new output paths only; it never silently overwrites a file or substitutes another template.

```sh
node skills/default-template-library/scripts/generate-template.mjs \
  --template-id artifact-template-project-kickoff \
  --output ./project-kickoff.pptx \
  --audit ./project-kickoff.audit.json
```

## Catalog boundary

`catalog.json` lists 20 useful document, presentation, and workbook intents. Six are ready, project-authored generators. A `planned` entry is a compatibility target, not an installed template. It must fail explicitly instead of falling back to a visually unrelated design.

The catalog ships no Office or preview binaries. A future ready entry must be designed in this repository, generated from its own source, and pass import/edit/export/second-import plus render QA before it changes status.

For a user-owned reference file, use [Template Creator](../template-creator/README.md). It retains that file locally under the user's selected template home; it is intentionally separate from this distributable source-free catalog.
