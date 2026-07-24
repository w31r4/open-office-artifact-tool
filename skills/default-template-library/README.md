# Default Template Library

This checked-in catalog provides 20 independently selectable, reference-backed Office templates: 7 documents, 7 presentations, and 6 spreadsheets.

The library is copied from the MIT-licensed `office-artifact-tool` repository at commit `256cb31bfe0a07b3cef0051b6b159342be381378` (`Add default Office template library`). The retained Office files and PNG previews are intentionally byte-for-byte source assets; their inventory and individual SHA-256 values are recorded in `integrity.json`. See [LICENSE.md](LICENSE.md).

## Layout

```text
skills/default-template-library/
├── manifest.json
├── assets/icon.svg
└── skills/artifact-template-<name>/
    ├── SKILL.md
    ├── artifact-template.json
    ├── agents/agent.yaml
    └── assets/
        ├── preview.png
        └── reference.docx | reference.pptx | reference.xlsx
```

Each nested skill retains its reference Office file and preview image. Its
schema-v2 `artifact-template.json` adds intended uses, avoid cases, audiences,
content shapes, visual traits, visual commitment, verified edit operations,
license/source provenance, and retained-asset hashes. OfficeKit can therefore
shortlist templates without loading all twenty Skill descriptions or opening
every Office file.

Use the named template skill to create a new artifact while preserving the
retained layout and formatting unless the request calls for a change.

These resources are repository-only and are intentionally excluded from the npm package tarball. Create a distinct output from a selected retained reference; never overwrite or mutate the checked-in reference file.

For a guarded working copy, run:

```sh
node skills/default-template-library/scripts/materialize-template.mjs \
  --template-id artifact-template-system-design \
  --output /absolute/path/system-design.docx
```

The materializer checks the retained source hash, refuses existing output and
audit paths, and writes a byte-identical working copy plus an audit record.
Use the matching Documents, Presentations, or Spreadsheets Skill to inspect,
edit, render, and verify that output. Complex source-bound Office graphs are
preserved only while unchanged; unsupported topology edits fail explicitly.
All seven retained PPTX templates expose at least one recognized SlidePart
placeholder whose existing text can be replaced through the bounded
Presentations workflow while its native identity, geometry, formatting, and
layout binding remain source-bound.
