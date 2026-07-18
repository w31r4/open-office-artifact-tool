# Default Template Library provenance

## Source and license

`skills/default-template-library/` is a repository-only import of the 20 Office templates committed in [`office-artifact-tool` `256cb31bfe0a07b3cef0051b6b159342be381378`](https://github.com/w31r4/office-artifact-tool/commit/256cb31bfe0a07b3cef0051b6b159342be381378), **Add default Office template library**.

That repository's root `LICENSE.md` is MIT, `Copyright (c) 2026 w31r4`. The full retained text lives beside the imported library at [`../skills/default-template-library/LICENSE.md`](../skills/default-template-library/LICENSE.md). This is a source/rights record, not legal advice.

## What is preserved

- 20 template Skills: 7 DOCX, 7 PPTX, and 6 XLSX;
- each original `reference.docx`, `reference.pptx`, or `reference.xlsx`;
- each original `preview.png`;
- source `SKILL.md`, `artifact-template.json`, `agents/agent.yaml`, manifest, and library icon;
- individual byte length and SHA-256 values, plus the deterministic binary aggregate, in `skills/default-template-library/integrity.json`.

The import changes only the repository-level adapter surface needed here: the plugin manifest, attribution/license record, integrity record, path policy, safety gates, and public-package exclusion. It does not regenerate, normalize, re-compress, or rewrite any retained Office or preview asset.

## Delivery boundary

The library is intentionally excluded from the npm tarball. It is available to Agents working from this repository, where a named template Skill uses the retained reference as a read-only starting point and writes a distinct output artifact. `scripts/materialize-template.mjs` first verifies the retained SHA-256, then atomically creates a byte-identical output plus a provenance audit and refuses to overwrite either destination. No source-free substitute or second template-generator fallback remains in this plugin.

The source files may contain rich or source-bound Office topology. A requested operation must use the matching Documents, Presentations, or Spreadsheets workflow and preserve the source boundary; when the public model cannot safely import or modify a graph, it must explain the limitation and fail closed rather than flattening the template or silently replacing its layout.

## Verification

`test/default-template-library.mjs` checks the canonical file inventory, secure relative paths, no symbolic links, JSON/YAML metadata, PNG structure, Office ZIP signatures, size budgets, every asset's SHA-256, and the source aggregate. It materializes all 20 templates, verifies overwrite refusal, runs every source through the public facade's import / unchanged export / second-import path, and, when LibreOffice plus Poppler are present, renders both source and processed files to non-empty native rasters. All seven PPTX templates additionally replace one visible SlidePart-placeholder title through `TextFrame.set()`, reimport it, prove placeholder identity/geometry plus paragraph/run formatting are unchanged, and render the edited output; a deliberate newline-topology change fails closed. It can additionally compare bytes against an explicitly supplied `OFFICE_TEMPLATE_SOURCE_ROOT` checkout. `DefaultTemplateLibraryCodecTests` adds native OpenChestnut no-op and bounded-edit coverage: slide-name metadata and recognized owner-local placeholder text for each PPTX, `updateFields` for each DOCX, ordinary string-cell edits for each XLSX, and source-bound rejection for the Financial Budget partial shared-formula range. `test/package-contents.mjs` and the clean-install package probe verify that no default-template-library file enters the npm tarball.
