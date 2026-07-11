# Handoff bundle index

## Main handoff

- `HANDOFF.md`

## Project goal

`open-office-artifact-tool` is intended to become a publishable open-source clean-room replacement for `office-artifact-tool`. The `reference-skills/` directories are part of the project target: they capture the agent workflows that should eventually operate the toolkit itself for full Documents / Spreadsheets / Presentations / PDF editing, rendering, inspection, and verification.

## Reference runtime package

- Project submodule path: `/Users/zfang/workspace/open-office-artifact-tool/reference/office-artifact-tool`
- Submodule remote: `https://github.com/w31r4/office-artifact-tool.git`
- Original local package path: `/Users/zfang/workspace/office-artifact-tool`
- Package: `office-artifact-tool@2.8.22`
- Main entry: `reference/office-artifact-tool/dist/artifact_tool.mjs`
- Manifest: `reference/office-artifact-tool/package.json`
- README: `reference/office-artifact-tool/README.md`

## Copied reference skills

The following directories were copied from `local reference source` into this project handoff bundle. Treat them as target workflow/spec material that should be clean-room adapted into runnable project skills, fixtures, and tests that call `open-office-artifact-tool` APIs.

### Documents

- Package directory: `reference-skills/documents/`
- Plugin manifest: `reference-skills/documents/.artifact-plugin/plugin.json`
- Main skill doc: `reference-skills/documents/skills/documents/SKILL.md`
- Important subdirectories:
  - `reference-skills/documents/skills/documents/tasks/`
  - `reference-skills/documents/skills/documents/ooxml/`
  - `reference-skills/documents/skills/documents/scripts/`
  - `reference-skills/documents/skills/documents/references/`
  - `reference-skills/documents/skills/documents/troubleshooting/`

### Spreadsheets

- Package directory: `reference-skills/spreadsheets/`
- Plugin manifest: `reference-skills/spreadsheets/.artifact-plugin/plugin.json`
- Main skill doc: `reference-skills/spreadsheets/skills/spreadsheets/SKILL.md`
- Additional live-control skill: `reference-skills/spreadsheets/skills/excel-live-control/SKILL.md`
- Important subdirectories/files:
  - `reference-skills/spreadsheets/skills/spreadsheets/API_QUICK_START.md`
  - `reference-skills/spreadsheets/skills/spreadsheets/charts.md`
  - `reference-skills/spreadsheets/skills/spreadsheets/domain_guidance/`
  - `reference-skills/spreadsheets/skills/spreadsheets/routing/`
  - `reference-skills/spreadsheets/skills/excel-live-control/`

### Presentations

- Package directory: `reference-skills/presentations/`
- Plugin manifest: `reference-skills/presentations/.artifact-plugin/plugin.json`
- Main skill doc: `reference-skills/presentations/skills/presentations/SKILL.md`
- Important subdirectories/files:
  - `reference-skills/presentations/skills/presentations/artifact_tool/API_QUICK_START.md`
  - `reference-skills/presentations/skills/presentations/container_tools/`
  - `reference-skills/presentations/skills/presentations/template_following_scripts/`
  - `reference-skills/presentations/skills/presentations/references/`
  - `reference-skills/presentations/skills/presentations/builtin_templates_support/`

### PDF

- Package directory: `reference-skills/pdf/`
- Plugin manifest: `reference-skills/pdf/.artifact-plugin/plugin.json`
- Main skill doc: `reference-skills/pdf/skills/pdf/SKILL.md`

## Clean-room note

Use these files to understand target behavior and acceptance gates. Do not copy reference implementation details into `/Users/zfang/workspace/open-office-artifact-tool`; implement with public standards and legally usable libraries/runtimes only.
