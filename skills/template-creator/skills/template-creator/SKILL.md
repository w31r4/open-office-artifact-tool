---
name: template-creator
description: Create or update a reusable local Office artifact template from a Word document, PowerPoint presentation, or Excel workbook. Use when the user asks to make a reusable template from a `.docx`, `.pptx`, or `.xlsx` reference, or explicitly asks to update an existing artifact-template skill. Do not use for one-off artifact creation from an existing template.
---

# Template Creator

Create or update a reference-backed local template. The source Office file stays in the template so later work can clone or import it faithfully.

## Routing

- Manage only direct-child template skills below `${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills`.
- Create a new template by default. Use a numbered name instead of overwriting an existing template.
- Update only when the user explicitly identifies exactly one existing `artifact-template-*` skill.
- Keep template creation local. Do not fetch remote templates or modify installed caches.

## Create workflow

1. Require exactly one `.docx`, `.pptx`, or `.xlsx` reference unless the user explicitly requests a batch. For a batch, complete this workflow separately for every file.
2. Infer a concise display name, intended-use description, and artifact kind from the reference and request. If there is enough evidence, also prepare compact selection metadata: intended uses, avoid cases, audiences, content shapes, visual traits, visual commitment, and provenance.
3. Create `preview.png` before packaging:
   - DOCX: render the reference and use a representative page PNG.
   - PPTX: render the reference and use a representative slide PNG.
   - XLSX: render the used range of the first visible non-empty sheet.
4. Inspect the PNG. Stop if it is blank, clipped, corrupted, or not representative of the reference.
5. Set `SKILL_DIR` to this skill directory and pass shell-escaped values directly to the creator:

```bash
node "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --reference-path "/absolute/path/reference.docx" \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners."
```

Pass selection metadata as one shell-escaped JSON value when it is known:

```bash
node "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --reference-path "/absolute/path/reference.pptx" \
  --preview-path "/absolute/path/preview.png" \
  --display-name "Quarterly Review" \
  --description "Review quarterly performance, decisions, risks, and outlook." \
  --selection-json '{"useWhen":["quarterly business review"],"avoidWhen":["project kickoff"],"audiences":["executive"],"contentShapes":["KPIs","decisions","risks"],"visualTraits":{"tone":["formal"],"density":"medium","colorMode":"light","structure":["sectioned"]},"visualCommitment":"neutral","editProfile":{"level":"copy-only","verifiedOperations":[]},"provenance":{"license":"user-provided","source":"local-user-reference"}}'
```

Do not claim a verified edit operation from visual inspection. Keep
`editProfile.level` as `copy-only` until a real import/edit/export/reimport test
proves a narrower or broader profile. If selection metadata is omitted, the
creator safely defaults to the intended-use description, an opinionated visual
commitment, and `copy-only`.

6. Read the JSON result. Verify that the generated directory contains `SKILL.md`, schema-v2 `artifact-template.json`, `agents/agent.yaml`, the retained `assets/reference.<ext>`, and `assets/preview.png`. Verify the recorded reference and preview hashes.

## Update workflow

1. Resolve the exact passed template and read its `SKILL.md`, `artifact-template.json`, `agents/agent.yaml`, retained reference, and preview. Stop if it is not a direct child of the local skills directory or if more than one target was passed.
2. Preserve the template folder name and every file or behavior the user did not ask to change.
3. For reference or visual changes, edit a temporary copy of the retained reference using the matching Office artifact workflow, render a new preview, and inspect it. For display-name or intended-use changes, retain the existing reference and preview unless they also change.
4. Pass every current or changed required value to the creator explicitly. Existing schema-v2 selection metadata is preserved when `--selection-json` is omitted; pass a complete replacement value when that metadata must change:

```bash
node "$SKILL_DIR/scripts/create-template-skill.mjs" \
  --mode "update" \
  --skill-name "artifact-template-standup" \
  --reference-path "/absolute/path/updated-reference.docx" \
  --preview-path "/absolute/path/updated-preview.png" \
  --display-name "Standup" \
  --description "Run a structured daily standup with updates, blockers, and owners."
```

5. The script accepts schema-v1 templates for migration, validates the existing template kind, preserves additional template-owned files, emits schema v2, and replaces the template atomically without changing its skill name.
6. Verify every requested change and confirm that no staging or backup directories remain.

## Response

Report the created or updated template's display name, artifact kind, and local path. State that the reference and preview remain with the template, and briefly describe how to invoke the returned template skill in the active agent environment. Do not emit product-specific cards, links, or sharing directives.

## Constraints

- Do not create an intermediary request file; pass creator inputs through command-line flags, including optional selection JSON.
- Do not delete or sanitize the retained reference; fidelity depends on retaining it verbatim.
- Do not change the artifact kind during an update.
- Do not mark a template `bounded-edit` or `composable` without repeatable capability evidence.
- Do not modify global skill metadata or protocol files.
