# Template Creator

Create or update reusable local Office templates from a reference `.docx`, `.pptx`, or `.xlsx` file and one representative PNG preview.

## Location and ownership

The creator writes user templates below:

```text
${OFFICE_ARTIFACT_HOME:-~/.office-artifact-tool}/skills/
```

Set `OFFICE_ARTIFACT_HOME` to choose another local root. Each template keeps a verbatim copy of its Office reference and its PNG preview; choose an appropriate local storage location before creating one.

## Create

Provide one supported Office reference, a valid PNG preview, a concise display name, and an intended-use description:

```sh
node skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs \
  --reference-path /absolute/path/reference.pptx \
  --preview-path /absolute/path/preview.png \
  --display-name "Quarterly business review" \
  --description "Create a quarterly business review from the saved deck layout."
```

The command returns JSON containing the created template name, artifact kind, and local path. It selects a numbered name rather than overwriting an existing template.

## Update

Updates require the exact template name and preserve other skill-owned files:

```sh
node skills/template-creator/skills/template-creator/scripts/create-template-skill.mjs \
  --mode update \
  --skill-name artifact-template-quarterly-business-review \
  --reference-path /absolute/path/updated-reference.pptx \
  --preview-path /absolute/path/updated-preview.png \
  --display-name "Quarterly business review" \
  --description "Create a quarterly business review from the updated deck layout."
```

The creator validates artifact kind consistency, stages changes beside the final directory, and replaces an updated template atomically with rollback if placement fails. A per-home write lock prevents concurrent template writes.

## Generated template layout

```text
$OFFICE_ARTIFACT_HOME/skills/artifact-template-<slug>/
├── SKILL.md
├── artifact-template.json
├── agents/agent.yaml
└── assets/
    ├── reference.docx | reference.pptx | reference.xlsx
    └── preview.png
```

`artifact-template.json` records the supported kind and paths to the retained reference and preview. The creator validates PNG chunk structure and CRCs before copying the preview.
