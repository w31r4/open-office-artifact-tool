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

Before it acquires a write lock or retains any bytes, the creator checks the
reference through the same bounded Office package inspector used by the public
facades. The extension must match a DOCX/PPTX/XLSX OPC package with its required
primary part, declared main content type, and exactly one root
`officeDocument` relationship; ZIP entry CRCs are also verified. A renamed text
file, a cross-family Office package, a broken root relationship, or a corrupt
archive fails closed and creates no template tree.

The generated `artifact-template.json` uses schema version 2. It records
selection metadata plus SHA-256 values for the retained Office file and PNG.
Without explicit selection metadata, new templates are intentionally
`copy-only` and visually `opinionated`; an Agent may recommend them but must not
claim that their content is safely editable.

Optional `--selection-json` accepts the complete selection profile in one JSON
value. Verified edit operations must come from a real
import/edit/export/reimport test, not visual inspection.

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

The creator validates artifact kind consistency, stages changes beside the
final directory, and replaces an updated template atomically with rollback if
placement fails. A per-home write lock prevents concurrent template writes.
Schema-v1 templates are migrated on update. Existing schema-v2 selection
metadata is preserved unless a complete `--selection-json` replacement is
provided.

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

`artifact-template.json` records the supported kind, paths, selection evidence,
edit profile, provenance, and retained-asset hashes. The creator validates PNG
chunk structure and CRCs before copying the preview, and applies the same
fail-closed Office-package admission check on every create or update.
