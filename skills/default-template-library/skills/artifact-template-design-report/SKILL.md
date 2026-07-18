---
name: artifact-template-design-report
description: Create a source-free Design Report DOCX with a decision summary, success criteria, approach boundary, risks, rollout, and verification plan. Use when the user needs a technical or product design record and does not supply a reference file to preserve.
---

# Design Report

Generate a new document from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-design-report \
  --output /absolute/path/design-report.docx \
  --audit /absolute/path/design-report.audit.json
```

Use the Documents workflow for requested edits. Keep the decision boundary visible: executive summary, success criteria, proposed approach, risks, rollout, verification, and open decisions. Do not present a generic rewrite as evidence that an external reference design was preserved.

Before delivery, import the generated or edited DOCX, inspect named blocks, verify it, export to a distinct final path, reimport, render it, and report the audit location.
