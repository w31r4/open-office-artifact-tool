---
name: artifact-template-project-kickoff
description: Create a source-free Project Kickoff PPTX with outcome, scope, delivery plan, owners, risks, and decision cadence. Use when the user asks to start or align a project and does not supply a reference deck to preserve.
---

# Project Kickoff

Generate a new presentation from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-project-kickoff \
  --output /absolute/path/project-kickoff.pptx \
  --audit /absolute/path/project-kickoff.audit.json
```

Use the Presentations workflow for bounded text and layout changes. Retain the three-slide narrative unless the user asks to change it: project outcome, scope and plan, then owners and decisions.

Before delivery, import the generated or edited PPTX, inspect named slides and shapes, verify it, export to a distinct final path, reimport, render all slides, and report the audit location. Do not claim that the presentation preserves an external reference: this template has no retained reference or preview asset.
