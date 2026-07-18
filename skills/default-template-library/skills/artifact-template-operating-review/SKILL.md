---
name: artifact-template-operating-review
description: Create a source-free Operating Review PPTX with an operating scorecard, delivery and risk lanes, decisions, and accountable owners. Use when the user needs a recurring operating review and does not supply a reference deck to preserve.
---

# Operating Review

Generate a new presentation from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-operating-review \
  --output /absolute/path/operating-review.pptx \
  --audit /absolute/path/operating-review.audit.json
```

Use the Presentations workflow for bounded text, numbers, and layout changes. Preserve the review narrative unless the request changes its purpose: scorecard, delivery and risks, then decisions and owners.

Before delivery, import the generated or edited PPTX, inspect the named slides and title shapes, verify it, export to a distinct final path, reimport, render all slides, and report the audit location.
