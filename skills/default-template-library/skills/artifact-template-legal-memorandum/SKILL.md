---
name: artifact-template-legal-memorandum
description: "Create a document using the Legal Memorandum template and its retained reference file. Use when the user selects or names Legal Memorandum. Draft legal memoranda with the issue, brief answer, relevant facts, analysis, and conclusion."
---

# Legal Memorandum

Create a new document from this template. Keep the reference file unchanged.

## Workflow

1. Read `artifact-template.json` and resolve its paths relative to this skill directory.
2. Use the matching Documents workflow with the retained reference file. If that workflow is unavailable, say so and stop; do not recreate or install a replacement.
3. Treat the user's prompt and available sources as the content input. Do not invent facts merely to fill a template slot.
4. Clone or import the reference instead of replacing its visual system with generic defaults.
5. Render and verify the finished document, then return the final artifact.

## Fidelity

Preserve page setup, sections, styles, lists, tables, headers, footers, and recurring page elements.

User instructions control requested content and explicit deviations. The retained reference controls layout and formatting where the user has not requested a change.
