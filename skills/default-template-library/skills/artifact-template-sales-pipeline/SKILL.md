---
name: artifact-template-sales-pipeline
description: "Create a spreadsheet using the Sales Pipeline template and its retained reference file. Use when the user selects or names Sales Pipeline. Track opportunities, stages, owners, deal sizes, probabilities, forecasts, next steps, and risks."
---

# Sales Pipeline

Create a new spreadsheet from this template. Keep the reference file unchanged.

## Workflow

1. Read `artifact-template.json` and resolve its paths relative to this skill directory.
2. Use the matching Spreadsheets workflow with the retained reference file. If that workflow is unavailable, say so and stop; do not recreate or install a replacement.
3. Treat the user's prompt and available sources as the content input. Do not invent facts merely to fill a template slot.
4. Clone or import the reference instead of replacing its visual system with generic defaults.
5. Render and verify the finished spreadsheet, then return the final artifact.

## Fidelity

Preserve sheet structure, formulas, names, number formats, dimensions, tables, charts, validation, conditional formatting, and frozen panes.

User instructions control requested content and explicit deviations. The retained reference controls layout and formatting where the user has not requested a change.
