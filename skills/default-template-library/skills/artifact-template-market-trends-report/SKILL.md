---
name: artifact-template-market-trends-report
description: "Create a presentation using the Market Trends Report template and its retained reference file. Use when the user selects or names Market Trends Report. Communicate market or industry trends, supporting evidence, implications, and recommended responses."
---

# Market Trends Report

Create a new presentation from this template. Keep the reference file unchanged.

## Workflow

1. Read `artifact-template.json` and resolve its paths relative to this skill directory.
2. Use the matching Presentations workflow with the retained reference file. If that workflow is unavailable, say so and stop; do not recreate or install a replacement.
3. Treat the user's prompt and available sources as the content input. Do not invent facts merely to fill a template slot.
4. Clone or import the reference instead of replacing its visual system with generic defaults.
5. Render and verify the finished presentation, then return the final artifact.

## Fidelity

Preserve source slides, layouts, masters, typography, geometry, images, charts, tables, and recurring slide chrome.

User instructions control requested content and explicit deviations. The retained reference controls layout and formatting where the user has not requested a change.
