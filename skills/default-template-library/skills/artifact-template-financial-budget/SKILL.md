---
name: artifact-template-financial-budget
description: Create a source-free Financial Budget XLSX with assumptions, a monthly plan, variance, and formula-driven checks. Use when the user asks for a budget workbook and does not supply a reference workbook to preserve.
---

# Financial Budget

Generate a new workbook from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-financial-budget \
  --output /absolute/path/financial-budget.xlsx \
  --audit /absolute/path/financial-budget.audit.json
```

Use the Spreadsheets workflow for bounded input, formula, format, and worksheet edits. Preserve the separate Assumptions, Monthly Plan, and Budget Summary sheets unless the user explicitly changes the model design.

Before delivery, recalculate, inspect formula and check records, verify the workbook, export to a distinct final path, reimport, render the relevant sheets, and report the audit location. Do not claim that the workbook preserves an external reference: this template has no retained reference or preview asset.
