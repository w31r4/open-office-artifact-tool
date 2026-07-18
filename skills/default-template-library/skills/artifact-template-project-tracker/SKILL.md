---
name: artifact-template-project-tracker
description: Create a source-free Project Tracker XLSX with a summary, work plan, risk and decision register, formula checks, owners, dates, and completion evidence. Use when the user needs an auditable delivery tracker and does not supply a reference workbook to preserve.
---

# Project Tracker

Generate a new workbook from the project-authored source-free template:

```sh
node ../../scripts/generate-template.mjs \
  --template-id artifact-template-project-tracker \
  --output /absolute/path/project-tracker.xlsx \
  --audit /absolute/path/project-tracker.audit.json
```

Use the Spreadsheets workflow for bounded input, status, formula, formatting, and worksheet edits. Preserve the Summary, Work Plan, and Risks and Decisions sheets unless the user explicitly changes the tracking model.

Before delivery, recalculate, inspect formula and status records, verify the workbook, export to a distinct final path, reimport, render the relevant sheets, and report the audit location.
