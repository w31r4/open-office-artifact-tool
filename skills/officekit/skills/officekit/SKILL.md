---
name: officekit
description: Use this skill to plan and coordinate broad, ambiguous, cross-format, or multi-deliverable Office work across Word/DOCX, Excel/XLSX/CSV, PowerPoint/PPTX, and PDF, including deciding whether to use zero or one available template. Load only the required installed domain Skills and preserve their own edit and QA rules. Do not use when the user explicitly invokes a single domain or template Skill for a self-contained task.
---

# OfficeKit

Turn one Office request into a small, explicit artifact workflow. Route each
output to its owning Skill, load only the instructions needed for that route,
and preserve the owning Skill's safety and QA rules.

## Respect explicit choices

- If the user explicitly invokes a domain, business, or template Skill for a
  self-contained task, use it directly.
- If an existing Office file is itself the output being edited, treat that file
  as the reference for that output. Do not search for a decorative template for
  it; decide independently for any other new outputs.
- Do not replace an unavailable Skill, template, provider, or live application
  with a different execution path. Report the missing component.

## Build the artifact route

1. Inventory every input file and requested output.
2. Identify whether the task is read-only, creates a new artifact, edits an
   existing artifact, or converts between formats.
3. Assign exactly one owning Skill to each output:
   - DOCX or Google Docs handoff: Documents
   - XLSX, CSV, TSV, or Google Sheets handoff: Spreadsheets
   - an already-open Excel workbook: Excel Live Control
   - PPTX or Google Slides handoff: Presentations
   - PDF: PDF
4. For multiple outputs, order the owners as a dependency graph. Pass facts,
   tables, images, and structured content between steps; never let two Skills
   mutate the same file.
5. Read [routing.md](references/routing.md) for cross-format work, live Excel,
   missing Skills, or an ambiguous owner.

## Load only the selected Skill

Load the installed `SKILL.md` for each chosen owner before using its package
APIs or scripts. Follow that Skill's import, edit, source-preservation, render,
verify, and publication rules without weakening them.

When OfficeKit is active, carry the routed task through the selected owners.
Do not send the user away to repeat the request through another Skill.

Do not preload every Office Skill. Do not copy domain API documentation into
this coordination layer.

## Decide whether a template helps

Make a template decision only for a new or substantially redesigned DOCX,
XLSX, or PPTX.

1. Prefer a user-provided reference.
2. Honor an explicitly named template.
3. Otherwise read [template-selection.md](references/template-selection.md).
4. Query available metadata with
   `scripts/query-templates.mjs`; do not inspect every template file.
5. Choose exactly one of `selected`, `ask`, or `none`.
6. Load previews only for the final one to three candidates.
7. Before selecting a template, load the owning domain Skill and confirm that
   the requested edits fit both the template's verified edit profile and the
   domain Skill's source-bound capabilities.

`none` means the owning Skill should compose the artifact from first
principles. It is a successful design decision, not an error or fallback.

Do not use the Office template catalog for a PDF-only task. When PDF is the
final form of an Office artifact, apply the template to the Office source step,
then let the PDF Skill inspect and verify the final PDF.

## Execute and verify

- Protect every input and retained template from overwrite.
- Complete each artifact under its owner's workflow.
- Reopen or reimport the result when the owner requires it.
- Run the owner's semantic and visual QA.
- For a multi-artifact task, verify shared facts, numbers, names, dates, and
  visual identity across outputs.
- Return the final files, the route used, the template decision, and any
  explicit capability limits.

Do not describe an unverified output as complete.
