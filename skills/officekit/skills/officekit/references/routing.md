# OfficeKit routing

Use this reference only when a request spans formats, has multiple outputs,
targets a live Excel session, or does not have an obvious owner.

## Installed Skill lookup

Agent Skill hosts do not define a portable nested-Skill call. Load the selected
Skill instructions from the installed filesystem instead.

Check these locations in order:

1. Flat installed siblings:
   - `../documents/SKILL.md`
   - `../spreadsheets/SKILL.md`
   - `../excel-live-control/SKILL.md`
   - `../presentations/SKILL.md`
   - `../pdf/SKILL.md`
   - `../template-creator/SKILL.md`
2. Repository plugin layout:
   - `../../../documents/skills/documents/SKILL.md`
   - `../../../spreadsheets/skills/spreadsheets/SKILL.md`
   - `../../../spreadsheets/skills/excel-live-control/SKILL.md`
   - `../../../presentations/skills/presentations/SKILL.md`
   - `../../../pdf/skills/pdf/SKILL.md`
   - `../../../template-creator/skills/template-creator/SKILL.md`

Resolve paths from the OfficeKit Skill directory, not the process working
directory. If the selected Skill is absent, stop with a broken-install message
that names the missing Skill. Do not call raw package APIs as a replacement.

## Ownership table

| Intent or output | Owner | Notes |
| --- | --- | --- |
| Read, create, or edit DOCX | Documents | It also owns Google Docs handoff guidance. |
| Read, create, or edit XLSX/CSV/TSV | Spreadsheets | Use file-based workflows. |
| Operate an already-open Excel workbook | Excel Live Control | Never substitute offline XLSX editing. |
| Read, create, or edit PPTX | Presentations | It also owns Google Slides handoff guidance. |
| Read, create, edit, sign, redact, repair, OCR, or verify PDF | PDF | Provider routing stays inside this Skill. |
| Register a reusable DOCX/XLSX/PPTX reference | Template Creator | This creates template metadata and retained assets, not the final business artifact. |

One output has one owner. A later owner may consume a completed artifact from
an earlier step, but it must not reopen that earlier file for unrelated edits.

## Route construction

Represent a cross-format request as a dependency-ordered list:

```text
step id -> owner -> inputs -> output -> required QA -> downstream consumers
```

Examples:

- Analyze an XLSX, create an executive deck, and deliver a PDF:
  `Spreadsheets -> Presentations -> PDF`
- Draft a DOCX report and attach supporting XLSX:
  `Spreadsheets -> Documents`, with separate final ownership for both files.
- Review a PDF and summarize findings in DOCX:
  `PDF -> Documents`
- Update an open Excel model and export a board deck:
  `Excel Live Control -> Presentations`

Pass only the evidence required by the next step: verified values, table
matrices, chart data, extracted text, images, citations, and provenance. Keep
the original source available for independent verification.

## Route boundaries

- A read-only request still loads the format owner, but it does not trigger
  template selection.
- Existing-file edits use the existing file as their design reference.
- A conversion does not grant the destination owner permission to alter the
  source.
- PDF provider installation, credentials, signing policy, rewrite policy, OCR,
  sanitization, and conformance remain governed by the PDF Skill.
- Google Docs, Sheets, and Slides transfer behavior remains governed by the
  corresponding domain Skill.
- If a requested operation exceeds an owner's published capability, expose the
  limit and ask for a changed outcome. Do not switch to an unselected tool.
