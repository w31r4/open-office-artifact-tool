# Spreadsheets

Spreadsheets is the file-type wrapper plugin for spreadsheet artifact workflows, Google Sheets-ready workbooks, and Microsoft Excel live control.

This installable Codex plugin is distributed with `open-office-artifact-tool`.

## Included Skills

- `Spreadsheets`: create, edit, analyze, visualize, render, and export spreadsheet files such as `.xlsx`, `.xls`, `.csv`, and `.tsv`, including Google Sheets-targeted workbooks that should be authored locally before import.
- `excel-live-control`: inspect, edit, verify, and control an open Microsoft Excel workbook through the ChatGPT add-in or a connected Excel session.

## Discoverability

Use this plugin for spreadsheet-oriented terms from the file-type naming model: sheet, sheets, Google Sheets, Excel, CSV, model, spreadsheet, spreadsheets, workbook, tracker, and `.xlsx`.

## Excel Live Control Boundary

The local `Spreadsheets` Skill runs against the public npm package. `excel-live-control` is also included so Codex can route live Excel requests correctly, but execution requires the host-provided connected-document app and an active Excel add-in session; those services are not implemented by the npm codec.

The core reference-style workbook example is tested end to end. Full parity with every API named in `API_QUICK_START.md` is still in progress and is tracked in `docs/reference-skills.md` in the source repository.

## Source

The plugin tree is versioned directly under `skills/spreadsheets` in the public repository.
