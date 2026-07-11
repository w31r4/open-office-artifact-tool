# Reference materials

This directory contains project-internal reference material for building `open-office-artifact-tool`.

## `office-artifact-tool` submodule

`reference/office-artifact-tool` is a Git submodule pointing at the authorized, non-open-source `office-artifact-tool` reference package:

- Remote: `https://github.com/w31r4/office-artifact-tool.git`
- Purpose: behavior/API/workflow reference for creating a publishable open-source clean-room replacement.

Use this submodule to observe the reference package's public package shape, exported API surface, smoke tests, examples, and observable behavior.

Do **not** vendor reference implementation internals into `open-office-artifact-tool`. Implement behavior independently using public standards, public libraries, OOXML/PDF specs, OpenXML SDK, Microsoft Office native automation, Playwright, LibreOffice, Poppler, PDF.js, sharp/canvas, and other legally usable technologies.

## Relationship to `handoff/reference-skills`

The reference skills under `handoff/2026-07-11/reference-skills/` capture the agent workflows that should eventually operate `open-office-artifact-tool` itself for Documents, Spreadsheets, Presentations, and PDF editing/rendering/verification.
