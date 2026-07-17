# Reference materials

This directory contains project-internal reference material for building `open-office-artifact-tool`.

## `office-artifact-tool` submodule

`reference/office-artifact-tool` is a Git submodule pointing at the public MIT-licensed `office-artifact-tool` reference package:

- Remote: `https://github.com/w31r4/office-artifact-tool.git`
- Purpose: behavior/API/workflow reference for creating a publishable open-source clean-room replacement.

Use this submodule to observe the reference package's public package shape, exported API surface, smoke tests, examples, and observable behavior.

The currently pinned reference revision is
`207ce094a55d82a37efdca42a1c5e9656f696962`, the remotely reachable
`origin/feat/sync-grid-layout-library-template-creator` commit. Its package
manifest remains `office-artifact-tool@2.8.24` and it is based on the prior
`2d0e249ea6b62f55cca22a343b832a38e8f7537c` runtime-sync revision.

This revision adds the neutral `grid-layout-library` naming and Template Creator
workflow. Both are adapted into this project's public Skills with independent
package, security, and workflow tests. Recording the exact remotely obtainable
commit preserves reproducible submodule checkout without importing the reference
runtime.

Do **not** vendor the reference package's runtime artifact, runtime module, runtime bindings, or implementation details into `open-office-artifact-tool`. Implement behavior independently using public standards, public libraries, OOXML/PDF specs, OpenXML SDK, Microsoft Office native automation, Playwright, LibreOffice, Poppler, PDF.js, sharp/canvas, and other legally usable technologies.

## Relationship to `handoff/reference-skills`

The reference skills under `handoff/2026-07-11/reference-skills/` capture the agent workflows that should eventually operate `open-office-artifact-tool` itself for Documents, Spreadsheets, Presentations, and PDF editing/rendering/verification.
