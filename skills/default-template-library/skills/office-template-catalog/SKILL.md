---
name: office-template-catalog
description: Route a request to a ready, source-free Office template from the bundled catalog. Use when the user asks which bundled template fits a document, presentation, or workbook task, or names a catalog entry. Do not use to import or recreate a third-party template file.
---

# Office Template Catalog

Read `../../catalog.json` before choosing a template. The catalog separates a useful template intent from a generated artifact that is actually available.

## Routing

1. Resolve one requested template ID or display name. If several fit, explain the small set of ready choices and ask the user to choose; do not silently select a visual design.
2. If its `status` is `ready`, run the matching template Skill. The generator must receive an explicit, distinct output path and should also write its audit JSON.
3. If its `status` is `planned`, say that no bundled source-free implementation exists yet. Do not substitute a nearby template and do not search installed caches or remote locations for an Office file.
4. If the user supplies their own DOCX, PPTX, or XLSX reference, route to Template Creator or the matching Office Skill. Keep the supplied reference local and distinguish it from this distributable catalog.

## Shared completion gate

For a generated artifact, inspect the generator audit, verify the matching model after import, make any requested bounded edit through the matching Office Skill, export to a distinct path, reimport, and render/review before delivery.

Every ready entry is project-authored and source-free. Its generator is the source of truth; there is no retained Office reference or preview asset to clone.
