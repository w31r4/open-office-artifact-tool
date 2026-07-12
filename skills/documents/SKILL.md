---
name: open-office-documents
description: Create, edit, inspect, render, and verify DOCX artifacts with open-office-artifact-tool and real LibreOffice/Poppler page QA.
---

# Open Office Documents

Use this project skill for standalone `.docx` artifact work. It is the clean-room Documents workflow for `open-office-artifact-tool`; it uses only the package's public facade and legally usable renderers.

## Contract

- Never import or copy the reference package, runtime artifact, runtime module, or runtime bindings.
- Preserve an imported document's content, styles, structure, and review state unless the user requests a redesign.
- For new documents, choose one coherent design preset before authoring. The current public facade ships `report` and `memo`; broader exact preset fidelity remains tracked in `docs/coverage.md`.
- Use real list items, tables, comments, hyperlinks, fields, citations, images, sections, and tracked changes rather than visual text imitations.
- Keep semantic and package inspection bounded. Store QA evidence in a temporary/output directory.
- Do not deliver a DOCX until semantic verification and page-image review pass.

## Authoring workflow

1. Create a `DocumentModel` or import an existing DOCX with `DocumentFile.importDocx`. After package-level OOXML patches, pass `{ preferNative: true }` so relationship-driven native parts take precedence over stale embedded model metadata.
2. Inspect the relevant blocks, styles, comments, and layout before editing.
3. Apply focused changes through public APIs.
4. Run `document.verify({ visualQa: true })` and fix every material issue.
5. Export DOCX and import the exported file again.
6. Render the real DOCX through LibreOffice to PDF, then Poppler to page PNGs.
7. Inspect every page image at full size. Comments also require structural inspection because headless renderers may omit them.
8. Once a model/native baseline is approved, compare the model preview and every native page on later runs; a page-count change is a visual regression until reviewed.

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({ name: "Decision brief", blocks: [] });
document.applyDesignPreset("report");
document.addParagraph("Decision brief", { styleId: "Title", name: "title" });
document.addParagraph("Recommendation", { styleId: "Heading1", name: "recommendation-heading" });
document.addParagraph("Proceed with the clean-room implementation.", { styleId: "Normal" });
document.addListItem("Validate the exported DOCX", { listType: "bullet" });
document.addHeader("Decision brief", { referenceType: "first", sectionIndex: 0 });
document.addFooter("Confidential", { referenceType: "even" });
document.addTable({
  name: "decision-table",
  styleId: "TableGrid",
  values: [["Area", "Status"], ["Semantic QA", "Pass"], ["Native render", "Required"]],
});

const output = await DocumentFile.exportDocx(document);
await output.save("decision-brief.docx");
```

## Verification commands

Verify any DOCX and write semantic/package/layout/model-preview/native-page evidence:

```sh
node skills/documents/scripts/verify-document.mjs \
  --input decision-brief.docx \
  --output-dir tmp/document-qa \
  --preview-format png \
  --native-render required
```

`--native-render auto` runs LibreOffice + Poppler when available and records a skip otherwise. Use `required` for final local delivery when those runtimes are installed, and `off` only for an explicitly documented structural-only check.

Run the checked-in fixture end to end:

```sh
node skills/documents/scripts/run-fixture.mjs \
  --fixture skills/documents/fixtures/business-brief.json \
  --output-dir tmp/document-skill-fixture \
  --native-render required
```

Create an approved model and native-page baseline:

```sh
node skills/documents/scripts/verify-document.mjs \
  --input decision-brief.docx \
  --output-dir tmp/document-baseline-run \
  --preview-format png \
  --native-render required \
  --baseline-dir tmp/document-baselines \
  --write-baseline true
```

Compare a later export against that baseline:

```sh
node skills/documents/scripts/verify-document.mjs \
  --input decision-brief.docx \
  --output-dir tmp/document-compare \
  --prefer-native true \
  --preview-format png \
  --native-render required \
  --baseline-dir tmp/document-baselines
```

## QA gates

- `DocumentFile.inspectDocx(...)` proves required package parts and relationships exist, including namespace-aware source XML `r:id`/`r:embed`/`r:link` resolution through the corresponding `.rels` part.
- `document.inspect(...)` proves agent-facing blocks, styles, anchors, comments, and default/first/even header/footer references plus zero-based section bindings survived roundtrip. Omit `sectionIndex` to target the final section. Native import follows `document.xml.rels` targets rather than assuming `header1.xml` or `footer1.xml`.
- `document.verify({ visualQa: true })` checks structural and modeled layout issues.
- Model SVG/Playwright preview catches facade-level layout regressions.
- LibreOffice PDF plus Poppler page PNGs are the native render gate on non-Windows hosts.
- PNG baselines compare the modeled preview and every native page through `visualQaArtifact(..., { pixelDiff: true })`; baseline page-count changes fail QA.
- Changed pages write PNG diff heatmaps into the QA output directory; dimension mismatches require a non-strict alignment mode.
- Use `--diff-alignment center|top-left|strict`, `--diff-color '#ff1848'`, and `--diff-unchanged-color '#334155'` to make dimension changes and review palettes explicit.
- For same-size renders with a known platform jitter, opt in to bounded registration with `--registration-offset 2`; use `--registration-improvement 0.1` to require at least 10% sampled mismatch improvement. QA records the chosen baseline translation and ignored edge pixels.
- Microsoft Office native automation remains the higher-fidelity Windows gate for Word-specific behavior.
- Deliver only the requested DOCX; previews and QA reports are internal unless requested.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Generic DOCX verifier: `scripts/verify-document.mjs`
