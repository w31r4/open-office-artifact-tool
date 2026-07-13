---
name: open-office-documents
description: Create, edit, inspect, render, and verify DOCX artifacts with open-office-artifact-tool and real LibreOffice/Poppler page QA.
---

# Open Office Documents

Use this project skill for standalone `.docx` artifact work. It is the clean-room Documents workflow for `open-office-artifact-tool`; it uses only the package's public facade and legally usable renderers.

## Contract

- Never import or copy the reference package's runtime artifact, runtime module, runtime bindings, or implementation details.
- Preserve an imported document's content, styles, structure, and review state unless the user requests a redesign.
- For new documents, choose one coherent design preset before authoring. The current public facade ships `report` and `memo`; broader exact preset fidelity remains tracked in `docs/coverage.md`.
- Use real list items, tables, comments, hyperlinks, fields, tagged bibliography sources/native `CITATION` fields, images, sections, and tracked changes rather than visual text imitations.
- Preserve direct and theme-backed run formatting. Theme fonts/colors and paired complex-script properties must survive a metadata-free native import before delivery.
- Keep semantic and package inspection bounded. Store QA evidence in a temporary/output directory.
- Do not deliver a DOCX until semantic verification and page-image review pass.

## Authoring workflow

1. Create a `DocumentModel` or import an existing DOCX with `DocumentFile.importDocx`. After package-level OOXML patches, pass `{ preferNative: true }` so relationship-driven native parts take precedence over stale embedded model metadata.
2. Inspect the relevant blocks, styles, comments, and layout before editing.
3. Apply focused changes through public APIs.
   For clean-room package surgery, `DocumentFile.patchDocx(...)` can create a Comments part and add matching block, paragraph, or table-cell anchors with `recipe: { kind: "comments", source: "word/document.xml", sourceReference: { anchors: [...] } }`. It can also relocate valid `commentsExtended`, `commentsIds`, `commentsExtensible`, and `people` parts with matching recipe kinds; semantic validation rejects orphan/multiple parts, unresolved or duplicate paragraph/durable identities, invalid UTC metadata, reply placeholders, and conflicting people.
   The same package API can create a Numbering part and assign declared `numId`/level pairs to target paragraphs with `recipe: { kind: "numbering", source: "word/document.xml", sourceReference: { assignments: [...] } }`.
   A Settings recipe can safely mutate an arbitrary relationship-backed Settings part with `sourceReference: { trackRevisions, updateFields, evenAndOddHeaders, mirrorMargins, documentProtection }`. Protection modes are passwordless `readOnly`, `comments`, `trackedChanges`, or `forms`; they discourage accidental editing but do not encrypt the DOCX.
4. Run `document.verify({ visualQa: true })` and fix every material issue.
5. Export DOCX and import the exported file again.
6. Render the real DOCX through LibreOffice to PDF, then Poppler to page PNGs.
7. Inspect every page image at full size. Comments also require structural inspection because headless renderers may omit them.
8. Once a model/native baseline is approved, compare the model preview and every native page on later runs; a page-count change is a visual regression until reviewed.

```js
import { DocumentFile, DocumentModel } from "open-office-artifact-tool";

const document = DocumentModel.create({
  name: "Decision brief",
  blocks: [],
  theme: {
    name: "Decision Theme",
    colors: { accent1: "#336699" },
    fonts: { major: "Source Serif 4", minor: "Aptos", majorEastAsia: "Noto Serif CJK SC", majorComplexScript: "Noto Naskh Arabic" },
  },
  defaultRunStyle: { fontFamily: "Aptos", fontSize: 22 },
});
document.applyDesignPreset("report");
document.styles.add("DecisionEmphasisBase", { type: "character", italic: true, themeColor: "accent1" });
document.styles.add("DecisionEmphasis", { type: "character", basedOn: "DecisionEmphasisBase", bold: true });
document.setSettings({ updateFields: true, documentProtection: { edit: "comments" } });
document.addParagraph("Decision brief", { styleId: "Title", name: "title" });
document.addParagraph("Recommendation", { styleId: "Heading1", name: "recommendation-heading" });
document.addParagraph("", { styleId: "Normal", runs: [
  { text: "Theme-backed decision", style: { runStyleId: "DecisionEmphasis", fontTheme: "majorHAnsi", themeTint: "80" } },
  { text: " العربية", style: { fontThemeComplexScript: "majorBidi", boldComplexScript: true, italicComplexScript: true, fontSizeComplexScript: 32 } },
] });
document.addParagraph("Proceed with the clean-room implementation.", { styleId: "Normal" });
document.addListItem("Validate the exported DOCX", { listType: "number", numberFormat: "upperLetter", start: 1, levelText: "%1)" });
document.setSectionSettings(0, { differentFirstPage: true });
document.addHeader("Decision brief", { referenceType: "first", sectionIndex: 0 });
document.addFooter("Confidential", { referenceType: "even" });
const decisionTable = document.addTable({
  name: "decision-table",
  styleId: "TableGrid",
  values: [["Area", "Status"], ["Semantic QA", "Pass"], ["Native render", "Required"]],
});
const review = document.addComment(decisionTable, "Verify the evidence table.", {
  author: "QA Agent",
  initials: "QA",
  date: "2026-07-11T00:00:00.000Z",
  resolved: true,
});
document.replyToComment(review, "Verified after native render.", { author: "Maintainer", initials: "MT" });

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
- `document.inspect(...)` proves agent-facing theme, settings, blocks, inspectable bookmark ranges, internal/external hyperlinks, multi-level list formats/start/level text and picture markers, bibliography sources plus native citation tags/results, styles, classic comment anchors, threaded/resolved state, durable identity, UTC metadata, people presence identity, and default/first/even header/footer references survived roundtrip. Each comment maps through the last classic-comment paragraph's `w14:paraId`; replies use `w15:paraIdParent`, Office 2019 `commentsIds` maps `paraId` to `durableId`, Office 2021 `commentsExtensible` carries UTC/follow-up metadata, and `people` binds author names to provider/user IDs. Native import follows `document.xml.rels` instead of assuming fixed theme/settings/styles/numbering/comment/header/footer/customXml filenames, resolves abstract numbering plus overrides, restores bookmark start/end targets, `w:anchor` links, external relationships, tooltip/history state, fields, and prefix-agnostic `b:Sources` entries keyed to `CITATION` fields. Treat a header/footer reference as a declaration, not proof that its variant is active: `differentFirstPage` controls first-page selection, `settings.evenAndOddHeaders` controls even-page selection, and a missing section reference inherits the previous section's same-type reference. `document.layoutJson()` must identify the effective reference type, source section, inherited state, and blank state for every modeled page. Omit `sectionIndex` only when intentionally targeting the final section.
- The checked-in business brief requires both a multi-block `RecommendationSection` bookmark and a row-major cross-cell `ReadinessEvidence` bookmark from `table.getCell(1, 0)` through `table.getCell(3, 2)`, each with a relationship-free internal jump through model, package, metadata-free import, second export, and real render gates. Its first section declares the running header/footer while its second section omits both references; native-preferred layout and rendered page text must prove that Word's same-type inheritance supplies them on page 2.
- The checked-in `package-comments.json` fixture creates an arbitrary-path Comments part through the public patch API, anchors one comment to a paragraph block and one to a table cell, then verifies both through native-preferred import and the real render gate.
- The checked-in `package-numbering.json` fixture creates an arbitrary-path Numbering part, binds two ordinary paragraphs to declared multilevel numbering definitions, and verifies their format/start/level metadata through native-preferred import and the real render gate.
- `document.addListItem(..., { pictureBullet })` accepts an embedded PNG/JPEG/GIF data URL or a non-fetched absolute URI with bounded point dimensions and alt text. Require `w:numPicBullet` before `w:abstractNum`, a matching `w:lvlPicBulletId`, and an image relationship owned by the Numbering part rather than `document.xml`; when relocating a Numbering part, relocate its `.rels` and adjust relative media targets too. Export uses the ISO-documented VML picture-symbol form for Word/LibreOffice compatibility, while import accepts VML and modern DrawingML forms. The business brief exercises embedded import, second export, SVG evidence, and a visibly colored LibreOffice/Poppler marker.
- The checked-in `package-settings.json` fixture creates an arbitrary-path Settings part, preserves unrelated compatibility markup, enables revision/field/header/margin settings, applies comments-only editing restrictions, and verifies the agent-facing state through native-preferred import and the real render gate.
- `document.verify({ visualQa: true })` checks structural and modeled layout issues.
- Model SVG/Playwright preview catches facade-level layout regressions.
- LibreOffice PDF plus Poppler page PNGs are the native render gate on non-Windows hosts.
- PNG baselines compare the modeled preview and every native page through `visualQaArtifact(..., { pixelDiff: true })`; baseline page-count changes fail QA. Supplying `--baseline-dir` is fail-closed: initialize it with `--write-baseline true`, because missing, empty, or non-contiguously numbered page sets are rejected.
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
