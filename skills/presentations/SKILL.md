---
name: open-office-presentations
description: Create, edit, inspect, render, baseline, and verify PPTX artifacts with open-office-artifact-tool and real LibreOffice/Poppler slide QA.
---

# Open Office Presentations

Use this project skill for standalone `.pptx` artifact work. It is the clean-room Presentations workflow for `open-office-artifact-tool`; it calls the package's public facade and legally usable renderers only.

## Contract

- Never import or copy the reference package's runtime artifact, runtime module, runtime bindings, or implementation details.
- Preserve an imported deck's content, theme, layouts, notes, comments, and object intent unless the user requests a redesign.
- Give each slide one narrative job and use audience-facing copy. Do not expose planning notes or production scaffolding on slides.
- Keep titles at least 50px on covers and 35px on content slides, mid-level text at least 24px, and body text at least 16px unless an inherited template requires otherwise.
- Use real tables, charts, images, connectors, notes, comments, themes, and layouts instead of flattening them into screenshots or visual imitations.
- Add connectors before nodes in the visual stack and route them through whitespace; never let a connector cross a label or unrelated node.
- Fix every unintended overlap, off-canvas element, clipping, and text overflow issue before delivery.
- Do not deliver a PPTX until semantic/package checks and full-size review of every rendered slide pass. A montage is an overview, not a substitute for per-slide review.

## Authoring workflow

1. Create a `Presentation` or import an existing PPTX with `PresentationFile.importPptx`.
2. Define the communication job and select a coherent theme/master/layout system before adding slides. Use `presentation.master` for the first-master compatibility path or `presentation.masters.add(...)` for multiple Slide Masters, then bind every layout with `masterId`; layout backgrounds/placeholders override only their owning master. Use `presentation.theme.setColors(...)`, `.setFonts(...)`, `.setTextStyles(...)`, and `.setColorMap(...)` for the shared native theme semantics; keep individual slide/shape styles for deliberate exceptions.
3. Inspect the relevant slides, text ranges, tables, charts, images, notes, and comments before editing.
4. Apply focused changes through public APIs and keep stable names on important objects.
   For bounded package surgery, `PresentationFile.patchPptx(...)` can attach caller-supplied image bytes or public chart XML to an existing slide with `recipe: { kind: "image"|"chart", source: "ppt/slides/slideN.xml", sourceReference: { objectId, name, alt, position } }`. Position is explicit pixels; the patcher owns DrawingML namespaces, relationship references, non-visual ID collision checks, deterministic replacement, and deletion cleanup.
5. Run `presentation.verify()` and `presentation.validateLayout()`; fix every material issue.
6. Export PPTX and import the exported file again.
7. Inspect native package parts with `PresentationFile.inspectPptx()`. Treat any notes/comments semantic issue as a delivery blocker: review relationships must originate from the correct slide/presentation part, roots/content types must match, and every legacy comment author/index must resolve through the singleton author registry.
8. Render every modeled slide with Playwright and render the original PPTX through LibreOffice to PDF plus Poppler PNGs.
9. Inspect every model and native slide image at full size. When a baseline is approved, compare PNG pixels on subsequent runs.

```js
import { Presentation, PresentationFile } from "open-office-artifact-tool";

const deck = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  master: {
    name: "Evidence Master",
    background: { fill: "bg1", mode: "reference", index: 1001 },
    placeholders: [{ type: "title", idx: 1, position: { left: 42, top: 36, width: 1196, height: 80 }, style: { fontSize: 42, bold: true, color: "accent1" } }],
  },
});
deck.theme.setColors({ accent1: "#3D8DFF", bg1: "#FFFFFF", tx1: "#000000" });
const slide = deck.slides.add({ name: "Evidence" });
const title = slide.shapes.add({
  name: "evidence-title",
  position: { left: 42, top: 36, width: 1196, height: 80 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Native PPTX parts preserve structure",
});
title.text.style = { fontFamily: "Arial", fontSize: 42, bold: true, color: "#000000" };
slide.tables.add({
  name: "evidence-table",
  position: { left: 42, top: 180, width: 1196, height: 360 },
  values: [["Gate", "Result"], ["Semantic", "Pass"], ["Visual", "Required"]],
  styleOptions: { headerRow: true },
});
slide.addNotes("Explain why package and visual evidence are complementary.");

await (await PresentationFile.exportPptx(deck)).save("evidence.pptx");
```

## Verification commands

Verify any PPTX and write inspect/package/layout/model/native evidence:

```sh
node skills/presentations/scripts/verify-presentation.mjs \
  --input evidence.pptx \
  --output-dir tmp/presentation-qa \
  --native-render required
```

Create an approved baseline:

```sh
node skills/presentations/scripts/verify-presentation.mjs \
  --input evidence.pptx \
  --output-dir tmp/presentation-baseline-run \
  --baseline-dir tmp/presentation-baselines \
  --write-baseline true \
  --native-render required
```

Compare a later render against it:

```sh
node skills/presentations/scripts/verify-presentation.mjs \
  --input evidence.pptx \
  --output-dir tmp/presentation-compare-run \
  --baseline-dir tmp/presentation-baselines \
  --native-render required
```

Run the checked-in fixture end to end:

```sh
node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/agent-readiness.json \
  --output-dir tmp/presentation-skill-fixture \
  --native-render required
```

`--native-render auto` runs LibreOffice + Poppler when available and records a skip otherwise. Use `required` for final local delivery when those runtimes are installed, and `off` only for an explicitly documented structural-only check.

## QA gates

- `PresentationFile.inspectPptx(...)` proves required slide, chart, media, notes, comments, comment-author registry, theme, master, and layout parts exist. It rejects wrong review-part roots/content types/sources, orphan or duplicate review relationships, multiple author registries, missing/duplicate author IDs, invalid or duplicate per-author comment indexes, and `lastIdx` values below the maximum used index. Native import must preserve per-comment/reply author identity even when notes, comments, or `commentAuthors.xml` use nonstandard relationship targets.
- `presentation.inspect(...)` proves agent-facing objects, master/layout identity, and review metadata survived roundtrip.
- Package evidence must include the presentation master list, master/layout parts, and the master↔layout plus slide→layout relationship chain when layouts are used.
- Theme evidence must include all 12 DrawingML color slots, major/minor Latin plus optional East-Asian/complex-script fonts, non-empty fill/line/effect/background format lists, a Slide Master `clrMap`, and title/body/other text styles. Re-import must restore the same agent-facing theme values.
- Master/layout evidence must restore native `p:bg` solid or scheme references and merge placeholders by `type` plus `idx`; slide backgrounds override layout backgrounds, which override the linked master. Inspect and render must report the same effective background.
- The checked-in `package-drawing.json` fixture generates a chart part through the public facade, attaches it and an arbitrary-path image to an existing slide through `patchPptx`, restores both as editable agent-facing objects, and passes the real render gate.
- The checked-in `package-notes-comments.json` fixture relocates notes, comments, and the singleton author registry to arbitrary valid paths, proves semantic validation reports zero issues, restores author identity and note text, and passes LibreOffice/Poppler rendering.
- `presentation.verify()` and `presentation.validateLayout()` catch structural, overlap, off-canvas, overflow, chart, table, image, placeholder, and dangling-comment issues.
- Per-slide Playwright PNGs catch facade/render regressions; the montage checks deck flow only.
- LibreOffice PDF plus Poppler slide PNGs are the native non-Windows render gate.
- Optional PNG baselines use `visualQaArtifact(..., { pixelDiff: true })`; approve baseline changes only after full-size review. Supplying `--baseline-dir` is fail-closed: initialize it with `--write-baseline true`, because missing, empty, or non-contiguously numbered model/native slide sets are rejected.
- Baseline approval replaces stale model/native slide files; later slide-count changes fail QA even if all remaining slides match.
- Changed slides write PNG diff heatmaps into the QA output directory; dimension mismatches require a non-strict alignment mode.
- Use `--diff-alignment center|top-left|strict`, `--diff-color '#ff1848'`, and `--diff-unchanged-color '#334155'` to make dimension changes and review palettes explicit.
- For same-size renders with a known platform jitter, opt in to bounded registration with `--registration-offset 2`; use `--registration-improvement 0.1` to require at least 10% sampled mismatch improvement. QA records the chosen baseline translation and ignored edge pixels.
- Microsoft Office native automation remains the higher-fidelity Windows gate for PowerPoint-specific behavior.
- Deliver only the requested PPTX; previews, baselines, and QA reports are internal unless requested.

## References

- Generated public API catalog: `../../docs/api.md`
- Current implementation coverage: `../../docs/coverage.md`
- Fixture runner: `scripts/run-fixture.mjs`
- Generic PPTX verifier: `scripts/verify-presentation.mjs`
