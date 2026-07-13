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
- Use `importPptxWithOpenXmlWasm` / `exportPptxWithOpenXmlWasm` when a workflow explicitly requests the bundled C# codec. Its current editable boundary is top-level rectangle/ellipse shapes and bounded paragraphs: ordered text, `{ field: { id?, type, text } }`, and `{ break: true }` inlines; strictly increasing left/center/right/decimal `tabStops`; level and left/center/right/justify alignment; character/auto-number/explicit-no-bullet markers; embedded PNG/JPEG/GIF/safe-SVG or non-fetched http(s) picture bullets; direct marker font, six-digit RGB or one of the 16 DrawingML theme colors, fixed/percentage/follow-text marker size; bold/italic, font size/family, solid RGB text color; and click links for absolute external URIs, internal slide IDs, or `nextSlide`/`previousSlide`/`firstSlide`/`lastSlide`/`endShow`. Imported text may be edited only without changing paragraph count, inline count, or text/field/break kinds; an empty imported `tabStops` list explicitly deletes the native list. Transformed/unresolved picture bullets, transformed or non-sRGB/non-theme marker colors, custom-show/mouse-over links, body/list styles, and other native properties stay source-bound; unknown field paragraph and break run properties are preserved. Replacing an unknown marker, marker-style choice, click action, or inline kind is unsupported. Pictures, charts, groups, connectors, content parts, notes, masters, layouts, themes, and recursive part graphs are read-only; preserve them unchanged and treat `unsupported_presentation_edit` or topology errors as hard stops, never as permission for a lossy fallback.
- Use `bulletImage` with a base64 PNG/JPEG/GIF/SVG data URL for native `a:buBlip` picture bullets in direct/grouped shapes, Slide Master text styles/placeholders, or Slide Layout placeholders; external URIs remain non-fetched and relationship-backed. Do not use a Unicode glyph as a silent substitute. Inspect inherited placeholders after import and require a renderer compatibility check before visual sign-off.
- For native charts, use `styleId`/`styleIndex` (1-48), `varyColors`, `barOptions` (`direction`, `grouping`, `gapWidth`, `overlap`), and `lineOptions` (`grouping`, `marker`, `smooth`) instead of approximating those choices with shapes. Use `chartType: "combo"` for a bar+line plot and set every series `chartType` to `bar` or `line`. Set any bar/line series to `axisGroup: "secondary"` (aliases `axis: "y2"` and `secondaryAxis: true`) and provide `axes.secondary.value.title` when the scales differ; keep at least one primary series. A bar, line, or combo chart may split one plot type across both axis groups, which emits separate native plot blocks while preserving global series order. Series accept `fill`/`color`, `line`/`stroke`, indexed `points` with fill/line overrides, and `dataLabels` overrides for value/category visibility and position; line series may also override `marker` and `smooth`. Use `trendline` for one or `trendlines` for several standard native trendlines (`linear`, `exponential`, `logarithmic`, `movingAverage`, `polynomial`, or `power`). All six types feed bounded deterministic SVG curve previews for vertical/horizontal bars, lines, and combo charts as well as native OOXML. Exponential and power previews require positive source values, while the non-positive portion of a logarithmic forecast domain is omitted; unsupported model portions remain native but never emit non-finite SVG. Use `errorBars` for x/y, both/minus/plus scalar fixed/percentage/standard-deviation/standard-error values or per-point custom plus/minus arrays, with optional end-cap and line styling. Custom ranges use `plusFormula`/`minusFormula` with optional cached `plusValues`/`minusValues` and `plusFormatCode`/`minusFormatCode`; the aliases `plusReference`/`minusReference` are accepted. Formula references require chart `externalData` containing embedded XLSX bytes (for example, `SpreadsheetFile.exportXlsx(workbook)`) or an absolute workbook URI, so the package contains a valid relationship-backed data source. Embedded bytes must pass the XLSX package inspector before export; external URIs remain non-fetched. Formula-only third-party references remain formula-only on roundtrip, while cached values drive deterministic SVG preview. SVG also previews primary/secondary scales and error bars.
- Add connectors before nodes in the visual stack and route them through whitespace; never let a connector cross a label or unrelated node.
- Fix every unintended overlap, off-canvas element, clipping, and text overflow issue before delivery.
- Do not deliver a PPTX until semantic/package checks and full-size review of every rendered slide pass. A montage is an overview, not a substitute for per-slide review.

## Authoring workflow

1. Create a `Presentation` or import an existing PPTX with `PresentationFile.importPptx`.
2. Define the communication job and select a coherent theme/master/layout system before adding slides. Use `presentation.master` for the first-master compatibility path or `presentation.masters.add(...)` for multiple Slide Masters, then bind every layout with `masterId`; layout backgrounds/placeholders override only their owning master. Use `presentation.theme.setColors(...)`, `.setFonts(...)`, `.setTextStyles(...)`, and `.setColorMap(...)` for deck defaults. Put reusable title/body/other list levels in a master's `textParagraphStyles`; placeholder `paragraphStyles`, layout styles, and direct slide paragraphs override that cascade. Keep individual slide/shape styles for deliberate exceptions.
3. Inspect the relevant slides, text ranges, tables, charts, images, notes, and comments before editing.
   Imported unsupported native objects appear in `presentation.inspect({ kind: "nativeObject" })` and resolve by ID. `contentPart`, OLE, SmartArt/diagram, and generic graphic-frame facades are deliberately `editable:false`: preserve them through ordinary modeled edits and second export, but do not claim semantic editing or replace their raw XML/part graph.
4. Apply focused changes through public APIs and keep stable names on important objects.
   For bounded package surgery, `PresentationFile.patchPptx(...)` can attach caller-supplied image bytes or public chart XML to an existing slide with `recipe: { kind: "image"|"chart", source: "ppt/slides/slideN.xml", sourceReference: { objectId, name, alt, position } }`. Position is explicit pixels; the patcher owns DrawingML namespaces, relationship references, non-visual ID collision checks, deterministic replacement, and deletion cleanup.
5. Run `presentation.verify()` and `presentation.validateLayout()`; fix every material issue.
6. Export PPTX and import the exported file again.
7. Inspect native package parts with `PresentationFile.inspectPptx()`. Treat any notes/comments semantic issue as a delivery blocker: review relationships must originate from the correct slide/presentation part, roots/content types must match, every legacy comment author/index must resolve through the singleton author registry, and every Office 2021 modern comment/reply GUID must resolve through the singleton modern Author part.
8. Render every modeled slide with Playwright and render the original PPTX through LibreOffice to PDF plus Poppler PNGs.
9. Inspect every model and native slide image at full size. When a baseline is approved, compare PNG pixels on subsequent runs.

```js
import { Presentation, PresentationFile } from "open-office-artifact-tool";

const deck = Presentation.create({
  slideSize: { width: 1280, height: 720 },
  commentFormat: "modern", // omit for legacy commentAuthors.xml compatibility
  master: {
    name: "Evidence Master",
    background: { fill: "bg1", mode: "reference", index: 1001 },
    theme: { name: "Evidence Theme", colors: { accent1: "#3D8DFF" } },
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
const evidenceList = slide.shapes.add({
  name: "evidence-list",
  position: { left: 42, top: 112, width: 1196, height: 58 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: [{ bulletCharacter: "◆", bulletFont: "Georgia", bulletColor: "#DC2626", bulletSizePercent: 1.25, marginLeft: 24, indent: -12, runs: [{ run: "Inspectable", textStyle: { bold: true } }, " structured text"] }],
});
evidenceList.text.style = { fontFamily: "Arial", fontSize: 18, color: "#334155" };
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

Exercise the source-built Open XML SDK WebAssembly path and its opaque graph preservation:

```sh
node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/openxml-wasm-preservation.json \
  --output-dir tmp/presentation-wasm-fixture \
  --native-render required
```

Run the Office 2021 comments fixture when author/reply/status fidelity matters:

```sh
node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/modern-comments.json \
  --output-dir tmp/presentation-modern-comments \
  --native-render required
```

`--native-render auto` runs LibreOffice + Poppler when available and records a skip otherwise. Use `required` for final local delivery when those runtimes are installed, and `off` only for an explicitly documented structural-only check.

## QA gates

- `PresentationFile.inspectPptx(...)` proves required slide, chart, media, notes, comments, comment-author registry, theme, master, layout, and preserved native-object parts exist. It rejects wrong review-part roots/content types/sources, orphan or duplicate review relationships, multiple author registries, missing/duplicate author IDs, invalid or duplicate per-author comment indexes, and `lastIdx` values below the maximum used index. Native import must preserve per-comment/reply author identity even when notes, comments, or `commentAuthors.xml` use nonstandard relationship targets. For `nativeObject` records, require the expected `nativeKind`, relationship count, preserved-part count, and `editable:false`; after re-export, inspect again and render the real PPTX.
- `presentation.inspect(...)` proves agent-facing objects, master/layout identity, and review metadata survived roundtrip.
- Package evidence must include the presentation master list, master/layout parts, and the master↔layout plus slide→layout relationship chain when layouts are used.
- Theme evidence must include all 12 DrawingML color slots, major/minor Latin plus optional East-Asian/complex-script fonts, non-empty fill/line/effect/background format lists, a Slide Master `clrMap`, and title/body/other text styles. Every master must relate to its effective Theme part; inherited masters may share the deck Theme, while distinct overrides require distinct parts. Re-import and model/native rendering must restore the same effective values.
- Paragraph evidence must retain ordered text, field, and line-break inlines; field IDs/types/cached text; break styling; literal tab characters plus left/center/right/decimal tab stops; levels 0–8; character, auto-number, or relationship-backed picture bullets; marker font/color/fixed-or-percentage size; explicit `bullet*FollowText` overrides; margins/indents; point or percentage spacing; and the master → placeholder → layout → slide cascade. Direct/grouped, master, and layout picture bullets accept embedded PNG/JPEG/GIF/SVG data URLs or non-fetched external URIs; inspect effective `paragraphs` plus master/layout styles, re-export, and visually confirm marker choice, typography, indentation, field display values, line boundaries, and tab alignment. Each master/layout owns its own image relationship even when media bytes are shared. Structured inlines accept `link: { uri, tooltip?, targetFrame?, history?, highlightClick? }` for a non-fetched absolute external target, `link: { slideId, ... }` for an internal slide jump, `link: { action, ... }` for `nextSlide`, `previousSlide`, `firstSlide`, `lastSlide`, or `endShow`, and `link: { customShow, returnToSlide? }` for a named sequence created with `presentation.customShows.add(...)`. Verify relationship-backed `a:hlinkClick` in the owning slide/master/layout part. Relative actions and custom shows require empty `r:id`; custom shows additionally require a matching numeric `p:custShowLst` ID and reuse presentation-to-slide relationships for their ordered slide list. Import without private metadata, re-export, and inspect `customShow` plus SVG `data-hyperlink` evidence. LibreOfficeDev renders `a:buBlip` inconsistently (the current owner-inheritance fixture shows master-placeholder and layout-placeholder markers but omits the Slide Master `p:txStyles` marker), so require Windows PowerPoint evidence before claiming cross-renderer picture-bullet fidelity.
- Master/layout evidence must restore native `p:bg` solid or scheme references and merge placeholders by `type` plus `idx`; slide backgrounds override layout backgrounds, which override the linked master. Inspect and render must report the same effective background.
- The checked-in `package-drawing.json` fixture generates a chart part through the public facade, attaches it and an arbitrary-path image to an existing slide through `patchPptx`, restores both as editable agent-facing objects, and passes the real render gate. The `agent-readiness.json` fixture additionally gates a style-10 primary-column plus split-primary/secondary-line combo chart with independent model scales, two native line plot blocks, and a right-side value axis; it also covers explicit gap/overlap, per-series category/value label positions, named linear and polynomial trendlines (including the polynomial model curve), styled custom formula-reference error bars with native caches, a relationship-backed embedded workbook, dashed series outlines, a diamond line marker, indexed yellow/red point override, an embedded direct-shape picture bullet, a Slide Master-owned picture bullet materialized through layout inheritance, an external URI run hyperlink, a relationship-free `nextSlide` action, and a returning two-slide `QA Evidence` custom show. Inspect, OOXML, model rendering, and LibreOffice/Poppler all run.
- The checked-in `package-notes-comments.json` fixture relocates notes, comments, and the singleton author registry to arbitrary valid paths, proves semantic validation reports zero issues, restores author identity and note text, and passes LibreOffice/Poppler rendering.
- The checked-in `openxml-wasm-preservation.json` fixture crosses JavaScript authoring → bundled C# Open XML SDK import/list-style-text/link/field/break/tab-stop/picture-marker edit/export → JavaScript package/semantic/render QA. It changes an embedded content-addressed picture bullet plus character and auto-number markers; explicit/follow-text font, RGB/theme color, and fixed/percentage size; an external URI plus relative action while explicitly deleting another modeled link; field cached text; line-break styling; and tab-stop position/alignment. It retains the old picture media/relationship together with its chart, ordinary image, notes, master, theme, and every other unmodeled part/relationship.
- The checked-in `modern-comments.json` fixture emits Office 2021 `p188` Comment and Author parts, creates a real native `p:grpSp` containing a shape, table, relationship-backed chart, and picture, and anchors reviews through `txMk`, `grpSpMk`, `graphicFrameMk`, and `picMk`. It preserves GUID identities, native reply nesting, dates, authors, range offsets, group transforms, relationship parts, and resolved state across metadata-free import/export, then passes semantic plus LibreOffice/Poppler gates. Nested elements use complete ancestor moniker paths. Imported content-part/OLE/diagram objects themselves and their recursive part graphs survive second export, but modern comments targeting those read-only objects still use `unknownAnchor`.
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
