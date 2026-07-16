---
name: open-office-presentations
description: Create, edit, inspect, render, and verify PPTX artifacts through the canonical OpenChestnut Office path.
---

# Open Office Presentations

Use this skill for standalone .pptx work. PresentationFile.importPptx and PresentationFile.exportPptx always use the bundled OpenChestnut C# WASM path. There is no JavaScript Office fallback or path selector.

## Contract

- Preserve an imported deck's package graph, content, masters, layouts, relationships, and unsupported objects unless the user explicitly requests a supported edit.
- Treat unsupported authoring, topology changes, missing source snapshots, and opaque-content edits as hard failures. Never retry through another writer.
- Use real text, tables, images, charts, and connectors rather than flattening them into screenshots.
- Give every slide one narrative job. Keep titles at least 35px on content slides, body text at least 16px, and all content inside the slide.
- Run semantic, package, model-render, and native-render checks before delivery. A montage is only an overview; inspect every slide at full size.

## Supported 0.2 boundary

OpenChestnut can create, import, edit, and re-export the following top-level presentation objects:

- ordinary rect, ellipse, roundRect, and textbox shapes;
- basic solid fill, solid line, transform, and outer-shadow properties;
- structured paragraphs and runs with bounded formatting, fields, line breaks, lists, tab stops, and supported external/internal/action hyperlinks;
- straight and elbow connectors with explicit endpoints or shape targets, RGB line color/width, and optional triangle start/end arrows;
- embedded PNG/JPEG images with rectangular frames and direct transforms;
- fixed-topology plain-text tables;
- source-free bar, line, and pie charts backed by literal categories and numeric values, with bounded title, legend, RGB series styling, line marker, axes, and data labels.

For an imported source-bound deck, keep the slide, element, paragraph, run, table, series, and category topology fixed. Existing text, frames, basic styles, connector arrows, image metadata/bytes of the same media type, table cell strings, and literal chart values may be edited only where the imported object reports a modeled boundary.

Advanced native objects remain source-bound and opaque. Masters and layouts are retained from the source package, but new or structurally changed masters/layouts are outside the 0.2 authoring boundary. Combo charts, formula/external-data charts, trendlines, error bars, grouped shapes, SmartArt, OLE, WordArt, modern comments, notes authoring, custom shows, animations, and other advanced objects must be preserved unchanged or rejected.

PresentationFile.inspectPptx and PresentationFile.patchPptx are explicit low-level package tools, not alternate writers. Use patchPptx only for a caller-requested, bounded image/chart package patch with an exact source reference.

## Workflow

1. Create a Presentation, or import the existing PPTX with PresentationFile.importPptx(input).
2. Inspect the deck and relevant slides/objects before editing.
3. Make only supported, focused changes through the public object model.
4. Run presentation.verify() and presentation.validateLayout(); fix every material issue.
5. Export with PresentationFile.exportPptx(presentation).
6. Import the exported PPTX again and confirm the intended semantic state.
7. Inspect the package with PresentationFile.inspectPptx(file).
8. Render every modeled slide with Playwright. For final local delivery, render the PPTX through LibreOffice to PDF and Poppler PNGs.
9. Review every slide image at full size and compare approved PNG baselines when available.

~~~js
import { Presentation, PresentationFile } from "open-office-artifact-tool";

const deck = Presentation.create({ slideSize: { width: 1280, height: 720 } });
const slide = deck.slides.add({ name: "Evidence" });

const source = slide.shapes.add({
  name: "source-card",
  geometry: "roundRect",
  position: { left: 70, top: 140, width: 300, height: 130 },
  fill: "#DBEAFE",
  line: { fill: "#2563EB", width: 2 },
  shadow: { color: "#000000", blurRadius: 8, distance: 4, direction: 45, opacity: 0.2 },
  text: "Create",
});
const target = slide.shapes.add({
  name: "target-textbox",
  geometry: "textbox",
  position: { left: 470, top: 140, width: 300, height: 130 },
  fill: "transparent",
  line: { fill: "transparent", width: 0 },
  text: "Verify",
});
slide.connectors.add({
  name: "evidence-flow",
  connectorType: "elbow",
  from: source,
  to: target,
  start: { x: 370, y: 205 },
  end: { x: 470, y: 220 },
  line: { fill: "#334155", width: 2, endArrow: "triangle" },
});
slide.tables.add({
  name: "evidence-table",
  position: { left: 70, top: 360, width: 380, height: 180 },
  values: [["Gate", "Result"], ["Semantic", "Pass"], ["Visual", "Review"]],
  styleOptions: { headerRow: true, bandedRows: true },
});
slide.charts.add("bar", {
  name: "evidence-chart",
  title: "Checks",
  position: { left: 560, top: 330, width: 560, height: 300 },
  categories: ["Semantic", "Package", "Visual"],
  series: [{ name: "Count", values: [8, 12, 6], color: "#2563EB" }],
  legend: false,
  dataLabels: { showValue: true },
});

const first = await PresentationFile.exportPptx(deck);
const imported = await PresentationFile.importPptx(first);
imported.slides.getItem(0).charts.items[0].series[0].values[2] = 7;
const final = await PresentationFile.exportPptx(imported);
await final.save("evidence.pptx");
~~~

## Verification commands

Run the checked-in core fixture end to end:

~~~sh
node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/agent-readiness.json \
  --output-dir tmp/presentation-skill-fixture \
  --native-render required
~~~

Run the create, import, edit, and second-export fixture:

~~~sh
node skills/presentations/scripts/run-fixture.mjs \
  --fixture skills/presentations/fixtures/open-chestnut-preservation.json \
  --output-dir tmp/presentation-roundtrip-fixture \
  --native-render required
~~~

Verify any PPTX:

~~~sh
node skills/presentations/scripts/verify-presentation.mjs \
  --input evidence.pptx \
  --output-dir tmp/presentation-qa \
  --native-render required
~~~

Create a visual baseline by adding --baseline-dir tmp/presentation-baselines --write-baseline true. Compare later runs with the same --baseline-dir and omit --write-baseline.

--native-render auto records a skip when LibreOffice/Poppler are unavailable. Use required for final local delivery when those runtimes are installed, and off only for an explicitly documented structural-only check.

## QA gates

- presentation.inspect(...) must expose the intended shapes, text, tables, images, charts, and connectors after the second import.
- presentation.verify() and presentation.validateLayout() must report no delivery-blocking structural, overlap, off-canvas, overflow, or dangling-reference issue.
- PresentationFile.inspectPptx(...) must report a valid package with the expected slide, media, chart, master, layout, and relationship parts.
- Source-bound unsupported parts and relationships must remain unchanged. Any attempted edit outside the modeled boundary must fail closed.
- Model and native renders must contain the expected slide count. Inspect each PNG at full size; baseline comparison must not silently ignore missing or extra slides.
- Microsoft PowerPoint remains the highest-fidelity optional gate for PowerPoint-specific behavior.

## References

- Generated public API catalog: ../../docs/api.md
- Current implementation coverage: ../../docs/coverage.md
- Fixture runner: scripts/run-fixture.mjs
- Generic PPTX verifier: scripts/verify-presentation.mjs
