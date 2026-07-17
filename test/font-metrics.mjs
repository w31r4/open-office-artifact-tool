import assert from "node:assert/strict";

import {
  clearOfficeFontDesignMetrics,
  DocumentModel,
  Presentation,
  registerScopedOfficeFontDesignMetrics,
  resolveOfficeFontDesignMetrics,
  setOfficeFontDesignMetrics,
  skiaPaintBaselineCompensationPx,
  Workbook,
} from "open-office-artifact-tool";

const metric = (family, weight, ascent, extra = {}) => ({
  family,
  weight,
  unitsPerEm: 1_000,
  ascent,
  descent: 200,
  lineGap: 0,
  ...extra,
});

clearOfficeFontDesignMetrics();
setOfficeFontDesignMetrics([
  metric("Alpha", 300, 700),
  metric("Alpha", 400, 710),
  metric("Alpha", 700, 720),
  metric("alpha", 400, 730, { style: "italic" }),
  { family: "invalid", weight: 400, unitsPerEm: 1_000, ascent: 800, descent: -200 },
]);
assert.deepEqual(resolveOfficeFontDesignMetrics({ family: ["ALPHA"] }), {
  family: "Alpha",
  familyKey: "alpha",
  weight: 400,
  unitsPerEm: 1_000,
  ascent: 710,
  descent: 200,
  lineGap: 0,
  style: "normal",
  width: "normal",
});
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"], weight: 350 }).weight, 300);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"], weight: 600 }).weight, 700);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"], style: "italic" }).ascent, 730);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Missing", "Alpha"] }), undefined, "resolution must not silently skip the requested primary family");

const disposeOlder = registerScopedOfficeFontDesignMetrics([metric("Alpha", 400, 810)]);
const disposeNewer = registerScopedOfficeFontDesignMetrics([metric("Alpha", 400, 820)]);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"] }).ascent, 820);
disposeOlder();
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"] }).ascent, 820);
disposeNewer();
disposeNewer();
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"] }).ascent, 710);
setOfficeFontDesignMetrics([metric("Beta", 400, 900)]);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Alpha"] }), undefined);
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Beta"] }).ascent, 900);
clearOfficeFontDesignMetrics();
assert.equal(resolveOfficeFontDesignMetrics({ family: ["Beta"] }), undefined);
assert.throws(() => setOfficeFontDesignMetrics({}), TypeError);
assert.throws(() => registerScopedOfficeFontDesignMetrics(undefined), TypeError);

for (const value of [-2.75, -0.5, 0, 0.25, 0.5, 1.75, 10.333333]) {
  assert.equal(skiaPaintBaselineCompensationPx(value), value - Math.round(value));
}
assert.equal(skiaPaintBaselineCompensationPx(Number.NaN), 0);
assert.equal(skiaPaintBaselineCompensationPx(undefined), 0);

const document = DocumentModel.create({
  blocks: [{
    kind: "paragraph",
    text: "Font inventory",
    runs: [
      { text: "Font ", style: { fontFamily: "Zulu" } },
      { text: "inventory", style: { fontFamily: " alpha " } },
    ],
  }],
  defaultRunStyle: { fontFamily: "Zulu" },
});
assert.deepEqual(document.fontFamilies, ["alpha", "Aptos", "Aptos Display", "Zulu"]);
const documentFonts = document.fontFamilies;
documentFonts.push("mutation");
assert.doesNotMatch(document.fontFamilies.join(","), /mutation/);

const workbook = Workbook.create();
const sheet = workbook.worksheets.add("Fonts");
sheet.getRange("A1").values = [["Primary"]];
sheet.getRange("A1").format = { font: { name: "Zulu" } };
sheet.getRange("B1").values = [["Secondary"]];
sheet.getRange("B1").format = { font: { name: "alpha" } };
sheet.getRange("C1").values = [["Duplicate"]];
sheet.getRange("C1").format = { font: { name: "zulu" } };
assert.deepEqual(workbook.fontFamilies, ["alpha", "Aptos", "Zulu"]);

const presentation = Presentation.create();
const slide = presentation.slides.add();
slide.shapes.add({
  text: [
    { runs: [{ text: "Primary", style: { fontFamily: "Zulu" } }] },
    { bulletFont: "alpha", runs: [{ text: "Secondary", style: { fontFamily: "Alpha" } }] },
  ],
  textStyle: { fontFamily: "Zulu" },
});
assert.deepEqual(presentation.fontFamilies, ["Alpha", "Zulu"]);

console.log("font metrics smoke ok");
