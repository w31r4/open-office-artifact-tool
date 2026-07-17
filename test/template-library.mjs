import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { pathToFileURL } from "node:url";

const packageRoot = path.resolve(import.meta.dirname, "..");
const libraryRoot = path.join(
  packageRoot,
  "skills/presentations/skills/presentations/assets/builtin_templates/grid-layout-library",
);
const composeRoot = path.join(libraryRoot, "artifact-tool-compose");
const registryPath = path.join(composeRoot, "template-registry.json");
const previewIndexPath = path.join(libraryRoot, "assets/index.json");
const previewDirectory = path.join(libraryRoot, "assets/previews");

try {
  await fs.access(libraryRoot);
} catch (error) {
  if (error?.code === "ENOENT") {
    console.log("template library smoke skipped: repository-only skills are not packaged");
    process.exit(0);
  }
  throw error;
}

const { buildPresentation, exportPresentation } = await import(
  "../skills/presentations/skills/presentations/builtin_templates_support/scripts/create-presentation.mjs"
);
const outputDirectory = await fs.mkdtemp(
  path.join(os.tmpdir(), "office-artifact-tool-grid-layout-library-"),
);
const outputPath = path.join(outputDirectory, "grid-layout-library.pptx");

function getSlideCount(presentation) {
  if (typeof presentation.slides?.count === "number") return presentation.slides.count;
  if (Array.isArray(presentation.slides?.items)) return presentation.slides.items.length;
  throw new Error("Presentation slide collection does not expose a count.");
}

try {
  const [registry, previewIndex, composeModule, previewEntries] = await Promise.all([
    fs.readFile(registryPath, "utf8").then(JSON.parse),
    fs.readFile(previewIndexPath, "utf8").then(JSON.parse),
    import(pathToFileURL(path.join(composeRoot, "index.mjs")).href),
    fs.readdir(previewDirectory, { withFileTypes: true }),
  ]);

  if (registry.templateSlug !== "grid-layout-library") {
    throw new Error(`Unexpected template slug: ${registry.templateSlug}`);
  }
  if (registry.templates.length !== 26) {
    throw new Error(`Expected 26 registry templates, found ${registry.templates.length}.`);
  }
  if (!Array.isArray(composeModule.builders) || composeModule.builders.length !== 26) {
    throw new Error("The layout index must export exactly 26 builders.");
  }
  if (previewIndex.previews.length !== 26) {
    throw new Error(`Expected 26 preview-index records, found ${previewIndex.previews.length}.`);
  }

  const previewFiles = previewEntries
    .filter((entry) => entry.isFile() && entry.name.endsWith(".png"))
    .map((entry) => entry.name)
    .sort();
  if (previewFiles.length !== 27 || !previewFiles.includes("layout-library.png")) {
    throw new Error(`Expected 27 preview PNGs including the contact sheet, found ${previewFiles.length}.`);
  }

  const expectedTemplateIds = Array.from(
    { length: 26 },
    (_, index) => `grid-layout-library#slide-${String(index + 1).padStart(2, "0")}`,
  );
  const actualTemplateIds = registry.templates.map((template) => template.templateId);
  if (JSON.stringify(actualTemplateIds) !== JSON.stringify(expectedTemplateIds)) {
    throw new Error("Registry template IDs are not the expected sequential grid-layout-library IDs.");
  }

  for (const template of registry.templates) {
    const componentPath = path.resolve(composeRoot, template.component.module);
    const component = await import(pathToFileURL(componentPath).href);
    if (typeof component[template.component.exportName] !== "function") {
      throw new Error(`Missing builder ${template.component.exportName} in ${template.component.module}.`);
    }
  }

  for (const preview of previewIndex.previews) {
    const previewPath = path.resolve(path.dirname(previewIndexPath), preview.path);
    const previewStat = await fs.stat(previewPath);
    if (!previewStat.isFile() || previewStat.size <= 0) {
      throw new Error(`Preview is missing or empty: ${preview.path}`);
    }
  }

  const presentation = await buildPresentation(libraryRoot);
  if (getSlideCount(presentation) !== 26) {
    throw new Error(`Expected reconstructed deck to contain 26 slides, found ${getSlideCount(presentation)}.`);
  }

  await exportPresentation(libraryRoot, outputPath);
  const outputStat = await fs.stat(outputPath);
  if (!outputStat.isFile() || outputStat.size <= 0) {
    throw new Error(`Layout-library reconstruction produced an empty PPTX: ${outputPath}`);
  }

  console.log("template library smoke ok");
} finally {
  await fs.rm(outputDirectory, { force: true, recursive: true });
}
