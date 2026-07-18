import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs";
import os from "node:os";
import path from "node:path";

const repoRoot = path.resolve(import.meta.dirname, "..");
const temporary = fs.mkdtempSync(path.join(os.tmpdir(), "open-office-artifact-tool-pack-"));

try {
  const packed = run("npm", ["pack", repoRoot, "--json", "--ignore-scripts", "--pack-destination", temporary], repoRoot);
  const report = JSON.parse(packed.stdout)[0];
  const tarball = path.join(temporary, report.filename);
  assert.ok(fs.existsSync(tarball), `npm pack did not create ${tarball}`);
  const dependencyTarballs = packProductionDependencies(temporary);
  // Exercise a real npm install without making a release gate depend on the
  // registry. Optional renderer peers are intentionally outside this core
  // OpenChestnut/PDF probe and remain covered by package metadata tests.
  run("npm", [
    "install", "--offline", "--ignore-scripts", "--no-audit", "--no-fund",
    "--omit=dev", "--legacy-peer-deps", "--no-save", tarball, ...dependencyTarballs,
  ], temporary);

  const probe = String.raw`
    import { spawnSync } from "node:child_process";
    import fs from "node:fs";
    import path from "node:path";
    import { pathToFileURL } from "node:url";

    import {
      DocumentFile, DocumentModel, FileBlob, PdfArtifact, PdfFile,
      Presentation, PresentationFile, SpreadsheetFile, Workbook,
    } from "open-office-artifact-tool";

    try {
      await import("open-office-artifact-tool/codecs/openxml-wasm");
      process.exit(40);
    } catch (error) {
      if (error?.code !== "ERR_PACKAGE_PATH_NOT_EXPORTED") process.exit(41);
    }

    const workbook = Workbook.create();
    const sheet = workbook.worksheets.add("Packaged");
    sheet.getRange("A1:B2").values = [["Label", "Value"], ["clean install", 7]];
    sheet.getRange("D1:E3").values = [["X", "Y"], [1, 3], [2, 8]];
    const scatter = sheet.charts.add("scatter", sheet.getRange("D1:E3"));
    scatter.title = "Packed scatter";
    scatter.series.items[0].marker = { symbol: "circle", size: 6, fill: "#0EA5E9" };
    sheet.getRange("G1:I3").values = [["X", "Y", "Size"], [1, 3, 4], [2, 8, 9]];
    const bubble = sheet.charts.add("bubble", sheet.getRange("G1:I3"));
    bubble.title = "Packed bubble";
    bubble.series.items[0].fill = "#38BDF8";
    const xlsx = await SpreadsheetFile.exportXlsx(workbook);
    if (xlsx.metadata.codec !== "open-chestnut" || xlsx.bytes[0] !== 0x50 || xlsx.bytes[1] !== 0x4b) process.exit(1);
    const importedWorkbook = await SpreadsheetFile.importXlsx(xlsx);
    if (importedWorkbook.worksheets.getItem("Packaged").getRange("B2").values[0][0] !== 7) process.exit(2);
    const importedScatter = importedWorkbook.worksheets.getItem("Packaged").charts.items[0];
    if (importedScatter.type !== "scatter" || importedScatter.xAxis.axisType !== "valueAxis") process.exit(4);
    if (JSON.stringify(importedScatter.series.items[0].xValues) !== "[1,2]") process.exit(5);
    const importedBubble = importedWorkbook.worksheets.getItem("Packaged").charts.items[1];
    if (importedBubble.type !== "bubble" || importedBubble.xAxis.axisType !== "valueAxis") process.exit(6);
    if (JSON.stringify(importedBubble.series.items[0].bubbleSizes) !== "[4,9]") process.exit(7);
    const xlsx2 = await SpreadsheetFile.exportXlsx(importedWorkbook, { recalculate: false });
    if ((await SpreadsheetFile.importXlsx(xlsx2)).worksheets.getItem("Packaged").getRange("A2").values[0][0] !== "clean install") process.exit(3);

    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    const docx = await DocumentFile.exportDocx(document);
    if (docx.metadata.codec !== "open-chestnut" || docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(10);
    const importedDocument = await DocumentFile.importDocx(docx);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(11);
    if ((await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDocument))).blocks[0].text !== "clean install DOCX") process.exit(12);

    const presentation = Presentation.create();
    presentation.slides.add({ name: "Packaged" }).shapes.add({
      name: "Title", geometry: "roundRect", text: "clean install PPTX",
      position: { left: 40, top: 40, width: 520, height: 80 },
    });
    const pptx = await PresentationFile.exportPptx(presentation);
    if (pptx.metadata.codec !== "open-chestnut" || pptx.bytes[0] !== 0x50 || pptx.bytes[1] !== 0x4b) process.exit(20);
    const importedPresentation = await PresentationFile.importPptx(pptx);
    if (importedPresentation.slides.getItem(0).shapes.items[0].text.value !== "clean install PPTX") process.exit(21);
    if ((await PresentationFile.importPptx(await PresentationFile.exportPptx(importedPresentation))).slides.count !== 1) process.exit(22);

    const installedPackage = path.join(process.cwd(), "node_modules", "open-office-artifact-tool");
    const duplicateWorkflowPath = path.join(
      installedPackage,
      "skills", "presentations", "skills", "presentations", "examples", "openchestnut-slide-duplicate-workflow.mjs",
    );
    if (!fs.existsSync(duplicateWorkflowPath)) process.exit(23);
    const cloneFixture = Presentation.create({ slideSize: { width: 640, height: 360 } });
    const cloneSource = cloneFixture.slides.add({
      name: "Packed clone source",
      notes: "Packaged closed-leaf clone notes.",
    });
    const cloneGroup = cloneSource.addGroup({
      name: "packed-cluster",
      position: { left: 48, top: 40, width: 320, height: 120 },
      childFrame: { left: 0, top: 0, width: 320, height: 120 },
    });
    const cloneLeft = cloneGroup.shapes.add({ name: "left", position: { left: 0, top: 20, width: 90, height: 42 }, text: "Left" });
    const cloneRight = cloneGroup.shapes.add({ name: "right", position: { left: 210, top: 20, width: 90, height: 42 }, text: "Right" });
    cloneGroup.connectors.add({
      name: "join", from: cloneLeft, to: cloneRight,
      start: { x: 90, y: 41 }, end: { x: 210, y: 41 }, line: { fill: "#64748B", width: 1 },
    });
    cloneSource.comments.addThread(undefined, "Packaged closed-leaf clone comment.", {
      author: "Package QA",
      created: "2026-07-18T03:05:00Z",
      position: { x: 360, y: 240 },
    });
    const cloneInput = path.join(process.cwd(), "packed-clone-source.pptx");
    const cloneOutput = path.join(process.cwd(), "packed-clone-output.pptx");
    const cloneAudit = path.join(process.cwd(), "packed-clone-audit.json");
    await (await PresentationFile.exportPptx(cloneFixture)).save(cloneInput);
    const { duplicatePptxSlide } = await import(pathToFileURL(duplicateWorkflowPath).href);
    const cloneResult = await duplicatePptxSlide({
      inputPath: cloneInput,
      outputPath: cloneOutput,
      auditPath: cloneAudit,
      expectedName: "Packed clone source",
      allowClosedLeaves: true,
    });
    if (
      cloneResult.audit.operation.clonePart !== "ppt/slides/slide2.xml" ||
      !cloneResult.audit.operation.closedLeaves.speakerNotes ||
      !cloneResult.audit.operation.closedLeaves.legacyComments ||
      !cloneResult.audit.validation.package.retainedSourcePartsByteIdentical ||
      !cloneResult.audit.validation.package.closedLeaves.speakerNotes?.notesXmlByteIdentical ||
      !cloneResult.audit.validation.package.closedLeaves.legacyComments?.commentsXmlByteIdentical ||
      !cloneResult.audit.validation.reimport.sourceAndCloneSemanticsEqual ||
      !cloneResult.audit.validation.reimport.sourceAndCloneClosedLeavesEqual ||
      !cloneResult.audit.validation.modelRender.visualEquivalent
    ) process.exit(24);
    const packedClone = await PresentationFile.importPptx(new FileBlob(await fs.promises.readFile(cloneOutput), {
      type: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
      name: "packed-clone-output.pptx",
    }));
    if (
      packedClone.slides.count !== 2 ||
      packedClone.slides.getItem(1).groups.items[0].connectors.items[0].startTargetId !== packedClone.slides.getItem(1).groups.items[0].shapes.items[0].id ||
      packedClone.slides.getItem(1).speakerNotes.text !== "Packaged closed-leaf clone notes." ||
      packedClone.slides.getItem(1).comments.items[0].comments[0].text !== "Packaged closed-leaf clone comment."
    ) process.exit(25);

    const pdf = PdfArtifact.create({ pages: [{ text: "clean install PDF" }] });
    const pdfFile = await PdfFile.exportPdf(pdf);
    if (pdfFile.bytes[0] !== 0x25 || pdfFile.bytes[1] !== 0x50 || pdfFile.bytes[2] !== 0x44 || pdfFile.bytes[3] !== 0x46) process.exit(30);
    const inspection = await PdfFile.inspectPdf(pdfFile);
    if (inspection.summary.pages !== 1 || !inspection.summary.tagged) process.exit(31);
    const importedPdf = await PdfFile.importPdf(pdfFile);
    if (!importedPdf.extractText().includes("clean install PDF")) process.exit(32);

    const creatorPath = path.join(
      installedPackage,
      "skills", "template-creator", "skills", "template-creator", "scripts", "create-template-skill.mjs",
    );
    if (!fs.existsSync(creatorPath)) process.exit(50);

    const fixtureDirectory = path.join(process.cwd(), "template-creator-fixture");
    const templateHome = path.join(process.cwd(), "template-creator-home");
    const referencePath = path.join(fixtureDirectory, "reference.xlsx");
    const previewPath = path.join(fixtureDirectory, "preview.png");
    fs.mkdirSync(fixtureDirectory, { recursive: true });
    fs.writeFileSync(referencePath, xlsx.bytes);
    fs.writeFileSync(
      previewPath,
      Buffer.from(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==",
        "base64",
      ),
    );
    const created = spawnSync(
      process.execPath,
      [
        creatorPath,
        "--reference-path", referencePath,
        "--preview-path", previewPath,
        "--display-name", "Packed workbook template",
        "--description", "Create a workbook from the clean-installed package fixture.",
      ],
      {
        cwd: process.cwd(),
        encoding: "utf8",
        env: { ...process.env, OFFICE_ARTIFACT_HOME: templateHome },
      },
    );
    if (created.status !== 0) {
      process.stderr.write(created.stderr);
      process.exit(51);
    }
    const template = JSON.parse(created.stdout);
    if (
      template.kind !== "spreadsheet" ||
      template.skillName !== "artifact-template-packed-workbook-template" ||
      path.dirname(template.skillPath) !== path.join(templateHome, "skills")
    ) process.exit(52);
    const sidecar = JSON.parse(fs.readFileSync(path.join(template.skillPath, "artifact-template.json"), "utf8"));
    if (
      sidecar.schemaVersion !== 1 ||
      sidecar.kind !== "spreadsheet" ||
      sidecar.reference !== "assets/reference.xlsx" ||
      sidecar.preview !== "assets/preview.png"
    ) process.exit(53);
    if (
      !fs.readFileSync(path.join(template.skillPath, sidecar.reference)).equals(Buffer.from(xlsx.bytes)) ||
      !fs.readFileSync(path.join(template.skillPath, sidecar.preview)).equals(fs.readFileSync(previewPath))
    ) process.exit(54);

    if (fs.existsSync(path.join(installedPackage, "skills", "default-template-library"))) process.exit(55);
  `;

  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin",
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("OpenChestnut, PDF, and Template Creator clean-install package smoke ok; repository-only templates are excluded");

function run(command, args, cwd, environment = {}) {
  const result = spawnSync(command, args, {
    cwd,
    encoding: "utf8",
    env: { ...process.env, ...environment },
    shell: false,
  });
  assert.equal(result.status, 0, `${command} ${args.join(" ")} failed\nSTDOUT:\n${result.stdout}\nSTDERR:\n${result.stderr}`);
  return result;
}

function packProductionDependencies(temporary) {
  const lock = JSON.parse(fs.readFileSync(path.join(repoRoot, "package-lock.json"), "utf8"));
  const destination = path.join(temporary, "dependency-tarballs");
  fs.mkdirSync(destination, { recursive: true });
  return Object.entries(lock.packages || {})
    .filter(([location, metadata]) => location.startsWith("node_modules/") && !metadata.dev && !metadata.optional && !metadata.peer)
    .map(([location]) => {
      const source = path.join(repoRoot, location);
      assert.ok(fs.existsSync(source), `npm ci production dependency is missing: ${location}`);
      const packed = run("npm", ["pack", source, "--json", "--ignore-scripts", "--pack-destination", destination], repoRoot);
      const report = JSON.parse(packed.stdout)[0];
      return path.join(destination, report.filename);
    });
}
