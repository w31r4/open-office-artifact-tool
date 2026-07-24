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
    import { createHash } from "node:crypto";
    import fs from "node:fs";
    import path from "node:path";
    import { pathToFileURL } from "node:url";

    import {
      DocumentFile, DocumentModel, FileBlob, PdfArtifact, PdfFile,
      Presentation, PresentationFile, SpreadsheetFile, Workbook,
    } from "open-office-artifact-tool";

    const canonicalCodec = await import("open-office-artifact-tool/codecs/open-chestnut");
    const legacyCodec = await import("open-office-artifact-tool/codecs/openxml-wasm");
    const canonicalWire = await import("open-office-artifact-tool/codecs/open-chestnut/wire");
    const legacyWire = await import("open-office-artifact-tool/codecs/openxml-wasm/wire");
    for (const name of Object.keys(canonicalCodec)) {
      if (legacyCodec[name] !== canonicalCodec[name]) process.exit(59);
    }
    if (
      legacyCodec.OPEN_XML_WASM_PROTOCOL_VERSION !== canonicalCodec.OPEN_CHESTNUT_PROTOCOL_VERSION ||
      legacyCodec.OpenXmlWasmCodecError !== canonicalCodec.OpenChestnutCodecError ||
      legacyCodec.exportXlsxWithOpenXmlWasm !== canonicalCodec.exportXlsxWithOpenChestnut ||
      legacyCodec.importXlsxWithOpenXmlWasm !== canonicalCodec.importXlsxWithOpenChestnut ||
      legacyCodec.invokeOpenXmlWasm !== canonicalCodec.invokeOpenChestnut ||
      legacyCodec.openXmlWasmStatus !== canonicalCodec.openChestnutStatus ||
      legacyWire.CodecRequestSchema !== canonicalWire.CodecRequestSchema ||
      legacyWire.CodecResponseSchema !== canonicalWire.CodecResponseSchema
    ) process.exit(60);

    const legacyWorkbook = Workbook.create();
    legacyWorkbook.worksheets.add("Legacy package import").getRange("A1").values = [["same runtime"]];
    const legacyXlsx = await legacyCodec.exportXlsxWithOpenXmlWasm(legacyWorkbook);
    const legacyImported = await legacyCodec.importXlsxWithOpenXmlWasm(legacyXlsx);
    if (
      legacyXlsx.metadata.codec !== "open-chestnut" ||
      legacyImported.worksheets.getItem("Legacy package import").getRange("A1").values[0][0] !== "same runtime"
    ) process.exit(61);

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
    {
      const packagedRoot = path.join(process.cwd(), "node_modules", "open-office-artifact-tool");
      const validationWorkflowPath = path.join(
        packagedRoot,
        "skills", "spreadsheets", "skills", "spreadsheets", "examples", "openchestnut-data-validation-workflow.mjs",
      );
      if (!fs.existsSync(validationWorkflowPath)) process.exit(8);
      const validationOutput = path.join(process.cwd(), "packed-data-validation.xlsx");
      const { createDataValidationWorkbook } = await import(pathToFileURL(validationWorkflowPath).href);
      const validationResult = await createDataValidationWorkbook(validationOutput);
      const validationRoundTrip = await SpreadsheetFile.importXlsx(await FileBlob.load(validationOutput));
      const validationSheet = validationRoundTrip.worksheets.getItem("Intake");
      if (
        validationResult.audit.provider.actual !== "open-chestnut" ||
        validationResult.audit.provider.fallbackUsed ||
        validationSheet.dataValidations.items.length !== 3 ||
        validationSheet.dataValidations.items[0].rule.prompt !== "Pick the current workflow state." ||
        validationSheet.dataValidations.items[0].rule.errorStyle !== "information" ||
        validationSheet.dataValidations.items[0].rule.showDropdown !== false
      ) process.exit(9);
    }

    const document = DocumentModel.create({ paragraphs: ["clean install DOCX"] });
    document.addInsertion("packaged accepted insertion", { author: "Package QA" });
    document.addDeletion("packaged removed deletion", { author: "Package QA" });
    document.setSettings({ trackRevisions: true, documentProtection: "comments" });
    document.addWatermark("PACKAGED DRAFT", { sectionIndex: 0 });
    document.addImage({
      dataUrl: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAACXBIWXMAAAPoAAAD6AG1e1JrAAAADUlEQVR4nGNgYGBgAAAABQABpfZFQAAAAABJRU5ErkJggg==",
      alt: "Packed floating image",
      widthPx: 80,
      heightPx: 60,
      placement: {
        type: "floating",
        horizontal: { relativeTo: "margin", offsetPx: 24 },
        vertical: { relativeTo: "paragraph", offsetPx: 0 },
        wrap: "square",
        wrapSide: "right",
        distanceFromTextPx: { top: 2, right: 8, bottom: 2, left: 8 },
      },
    });
    const docx = await DocumentFile.exportDocx(document);
    if (docx.metadata.codec !== "open-chestnut" || docx.bytes[0] !== 0x50 || docx.bytes[1] !== 0x4b) process.exit(10);
    const importedDocument = await DocumentFile.importDocx(docx);
    if (importedDocument.blocks[0].text !== "clean install DOCX") process.exit(11);
    if (importedDocument.settings.documentProtection?.edit !== "comments") process.exit(42);
    if (importedDocument.watermarks.length !== 1 || importedDocument.watermarks[0].text !== "PACKAGED DRAFT") process.exit(46);
    const importedImage = importedDocument.blocks.find((block) => block.kind === "image");
    if (
      importedImage?.placement?.type !== "floating" ||
      importedImage.placement.horizontal.relativeTo !== "margin" ||
      importedImage.placement.vertical.relativeTo !== "paragraph" ||
      importedImage.placement.wrap !== "square" ||
      importedImage.placement.wrapSide !== "right"
    ) process.exit(44);
    importedImage.placement = {
      type: "floating",
      horizontal: { relativeTo: "page", offsetPx: 36 },
      vertical: { relativeTo: "paragraph", offsetPx: 0 },
      wrap: "topAndBottom",
      distanceFromTextPx: { top: 4, right: 0, bottom: 4, left: 0 },
    };
    importedDocument.watermarks[0].text = "PACKAGED REVIEW";
    const packagedDocument2 = await DocumentFile.importDocx(await DocumentFile.exportDocx(importedDocument));
    if (packagedDocument2.blocks[0].text !== "clean install DOCX") process.exit(12);
    const packagedImage2 = packagedDocument2.blocks.find((block) => block.kind === "image");
    if (packagedImage2?.placement?.wrap !== "topAndBottom" || packagedImage2.placement.horizontal.relativeTo !== "page") process.exit(45);
    if (packagedDocument2.watermarks.length !== 1 || packagedDocument2.watermarks[0].text !== "PACKAGED REVIEW") process.exit(47);
    const docxSourceHash = createHash("sha256").update(docx.bytes).digest("hex");
    const finalizedDocx = await DocumentFile.finalizeRevisions(docx, {
      mode: "accept",
      expectedSourceSha256: docxSourceHash,
    });
    const finalization = finalizedDocx.metadata.revisionFinalization;
    if (finalization.sourceSha256 !== docxSourceHash || finalization.insertionCount !== 1 || finalization.deletionCount !== 1) process.exit(13);
    if (finalization.trackingBefore !== true || finalization.trackingAfter !== false) process.exit(14);
    if (JSON.stringify(finalization.changedParts) !== JSON.stringify(["word/document.xml", "word/settings.xml"])) process.exit(15);
    const finalizedDocument = await DocumentFile.importDocx(finalizedDocx);
    if (finalizedDocument.blocks.some((block) => block.kind === "change") || finalizedDocument.settings.trackRevisions) process.exit(16);
    if (finalizedDocument.settings.documentProtection?.edit !== "comments") process.exit(43);
    if (!finalizedDocument.blocks.some((block) => block.text === "packaged accepted insertion") || finalizedDocument.blocks.some((block) => block.text === "packaged removed deletion")) process.exit(17);
    if (createHash("sha256").update(docx.bytes).digest("hex") !== docxSourceHash) process.exit(18);

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
    cloneSource.shapes.add({
      name: "packed-clone-links",
      geometry: "textbox",
      position: { left: 48, top: 190, width: 420, height: 64 },
      fill: "transparent",
      line: { fill: "transparent", width: 0 },
      text: [{ runs: [
        { text: "Guide ", link: { uri: "https://example.com/packed-clone" } },
        { text: "Next ", link: { action: "nextSlide" } },
        { text: "Review route", link: { customShow: "Packed route", returnToSlide: true } },
      ] }],
    });
    cloneFixture.customShows.add({ name: "Packed route", nativeId: 23, slides: [cloneSource] });
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
      cloneResult.audit.operation.runHyperlinks.relationshipCount !== 1 ||
      cloneResult.audit.operation.runHyperlinks.actionOnlyCount !== 2 ||
      cloneResult.audit.operation.runHyperlinks.customShowCount !== 1 ||
      !cloneResult.audit.validation.package.runHyperlinks.exactSourceGraphRetained ||
      !cloneResult.audit.validation.package.customShows.exactSourceMembershipRetained ||
      !cloneResult.audit.validation.reimport.customShowMembershipRetained ||
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
      packedClone.slides.getItem(1).comments.items[0].comments[0].text !== "Packaged closed-leaf clone comment." ||
      packedClone.slides.getItem(1).shapes.items.find((shape) => shape.name === "packed-clone-links").text.paragraphs[0].runs[2].link.customShow !== "Packed route" ||
      JSON.stringify(packedClone.customShows.getItem("Packed route").slideIds) !== JSON.stringify([packedClone.slides.getItem(0).id])
    ) process.exit(25);

    const pdf = PdfArtifact.create({ pages: [{ text: "clean install PDF" }] });
    const pdfFile = await PdfFile.exportPdf(pdf);
    if (pdfFile.bytes[0] !== 0x25 || pdfFile.bytes[1] !== 0x50 || pdfFile.bytes[2] !== 0x44 || pdfFile.bytes[3] !== 0x46) process.exit(30);
    const inspection = await PdfFile.inspectPdf(pdfFile);
    if (inspection.summary.pages !== 1 || !inspection.summary.tagged) process.exit(31);
    const importedPdf = await PdfFile.importPdf(pdfFile);
    if (!importedPdf.extractText().includes("clean install PDF")) process.exit(32);
    const pyhankoProviderPath = path.join(
      installedPackage,
      "skills", "pdf", "skills", "pdf", "scripts", "pyhanko_provider.py",
    );
    if (
      !fs.existsSync(pyhankoProviderPath) ||
      !fs.readFileSync(pyhankoProviderPath, "utf8").includes("open-office-artifact-tool.pyhanko-verify.v1")
    ) process.exit(33);
    const verapdfProviderPath = path.join(
      installedPackage,
      "skills", "pdf", "skills", "pdf", "scripts", "verapdf_provider.py",
    );
    if (
      !fs.existsSync(verapdfProviderPath) ||
      !fs.readFileSync(verapdfProviderPath, "utf8").includes("open-office-artifact-tool.verapdf-validation.v1")
    ) process.exit(34);

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
    const retainedReference = fs.readFileSync(path.join(template.skillPath, sidecar.reference));
    const retainedPreview = fs.readFileSync(path.join(template.skillPath, sidecar.preview));
    if (
      template.schemaVersion !== 2 ||
      sidecar.schemaVersion !== 2 ||
      sidecar.id !== "artifact-template-packed-workbook-template" ||
      sidecar.displayName !== "Packed workbook template" ||
      sidecar.kind !== "spreadsheet" ||
      sidecar.reference !== "assets/reference.xlsx" ||
      sidecar.preview !== "assets/preview.png" ||
      sidecar.visualCommitment !== "opinionated" ||
      sidecar.editProfile?.level !== "copy-only" ||
      sidecar.provenance?.referenceSha256 !== createHash("sha256").update(retainedReference).digest("hex") ||
      sidecar.provenance?.previewSha256 !== createHash("sha256").update(retainedPreview).digest("hex")
    ) process.exit(53);
    if (
      !retainedReference.equals(Buffer.from(xlsx.bytes)) ||
      !retainedPreview.equals(fs.readFileSync(previewPath))
    ) process.exit(54);

    if (fs.existsSync(path.join(installedPackage, "skills", "default-template-library"))) process.exit(55);

    const officeKitQueryPath = path.join(
      installedPackage,
      "skills", "officekit", "skills", "officekit", "scripts", "query-templates.mjs",
    );
    if (!fs.existsSync(officeKitQueryPath)) process.exit(56);
    const queried = spawnSync(
      process.execPath,
      [
        officeKitQueryPath,
        "--kind", "spreadsheet",
        "--root", path.join(templateHome, "skills"),
        "--id", template.skillName,
      ],
      {
        cwd: process.cwd(),
        encoding: "utf8",
      },
    );
    if (queried.status !== 0) {
      process.stderr.write(queried.stderr);
      process.exit(57);
    }
    const catalogResult = JSON.parse(queried.stdout);
    if (
      catalogResult.selectionMade !== false ||
      catalogResult.invalid.length !== 0 ||
      catalogResult.candidates.length !== 1 ||
      catalogResult.candidates[0].id !== template.skillName ||
      catalogResult.candidates[0].editProfile?.level !== "copy-only" ||
      catalogResult.candidates[0].skillPath !==
        path.join(template.skillPath, "SKILL.md") ||
      catalogResult.candidates[0].referencePath !==
        path.join(template.skillPath, "assets", "reference.xlsx")
    ) process.exit(58);
  `;

  run(process.execPath, ["--input-type=module", "-e", probe], temporary, {
    PATH: process.platform === "win32" ? "C:\\Windows\\System32" : "/usr/bin:/bin",
  });
} finally {
  fs.rmSync(temporary, { force: true, recursive: true });
}

console.log("OpenChestnut, PDF, OfficeKit, and Template Creator clean-install package smoke ok; repository-only templates are excluded");

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
