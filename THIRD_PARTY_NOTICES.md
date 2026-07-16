# Third-party notices

`open-office-artifact-tool` is an MIT-licensed implementation. It does not contain or redistribute the reference package's private runtime artifacts, compiled implementation modules, WebAssembly payloads, or native bindings. The npm package does include adapted MIT-licensed reference Skill workflow text, scripts, and assets under `skills/`; the Office codec itself is independently source-built as OpenChestnut.

The npm package declares or optionally integrates the following public libraries. Versions are pinned by `package-lock.json`; consult each installed package's `LICENSE` file for complete license text.

| Component | Role | License |
| --- | --- | --- |
| Buf Protobuf (`@bufbuild/protobuf`) | Public protobuf wire-schema runtime for JavaScript | Apache-2.0 AND BSD-3-Clause |
| JSZip | OOXML ZIP package reading/writing | MIT option from `(MIT OR GPL-3.0-or-later)` |
| PDF.js (`pdfjs-dist`) | Optional arbitrary-PDF parsing | Apache-2.0 |
| Playwright | Optional deterministic browser rendering | Apache-2.0 |
| sharp | Optional raster conversion and JPEG/WebP pixel decoding | Apache-2.0 |
| node-canvas (`canvas`) | Optional canvas raster adapter | MIT |
| Microsoft Open XML SDK (`DocumentFormat.OpenXml`) | OOXML package codec compiled from source into the bundled WebAssembly runtime | MIT |
| Google Protobuf for .NET | Public protobuf wire-schema runtime compiled from source into the bundled WebAssembly runtime | BSD-3-Clause |
| .NET 8 WebAssembly runtime | Bundled execution runtime for the OpenChestnut codec | MIT plus the upstream third-party notices shipped under `runtime/open-chestnut/` |
| Reference file-type Skill bundles | Agent workflow text, helper scripts, and visual assets adapted for the public package | MIT, covered by the repository license and retained Documents Skill license copy |

The PDF Skill also ships thin Python scripts that can call the following separately installed providers. These Python packages and binaries are not npm dependencies, are not copied into this repository, and are not included in the npm tarball:

| Optional PDF provider | Skill role | Upstream license |
| --- | --- | --- |
| ReportLab | Greenfield layout-oriented PDF creation | BSD |
| pdfplumber | Read-only text/geometry/table extraction | MIT |
| pypdf | Basic PDF structure, AcroForm, annotation, rewrite, and incremental operations | BSD-3-Clause |
| PyMuPDF / MuPDF | Advanced imported-PDF editing, redaction, scrub, and strict residue inspection | GNU AGPL-3.0 or an Artifex commercial license |

This project has explicitly approved the PyMuPDF GNU AGPL provider route for its own workflow. The npm package itself remains MIT licensed because it does not contain, link, install, or redistribute PyMuPDF/MuPDF; the Skill requires a separately selected Python environment and an explicit `agpl` or `commercial` acknowledgement before invoking that provider. Downstream installation, network deployment, modification, and redistribution must comply with the selected PyMuPDF/MuPDF license. This notice is not a substitute for the upstream license text or legal advice.

Buf CLI, `protoc-gen-es`, `Grpc.Tools`, and the .NET SDK/WebAssembly workload are build-only tools. Their generated protocol bindings or compiled outputs are shipped, but the tools themselves are not included in the npm tarball.

Sharp may install platform-specific libvips binary packages under LGPL-3.0-or-later, sometimes combined with Apache-2.0/MIT components. These are optional dynamically linked runtime dependencies and are not copied into this repository or npm tarball. Downstream distributors remain responsible for the corresponding LGPL notices and relinking/source obligations when redistributing those binaries.

The following external programs are invoked only when installed separately. They are not bundled in the npm package:

- Chromium or another Playwright-managed browser: see the browser distribution's own third-party notices.
- LibreOffice: MPL-2.0 and LGPL-3.0-or-later licensing applies to the separately installed application.
- Poppler command-line tools: GPL licensing applies to the separately installed binaries.
- qpdf: optional separately installed structural inspection/rewrite CLI; retain its Apache-2.0 and applicable embedded-component notices.
- pikepdf: planned optional qpdf-based Python provider; no mutation adapter is shipped in this release.
- pyHanko: optional separately installed PDF signing and signature-validation provider; retain its package and transitive cryptography notices.
- veraPDF: optional separately installed PDF/A and PDF/UA validation distribution; retain the notices shipped by the selected veraPDF components.
- OCRmyPDF and Tesseract: optional OCR providers; no OCRmyPDF adapter is shipped in this release, while strict image residue checks can use a separately installed Tesseract through PyMuPDF.
- .NET 8 SDK: used to build the WebAssembly codec and optional Office bridge; the SDK is not bundled. The source-built .NET WebAssembly runtime needed by consumers is bundled with its upstream license and notices.
- Microsoft Office: reference software required only for optional Windows native automation; users must supply a valid installation and license.

OOXML, Open Packaging Conventions, PDF, and related file-format specifications are used as public interoperability standards. OpenChestnut codec behavior is implemented from source with the public Open XML SDK and the repository's versioned wire schema; no private reference runtime binary is shipped.
