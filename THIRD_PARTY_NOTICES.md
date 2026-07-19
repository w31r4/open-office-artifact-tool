# Third-party notices

`open-office-artifact-tool` is licensed under GNU AGPL-3.0-or-later. It does not contain or redistribute the reference package's private runtime artifacts, compiled implementation modules, WebAssembly payloads, or native bindings. The npm package includes adapted reference Skill workflow text, scripts, and assets under `skills/`; their retained upstream MIT notices remain in force for the original material, while this project's modifications and combined distribution use the repository's AGPL license. The Office codec itself is independently source-built as OpenChestnut.

The npm package declares or optionally integrates the following public libraries. Versions are pinned by `package-lock.json`; consult each installed package's `LICENSE` file for complete license text.

| Component | Role | License |
| --- | --- | --- |
| Buf Protobuf (`@bufbuild/protobuf`) | Public protobuf wire-schema runtime for JavaScript | Apache-2.0 AND BSD-3-Clause |
| JSZip | OOXML ZIP package reading/writing | MIT option from `(MIT OR GPL-3.0-or-later)` |
| MuPDF.js (`mupdf`) | Required, runtime-lazy arbitrary-PDF parsing, inspection, rendering, and bounded native editing | GNU AGPL-3.0-or-later |
| PDF.js (`pdfjs-dist`) | Optional arbitrary-PDF parsing | Apache-2.0 |
| Playwright | Optional deterministic browser rendering | Apache-2.0 |
| sharp | Optional raster conversion and JPEG/WebP pixel decoding | Apache-2.0 |
| node-canvas (`canvas`) | Optional canvas raster adapter | MIT |
| Microsoft Open XML SDK (`DocumentFormat.OpenXml`) | OOXML package codec compiled from source into the bundled WebAssembly runtime | MIT |
| Google Protobuf for .NET | Public protobuf wire-schema runtime compiled from source into the bundled WebAssembly runtime | BSD-3-Clause |
| .NET 8 WebAssembly runtime | Bundled execution runtime for the OpenChestnut codec | MIT plus the upstream third-party notices shipped under `runtime/open-chestnut/` |
| Reference file-type Skill bundles | Agent workflow text, helper scripts, and visual assets adapted for the public package | Retained upstream MIT notices for original material; project modifications and combined distribution are AGPL-3.0-or-later |

## Repository-only Default Template Library

`skills/default-template-library/` is a repository-only import of the 20 Office template skills introduced by [`office-artifact-tool` commit `256cb31bfe0a07b3cef0051b6b159342be381378`](https://github.com/w31r4/office-artifact-tool/commit/256cb31bfe0a07b3cef0051b6b159342be381378), **Add default Office template library**. That source repository declares the following MIT license and copyright:

> MIT License — Copyright (c) 2026 w31r4

The full retained notice is at [`skills/default-template-library/LICENSE.md`](skills/default-template-library/LICENSE.md). The source Office files and PNG previews remain byte-for-byte copies; `skills/default-template-library/integrity.json` records the source commit plus individual and aggregate SHA-256 values. This repository-only directory is intentionally excluded from the npm tarball, so the package's AGPL distribution does not redistribute these assets.

The PDF Skill also ships thin Python scripts that can call the following separately installed providers. These Python packages and binaries are not npm dependencies, are not copied into this repository, and are not included in the npm tarball:

| Optional PDF provider | Skill role | Upstream license |
| --- | --- | --- |
| ReportLab | Greenfield layout-oriented PDF creation | BSD |
| pdfplumber | Read-only text/geometry/table extraction | MIT |
| pypdf | Basic PDF structure, AcroForm, annotation, rewrite, and incremental operations | BSD-3-Clause |
| PyMuPDF | Optional specialist strict scrub, legacy high-level imported-PDF edits, and residue inspection not yet covered by the JavaScript path | GNU AGPL-3.0 or an Artifex commercial license |
| pyHanko core and pyhanko-certvalidator | Source-bound read-only PDF signature integrity, trust, difference, timestamp, DocMDP, and FieldMDP validation | MIT |
| veraPDF 1.30.x CLI | Source-bound read-only PDF/A and PDF/UA machine-rule validation | MPL-2.0-or-later and GPL-3.0-or-later options, plus distribution notices |
| OCRmyPDF 17.8.x, Tesseract 5.x, and the selected OCRmyPDF runtime dependencies | Source-bound complete-document searchable-layer OCR through a shipped thin adapter | OCRmyPDF: MPL-2.0; Tesseract: Apache-2.0; pypdfium2: BSD-3-Clause/Apache-2.0; pikepdf: MPL-2.0; fpdf2: LGPL-3.0-only; retain every selected distribution's transitive notices |

The official `mupdf` npm package is a required direct dependency. A normal npm installation resolves it alongside this package, although its bytes remain in its own dependency tarball rather than being copied into this project's `.tgz`; the WASM runtime initializes lazily on the first PDF operation. There is no lifecycle hook or standalone downloader. Optional Python providers remain separately installed tools. Downstream installation, network deployment, modification, and redistribution must comply with the applicable GNU AGPL v3-or-later obligations. This notice is not a substitute for the upstream license text or legal advice.

Buf CLI, `protoc-gen-es`, `Grpc.Tools`, and the .NET SDK/WebAssembly workload are build-only tools. Their generated protocol bindings or compiled outputs are shipped, but the tools themselves are not included in the npm tarball.

Sharp may install platform-specific libvips binary packages under LGPL-3.0-or-later, sometimes combined with Apache-2.0/MIT components. These are optional dynamically linked runtime dependencies and are not copied into this repository or npm tarball. Downstream distributors remain responsible for the corresponding LGPL notices and relinking/source obligations when redistributing those binaries.

The following external programs are invoked only when installed separately. They are not bundled in the npm package:

- Chromium or another Playwright-managed browser: see the browser distribution's own third-party notices.
- LibreOffice: MPL-2.0 and LGPL-3.0-or-later licensing applies to the separately installed application.
- Poppler command-line tools: GPL licensing applies to the separately installed binaries.
- qpdf 11+: optional separately installed structural inspection/recovery/linearization CLI used through the shipped thin provider script; retain its Apache-2.0 and applicable embedded-component notices.
- pikepdf: planned optional qpdf-based Python provider; no mutation adapter is shipped in this release.
- pyHanko: optional separately installed PDF provider. The shipped thin adapter
  uses pyHanko core `>=0.35,<0.36` for read-only signature validation; signing,
  timestamp, and LTV command workflows additionally use the separately packaged
  `pyhanko-cli`. Neither distribution is bundled by npm. Retain their MIT
  license and transitive cryptography notices.
- veraPDF 1.30.x: optional separately installed PDF/A and PDF/UA validation distribution used through the shipped bounded adapter; retain its MPL-2.0-or-later/GPL-3.0-or-later choice and the notices shipped by the selected components.
- OCRmyPDF `>=17.8,<17.9` and Tesseract 5.x: optional separately installed OCR providers used by the shipped source-bound complete-document searchable-layer adapter. The adapter also requires qpdf 11+, Poppler `pdftotext`, and OCRmyPDF's separately installed fpdf2/pypdfium runtime; none is bundled by npm. OCRmyPDF is MPL-2.0, Tesseract is Apache-2.0, and the transitive runtime components retain their own notices. Strict image-residue checks remain the separate PyMuPDF/Tesseract sanitize gate.
- .NET 8 SDK: used to build the WebAssembly codec and optional Office bridge; the SDK is not bundled. The source-built .NET WebAssembly runtime needed by consumers is bundled with its upstream license and notices.
- Microsoft Office: reference software required only for optional Windows native automation; users must supply a valid installation and license.

OOXML, Open Packaging Conventions, PDF, and related file-format specifications are used as public interoperability standards. OpenChestnut codec behavior is implemented from source with the public Open XML SDK and the repository's versioned wire schema; no private reference runtime binary is shipped.

## Repository-only veraPDF test fixture

`test/fixtures/pdf/verapdf-pdfa1b-pass.pdf` is copied byte-for-byte from the
official [`veraPDF-corpus`](https://github.com/veraPDF/veraPDF-corpus) at commit
`49de56cd987929932c9e4fbbbe67d052bf44ef83`, path
`PDF_A-1b/6.1 File structure/6.1.2 File header/veraPDF test suite 6-1-2-t01-pass-a.pdf`.
That corpus is licensed under Creative Commons Attribution 4.0. The retained
fixture SHA-256 is
`66077f449d472a048e3bbf7192aa6d2b0b0ebd6b6d8a6f878f776f69424b6deb`.
It is used only by repository tests and is excluded from the npm package.
