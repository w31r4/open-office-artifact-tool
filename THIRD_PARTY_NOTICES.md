# Third-party notices

`open-office-artifact-tool` is an MIT-licensed clean-room implementation. It does not contain or redistribute the reference package, its runtime artifacts, runtime module modules, or runtime bindings.

The npm package declares or optionally integrates the following public libraries. Versions are pinned by `package-lock.json`; consult each installed package's `LICENSE` file for complete license text.

| Component | Role | License |
| --- | --- | --- |
| JSZip | OOXML ZIP package reading/writing | MIT option from `(MIT OR GPL-3.0-or-later)` |
| PDF.js (`pdfjs-dist`) | Optional arbitrary-PDF parsing | Apache-2.0 |
| Playwright | Optional deterministic browser rendering | Apache-2.0 |
| sharp | Optional raster conversion and JPEG/WebP pixel decoding | Apache-2.0 |
| node-canvas (`canvas`) | Optional canvas raster adapter | MIT |

Sharp may install platform-specific libvips binary packages under LGPL-3.0-or-later, sometimes combined with Apache-2.0/MIT components. These are optional dynamically linked runtime dependencies and are not copied into this repository or npm tarball. Downstream distributors remain responsible for the corresponding LGPL notices and relinking/source obligations when redistributing those binaries.

The following external programs are invoked only when installed separately. They are not bundled in the npm package:

- Chromium or another Playwright-managed browser: see the browser distribution's own third-party notices.
- LibreOffice: MPL-2.0 and LGPL-3.0-or-later licensing applies to the separately installed application.
- Poppler command-line tools: GPL licensing applies to the separately installed binaries.
- .NET 8: used to build/run the optional Office bridge; it is not bundled.
- Microsoft Office: reference software required only for optional Windows native automation; users must supply a valid installation and license.

OOXML, Open Packaging Conventions, PDF, and related file-format specifications are used as public interoperability standards. Reference package observations are limited to public package shape, exports, examples, smoke tests, and observable behavior.
