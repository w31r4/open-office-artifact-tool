# Provider setup and probes

## Default: MuPDF.js

`mupdf@1.28.0` is the sole required PDF runtime. A normal npm installation is
enough; its WASM starts only for a MuPDF-backed operation:

```bash
npm install open-office-artifact-tool
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
```

There is no PDF `postinstall`, standalone downloader, virtual-environment
bootstrapper, or global mutation. `PdfFile.importPdf` and the thin CLI enforce
input/image/page/object/render limits, source/output separation (including
symlink aliases), and atomic output publication.

## Resolve before preparing a specialist runtime

Use the public subpath rather than guessing a local install command:

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");

const resolution = await PdfProviders.resolve({
  task: "repair",
  provider: "qpdf",
  inspection,
  savePolicy: "rewrite",
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
});
console.log(resolution.status, resolution.reason, resolution.installPlan);
```

`resolve` is read-only. It describes one selected route only and returns
`ready`, `installable`, or `blocked`; it never initializes MuPDF, downloads,
changes `PATH`, chooses a fallback, or acquires secrets. Pass the resulting
object unchanged to `PdfProviders.ensure({ resolution })` only when it is
`installable` and project policy explicitly permits a managed installation.
`PdfProviders.probe({ provider, policyPath })` checks exactly one provider and does not
download or mutate anything.

## Project policy

The conventional policy file is
`.open-office-artifact-tool/pdf-providers.json`. Its missing default is:

```json
{ "installPolicy": "disabled" }
```

To authorize managed packs, commit or otherwise provision an explicit policy
appropriate for the project. This example authorizes only qpdf and fixes an
upper byte budget; it does not authorize another provider:

```json
{
  "installPolicy": "managed",
  "allowedProviders": ["qpdf"],
  "allowedPacks": ["qpdf"],
  "acceptedLicenses": [],
  "allowedOcrLanguages": ["eng", "chi_sim"],
  "maxDownloadBytes": 250000000,
  "maxUnpackedBytes": 750000000
}
```

`managed` requires an exact catalogue artifact for the current platform, every
pack in the dependency closure, licence acknowledgement where declared, and
both budgets. Initial managed targets are `darwin-arm64` and `linux-x64`.
`eng` and `chi_sim` are the initial OCR policy defaults; any other language must
appear in both policy and the immutable language-pack catalogue.

The cache is always private to the project:

```text
.open-office-artifact-tool/providers/<pack>/<version>/<platform>/
```

The installer accepts only declared, hash-pinned, versioned HTTPS release
assets (or an enterprise mirror that serves identical hash-pinned bytes). It
uses lock files, temporary downloads, exact size/SHA-256 verification, safe
archive extraction, atomic publication, and receipts. It rejects `latest`,
package-manager installs, global pip, dynamic npm installation, lifecycle
hooks, undeclared URLs, absolute paths, `..`, symlink/hardlink archive entries,
and cache paths that escape the project directory.

After `ensure` returns `ready`, use only the exact paths in
`result.runtime.managed`: pass its `pythonPath`, `commandPaths`, and
`environment` to the selected adapter. Do not reconstruct a cache path or mix
it with a system runtime. For managed OCR, `environment` includes
`OPEN_OFFICE_PDF_TESSDATA_DIRS`, a platform path-list of the selected verified
language-pack directories. `ocrmypdf_provider.py` copies regular, unlinked
`.traineddata` files into a per-operation private directory and sets its own
`TESSDATA_PREFIX`; never point `TESSDATA_PREFIX` directly at the project cache.

**Current catalog state:** qpdf, `python-foundation`, and
`python-specialists` `3.13.14-oat.1`, plus veraPDF/JRE `1.30.2-oat.1`, have
published attested assets for `darwin-arm64` and `linux-x64`. A
policy-authorized qpdf route, a ReportLab/pdfplumber/pypdf route with the
foundation pack, a PyMuPDF/pikepdf/pyHanko route with `python-specialists` and
qpdf, or a veraPDF conformance route can resolve as `installable` and be passed
unchanged to `ensure`. The foundation contains isolated CPython plus ReportLab,
pdfplumber, pypdf, and Pillow. Specialists contain PyMuPDF, pikepdf, pyHanko,
and certificate validation, and require the catalogued AGPL-or-commercial
acknowledgement. The veraPDF pack carries its own JRE. OCR core/language and
Poppler packs remain unpublished, so those routes are `blocked` with a precise
reason even under a permissive policy. Do not substitute a hand-written
download URL or claim that `ensure` installed a future pack.

## Existing controlled runtime: system-only

`system-only` lets a deployment select its already-provisioned runtime without
granting any download authority. Whitelist exactly the provider and set its
documented runtime path, for example:

```json
{
  "installPolicy": "system-only",
  "allowedProviders": ["qpdf"],
  "allowedPacks": [],
  "acceptedLicenses": [],
  "allowedOcrLanguages": ["eng", "chi_sim"],
  "maxDownloadBytes": 0,
  "maxUnpackedBytes": 0
}
```

```bash
export OPEN_OFFICE_PDF_QPDF="/controlled/runtime/bin/qpdf"
python3 scripts/pdf_provider.py check --provider qpdf --require
python3 scripts/qpdf_provider.py probe
```

Use the analogous selected runtime variables only for the provider actually
chosen by the resolver: `OPEN_OFFICE_PDF_PROVIDER_PYTHON`,
`OPEN_OFFICE_PDF_VERAPDF`, `OPEN_OFFICE_PDF_OCRMYPDF`,
`OPEN_OFFICE_PDF_GS`, or the Poppler paths.
For Python, point `OPEN_OFFICE_PDF_PROVIDER_PYTHON` to the virtual-environment
executable itself (`bin/python` or `Scripts/python.exe`), not its resolved base
interpreter. Every shipped Python entry point re-executes through that exact
link so `pyvenv.cfg` and installed modules remain authoritative. A missing or
wrong version is blocked; no automatic `PATH` search can switch an explicitly
selected managed/system route.

Existing deployment provisioning remains the deployment owner's responsibility.
The Skill deliberately does not prescribe `brew`, `apt`, global pip, or another
package manager as an implicit repair path.

## Credentials, licences, and operation evidence

The pyHanko runtime can be selected by policy, but P12/private keys, HSMs,
remote-signing credentials, TSA/LTV access, and trust roots are always
caller-provided. The passphrase travels only on stdin. The qpdf encryption
task likewise requires caller-owned user/owner password files; the resolver
records only the credential type and never receives their values. PyMuPDF and any other
catalogue licence requiring acknowledgement must be listed in
`acceptedLicenses`; availability never waives that acknowledgement.

After a provider is `ready`, run its task-specific probe and plan, preserve the
exact source hash, select `read-only`, `rewrite`, `incremental`, or `sanitize`,
and use a distinct transactional output. Follow the linked task guide for the
provider's source-binding, signature-invalidation, OCR, output, and QA rules:

- [repair and linearization](repair_linearize.md)
- [AES-256 delivery-copy encryption](encryption.md)
- [active/auxiliary cleanup](structure_clean.md)
- [scanned-PDF OCR](ocr.md)
- [sign and verify](sign_verify.md)
- [accessibility and conformance](accessibility.md)
- [redaction and sanitize](redact.md)

The [provider matrix](../references/PROVIDER_MATRIX.md) explains what each
route can and cannot do. The public catalog, not this guide, is the source of
versions, hashes, package sizes, artefact URLs, and platform availability.
