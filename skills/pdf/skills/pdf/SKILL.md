---
name: "pdf"
description: "Create, inspect, edit, redact, sign, render, and verify PDF files through explicit, fail-closed provider routes. Use for new tagged documents, imported PDFs, forms, annotations, OCR, conformance, and visual QA."
---

# PDF Skill

## What this Skill does

This Skill gives an agent bounded, auditable PDF primitives. PDF is independent
from the OpenChestnut DOCX/XLSX/PPTX codec: do not add a PDF protobuf/WASM codec
or reconstruct an imported PDF through `PdfArtifact` or PDF.js and call that a
fidelity-preserving edit.

`PdfArtifact` is for new semantic/tagged documents. `PdfFile` plus required,
runtime-lazy MuPDF.js is the default for arbitrary PDF inspect, render, and
bounded direct-original edits. Every specialist route is explicit. A failed
provider is an error, never a silent fallback.

Use `createPdfjsParser()` only as an optional read-only PDF.js adapter for page
geometry, positioned text, heuristic tables, and bounded image evidence; it is
never an imported-PDF edit representation.

## Route every job first

Use this sequence for every request:

1. Preserve original bytes and inspect with MuPDF.js.
2. Resolve the declared intent with `PdfProviders.resolve`.
3. If the result is `installable`, obtain explicit project-policy authority and
   call `ensure`; otherwise use only the returned `ready` provider.
4. Probe the selected provider, perform one explicit save strategy, then
   inspect, audit, render, and review the output.

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");
let resolution = await PdfProviders.resolve({
  task: "edit-content",
  inspection,
  savePolicy: "rewrite",
  mutationAuthorized: true,
});
if (resolution.status === "installable") {
  resolution = await PdfProviders.ensure({ resolution });
}
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
```

The public resolver returns only `ready`, `installable`, or `blocked`, with the
selected provider/pack, platform, pinned version/artifact information, download
and unpack estimates, licence acknowledgement, runtime paths, prerequisites,
and operation boundary. It never selects an alternate provider.
`PdfProviders.ensure({ resolution })` never acquires P12/private keys,
HSM/remote-signing credentials, TSA/LTV access, or other secrets.
`PdfProviders.probe({ provider })` checks exactly that provider without a
download or fallback.

Run the default native CLI before an imported-PDF operation:

```bash
node scripts/mupdf.mjs probe
node scripts/mupdf.mjs inspect input.pdf
node scripts/mupdf.mjs render input.pdf tmp/pdfs/page-1.png --page 1 --dpi 144
```

The CLI budgets input/page/object/render work, refuses source overwrite
including symlink aliases, writes atomically, and never falls back.

## Capability-pack policy

The only required npm runtime is `mupdf@1.28.0`; it initializes only on the
first MuPDF-backed PDF operation. Root import and provider resolution do not
start WASM, download a pack, or modify the filesystem.

External capability packs live in the project-private
`.open-office-artifact-tool/providers/` cache. The conventional policy is
`.open-office-artifact-tool/pdf-providers.json`; a missing file means
`installPolicy: "disabled"`. An agent may install only after the user/project
sets `managed`, whitelists every provider and pack, accepts required licences,
allows the requested OCR languages, and supplies finite byte budgets:

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

Only hash-pinned, versioned project release assets may be installed; an
enterprise mirror must serve identical bytes. The installer uses a lock,
temporary download, exact size/hash checks, safe archive extraction, atomic
publication, and a receipt. It rejects `latest`, package-manager/global-pip
installs, lifecycle hooks, path traversal, links, and undeclared URLs. The
initial managed targets are `darwin-arm64` and `linux-x64`.

**Current release-catalog state:** qpdf and `python-foundation`
`3.13.14-oat.1` have published, attested `darwin-arm64` and `linux-x64`
assets. qpdf covers repair, linearization, inspection, and the bounded AES-256
delivery-copy route. The foundation is isolated CPython with ReportLab,
pdfplumber, pypdf, and Pillow only. `python-specialists` (PyMuPDF/pikepdf/
pyHanko), OCR, Poppler, and veraPDF/JRE packs remain unpublished and
deliberately resolve as `blocked`, not as downloadable promises. Use a selected
`system-only` policy with an explicitly configured local runtime when one is
already managed by the deployment. Once a task chooses managed or system-only,
do not automatically switch routes.

See [provider setup and probes](tasks/provider_setup.md) for the full policy,
system-runtime, probe, and failure contract. The [provider matrix](references/PROVIDER_MATRIX.md)
states human-readable capability boundaries; the public catalog is the only
source of pack versions, hashes, sizes, and installation facts.

## Choose the narrowest provider

| Need | Explicit route | Detailed task |
| --- | --- | --- |
| New tagged/semantic PDF | `PdfArtifact` | [create](tasks/create.md) |
| New layout-oriented PDF | ReportLab | [create](tasks/create.md) |
| Imported-PDF inspect/render/bounded edit | MuPDF.js / `scripts/mupdf.mjs` | [read](tasks/read_review.md), [edit](tasks/edit_existing.md) |
| Text/table geometry evidence | pdfplumber | [read](tasks/read_review.md) |
| Attachments, complex forms, merge/reorder/stamp | pypdf | [forms](tasks/forms_annotations.md), [transform](tasks/transform.md) |
| Strict scrub, residue/OCR redaction, advanced bounded edit | PyMuPDF | [redact](tasks/redact.md) |
| Repair or linearize | `scripts/qpdf_provider.py` | [repair](tasks/repair_linearize.md) |
| Create an AES-256 encrypted delivery copy | `scripts/qpdf_provider.py` | [encryption](tasks/encryption.md) |
| Active/auxiliary structure cleanup | `scripts/pikepdf_provider.py` | [structure cleanup](tasks/structure_clean.md) |
| Searchable-layer OCR | `scripts/ocrmypdf_provider.py` | [OCR](tasks/ocr.md) |
| Local PKCS#12 sign or signature validation | `scripts/pyhanko_sign_provider.py`, `scripts/pyhanko_provider.py` | [sign](tasks/sign_verify.md) |
| PDF/A or PDF/UA machine checks | `scripts/verapdf_provider.py` | [accessibility](tasks/accessibility.md) |
| Independent native visual QA | Poppler | [render review](tasks/render_review.md) |

## Imported-PDF invariants

Keep the source immutable, bind its SHA-256, choose `read-only`, `rewrite`,
`incremental`, or `sanitize` before mutation, and publish a distinct output.
Inspect signatures, ByteRange, DocMDP/FieldMDP, encryption, forms, annotations,
attachments, metadata, active content, page boxes, and page count first.
`incremental` preserves old bytes; it is not signature authorization.

MuPDF.js supports only bounded direct-original operations. Its inspect output
keeps raw `mediaBox`/`cropBox` as unrotated PDF-space facts and emits a rotated
effective `mupdf-page-space` bbox for 0/90/180/270-degree placement. Use its
returned `sourceSha256`, `mupdf-link`/annotation/form locators, page snapshot,
and `appearanceBbox`; re-inspect after every output.

- `add_text_annotation`, `add_text_highlight`, `add_link`, `delete_annotation`,
  `update_annotation`, `delete_link`, `update_link`, and `update_form_field`
  are source-bound operations; placement uses the inspected page snapshot.
  `add_text_annotation` takes a visible pin and rewrite; a highlight requires a
  unique native text selection and rewrite.
- `duplicate_page` requires source SHA-256, is the only operation in a full
  rewrite, and requires Poppler pixel identity after reinspection.
- `set_page_crop` is raw unrotated CropBox visibility only, not redaction.
  `rotate_page` sets an absolute right-angle `/Rotate`; neither enables content
  reflow. Delete/redaction operations cannot be incremental.
- General Word-style reflow, arbitrary text replacement, Dynamic XFA, complex
  JavaScript, 3D, and RichMedia are not made safe by these primitives. Preserve
  them opaquely when the selected operation allows it, otherwise fail closed.

For the complete operation schemas and edge cases, read [edit existing](tasks/edit_existing.md)
and [forms and annotations](tasks/forms_annotations.md), not this overview.

## Specialist safety boundaries

- qpdf receives a source SHA-256 before repair or linearize. Its separate
  `encrypt` primitive creates one AES-256 copy from an unencrypted source using
  caller-owned restricted password files and private argument files; it does
  not open/decrypt/re-encrypt existing encrypted PDFs or edit permissions. Both
  routes are structural/full-rewrite operations, not redaction or sanitize;
  use [repair](tasks/repair_linearize.md) or [encryption](tasks/encryption.md).
- pikepdf offers only `active-content` and `active-and-auxiliary` profiles. It
  is not redaction, metadata cleanup, or XFA cleanup; use
  [structure cleanup](tasks/structure_clean.md).
- PyMuPDF `redact_ocr_text` is sanitize-only: require exact
  `expected_rotation` (0/90/180/270), a named language, match count, bounded
  raster work, residue evidence, rewrite, and invalidation acknowledgement.
  Coordinates remain unrotated PyMuPDF page space. A complete imported PDF
  OCR workflow is not a sanitizer; see [OCR](tasks/ocr.md).
- pyHanko local PKCS#12 signing uses a passphrase on stdin only. `pyhanko_provider.py`
  validates under an explicit trust root. Timestamp, LTV, PKCS#11/HSM, remote
  signing, and network evidence are external workflows, never auto-installed.
- veraPDF's `verapdf_provider.py` is a machine-rule gate, not repair or a substitute for
  human review of PDF/UA. See [accessibility](tasks/accessibility.md).

When a Python specialist is selected, one configured virtual environment
executable remains provider identity so its `pyvenv.cfg` is preserved. Do not
retry via a different Python interpreter after a failed probe.

## Delivery gate

1. Record selected provider, version, policy/receipt (if any), source hash,
   save strategy, and no-fallback evidence in the canonical audit envelope.
2. Reopen and verify the intended change; use `scripts/pdf_audit.py validate`.
3. Run the requested specialist evidence. Sanitization must pass its residue
   and single-revision gates; signatures and PDF/UA retain their separate
   validation/human boundaries.
4. Render every final page with MuPDF.js or independently with Poppler and
   inspect clipping, overlaps, glyphs, images, fields, annotations, signatures,
   redactions, and page geometry before delivery.

The project and its required MuPDF.js dependency are GNU AGPL-3.0-or-later.
Managed and system providers retain their own licences; the resolver exposes
the acknowledgement required before any installation or operation.
