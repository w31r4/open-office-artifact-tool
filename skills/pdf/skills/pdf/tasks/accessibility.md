# Accessibility and archival conformance

## Greenfield tagged authoring

`PdfArtifact` can emit explicit headings, reading order, language/title, Table/TR/TH/TD structure, Figure alternative text, decorative artifacts, and tagged marked content. Run `pdf.verify()` and `PdfFile.inspectPdf(...)` before export delivery.

This modeled profile is not a claim of full PDF/UA conformance.

## PDF/UA and PDF/A validation

Use veraPDF for machine-verifiable rules:

```bash
verapdf --format json output.pdf > tmp/pdfs/verapdf-report.json
```

Choose the intended profile explicitly when the default does not match the requirement. Preserve the exact veraPDF version, profile, exit status, and report.

veraPDF documents that PDF/UA includes human checkpoints beyond machine-verifiable rules. Manually review document purpose/title, meaningful reading sequence, heading semantics, link purpose, table interpretation, alternative text quality, color/contrast, and keyboard/assistive-technology behavior where relevant.

## Imported PDFs

Do not rebuild an arbitrary PDF through `PdfArtifact` merely to add a tagged wrapper. Use a provider that can edit the native structure tree or report the unsupported remediation boundary. Validate the final native file, not an intermediate model.
