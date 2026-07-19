# Accessibility and archival conformance

## Greenfield tagged authoring

`PdfArtifact` can emit explicit headings, reading order, language/title, Table/TR/TH/TD structure including constrained logical tables spanning consecutive pages, Figure alternative text, meaningful Link annotations with OBJR association, decorative/running artifacts, and tagged marked content. Run `pdf.verify()` and `PdfFile.inspectPdf(...)` before export delivery.

Use [the accessible board report example](../examples/accessible-board-report.mjs) as an executable pattern. It deliberately reports three separate results: project-modeled verification, optional veraPDF machine-rule evidence, and still-required human PDF/UA judgment.

This modeled profile is not a claim of full PDF/UA conformance.

## PDF/UA and PDF/A validation

Use the shipped adapter with a separately installed veraPDF 1.30.x CLI for machine-verifiable rules. Bind the final bytes and choose one supported built-in profile explicitly:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
SOURCE_SHA256="$(shasum -a 256 output.pdf | awk '{print $1}')"
"$PYTHON_BIN" scripts/pdf_provider.py check --provider verapdf --require
"$PYTHON_BIN" scripts/verapdf_provider.py probe
"$PYTHON_BIN" scripts/verapdf_provider.py validate output.pdf \
  --expected-sha256 "$SOURCE_SHA256" \
  --flavour ua1 \
  --require-compliant \
  > tmp/pdfs/verapdf-ua1.json
```

Supported profiles are `1a`, `1b`, `2a`, `2b`, `2u`, `3a`, `3b`, `3u`, `4`, `4e`, `4f`, `ua1`, and `ua2`. The adapter rejects automatic/default selection, custom profiles, passwords, directories, and arbitrary veraPDF flags. It validates a private snapshot, re-proves the source identity, bounds execution and report size, and returns a typed report without retaining the raw provider JSON. Preserve the selected profile, veraPDF component versions, exact source SHA-256, rule counts, and bounded failed-rule evidence.

Without `--require-compliant`, a noncompliant PDF is a successful validation operation whose `machineRuleCompliant` value is `false`. Use the flag when compliance is a delivery requirement; the same structured report is emitted on the failure path.

veraPDF documents that PDF/UA includes human checkpoints beyond machine-verifiable rules. The adapter always marks PDF/UA reports as requiring human review. Manually review document purpose/title, meaningful reading sequence, heading semantics, link purpose, table interpretation, alternative text quality, color/contrast, and keyboard/assistive-technology behavior where relevant.

## Imported PDFs

Do not rebuild an arbitrary PDF through `PdfArtifact` merely to add a tagged wrapper. Use a provider that can edit the native structure tree or report the unsupported remediation boundary. Validate the final native file, not an intermediate model.
