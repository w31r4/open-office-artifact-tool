# PDF operation audit schema

Every imported-PDF mutation and security-sensitive read-only extraction emits one canonical `open-office-artifact-tool.pdf-audit.v1` JSON record. Do not invent aliases such as `outputs.pdf`, `actual_provider`, `provider_version`, `save_strategy`, or `silent_fallback`; downstream Agent and evaluator code reads the stable camelCase fields below.

Required success shape:

```json
{
  "schema": "open-office-artifact-tool.pdf-audit.v1",
  "status": "succeeded",
  "source": { "path": "/absolute/input.pdf", "bytes": 123, "sha256": "..." },
  "output": { "path": "/absolute/output.pdf", "bytes": 123, "sha256": "..." },
  "provider": {
    "actual": "pymupdf",
    "version": "1.27.2.3",
    "licenseChoice": "agpl",
    "silentFallback": false
  },
  "savePolicy": { "strategy": "sanitize" },
  "preflight": { "probeCompleted": true, "planCompleted": true },
  "operation": { "type": "replace_text" },
  "validation": {}
}
```

`source.sha256` and `output.sha256` identify the exact delivered bytes. `provider.actual`, `provider.version`, `provider.silentFallback`, `savePolicy.strategy`, and `operation.type` are never inferred from prose. Provider-specific evidence, fit checks, signature policy, residue scans, Poppler results, warnings, and task-specific assertions belong in additional fields without renaming the canonical envelope.

For a multi-source operation, keep `source` as the exact operation manifest and add an `inputs` array containing the same file-evidence shape for every imported PDF. The validator requires a matching repeated `--input` path for every record and recomputes each byte count and SHA-256; ordering is not significant, but missing, duplicate, or extra records fail.

```json
{
  "source": { "path": "/absolute/merge-stamp.json", "bytes": 456, "sha256": "..." },
  "inputs": [
    { "path": "/absolute/cover.pdf", "bytes": 123, "sha256": "..." },
    { "path": "/absolute/report.pdf", "bytes": 234, "sha256": "..." }
  ]
}
```

For attachment quarantine, use `savePolicy.strategy: "read-only"`, `operation.type: "extract-attachments"`, and bind `output` to the delivered `attachments.json` manifest. The quarantine file hashes and contained paths remain task-specific validation evidence inside that manifest and the audit `validation` object.

For bounded qpdf encryption, use `savePolicy.strategy: "rewrite"` and
`operation.type: "qpdf-encrypt-aes-256"`. Record the catalog credential
declaration, qpdf version, AES-256/key-bit evidence, signature decision,
authorized `checkAfter`, topology comparison, and password-safe render review.
Never record either password, a password hash/fingerprint, a secret-file path,
private argument-file contents, or a command line containing a secret.

For `failed_closed`, set `output` to `null`, include a non-empty `reason`, keep the source/provider/save-policy/operation evidence, record `preflight.probeCompleted` and `preflight.planCompleted` truthfully (either may be `false` when that gate caused the refusal), and do not leave a partial modified PDF at the requested output path. A `succeeded` record requires both preflight fields to be `true`.

Validate before delivery:

```bash
python3 scripts/pdf_audit.py validate outputs/audit.json \
  --source inputs/source.pdf \
  --artifact outputs/modified.pdf \
  --require-operation replace_text
```

For a merge, use the manifest as `--source` and repeat `--input` for every source PDF.

The validator recomputes source and artifact byte counts and SHA-256 hashes. It does not replace semantic, signature, residue, conformance, or render verification; those results remain required entries under `validation` when applicable. The machine-readable envelope is in [`pdf-audit-v1.schema.json`](pdf-audit-v1.schema.json).
