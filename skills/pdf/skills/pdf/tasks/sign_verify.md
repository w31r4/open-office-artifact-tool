# Sign and verify PDFs

Use pyHanko for digital signatures, timestamps, trust validation, and LTV/PAdES workflows. PyMuPDF and pypdf are not complete signature backends.

## Before signing

- Inspect existing signatures, fields, `/Perms`, DocMDP, and field locks.
- Choose the certificate/key source and trust policy deliberately.
- Keep private keys and passphrases outside scripts, logs, shell history, and repository files.
- Finalize content and visual QA before applying a restrictive certification signature.

## Add a signature field

```bash
pyhanko sign addfields \
  --field '1/72,72,260,120/Sig1' \
  input.pdf tmp/pdfs/with-signature-field.pdf
```

## Sign

```bash
pyhanko sign addsig --field Sig1 pemder \
  --key /secure/key.pem \
  --cert /secure/cert.pem \
  input.pdf output.pdf
```

Use PKCS#12, PKCS#11, timestamp, revocation, and PAdES options only after reading the corresponding pyHanko configuration. Do not invent trust roots or disable validation warnings to make a gate pass.

## Validate

```bash
pyhanko sign validate --pretty-print output.pdf
```

Validation must use the intended trust roots, time, revocation policy, and network/cache policy. Record cryptographic integrity, signer trust, timestamps, modification/difference analysis, and any unsupported or experimental checks separately.

After any later incremental form or annotation update, validate again and review whether DocMDP permits the exact change.
