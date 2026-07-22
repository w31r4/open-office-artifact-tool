# Create a bounded AES-256 encrypted PDF copy

Use `scripts/qpdf_provider.py encrypt` only to create one distinct AES-256
copy from a freshly inspected **unencrypted source** PDF. It is an Agent primitive for
an explicit “protect this new delivery copy” request, not a general password
workflow. It never opens, decrypts, re-encrypts, or changes permissions on an
already encrypted input.

The route needs qpdf `>=11.7.0`, a selected qpdf capability, explicit rewrite
and signature-invalidation authority, and two caller-owned password files. A
missing provider, unavailable managed release asset for the current platform,
unsupported qpdf version, encrypted input, stale hash, signed-input
acknowledgement omission, output collision, or password-policy failure stops
the operation without fallback.

## Resolve and inspect first

The public capability resolver does not accept password values. It records only
that the caller has declared the required secret channel:

```js
import { PdfFile } from "open-office-artifact-tool";
import { PdfProviders } from "open-office-artifact-tool/pdf/providers";

const inspection = await PdfFile.inspectPdf("input.pdf");
const resolution = await PdfProviders.resolve({
  task: "encrypt",
  provider: "qpdf",
  inspection,
  savePolicy: "rewrite",
  mutationAuthorized: true,
  invalidateSignaturesAuthorized: true,
  credentials: ["caller-owned-user-and-owner-password-files"],
});
if (resolution.status !== "ready") throw new Error(resolution.reason.message);
```

For a deployment-owned qpdf runtime, select `system-only` policy first. For a
managed runtime, `ensure` is permitted only when the published, hash-pinned
qpdf pack is explicitly authorized. Do not download qpdf ad hoc or replace a
selected managed route with a system route. See [provider setup](provider_setup.md).

Then obtain a fresh qpdf report. Copy only its SHA-256 value, never its source
path into an output name:

```bash
PYTHON_BIN="${OPEN_OFFICE_PDF_PROVIDER_PYTHON:-python3}"
"$PYTHON_BIN" scripts/qpdf_provider.py inspect input.pdf \
  > tmp/pdfs/qpdf-inspect.json
```

## Password-file contract

Provide an absolute or project-relative **regular, non-symlink** user-password
file and owner-password file. On POSIX both must be private to their owner
(`0600` or stricter; no group/world permissions), contain at most 4096 bytes,
and contain exactly one non-empty UTF-8 line. A final line ending is accepted;
embedded newline, carriage return, NUL, invalid UTF-8, an empty value, or equal
user/owner values are rejected. On Windows, the caller must enforce equivalent
ACL restrictions.

The wrapper copies password values only into private `0700` transaction-local
qpdf argument files (`0600`) and invokes qpdf with `@argument-file` paths.
Passwords never appear in qpdf argv, environment variables, JSON audit output,
or diagnostic text. This limits ordinary process-list/log exposure; it is not a
claim that filesystem deletion securely erases every storage layer.

Do not pass a password as a shell argument, environment variable, repository
file, prompt transcript, or audit field.

## Encrypt into a new delivery copy

Declare the credential *type* to the generic planner, then give only file paths
to the task-specific wrapper. The planner does not receive password values.

```bash
"$PYTHON_BIN" scripts/pdf_provider.py plan \
  --task encrypt --provider qpdf --strategy rewrite \
  --input input.pdf --output outputs/protected.pdf \
  --invalidate-signatures \
  --credential-declaration caller-owned-user-and-owner-password-files \
  --require-provider

"$PYTHON_BIN" scripts/qpdf_provider.py encrypt \
  input.pdf outputs/protected.pdf \
  --expected-sha256 '<sha256-from-qpdf-inspect>' \
  --user-password-file /secure/channel/pdf-user-password.txt \
  --owner-password-file /secure/channel/pdf-owner-password.txt \
  --invalidate-signatures \
  > tmp/pdfs/qpdf-encrypt.json
```

The resolver and generic planner require signature-invalidation authorization
up front because this task always writes a new byte representation. The wrapper
independently refuses any observed signature field, ByteRange, `/Perms`, DocMDP,
or FieldMDP evidence unless `--invalidate-signatures` is supplied. That flag
acknowledges that the old signature does not approve the new encrypted bytes;
it does not remove the need for pyHanko/DocMDP review or, where required, a new
signature.

The wrapper makes a read-only source snapshot, writes only the candidate,
rejects retention of the complete unencrypted source byte prefix, requires
clean independent user-password and owner-password reinspections, checks page/form/annotation/
attachment/outline/tagging counts, re-proves source identity, fsyncs, and
publishes with no replacement. Its result records AES-256 evidence and the
password channel without returning either password or source secret-file path.

## Boundary and delivery gates

- This route requests qpdf's default permission profile. It does not expose a
  permission editor, and PDF viewer restrictions are advisory rather than a
  substitute for access control.
- It does not decrypt, inspect an encrypted input with a supplied password,
  remove encryption, rotate credentials, preserve a signature, sanitize, redact,
  validate conformance, or promise universal viewer compatibility.
- A plain `qpdf_provider.py inspect protected.pdf` intentionally fails because
  the public inspect primitive accepts no password. Use the encryption report's
  authorized reinspection, then a password-safe viewer or a separately approved
  secure renderer for final review.
- Preserve the source and the secret files. Deliver only the distinct output
  after recording the exact output hash, authorized-open result, signature
  decision, and independent visual review in the PDF audit.
