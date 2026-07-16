# Agent black-box evaluations

`open-office-agent-promptbench-v1` measures whether an Agent can turn a realistic Office/PDF request into a verified artifact through the published Skills and npm package. It evaluates the complete workflow, not isolated API calls:

```text
prompt + immutable inputs
  -> selected native Skill
  -> packed open-office-artifact-tool candidate
  -> explicit provider and save policy
  -> exact outputs + audit
  -> semantic, visual, security, and trace grading
```

The suite is evaluator-side repository infrastructure. `evals/`, the runner, locked assets, trial traces, and grading oracles are deliberately excluded from the npm consumer package.

## Current suite

- 26 cases: 19 PDF, 2 Documents, 3 Spreadsheets, and 2 Presentations.
- 7 `ready` PDF cases use deterministic generated or inline inputs.
- 19 `asset-required` cases define fixture specifications but do not run until licensed, version-pinned corpus files or test PKI are placed under `evals/assets/`.
- Every family has both a success and a fail-closed case. Some advanced PDF cases accept either verified success or an explicit safe refusal.
- The default policy uses three trials per subject. Trial count is recorded per case rather than silently inferred by the Agent.

The committed JSONL contains evaluator-only expected outcomes, sources, and grade specifications. The Agent receives only `PROMPT.md`, declared inputs, the selected Skill, the installed candidate tarball, and exact deliverable paths.

## Commands

```sh
npm run eval:agents -- validate
npm run eval:agents -- list --family pdf --status ready
npm run eval:agents -- show pdf-bounded-contract-id-replace
npm run eval:agents -- prepare pdf-bounded-contract-id-replace --subject candidate --trial 1
npm run eval:agents -- prepare pdf-bounded-contract-id-replace --subject reference --trial 1
npm run eval:agents -- run pdf-overflow-replace-refusal --subject candidate --trial 1
npm run eval:agents -- score pdf-overflow-replace-refusal --trial-root /absolute/trial/path
```

Generated PDF fixtures require Python with ReportLab and pypdf. The three implemented PDF case graders additionally require pdfplumber and Pillow; the bounded-replacement and active-content visual oracles require `pdftoppm`. Set `OPEN_OFFICE_AGENT_EVAL_PYTHON` to that interpreter and, only when it is not on `PATH`, set `OPEN_OFFICE_AGENT_EVAL_PDFTOPPM`. In Codex, use the Python executable returned by `load_workspace_dependencies`; `OPEN_OFFICE_PDF_PROVIDER_PYTHON` remains a compatible fallback.

`--subject candidate` installs `skills/<family>/skills/<skill>`. `--subject reference` copies the matching handoff reference Skill into that trial and changes only its `office-artifact-tool` package-name occurrences inside the isolated copy. Both subjects install the same freshly packed `open-office-artifact-tool` tarball, so the comparison changes the Skill instructions rather than the product candidate.

## Isolation and provenance

Each preparation creates a fresh trial tree outside the repository by default:

```text
<run-root>/<case>/<subject>-trial-<n>/
  workspace/
    .agents/skills/<skill>/
    inputs/
    outputs/
    PROMPT.md
    node_modules/open-office-artifact-tool/
  evaluator/
    package/<candidate.tgz>
    run.json
    trace.jsonl
    final.txt
    report.json
```

`run.json` records the git commit, whether the source worktree was dirty, hashes of its status and tracked diff, the actual packed tarball hash, Skill-tree hash, input fingerprints, reference-only package-name patches, and an oracle fingerprint. The tarball SHA-256 is the authoritative identity of the candidate bytes, including when a local WIP is intentionally evaluated; a recorded HEAD is never presented as if it alone identified a dirty trial. The record also fingerprints `PROMPT.md`, the workspace package/lock files, `.agents`, and the complete installed `node_modules` tree, including file modes and symlink targets; changing the Skill or executed dependency graph therefore fails a hard gate. Skill copying filters ignored Python bytecode, the installed tree is read-only, and Codex runs with Python bytecode generation disabled, so local cache noise cannot enter the immutable Skill subject. File inputs are made read-only and SHA-256 checked after execution. Directory inputs such as test PKI are recursively made read-only and hashed as a deterministic tree, so added, removed, renamed, mode-changed, or content-changed files fail the source-immutability gate.

The runner starts Codex with an ephemeral, ignored-config `workspace-write` sandbox. This is useful process isolation, but it is not an oracle-confidentiality boundary on every host. A production benchmark must mount only the trial workspace into a no-network container or VM and keep the evaluator directory and repository unavailable to the Agent.

## Scoring boundary

Current generic hard gates verify:

- oracle version, completed execution, trace, and final response;
- immutable prompt, Skill, installed dependency graph, and file/directory inputs;
- exact required artifacts, non-empty bytes, basic PDF/OOXML magic, and valid JSON deliverables;
- no modified artifact for fail-closed outcomes;
- a valid `failed_closed` audit when one is emitted;
- correct branching for `success-or-safe-refusal` cases.
- regular-file outputs only and no undeclared success deliverables.

Generic gates alone do **not** constitute a passing task score. Reports for cases without a case grader continue to say `partial-generic-only` or `generic-refusal-gates` and keep their semantic, visual, security, and trace evidence in `pending`.

Three PDF cases now have complete case-specific grading:

| Case | Machine | Visual | Security | Trace |
| --- | --- | --- | --- | --- |
| `pdf-bounded-contract-id-replace` | Independent page/text/font/box assertions | Poppler renders all five pages; non-target pages must be pixel-identical and the page-3 diff must stay inside the source text mask | Raw, extracted, decoded-stream, and metadata residue scan; one `startxref`/`%%EOF`; canonical audit hashes must match final bytes | PyMuPDF/version, explicit `sanitize`, no fallback, shipped probe before edit, typed mutation, and no low-level stream mutation |
| `pdf-overflow-replace-refusal` | Independent ReportLab font metric plus pdfplumber 70pt-box proof and structured audit geometry | Not applicable | No partial artifact, no mutation claim, and source provenance | PyMuPDF capability evidence, no fallback, and no mutation command after failed preflight |
| `pdf-active-content-public-sanitize` | pypdf proves the fixture contains root/additional JavaScript, Launch/SubmitForm actions, attachments, invisible text, a comment, populated form values, and personal metadata, then proves every channel is inert | Poppler renders every page and permits changes only inside the removed widget-value/comment masks | All canaries absent from object strings, streams, attachments, annotations/forms, text, and metadata; one revision; original prefix absent; canonical audit hashes match | Probe and route plan before typed scrub, then standalone inert residue scan, Poppler render, audit byte validation, no fallback or low-level mutation bypass |

The category weights are machine 45, visual 25, security 20, and trace 10. A category earns its weight only when every check in it passes; a not-applicable category is removed from the denominator. `rawScorePercent` preserves the evidence-weighted result; a failed safety hard gate forces `scorePercent` to zero. `taskPassed` is true only when generic hard gates and every applicable case category pass. Missing evaluator dependencies produce `grader-unavailable` with an infrastructure error, never a candidate failure or an inferred pass. A present but unreadable or unrenderable candidate PDF is instead a definitive graded failure.

The PDF oracle is evaluator-side and never copied into the Agent workspace. It uses pypdf/pdfplumber for semantic and structural evidence and Poppler/Pillow for visual evidence, independently of the PyMuPDF mutation provider. Audit claims and Codex command traces are graded separately from final bytes.

## Pilot findings

The first generated PDF pilots produced two useful product signals:

1. The overflow replacement correctly failed closed, preserved the source hash, and emitted no modified PDF.
2. The first bounded equal-length contract-ID replacement was semantically and visually correct, including five-page count and unchanged non-target pages. However, the shipped `pymupdf_edit.py replace_text` primitive rejected the replacement because its fit calculation exceeded the box by roughly `0.00002pt`; the Agent then bypassed the typed primitive with a direct content-stream replacement. The grader scores that historical result 90/100: machine, visual, and security pass, while save-policy, typed-primitive, and low-level-bypass trace checks fail. It is therefore not a passing task.
3. The typed primitive now preserves the source baseline/default Base14 style and tolerates only provider/search-box float quantization, capped at `0.0005pt`; genuine overflow and rotated/cross-span input still fail closed. A second trial exposed an underspecified audit envelope, so the Skill now publishes `open-office-artifact-tool.pdf-audit.v1`, a JSON Schema, and `pdf_audit.py validate` to bind canonical provider/save-policy/operation evidence to exact source/output bytes.
4. With those changes fixed, one candidate trial scored 100/100 across machine, visual, security, and trace categories. A reference-subject trial using the byte-identical candidate tarball also scored 100/100: its different reference Skill discovered and used the published candidate PDF tasks/scripts. This supports package discoverability and reference-workflow compatibility; it does not establish candidate-Skill superiority, and the default three repeat trials per fixed subject have not yet been run.
5. The active-content case exposed a provider-contract gap: PyMuPDF's stock scrub left root actions, invisible text, a comment, and a populated form value in the generated file. The typed adapter now removes the bounded inert-publication surface, fails when invisible text overlaps visible text, and runs a structural `--require-inert` gate after the full rewrite. A first live trial safely refused because the Skill hard-coded system `python3` instead of the runner-configured provider interpreter. After making interpreter selection explicit, two fixed-candidate trials passed 100/100. A third repeat exposed a narrower gap: PyMuPDF represented removed `/OpenAction`, `/AA`, `/JavaScript`, and `/EmbeddedFiles` names as `null`; the Agent repaired the final bytes through direct object mutation, so the trace hard gate correctly reduced a 90 raw score to zero. The official scrub primitive now physically removes those null active-content names, reports their xrefs, and fails closed on unfamiliar object serialization. The repeat matrix must be restarted on the fixed tarball; the historical bypass is not counted as a pass.

An earlier pilot missed the exact output filename because the runner prompt omitted declared inputs/deliverables. `PROMPT.md` now includes both, so that historical naming failure is evaluator noise rather than a product defect.

Command-trace grading detects completed inline shell commands and shipped primitive invocations; it is not operating-system syscall attestation. A production benchmark that must resist a deliberately evasive Agent should add container-level process/filesystem tracing in addition to the current no-network mount boundary.
