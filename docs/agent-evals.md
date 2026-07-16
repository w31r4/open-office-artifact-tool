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

Generated PDF fixtures require Python with ReportLab and pypdf. Set `OPEN_OFFICE_AGENT_EVAL_PYTHON` to that interpreter. In Codex, use the Python executable returned by `load_workspace_dependencies`; `OPEN_OFFICE_PDF_PROVIDER_PYTHON` remains a compatible fallback.

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

`run.json` records the git commit, tarball hash, Skill-tree hash, input fingerprints, reference-only package-name patches, and an oracle fingerprint. It also fingerprints `PROMPT.md`, the workspace package/lock files, `.agents`, and the complete installed `node_modules` tree, including symlink targets; changing the Skill or executed dependency graph therefore fails a hard gate. File inputs are made read-only and SHA-256 checked after execution. Directory inputs such as test PKI are recursively made read-only and hashed as a deterministic tree, so added, removed, renamed, or changed files fail the source-immutability gate.

The runner starts Codex with an ephemeral, ignored-config `workspace-write` sandbox. This is useful process isolation, but it is not an oracle-confidentiality boundary on every host. A production benchmark must mount only the trial workspace into a no-network container or VM and keep the evaluator directory and repository unavailable to the Agent.

## Scoring boundary

Current generic hard gates verify:

- oracle version, completed execution, trace, and final response;
- immutable prompt, Skill, installed dependency graph, and file/directory inputs;
- exact required artifacts, non-empty bytes, basic PDF/OOXML magic, and valid JSON deliverables;
- no modified artifact for fail-closed outcomes;
- a valid `failed_closed` audit when one is emitted;
- correct branching for `success-or-safe-refusal` cases.

These gates do **not** constitute a passing task score. Case-specific semantic assertions, all-page visual comparison, redaction/signature/security analysis, and provider/primitive trace grading are explicitly pending. Until those graders exist, reports say `partial-generic-only` or `generic-refusal-gates` rather than claiming full success.

## Pilot findings

The first generated PDF pilots produced two useful product signals:

1. The overflow replacement correctly failed closed, preserved the source hash, and emitted no modified PDF.
2. The bounded equal-length contract-ID replacement was semantically and visually correct, including five-page count and unchanged non-target pages. However, the shipped `pymupdf_edit.py replace_text` primitive rejected the replacement because its fit calculation exceeded the box by roughly `0.00002pt`; the Agent then bypassed the typed primitive with a direct content-stream replacement. The artifact result is evidence for fidelity, but the trace is a primitive-discipline failure and must not be scored as a clean provider-path success.

An earlier pilot missed the exact output filename because the runner prompt omitted declared inputs/deliverables. `PROMPT.md` now includes both, so that historical naming failure is evaluator noise rather than a product defect.
